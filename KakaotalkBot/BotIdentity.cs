using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KakaotalkBot
{
    // 이 파일의 키는 PC에 종속되지 않습니다. 비공개 저장소와 실행 폴더에 함께 보관합니다.
    internal sealed class BotIdentityConfiguration
    {
        public int Version = 1;
        public string Key { get; set; }
        public string SpreadsheetId { get; set; }
        public string PreviousSpreadsheetId { get; set; }
        public string SuperUserId { get; set; }
    }

    internal sealed class BotIdentity
    {
        private readonly byte[] encryptionKey, authenticationKey;
        private readonly HashSet<string> knownUsers = new HashSet<string>();
        private readonly object gate = new object();
        private readonly ConcurrentDictionary<long, string> encodedCache = new ConcurrentDictionary<long, string>();
        private readonly ConcurrentDictionary<string, long> decodedCache = new ConcurrentDictionary<string, long>();
        internal readonly string SpreadsheetId;
        // PC를 바꿔도 같은 설정 파일의 봇 ID로 권한을 유지합니다. 읽기·복호화는 최초 한 번만 합니다.
        private static readonly Lazy<long> superUserId = new Lazy<long>(() =>
        {
            var config = JsonConvert.DeserializeObject<BotIdentityConfiguration>(File.ReadAllText(ConfigurationPath));
            if (config == null) throw new InvalidDataException("권한 설정 오류");
            return string.IsNullOrWhiteSpace(config.SuperUserId) ? 0 : new BotIdentity(config).Decode(config.SuperUserId);
        });
        internal static long SuperUserId { get { return superUserId.Value; } }
        internal static string ConfigurationPath { get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "BotIdentity.json"); } }
        internal static string ResolveSpreadsheet(string current)
        {
            if (!File.Exists(ConfigurationPath)) return current;
            var config = JsonConvert.DeserializeObject<BotIdentityConfiguration>(File.ReadAllText(ConfigurationPath));
            if (config == null || string.IsNullOrWhiteSpace(config.SpreadsheetId)) throw new InvalidDataException("고정 ID 연결 설정 오류");
            return string.IsNullOrEmpty(current) || current == "<YOUR_SPREADSHEET_ID>" || current == config.PreviousSpreadsheetId ? config.SpreadsheetId : current;
        }
        internal static BotIdentity Load(string sheetId)
        {
            if (!File.Exists(ConfigurationPath)) throw new InvalidOperationException("BotIdentity.json이 없습니다. 기존 PC의 고정 ID 설정을 복사해 주세요. 새 키를 만들면 안 됩니다.");
            var identity = new BotIdentity(JsonConvert.DeserializeObject<BotIdentityConfiguration>(File.ReadAllText(ConfigurationPath)));
            if (identity.SpreadsheetId != sheetId) throw new InvalidOperationException("고정 ID 설정과 스프레드시트가 다릅니다. 새 DB의 시트 ID를 사용해 주세요.");
            return identity;
        }
        internal BotIdentity(BotIdentityConfiguration configuration)
        {
            if (configuration == null || configuration.Version != 1) throw new InvalidDataException("고정 ID 설정 버전 오류");
            byte[] key = Convert.FromBase64String(configuration.Key ?? "");
            if (key.Length != 64) throw new InvalidDataException("고정 ID 키 형식 오류");
            encryptionKey = key.Take(32).ToArray(); authenticationKey = key.Skip(32).ToArray();
            SpreadsheetId = configuration.SpreadsheetId;
        }
        internal string KeyCheck { get { return Encode(1); } }
        private byte[] Mac(byte[] bytes) { using (var h = new HMACSHA256(authenticationKey)) return h.ComputeHash(bytes); }
        internal string Encode(long id)
        {
            if (id == 0) return "0";
            if (id < 0) throw new InvalidDataException("사용자 ID는 양수여야 합니다.");
            return encodedCache.GetOrAdd(id, Encrypt);
        }
        private string Encrypt(long id)
        {
            byte[] plain = Encoding.ASCII.GetBytes(id.ToString(CultureInfo.InvariantCulture));
            // 사용자 식별자는 동일성을 유지해야 하므로 전용 키로 결정적 IV를 만듭니다.
            // 암호문에 별도 인증 태그를 붙여 다른 키나 변조된 ID의 사용을 차단합니다.
            byte[] iv = Mac(Encoding.ASCII.GetBytes("iv:" + id.ToString(CultureInfo.InvariantCulture))).Take(16).ToArray();
            byte[] cipher;
            using (var aes = Aes.Create())
            { aes.Key = encryptionKey; aes.IV = iv; using (var transform = aes.CreateEncryptor()) cipher = transform.TransformFinalBlock(plain, 0, plain.Length); }
            byte[] body = iv.Concat(cipher).ToArray();
            return "bot_" + Convert.ToBase64String(body.Concat(Mac(body).Take(16)).ToArray()).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        }
        internal long Decode(string value)
        {
            if (value == "0") return 0;
            if (value == null || !value.StartsWith("bot_", StringComparison.Ordinal)) throw new InvalidDataException("DB에 봇 ID가 아닌 값이 있습니다. 저장을 중단합니다.");
            return decodedCache.GetOrAdd(value, Decrypt);
        }
        private long Decrypt(string value)
        {
            string encoded = value.Substring(4).Replace('-', '+').Replace('_', '/');
            byte[] data = Convert.FromBase64String(encoded.PadRight((encoded.Length + 3) / 4 * 4, '='));
            if (data.Length < 48 || data.Length % 16 != 0) throw new InvalidDataException("봇 ID 형식 오류");
            byte[] body = data.Take(data.Length - 16).ToArray(), expected = Mac(body);
            int difference = 0;
            for (int i = 0; i < 16; i++) difference |= expected[i] ^ data[body.Length + i];
            if (difference != 0) throw new InvalidDataException("봇 ID의 키가 다르거나 값이 변조되었습니다.");
            byte[] plain;
            using (var aes = Aes.Create())
            { aes.Key = encryptionKey; aes.IV = body.Take(16).ToArray(); using (var transform = aes.CreateDecryptor()) plain = transform.TransformFinalBlock(body, 16, body.Length - 16); }
            long id = long.Parse(Encoding.ASCII.GetString(plain), CultureInfo.InvariantCulture);
            if (id <= 0 || Encode(id) != value) throw new InvalidDataException("봇 ID 내용 오류");
            return id;
        }
        private string Id(string value, bool encode)
        {
            if (!encode) { long native = Decode(value); Register(native.ToString(CultureInfo.InvariantCulture)); return native.ToString(CultureInfo.InvariantCulture); }
            long id = long.Parse(value, CultureInfo.InvariantCulture); Register(value); return Encode(id);
        }
        private void Register(string value) { if (value != "0") lock (gate) knownUsers.Add(value); }
        internal string Scrub(string text)
        {
            if (text == null) return null;
            lock (gate) return Regex.Replace(text, @"bot_[A-Za-z0-9_-]+|(?<![0-9])[0-9]+(?![0-9])", m => knownUsers.Contains(m.Value) ? Encode(long.Parse(m.Value, CultureInfo.InvariantCulture)) : m.Value);
        }
        private string Text(string text, bool encode)
        {
            // 사용자가 직접 쓴 bot_ 문구와 자동 변환한 식별자를 구분합니다.
            return encode ? Scrub((text ?? "").Replace("bot_", "bot_literal_")) :
                Regex.Replace(text ?? "", @"bot_(?:[A-Za-z0-9_-]{86}|[A-Za-z0-9_-]{64})(?![A-Za-z0-9_-])", m => DecodeText(m.Value)).Replace("bot_literal_", "bot_");
        }
        private string DecodeText(string text)
        {
            // 닉네임이나 메모에 우연히 bot_ 문자열이 있어도 식별자 컬럼처럼 거부하지 않습니다.
            try { return Decode(text).ToString(CultureInfo.InvariantCulture); }
            catch (FormatException) { return text; }
            catch (InvalidDataException) { return text; }
            catch (CryptographicException) { return text; }
        }
        private string Key(string text, int position, bool encode)
        {
            var parts = text.Split(':'); if (position < 0) position = parts.Length - 1;
            if (position >= parts.Length) throw new InvalidDataException("사용자 이력 키 형식 오류");
            parts[position] = Id(parts[position], encode); return string.Join(":", parts);
        }
        internal IList<IList<object>> Transform(string title, IList<IList<object>> rows, bool encode, bool header)
        {
            if (rows == null) return null;
            int column = title == "DB_UserId" ? 7 : title == "DB_RoomUsers" || title == "DB_RoomOperators" ? 1 : title == "DB_RoomEvents" || title == "DB_NicknameHistory" ? 2 : -1;
            var result = rows.Select(r => (IList<object>)r.ToList()).ToList();
            // 먼저 전체 사용자를 등록해 앞쪽 행의 메모에서도 뒤쪽 사용자 ID를 숨깁니다.
            if (encode && column >= 0) foreach (var row in result.Skip(header ? 1 : 0)) if (row.Count > column) Register(Convert.ToString(row[column], CultureInfo.InvariantCulture));
            for (int index = 0; index < result.Count; index++)
            {
                var row = result[index];
                if (header && index == 0)
                {
                    if (column >= 0)
                    {
                        string expected = encode ? "user_id" : "bot_user_id";
                        if (row.Count <= column || Convert.ToString(row[column]) != expected) throw new InvalidDataException(title + " ID 헤더가 올바르지 않습니다.");
                        row[column] = encode ? "bot_user_id" : "user_id";
                    }
                    continue;
                }
                if (row.All(v => string.IsNullOrEmpty(Convert.ToString(v)))) continue;
                if (column >= 0)
                {
                    row[column] = Id(Convert.ToString(row[column], CultureInfo.InvariantCulture), encode);
                    if (title == "DB_RoomEvents") row[0] = Key(Convert.ToString(row[0]), 2, encode);
                    int[] textColumns = title == "DB_UserId" ? new[] { 0 } : title == "DB_NicknameHistory" ? new[] { 3, 4 } : title == "DB_RoomEvents" ? new[] { 6 } : new[] { 2 };
                    foreach (int c in textColumns) if (row.Count > c) row[c] = Text(Convert.ToString(row[c]), encode);
                }
                else if (title == "DB_Operations" || title == "DB_RoomTitles" || title == "DB_DailyQuests")
                {
                    int jsonColumn = title == "DB_Operations" ? 2 : 1;
                    var json = JsonConvert.DeserializeObject<JObject>(Convert.ToString(row[jsonColumn]), new JsonSerializerSettings { DateParseHandling = DateParseHandling.None });
                    foreach (string field in new[] { "UserId", "AuthorId", "RevokedBy" })
                        if (json[field] != null)
                        { string id = Id(json[field].ToString(), encode); json[field] = encode ? (JToken)new JValue(id) : new JValue(long.Parse(id, CultureInfo.InvariantCulture)); }
                    string kind = (string)json["Kind"], key = (string)json["Key"];
                    bool keyed = title == "DB_DailyQuests" || title == "DB_RoomTitles" && key.StartsWith("monthly:", StringComparison.Ordinal) || title == "DB_Operations" && new[] { "monthly", "popularity", "point_reset", "reentry" }.Contains(kind);
                    if (keyed) json["Key"] = Key(key, -1, encode);
                    if (title == "DB_DailyQuests" && json["Rewards"] is JObject)
                        foreach (var reward in ((JObject)json["Rewards"]).Properties()) reward.Value["Key"] = Key((string)reward.Value["Key"], 1, encode);
                    foreach (string field in new[] { "Text", "Name", "Title", "Reason" }) if (json[field] != null && json[field].Type == JTokenType.String) json[field] = Text((string)json[field], encode);
                    row[jsonColumn - 1] = (string)json["Key"];
                    row[jsonColumn] = json.ToString(Formatting.None);
                }
                else if (encode) for (int c = 0; c < row.Count; c++) if (row[c] is string) row[c] = Scrub((string)row[c]);
            }
            return result;
        }
    }
}
