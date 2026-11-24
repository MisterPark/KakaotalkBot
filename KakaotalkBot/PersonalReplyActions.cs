using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Newtonsoft.Json.Linq;

namespace KakaotalkBot
{
    internal sealed class PersonalEmoticon
    {
        internal string Pack, Title, Meaning;
        internal int Index;
        internal string Id { get { return Pack + ":" + Index; } }
    }
    internal static class PersonalReplyActions
    {
        // 실제 PC 선택창의 ReactionPickerItemData(kind=2)와 로컬 아이콘을 대조한 기본 공감.
        internal static readonly IReadOnlyDictionary<string, string> Reactions = new Dictionary<string, string> {
            { "1200509_021", "좋아요·동의" }, { "1200509_029", "하트·고마움·애정" },
            { "1200509_037", "확인·알겠음" }, { "1200509_001", "웃음·즐거움" },
            { "1200509_002", "놀람" }, { "1200509_003", "슬픔·위로" }
        };
        internal static bool Allowed(PersonalReplyDecision decision, IEnumerable<PersonalEmoticon> items, long target)
        {
            if (decision == null) return false;
            if (decision.Action == "reply" || decision.Action == "hold") return string.IsNullOrEmpty(decision.Choice);
            if (!string.IsNullOrEmpty(decision.Text)) return false;
            if (decision.Action == "reaction") return target > 0 && Reactions.ContainsKey(decision.Choice ?? "");
            return decision.Action == "emoticon" && (items ?? Enumerable.Empty<PersonalEmoticon>()).Any(x => x.Id == decision.Choice);
        }
        internal static string Preview(PersonalReplyDecision decision, IEnumerable<PersonalEmoticon> items)
        {
            if (decision.Action == "reply") return decision.Text;
            if (decision.Action == "reaction") return "[공감] " + Reactions[decision.Choice];
            if (decision.Action == "emoticon") { var item = items.Single(x => x.Id == decision.Choice); return "[이모티콘] " + item.Title + " #" + item.Index + " · " + item.Meaning; }
            return "[직접 확인 필요] " + decision.Reason;
        }
        private sealed class RankedEmoticon { internal PersonalEmoticon Item; internal int Score; internal uint Order; }
        private static uint Order(string id, long seed)
        { unchecked { uint value = 2166136261U ^ (uint)seed ^ (uint)(seed >> 32); foreach (char c in id) value = (value ^ c) * 16777619U; return value; } }
        internal static List<PersonalEmoticon> Candidates(IReadOnlyList<PersonalEmoticon> all, IReadOnlyList<PersonalTurn> history, JObject usage)
        {
            const int maximum = 240;
            if (all.Count <= maximum) return all.ToList();
            var latest = history.LastOrDefault(t => !t.Own);
            string context = string.Join(" ", history.Where(t => !t.Own && t.EmoticonId == null).Reverse().Take(3).Select(t => t.Text));
            var terms = new HashSet<string>(System.Text.RegularExpressions.Regex.Matches(context.ToLowerInvariant(), @"[\p{L}]{2,}").Cast<System.Text.RegularExpressions.Match>().Select(m => m.Value));
            // 후보를 줄이는 검색어 확장입니다. 최종 의미 판단은 전체 최근 대화를 받은 AI가 합니다.
            foreach (var group in new[] { "ㅋㅋ|ㅎㅎ|웃|하하|재밌", "고마|감사|고맙", "잘자|굿나잇|졸려|자러|잠|꿈", "안녕|하이|인사|잘가",
                "미안|죄송|사과", "축하|생일|생축", "힘내|힘들|위로|괜찮", "사랑|하트|보고싶|애정", "알겠|확인|오케이|ㅇㅋ|응|네", "헐|놀람|대박" })
            {
                var words = group.Split('|'); if (words.Any(context.Contains)) foreach (var word in words) terms.Add(word);
            }
            var favorites = new Dictionary<string, int>();
            foreach (var row in (usage["owner_favorites"] as JArray ?? new JArray()).OfType<JObject>())
                favorites[(string)row["id"]] = Math.Min(100, (int)row["count"]);
            var pairs = new Dictionary<string, int>();
            foreach (var row in (usage["reply_pairs"] as JArray ?? new JArray()).OfType<JObject>())
                if (latest != null && latest.EmoticonId != null && (string)row["contact"] == latest.EmoticonId)
                    pairs[(string)row["owner"]] = Math.Min(100, (int)row["count"]);
            var ranked = all.Select(item => {
                int favorite, pair; favorites.TryGetValue(item.Id, out favorite); pairs.TryGetValue(item.Id, out pair);
                string labels = (item.Meaning + " " + item.Title).ToLowerInvariant();
                return new RankedEmoticon { Item = item, Score = terms.Count(labels.Contains) * 30 + favorite + (pair > 0 ? 500 + pair : 0), Order = Order(item.Id, latest == null ? 0 : latest.LogId) };
            }).OrderByDescending(r => r.Score).ThenBy(r => r.Order).ThenBy(r => r.Item.Id, StringComparer.Ordinal).ToList();
            var result = new List<PersonalEmoticon>(); var seen = new HashSet<string>(); var perPack = new Dictionary<string, int>();
            foreach (var entry in ranked.Where(r => r.Score > 0))
            {
                int count; perPack.TryGetValue(entry.Item.Pack, out count); if (count >= 12) continue;
                if (seen.Add(entry.Item.Id)) { result.Add(entry.Item); perPack[entry.Item.Pack] = count + 1; }
                if (result.Count >= 80) break;
            }
            // 나머지는 팩별로 번갈아 채워, 앞쪽 팩이나 큰 팩 하나가 후보를 독점하지 않게 합니다.
            var packs = ranked.GroupBy(r => r.Item.Pack).Select(g => new Queue<RankedEmoticon>(g)).ToList();
            while (result.Count < maximum && packs.Any(q => q.Count > 0))
                foreach (var pack in packs)
                {
                    while (pack.Count > 0 && seen.Contains(pack.Peek().Item.Id)) pack.Dequeue();
                    if (pack.Count == 0) continue;
                    var item = pack.Dequeue().Item; seen.Add(item.Id); result.Add(item);
                    if (result.Count == maximum) break;
                }
            return result;
        }
        internal static List<PersonalEmoticon> Load(ChatRoomInfo room, CancellationToken cancellation)
        {
            string source = Path.Combine(room.AccountPath, "emoticon.edb");
            string folder = Path.Combine(Path.GetTempPath(), "KakaotalkBot-Emoticons-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            string output = Path.Combine(folder, "catalog.sqlite");
            try
            {
                var decryptor = new KakaoTalkDecryptor();
                using (var key = decryptor.DiscoverKey(room.ProcessId, decryptor.ReadDatabaseInfo(source)))
                {
                    cancellation.ThrowIfCancellationRequested();
                    if (key == null) throw new InvalidOperationException("이모티콘 목록 키를 확인하지 못했습니다.");
                    var snapshot = new ChatDatabaseSnapshot(source, output, decryptor, key);
                    snapshot.Refresh(cancellation);
                    if (snapshot.RefreshDeferred || !File.Exists(output)) throw new IOException("이모티콘 목록을 갱신 중입니다.");
                    using (var db = new ChatSqlite(output))
                    {
                        var labels = db.Query("SELECT keywordId,keywordName FROM SuggestionDictionary").ToDictionary(r => Convert.ToInt64(r[0]), r => Convert.ToString(r[1]));
                        return ReadCatalog(db.Query("SELECT itemCode,title,itemType,itemSubType,expiredAt,emotsJson,kids FROM KeyboardItems WHERE ownedItem=1"), labels,
                            DateTimeOffset.UtcNow.ToUnixTimeSeconds());
                    }
                }
            }
            finally
            {
                // 이 작업에서 만든 임시 파일만 지웁니다. 원본 DB와 사용자 설정은 수정하지 않습니다.
                foreach (string path in new[] { output, output + ".building" }) if (File.Exists(path)) File.Delete(path);
                if (!Directory.EnumerateFileSystemEntries(folder).Any()) Directory.Delete(folder);
            }
        }
        internal static List<PersonalEmoticon> ReadCatalog(IEnumerable<object[]> rows, IDictionary<long, string> labels, long now)
        {
            var result = new List<PersonalEmoticon>(); var seen = new HashSet<string>();
            foreach (var row in rows)
            {
                string pack = Convert.ToString(row[0]), title = Convert.ToString(row[1]), type = Convert.ToString(row[2]);
                if (pack.Length < 1 || pack.Length > 16 || pack.Any(c => c < '0' || c > '9') ||
                    type != "STICKER" && type != "STICKER_ANI" || Convert.ToInt64(row[4] ?? 0) > 0 && Convert.ToInt64(row[4]) <= now) continue;
                JArray emots, kids;
                try { emots = JArray.Parse(Convert.ToString(row[5])); kids = JArray.Parse(Convert.ToString(row[6])); }
                catch (Newtonsoft.Json.JsonException) { continue; }
                foreach (var emot in emots.OfType<JObject>())
                {
                    int index;
                    var indexToken = emot["emotIdx"]; var codeToken = emot["itemCode"]; var soundToken = emot["soundPath"];
                    if (indexToken == null || indexToken.Type != JTokenType.Integer || !int.TryParse(indexToken.ToString(), out index) ||
                        codeToken == null || codeToken.Type != JTokenType.String || (string)codeToken != pack || index < 1 || index > 200 ||
                        soundToken != null && soundToken.Type != JTokenType.Null && (soundToken.Type != JTokenType.String || !string.IsNullOrEmpty((string)soundToken))) continue;
                    var tags = index <= kids.Count ? kids[index - 1] as JArray : null;
                    string meaning = tags == null ? "" : string.Join("·", tags.Where(t => t.Type == JTokenType.Integer && labels.ContainsKey((long)t)).Select(t => labels[(long)t]).Distinct().Take(8));
                    // 의미를 모르는 그림이나 필터에 걸리는 상품은 모델에 선택지로 제공하지 않습니다.
                    if (string.IsNullOrWhiteSpace(meaning) || ChatProfanityFilter.Mask(title + meaning) != title + meaning) continue;
                    var item = new PersonalEmoticon { Pack = pack, Index = index, Title = PersonalReplyAi.Limit(title, 80), Meaning = PersonalReplyAi.Limit(meaning, 120) };
                    if (seen.Add(item.Id)) result.Add(item);
                }
            }
            return result;
        }
    }
}
