using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace KakaotalkBot
{
    internal sealed class OperatorNotice
    {
        public string Key, Account, Body, State, Error;
        public long RoomId, SourceId, At;
    }

    // 전송 직전에 시도 상태를 디스크에 확정합니다. 전송 결과가 불명확하면 재시작해도 다시 보내지 않습니다.
    internal sealed class OperatorNotificationQueue
    {
        private readonly string path;
        internal readonly List<OperatorNotice> Items = new List<OperatorNotice>();
        internal string Status = "운영진 알림 꺼짐";
        internal OperatorNotificationQueue(string path)
        {
            this.path = path;
            if (!File.Exists(path)) return;
            var rows = JsonConvert.DeserializeObject<List<OperatorNotice>>(File.ReadAllText(path));
            if (rows == null || rows.Any(r => r == null || string.IsNullOrWhiteSpace(r.Key) || r.RoomId <= 0 || string.IsNullOrWhiteSpace(r.Account) ||
                !new[] { "pending", "blocked", "cancelled", "attempted", "unknown", "submitted" }.Contains(r.State)) ||
                rows.Select(r => r.Key).Distinct().Count() != rows.Count) throw new FormatException("운영진 전송 기록이 손상되었습니다. 자동 전송을 중단합니다.");
            Items.AddRange(rows);
            foreach (var item in Items.Where(r => r.State == "attempted")) { item.State = "unknown"; item.Error = "이전 실행에서 결과를 확인하지 못했습니다. 방에서 직접 확인하세요."; }
        }
        internal void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temp = path + ".tmp";
            File.WriteAllText(temp, JsonConvert.SerializeObject(Items, Formatting.Indented));
            if (File.Exists(path)) File.Replace(temp, path, path + ".bak"); else File.Move(temp, path);
        }
        internal void Enqueue(string key, string account, long room, long source, string body, long at)
        {
            if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(account) || room <= 0 || source <= 0 || room == source)
                throw new ArgumentException("알림의 원본 방과 전송 대상이 올바르지 않습니다.");
            if (Items.Any(r => r.Key == key)) return;
            if (string.IsNullOrWhiteSpace(body)) return;
            if (body.Length > 3000) body = body.Substring(0, char.IsHighSurrogate(body[2999]) ? 2999 : 3000) + "\n… 전체 내용은 봇 운영 화면에서 확인하세요.";
            Items.Add(new OperatorNotice { Key = key, Account = account, RoomId = room, SourceId = source, Body = body, At = at, State = "pending" });
            Save();
        }
        internal void Process(string account, long destination, Func<OperatorNotice, string> validate, Action<OperatorNotice> send)
        {
            var item = Items.FirstOrDefault(r => r.State == "pending");
            if (item == null) return;
            if (item.Account != account || item.RoomId != destination)
            { item.State = "cancelled"; item.Error = "알림방 설정이 변경되었습니다."; Save(); return; }
            string error = validate(item);
            if (error != null) { item.State = "blocked"; item.Error = error; Status = "전송 차단: " + error; Save(); return; }
            item.State = "attempted"; Save();
            try { send(item); item.State = "submitted"; item.Error = "서버 수신은 확인하지 않음"; Status = "알림 입력창 전송 완료 (서버 수신 미확인)"; }
            catch (Exception ex) { item.State = "unknown"; item.Error = ex.GetBaseException().Message; Status = "알림 결과 미확인: " + item.Error; }
            Save();
        }
        internal void RetryBlocked()
        { foreach (var item in Items.Where(r => r.State == "blocked")) { item.State = "pending"; item.Error = null; } Save(); }

        internal void Collect(Database database, ChatRoomInfo source, Settings config)
        {
            if (!config.OperatorAlertsEnabled || source == null || source.ChatId == config.OperatorAlertRoomId ||
                source.AccountPath != config.OperatorAlertAccount) return;
            var rows = database.Operations.Rows.Values.ToArray();
            foreach (var row in rows.Where(r => r.ChatId == source.ChatId && r.At >= config.OperatorAlertSince && (r.Kind == "reentry" || r.Kind == "memo")))
            {
                if (row.Kind == "reentry" && !config.AlertReentry || row.Kind == "memo" && !config.AlertMemo) continue;
                User user; database.FindUser(row.UserId, out user);
                string body = "[" + source.Name + " · " + (row.Kind == "reentry" ? "재입장" : "메모 등록") + "]\n사용자: " +
                    database.ChatUserName(row.ChatId, row.UserId, row.Name);
                if (row.Kind == "reentry") body += "\n" + row.Text + "\n퇴장 " + (user == null ? 0 : user.LeaveCount) + "회 · 강퇴 " + (user == null ? 0 : user.KickCount) + "회";
                else body += "\n작성자: " + (row.Source == "local_operator" ? "로컬 운영자" : database.ChatUserName(row.ChatId, row.AuthorId));
                if (config.AlertMemoBody && row.Kind == "memo") body += "\n" + row.Text;
                else body += "\n메모 본문은 봇 운영 화면에서 확인하세요.";
                Enqueue(row.Key, config.OperatorAlertAccount, config.OperatorAlertRoomId, source.ChatId, body, row.At);
            }
            foreach (var group in rows.Where(r => r.Kind == "point_reset" && r.At >= config.OperatorAlertSince).GroupBy(r => r.Month))
            {
                if (config.AlertReset) Enqueue("reset-notice:" + group.Key, config.OperatorAlertAccount, config.OperatorAlertRoomId, source.ChatId,
                    "[포인트 초기화 완료]\n적용 월: " + group.Key + " (한국 시간)\n전 사용자 " + group.Count() + "명의 포인트를 0으로 초기화했습니다.\n이전 잔액은 DB에 보관했습니다.", group.First().At);
                if (config.AlertMonthly)
                {
                    string month = DateTime.ParseExact(group.Key, "yyyy-MM", System.Globalization.CultureInfo.InvariantCulture).AddMonths(-1).ToString("yyyy-MM");
                    Enqueue("monthly-notice:" + source.ChatId + ":" + group.Key, config.OperatorAlertAccount, config.OperatorAlertRoomId, source.ChatId,
                        "[" + source.Name + " · 월간 집계]\n" + database.RoomStatistics(source.ChatId, month) + "\n\n" + database.Operations.Ranking(source.ChatId, month) + "\n\n" + database.Operations.Ranking(source.ChatId, month, true), group.First().At);
                }
            }
        }
    }
}
