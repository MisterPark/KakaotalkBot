using System;
using System.Collections.Generic;
using System.Globalization;

namespace KakaotalkBot
{
    public class User
    {
        // 닉네임은 표시용이며, 개인 식별에는 변경할 수 없는 ID만 사용한다.
        public long UserId { get; private set; }
        public string Nickname = string.Empty;
        public bool TakeAttendance = false;
        public DateTime AttendanceAt = DateTime.MinValue;
        public int Point = 0;
        public int Popularity = 0;
        public string Password = "0000";
        public int Contribution = 0;
        public int LeaveCount = 0;
        public int KickCount = 0;
        public long Experience = 0;
        // 레벨 L의 누적 기준은 100 * (L-1) * L / 2입니다. 레벨은 경험치에서 계산합니다.
        public int Level
        {
            get
            {
                int low = 1, high = int.MaxValue;
                while (low < high)
                {
                    int middle = low + (int)(((long)high - low + 1) / 2);
                    if (100m * (middle - 1) * middle / 2 <= Experience) low = middle;
                    else high = middle - 1;
                }
                return low;
            }
        }
        public decimal NextLevelExperience { get { return 100m * Level * (Level + 1m) / 2; } }

        public User(long userId)
        {
            if (userId <= 0) throw new ArgumentOutOfRangeException("userId");
            UserId = userId;
        }

        public List<object> ToRow()
        {
            return new List<object>() { Nickname, TakeAttendance, AttendanceAt.ToString("O", CultureInfo.InvariantCulture),
                Point, Popularity, Password, Contribution, UserId.ToString(CultureInfo.InvariantCulture), LeaveCount, KickCount,
                Experience.ToString(CultureInfo.InvariantCulture), Level };
        }

        public static User ToUser(List<string> list)
        {
            long id;
            if (list == null || list.Count < 8 || list.Count > 12 || !long.TryParse(list[7], NumberStyles.None, CultureInfo.InvariantCulture, out id) || id <= 0)
                throw new FormatException("user_id는 반올림되지 않은 양의 정수 문자열이어야 합니다.");
            User user = new User(id);
            user.Nickname = list[0];
            user.TakeAttendance = Convert.ToBoolean(list[1]);
            user.AttendanceAt = DateTime.Parse(list[2], CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
            user.Point = int.Parse(list[3], CultureInfo.InvariantCulture);
            user.Popularity = int.Parse(list[4], CultureInfo.InvariantCulture);
            user.Password = list[5];
            user.Contribution = int.Parse(list[6], CultureInfo.InvariantCulture);
            // 기존 ID 계정의 새 컬럼이 비어 있으면 0부터 집계한다.
            user.LeaveCount = ReadCount(list, 8);
            user.KickCount = ReadCount(list, 9);
            if (list.Count > 10 && !string.IsNullOrWhiteSpace(list[10]) &&
                !long.TryParse(list[10], NumberStyles.None, CultureInfo.InvariantCulture, out user.Experience))
                throw new FormatException("경험치는 0 이상의 정수여야 합니다.");
            if (user.Point < 0 || user.Contribution < 0) throw new FormatException("포인트와 기여도는 음수일 수 없습니다.");

            return user;
        }

        private static int ReadCount(List<string> row, int index)
        {
            if (row.Count <= index || string.IsNullOrWhiteSpace(row[index])) return 0;
            int count;
            if (!int.TryParse(row[index], NumberStyles.None, CultureInfo.InvariantCulture, out count))
                throw new FormatException("퇴장·강퇴 횟수는 0 이상의 정수여야 합니다.");
            return count;
        }
    }
}
