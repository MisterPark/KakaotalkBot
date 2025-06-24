using System;
using System.Collections.Generic;

namespace KakaotalkBot
{
    public class User
    {
        public string Nickname = string.Empty;
        public bool TakeAttendance = false;
        public DateTime AttendanceAt = DateTime.MinValue;
        public int Point = 0;

        public List<object> ToRow()
        {
            return new List<object>() { Nickname, TakeAttendance, AttendanceAt, Point };
        }

        public static User ToUser(List<string> list)
        {
            User user = new User();
            user.Nickname = list[0];
            user.TakeAttendance = Convert.ToBoolean(list[1]);
            user.AttendanceAt = DateTime.Parse(list[2]);
            user.Point = Convert.ToInt32(list[3]);

            return user;
        }
    }
}
