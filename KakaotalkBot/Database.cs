using System;
using System.Collections.Generic;
using System.Linq;

namespace KakaotalkBot
{
    public class Database
    {
        private GoogleSheetHelper keywordSheet;
        private List<string> keywords = new List<string>();
        private List<List<string>> commands = new List<List<string>>();
        private List<User> userTable = new List<User>();

        public List<List<string>> Commands { get { return commands; } }
        public List<string> Keywords { get { return keywords; } }
        public List<User> UserTable { get { return userTable; } }

        public Database(string applicationName, string spreadsheetId)
        {
            keywordSheet = new GoogleSheetHelper(applicationName, spreadsheetId);
            commands = GetCommanads();
            keywords = GetKeywords();
            userTable = GetUserTable();
        }

        public void UpdateCommands()
        {
            commands = GetCommanads();
            keywords = GetKeywords();
        }

        public void ResetAttendance()
        {
            List<List<object>> users = new List<List<object>>();
            foreach (User user in userTable)
            {
                user.TakeAttendance = false;
                users.Add(user.ToRow());
            }
            keywordSheet.WriteToSheetAll("DB", users);
        }

        public void UpdateUserTable()
        {
            List<List<object>> users = new List<List<object>>();
            foreach (User user in userTable)
            {
                users.Add(user.ToRow());
            }
            keywordSheet.WriteToSheetAll("DB", users);
        }

        public void AddUser(string username)
        {
            if (FindUser(username)) return;

            userTable.Add(new User() {Nickname = username });
        }

        public bool FindUser(string username) 
        {
            foreach (User user in userTable)
            {
                if(user.Nickname == username)
                {
                    return true;
                }
            }
            return false;
        }

        public bool FindUser(string username, out User user)
        {
            foreach (User u in userTable)
            {
                if (u.Nickname == username)
                {
                    user = u;
                    return true;
                }
            }
            user = null;
            return false;
        }

        public bool CheckAttendance(string username)
        {
            AddUser(username);

            foreach (User user in userTable)
            {
                if (user.Nickname == username)
                {
                    if(user.TakeAttendance)
                    {
                        return true;
                    }
                    else
                    {
                        user.TakeAttendance = true;
                        user.AttendanceAt = DateTime.Now;
                        user.Point += 10;
                        return false;
                    }
                }
            }

            return false;
        }

        public List<User> GetPopularityRank()
        {
            return userTable.OrderByDescending(x => x.Popularity).ToList();
        }

        private List<User> GetUserTable()
        {
            var db = keywordSheet.ReadAllFromSheet("DB");
            List<User> users = new List<User>();
            foreach (var row in db)
            {
                users.Add(User.ToUser(row));
            }
            return users;
        }

        private List<List<string>> GetCommanads()
        {
            var keywords = keywordSheet.ReadAllFromSheet("키워드");
            return keywords;
        }

        public List<string> GetKeywords()
        {
            List<string> keywordList = new List<string>();
            foreach (var row in commands)
            {
                keywordList.Add(row[0]);
            }

            return keywordList;
        }

        public string GetAnswer(string keyword)
        {
            var keywords = commands;
            foreach (var row in keywords)
            {
                if (row[0] == keyword && row.Count > 1) 
                {
                    return row[1];
                }
            }

            return null;
        }

        public string[] GetAnswers(string keyword)
        {
            var keywords = commands;
            foreach (var row in keywords)
            {
                if (row[0] == keyword && row.Count > 1)
                {
                    return row[1].Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                }
            }

            return null;
        }
    }
}
