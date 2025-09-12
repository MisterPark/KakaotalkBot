using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Data.SQLite;

namespace KakaotalkBot
{
    public class Database
    {
        private GoogleSheetHelper keywordSheet;
        private List<string> keywords = new List<string>();
        private List<List<string>> commands = new List<List<string>>();
        private List<User> userTable = new List<User>();
        private List<Quiz> commonSenses = new List<Quiz>();
        private List<Topic> topics = new List<Topic>();

        private RandomNumberGenerator random;
        private byte[] randomBytes = new byte[4];
        private int currentAnswerIndex = -1;
        private string[] categoryOrder = new string[]
        {
            "🗳정치", "🎞역사", "📊경제", "👨‍🏫인문학", "🌍기타"
        };

        public List<List<string>> Commands { get { return commands; } }
        public List<string> Keywords { get { return keywords; } }
        public List<User> UserTable { get { return userTable; } }
        public List<Quiz> CommonSenses { get { return commonSenses; } }
        public int CurrentAnswerIndex { get {  return currentAnswerIndex; }  set { currentAnswerIndex = value; } }
        public List<Topic> Topics { get { return topics; } }
        public string[] CategoryOrder { get { return categoryOrder; } }

        public Database(string applicationName, string spreadsheetId)
        {
            random = RandomNumberGenerator.Create();
            keywordSheet = new GoogleSheetHelper(applicationName, spreadsheetId);
            commands = GetCommanads();
            keywords = GetKeywords();
            userTable = GetUserTable();
            commonSenses = GetCommonSenses();
            topics = GetTopic();
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

        public void UpdateCommonSenses()
        {
            commonSenses = GetCommonSenses();
        }

        public void UpdateTopic()
        {
            topics = GetTopic();
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

        private List<Topic> GetTopic()
        {
            var db = keywordSheet.ReadAllFromSheet("Topic");
            List<Topic> users = new List<Topic>();
            for (int i = 0; i < db.Count; i++)
            {
                if (i == 0) continue;
                var row = db[i];
                users.Add(Topic.ToTopic(row));
            }
            return users;
        }

        private List<Quiz> GetCommonSenses()
        {
            var db = keywordSheet.ReadAllFromSheet("상식퀴즈");
            List<Quiz> users = new List<Quiz>();
            for (int i = 0; i < db.Count; i++) 
            {
                if (i == 0) continue;
                var row = db[i];
                users.Add(Quiz.ToCommonSense(row));
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

        public void SetNextCommonSense()
        {
            random.GetBytes(randomBytes);
            int randomValue = BitConverter.ToInt32(randomBytes, 0);
            randomValue = Math.Abs(randomValue);
            currentAnswerIndex = randomValue % commonSenses.Count;
        }

        public string GetCommonSenseText()
        {
            if(currentAnswerIndex < 0 )
            {
                return string.Empty;
            }
            Quiz cs = commonSenses[currentAnswerIndex];
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("[상식퀴즈]");
            sb.AppendLine(cs.Question);
            sb.AppendLine($"분류: {cs.Category}");
            sb.AppendLine($"난이도: {cs.Difficulty}");
            sb.Append($"힌트: {cs.Hint}");
            
            return sb.ToString();
        }

        public Quiz GetCurrentQuiz()
        {
            if (currentAnswerIndex < 0) return null;

            Quiz cs = commonSenses[currentAnswerIndex];
            return cs;
        }

        public List<Topic> GetOrderedTopics()
        {
            return topics.OrderByDescending(x => x.CreatedAt).OrderBy(x => Array.IndexOf(categoryOrder, x.Category)).ToList();
        }
    }
}
