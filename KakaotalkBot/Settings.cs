using System.IO;
using Newtonsoft.Json;

namespace KakaotalkBot
{
    public class Settings
    {
        private static Settings instance;
        public static Settings Instance
        {
            get 
            {
                if (instance == null)
                {
                    instance = Load();
                }

                return instance; 
            }
        }

        private Settings()
        {

        }
        public string SpreadsheetId { get; set; }
        public string ApplicationName { get; set; }
        public string ChatAccountPath { get; set; }
        public long ChatRoomId { get; set; }
        public long ChatOwnAuthorId { get; set; }
        public string ChatSendRoomName { get; set; }
        public bool OperatorAlertsEnabled { get; set; }
        public string OperatorAlertAccount { get; set; }
        public long OperatorAlertRoomId { get; set; }
        public long OperatorAlertSince { get; set; }
        public bool AlertReentry { get; set; } = true;
        public bool AlertMemo { get; set; } = true;
        public bool AlertHistory { get; set; } = true;
        public bool AlertMonthly { get; set; } = true;
        public bool AlertReset { get; set; } = true;
        public bool AlertMemoBody { get; set; }

        public static void Save(Settings settings)
        {
            string json = JsonConvert.SerializeObject(settings);
            string path = Path.GetFullPath("settings.json"), temp = path + ".tmp";
            File.WriteAllText(temp, json);
            if (File.Exists(path)) File.Replace(temp, path, path + ".bak"); else File.Move(temp, path);
        }

        public static Settings Load()
        {
            if (!File.Exists("settings.json"))
            {
                var defaultSettings = new Settings
                {
                    SpreadsheetId = "<YOUR_SPREADSHEET_ID>",
                    ApplicationName = "KakaoChatLogger"
                };
                Save(defaultSettings);
                return defaultSettings;
            }
            string json = File.ReadAllText("settings.json");
            return JsonConvert.DeserializeObject<Settings>(json);
        }
    }
}
