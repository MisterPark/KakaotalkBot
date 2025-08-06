using System.IO;
using Newtonsoft.Json;

namespace GoogleSheet
{
    public class Settings
    {
        public string SpreadsheetId { get; set; }
        public string ApplicationName { get; set; }

        public static void Save(Settings settings)
        {
            string json = JsonConvert.SerializeObject(settings);
            File.WriteAllText("settings.json", json);
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
