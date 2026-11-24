using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace KakaotalkBot
{
    internal sealed class PersonalContactSettings
    {
        public string AccountPath { get; set; }
        public long ChatId { get; set; }
        public long UserId { get; set; }
        public string Notes { get; set; } = "";
    }

    internal sealed class PersonalReplySettings
    {
        public string Model { get; set; } = "gpt-6.1-terra";
        public string ProtectedApiKey { get; set; }
        public string OwnerNotes { get; set; } = "짧고 자연스럽게 대화한다. 상대에게 평소 쓰던 존댓말/반말과 이모티콘 사용을 따른다.";
        public bool ImportHistory { get; set; } = true;
        public List<PersonalContactSettings> Contacts { get; set; } = new List<PersonalContactSettings>();
        internal static string FilePath { get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PersonalAI", "settings.json"); } }
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("KakaotalkBot.PersonalAI.v1");

        internal void SetApiKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length > 4096 || value.IndexOfAny(new[] { '\r', '\n' }) >= 0)
                throw new ArgumentException("올바른 OpenAI API 키를 입력하세요.");
            ProtectedApiKey = Convert.ToBase64String(ProtectedData.Protect(Encoding.UTF8.GetBytes(value.Trim()), Entropy, DataProtectionScope.CurrentUser));
        }
        internal string GetApiKey()
        {
            if (string.IsNullOrWhiteSpace(ProtectedApiKey)) throw new InvalidOperationException("OpenAI API 키를 먼저 저장하세요.");
            try { return Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(ProtectedApiKey), Entropy, DataProtectionScope.CurrentUser)); }
            catch (Exception error) when (error is CryptographicException || error is FormatException)
            { throw new InvalidOperationException("이 Windows 계정에서 API 키를 읽을 수 없습니다. 키를 다시 저장하세요."); }
        }
        internal static PersonalReplySettings Load()
        {
            if (!File.Exists(FilePath)) return new PersonalReplySettings();
            var value = JsonConvert.DeserializeObject<PersonalReplySettings>(File.ReadAllText(FilePath));
            if (value == null || value.Contacts == null) throw new InvalidDataException("개인 AI 설정을 읽을 수 없습니다.");
            return value;
        }
        internal void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
            string temporary = FilePath + ".tmp";
            File.WriteAllText(temporary, JsonConvert.SerializeObject(this, Formatting.Indented));
            if (File.Exists(FilePath)) File.Replace(temporary, FilePath, null); else File.Move(temporary, FilePath);
        }
    }
}
