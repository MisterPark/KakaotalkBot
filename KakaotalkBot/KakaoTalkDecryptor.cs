using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace KakaotalkBot
{
    public class KakaoTalkDecryptor
    {
        private const int PageSize = 4096;

        public static void DecryptKakaoEdb(string edbInputPath, string edbOutputPath, string userId, byte[] hardcodedKey)
        {
            // 1. 레지스트리 값 읽기
            string sysUuid = GetRegistryValue("sys_uuid");
            string hddModel = GetRegistryValue("hdd_model");
            string hddSerial = GetRegistryValue("hdd_serial");

            if (string.IsNullOrEmpty(sysUuid) || string.IsNullOrEmpty(hddModel) || string.IsNullOrEmpty(hddSerial))
                throw new Exception("필요한 레지스트리 값을 찾을 수 없습니다.");

            // 2. Pragma 문자열 생성
            string pragma = $"{sysUuid}|{hddModel}|{hddSerial}";
            byte[] pragmaBytes = Encoding.UTF8.GetBytes(pragma);

            // 3. AES-CBC로 암호화 (Padding 없음, IV=0)
            byte[] encryptedPragma;
            using (Aes aes = Aes.Create())
            {
                aes.Key = hardcodedKey;
                aes.IV = new byte[16]; // all zeros
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;

                using (ICryptoTransform encryptor = aes.CreateEncryptor())
                {
                    encryptedPragma = encryptor.TransformFinalBlock(pragmaBytes, 0, pragmaBytes.Length);
                }
            }

            string base64Encrypted = Convert.ToBase64String(encryptedPragma);

            // 4. SHA512 해싱 후 base64 인코딩
            string finalPragma;
            using (SHA512 sha = SHA512.Create())
            {
                byte[] hashed = sha.ComputeHash(Encoding.UTF8.GetBytes(base64Encrypted));
                finalPragma = Convert.ToBase64String(hashed);
            }

            // 5. 키와 IV 생성
            string combined = finalPragma + userId;
            while (combined.Length < 512)
            {
                combined += combined;
            }
            combined = combined.Substring(0, 512);

            byte[] key, iv;
            using (MD5 md5 = MD5.Create())
            {
                key = md5.ComputeHash(Encoding.UTF8.GetBytes(combined));
                iv = md5.ComputeHash(Encoding.UTF8.GetBytes(Convert.ToBase64String(key)));
            }

            // 6. edb 파일 복호화
            DecryptFile(edbInputPath, edbOutputPath, key, iv);
        }

        private static void DecryptFile(string inputPath, string outputPath, byte[] key, byte[] iv)
        {
            using (FileStream input = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            using (FileStream output = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;

                using (ICryptoTransform decryptor = aes.CreateDecryptor())
                {
                    byte[] buffer = new byte[PageSize];
                    int read;

                    while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        byte[] decrypted = decryptor.TransformFinalBlock(buffer, 0, read);
                        output.Write(decrypted, 0, decrypted.Length);
                    }
                }
            }
        }

        private static string GetRegistryValue(string name)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Kakao\KakaoTalk\DeviceInfo\20241213-091852-157"))
            {
                return key?.GetValue(name)?.ToString() ?? "";
            }
        }
    }
}
