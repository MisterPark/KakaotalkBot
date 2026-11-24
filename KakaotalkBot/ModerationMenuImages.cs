using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using OpenChatRoomCleaner;

namespace KakaotalkBot
{
    internal sealed class ModerationMenuImages
    {
        private readonly string directory;
        private readonly string fixedAssets;
        private readonly Dictionary<string, Pixels> templates = new Dictionary<string, Pixels>();
        internal ModerationMenuImages(string directory) : this(directory, null) { }
        internal ModerationMenuImages(string directory, string fixedAssets) { this.directory = directory; this.fixedAssets = fixedAssets; }
        private static string Key(string text)
        {
            switch (ModerationMenu.Normalize(text))
            {
                case "오픈채팅관리": return "openchat-management";
                case "내보낸상대관리": return "banned-member-management";
                default: return null;
            }
        }
        internal bool Find(string text, Pixels image, out Rectangle bounds)
        {
            bounds = Rectangle.Empty; string key = Key(text); Pixels template;
            if (key == null) return false;
            if (!Load(key, Path.Combine(directory, key + ".png"), out template)) return false;
            return FindUnique(image, template, new Rectangle(0, 0, image.Width, image.Height), out bounds);
        }
        internal bool FindIcon(string text, Pixels image, out Rectangle bounds)
        {
            bounds = Rectangle.Empty;
            if (fixedAssets == null || Key(text) != "openchat-management") return false;
            Pixels template;
            if (!Load("openchat-management-icon", Path.Combine(fixedAssets, "openchat-management-icon.png"), out template)) return false;
            // 사용자가 제공한 아이콘을 EVA_Menu의 왼쪽 아이콘 열에서만 찾습니다.
            var area = new Rectangle(0, 0, Math.Min(image.Width, Math.Max(48, template.Width * 3)), image.Height);
            return FindUnique(image, template, area, out bounds);
        }
        internal bool FindButton(string text, Pixels image, out Rectangle bounds)
        {
            bounds = Rectangle.Empty;
            if (fixedAssets == null || Key(text) != "banned-member-management") return false;
            Pixels template;
            if (!Load("banned-member-management-button", Path.Combine(fixedAssets, "banned-member-management-button.png"), out template)) return false;
            return FindUnique(image, template, new Rectangle(0, 0, image.Width, image.Height), out bounds);
        }
        private bool Load(string key, string path, out Pixels template)
        {
            if (!templates.TryGetValue(key, out template))
            {
                if (!File.Exists(path)) return false;
                try { using (var bitmap = new Bitmap(path)) template = new Pixels(bitmap); }
                catch (IOException) { return false; }
                catch (ArgumentException) { return false; }
                catch (ExternalException) { return false; }
                templates[key] = template;
            }
            return true;
        }
        private static bool FindUnique(Pixels image, Pixels template, Rectangle area, out Rectangle bounds)
        {
            if (!image.Find(template, area, out bounds)) return false;
            Rectangle duplicate;
            if (image.Find(template, new Rectangle(area.Left, bounds.Bottom, area.Width, area.Bottom - bounds.Bottom), out duplicate) ||
                image.Find(template, new Rectangle(bounds.Right, bounds.Top, area.Right - bounds.Right, area.Bottom - bounds.Top), out duplicate))
            { bounds = Rectangle.Empty; return false; }
            return true;
        }
        internal void Remember(string text, Bitmap image, Rectangle bounds)
        {
            string key = Key(text);
            if (key == null || bounds.Width < 32 || bounds.Height < 8 || bounds.Height > 128 || !new Rectangle(Point.Empty, image.Size).Contains(bounds)) return;
            bounds.Inflate(2, 2); bounds.Intersect(new Rectangle(Point.Empty, image.Size));
            using (var crop = image.Clone(bounds, PixelFormat.Format32bppArgb))
            {
                templates[key] = new Pixels(crop);
                try { Directory.CreateDirectory(directory); crop.Save(Path.Combine(directory, key + ".png"), ImageFormat.Png); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
                catch (ExternalException) { }
            }
        }
    }
}
