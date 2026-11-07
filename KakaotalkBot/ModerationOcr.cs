using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace KakaotalkBot
{
    internal sealed class OcrLabel
    {
        internal string Text;
        internal Rectangle Bounds;
    }
    internal static class ModerationOcr
    {
        internal static string Normalize(string text) { return new string((text ?? "").Where(c => !char.IsWhiteSpace(c)).ToArray()); }
        internal static List<OcrLabel> Read(Bitmap bitmap)
        {
            var engine = OcrEngine.TryCreateFromLanguage(new Language("ko-KR"));
            if (engine == null) throw new InvalidOperationException("Windows 한국어 OCR 언어 구성 요소가 필요합니다.");
            using (var encoded = new MemoryStream())
            using (var stream = new InMemoryRandomAccessStream())
            {
                float scale = Math.Min(3f, OcrEngine.MaxImageDimension / (float)Math.Max(bitmap.Width, bitmap.Height));
                using (var enlarged = new Bitmap((int)(bitmap.Width * scale), (int)(bitmap.Height * scale)))
                {
                    using (var graphics = Graphics.FromImage(enlarged))
                    { graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic; graphics.DrawImage(bitmap, 0, 0, enlarged.Width, enlarged.Height); }
                    enlarged.Save(encoded, ImageFormat.Png);
                }
                using (var writer = new DataWriter(stream))
                {
                    writer.WriteBytes(encoded.ToArray()); writer.StoreAsync().AsTask().GetAwaiter().GetResult(); writer.DetachStream();
                }
                stream.Seek(0);
                var decoder = BitmapDecoder.CreateAsync(stream).AsTask().GetAwaiter().GetResult();
                using (var pixels = decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied).AsTask().GetAwaiter().GetResult())
                {
                    var result = engine.RecognizeAsync(pixels).AsTask().GetAwaiter().GetResult();
                    return result.Lines.Select(line => new OcrLabel { Text = line.Text, Bounds = line.Words.Select(word => Rectangle.Round(new RectangleF((float)word.BoundingRect.X / scale, (float)word.BoundingRect.Y / scale, (float)word.BoundingRect.Width / scale, (float)word.BoundingRect.Height / scale))).Aggregate(Rectangle.Union) }).ToList();
                }
            }
        }
    }
}
