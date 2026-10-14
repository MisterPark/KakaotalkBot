using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace OpenChatRoomCleaner
{
    // 원본 화면 배율에서 매칭합니다. 배율/테마가 다르면 클릭하지 않고 재인식합니다.
    sealed class Pixels
    {
        internal readonly int Width, Height;
        internal readonly byte[] Data;
        internal Pixels(Bitmap source)
        {
            Width = source.Width; Height = source.Height;
            using (var copy = new Bitmap(Width, Height, PixelFormat.Format32bppArgb))
            {
                using (var g = Graphics.FromImage(copy)) g.DrawImageUnscaled(source, 0, 0);
                var bits = copy.LockBits(new Rectangle(0, 0, Width, Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                try { Data = new byte[Width * Height * 4]; Marshal.Copy(bits.Scan0, Data, 0, Data.Length); }
                finally { copy.UnlockBits(bits); }
            }
        }
        internal bool Find(Pixels template, Rectangle area, out Rectangle found)
        {
            found = Rectangle.Empty;
            area.Intersect(new Rectangle(0, 0, Width, Height));
            int lastX = area.Right - template.Width, lastY = area.Bottom - template.Height;
            for (int y = area.Top; y <= lastY; y++)
                for (int x = area.Left; x <= lastX; x++)
                    if (Fits(template, x, y, 4) && Fits(template, x, y, 1))
                    { found = new Rectangle(x, y, template.Width, template.Height); return true; }
            return false;
        }
        private bool Fits(Pixels t, int x, int y, int step)
        {
            int bad = 0, total = ((t.Width + step - 1) / step) * ((t.Height + step - 1) / step);
            for (int ty = 0; ty < t.Height; ty += step)
                for (int tx = 0; tx < t.Width; tx += step)
                {
                    int a = ((y + ty) * Width + x + tx) * 4, b = (ty * t.Width + tx) * 4;
                    if (Math.Abs(Data[a] - t.Data[b]) > 22 || Math.Abs(Data[a+1] - t.Data[b+1]) > 22 || Math.Abs(Data[a+2] - t.Data[b+2]) > 22)
                        if (++bad > total / 30) return false;
                }
            return true;
        }
        // 배경색은 비교하지 않고 흰색 아이콘의 전경/배경 모양을 함께 비교합니다.
        internal bool FindWhiteGlyph(Pixels glyph,Rectangle area,out Rectangle found)
        {
            found=Rectangle.Empty;area.Intersect(new Rectangle(0,0,Width,Height));
            int maxBad=glyph.Width*glyph.Height/20;
            for(int y=area.Top;y<=area.Bottom-glyph.Height;y++)
                for(int x=area.Left;x<=area.Right-glyph.Width;x++)
                {
                    int bad=0;
                    for(int gy=0;gy<glyph.Height && bad<=maxBad;gy++)
                        for(int gx=0;gx<glyph.Width;gx++)
                            if(IsWhite(((y+gy)*Width+x+gx)*4)!=glyph.IsWhite((gy*glyph.Width+gx)*4))
                                if(++bad>maxBad)break;
                    if(bad<=maxBad)
                    {
                        var hit=new Rectangle(x,y,glyph.Width,glyph.Height);
                        // 서로 다른 두 버튼이 후보이면 추측하지 않습니다.
                        if(!found.IsEmpty && !found.IntersectsWith(hit)){found=Rectangle.Empty;return false;}
                        if(found.IsEmpty)found=hit;
                    }
                }
            return !found.IsEmpty;
        }
        bool IsWhite(int offset)
        {
            int low=Math.Min(Data[offset],Math.Min(Data[offset+1],Data[offset+2]));
            int high=Math.Max(Data[offset],Math.Max(Data[offset+1],Data[offset+2]));
            return low>220 && high-low<30;
        }
        internal bool FindRoundBadge(Pixels badge,Rectangle area,out Rectangle found)
        {
            found=Rectangle.Empty;area.Intersect(new Rectangle(0,0,Width,Height));
            int cx=badge.Width/2,cy=badge.Height/2,radius=Math.Min(badge.Width,badge.Height)/2-1;
            for(int y=area.Top;y<=area.Bottom-badge.Height;y++)for(int x=area.Left;x<=area.Right-badge.Width;x++)
            {
                int bad=0,total=0;
                for(int by=0;by<badge.Height;by++)for(int bx=0;bx<badge.Width;bx++)
                {
                    if((bx-cx)*(bx-cx)+(by-cy)*(by-cy)>radius*radius)continue;
                    total++;int a=((y+by)*Width+x+bx)*4,b=(by*badge.Width+bx)*4;
                    if(Math.Abs(Data[a]-badge.Data[b])>30||Math.Abs(Data[a+1]-badge.Data[b+1])>30||Math.Abs(Data[a+2]-badge.Data[b+2])>30)bad++;
                }
                if(total>0&&bad<=total/8){found=new Rectangle(x,y,badge.Width,badge.Height);return true;}
            }
            return false;
        }
        internal int Ink(Rectangle area)
        {
            area.Intersect(new Rectangle(0,0,Width,Height)); int count = 0;
            for (int y=area.Top;y<area.Bottom;y++) for(int x=area.Left;x<area.Right;x++)
            { int p=(y*Width+x)*4; if(Data[p]<150 && Data[p+1]<150 && Data[p+2]<150) count++; }
            return count;
        }
        internal ulong Fingerprint(Rectangle area)
        {
            area.Intersect(new Rectangle(0,0,Width,Height)); ulong hash=1469598103934665603UL;
            for(int y=area.Top;y<area.Bottom;y+=2) for(int x=area.Left;x<area.Right;x+=2)
            { int p=(y*Width+x)*4; hash=unchecked((hash^(uint)(Data[p]/16+Data[p+1]/16*16+Data[p+2]/16*256))*1099511628211UL); }
            return hash;
        }
    }
    sealed class Row { internal Point Click; internal bool Protected, Uncertain; }
    sealed class Observation
    {
        internal Rectangle Members, Footer, ProfileMenu, Remove, Confirm;
        internal List<Row> Rows = new List<Row>();
        internal bool ListValid, ListViewportValid, NativeProfile;
        internal string ListFailure;
        internal ulong ListHash;
        internal Point ScrollPoint;
    }
    sealed class Vision
    {
        readonly Dictionary<string, Pixels> templates = new Dictionary<string, Pixels>();
        internal Vision(string directory)
        {
            foreach (var file in Directory.GetFiles(directory, "*.png"))
                using (var image = new Bitmap(file)) templates[Path.GetFileNameWithoutExtension(file)] = new Pixels(image);
            foreach(var key in new[]{"members","list-footer","list-search","owner","manager","manager2","self","profile-menu","profile-arrow","remove","confirm-title","confirm-body","confirm-ok"})
                if(!templates.ContainsKey(key)) throw new InvalidOperationException("이미지 누락: " + key);
        }
        bool Find(Pixels image,string name,Rectangle area,out Rectangle at) { return image.Find(templates[name],area,out at); }
        internal Rectangle FindProfileAction(Bitmap bitmap)
        {
            var pixels=new Pixels(bitmap);Rectangle found;
            // 프로필의 하단 버튼 띠만 검색합니다. 고정된 좌/중/우 슬롯은 사용하지 않습니다.
            var area=new Rectangle(8,Math.Max(0,bitmap.Height-120),Math.Max(0,bitmap.Width-16),Math.Min(98,bitmap.Height));
            return pixels.FindWhiteGlyph(templates["profile-arrow"],area,out found)?found:Rectangle.Empty;
        }
        internal static int AlignRowTop(int baseline,int badgeRowTop)
        {
            // 휠 스크롤은 행 높이(52) 단위로 멈추지 않습니다. 배지에서 실제 행 위상을
            // 구하고 검색칸 아래에 완전히 들어오는 첫 행으로 이동합니다.
            int offset=((badgeRowTop-baseline)%52+52)%52;
            // 테두리/배지 매칭의 1~3픽셀 오차로 첫 행 전체를 건너뛰지 않습니다.
            return baseline+(offset>=49?offset-52:offset);
        }
        internal Observation Read(Bitmap bitmap,bool readParticipantList=true,Stage? stage=null)
        {
            var p=new Pixels(bitmap); var full=new Rectangle(0,0,p.Width,p.Height); var o=new Observation(); Rectangle ignored;
            if(!stage.HasValue || stage==Stage.OpenList || stage==Stage.ReturnToList) Find(p,"members",new Rectangle(0,0,p.Width,Math.Min(90,p.Height)),out o.Members);
            if(!stage.HasValue || stage==Stage.Profile || stage==Stage.ReturnToList) Find(p,"profile-menu",full,out o.ProfileMenu);
            if(!stage.HasValue || stage==Stage.RemoveMenu || stage==Stage.ReturnToList) Find(p,"remove",full,out o.Remove);
            // 일반 확인 버튼만 찾지 않고, 내보내기 제목과 경고 문구까지 함께 확인합니다.
            if((!stage.HasValue || stage==Stage.Confirm) && Find(p,"confirm-title",full,out ignored) && Find(p,"confirm-body",full,out ignored)) Find(p,"confirm-ok",full,out o.Confirm);
            if(!readParticipantList || (stage.HasValue && stage!=Stage.OpenList && stage!=Stage.ReturnToList && stage!=Stage.SeekTop && stage!=Stage.Scan))return o;
            if(!Find(p,"list-footer",full,out o.Footer)){o.ListFailure="목록 하단 기준 이미지 없음";return o;}
            int rawLeft=o.Footer.X-80, left=rawLeft, bottom=o.Footer.Y-12;
            // 캡처 테두리와 가운데 정렬 반올림으로 1~몇 픽셀 차이가 생길 수 있습니다.
            // 좁은 허용 범위에서만 창 안으로 보정하고, 다른 레이아웃은 진단 후 대기합니다.
            if(left<0 && left>=-4)left=0;
            if(left+265>p.Width && left+265-p.Width<=4 && p.Width>=265)left=p.Width-265;
            // 큰 목록은 검색칸 아래에서 행이 시작합니다. 행 위치 보정은 이 기준을
            // 유지한 채 수행해야 하며, 절대 위치를 52로 나누어 상단으로 당기지 않습니다.
            Rectangle search;
            int top=Find(p,"list-search",new Rectangle(Math.Max(0,left),0,270,Math.Min(100,p.Height)),out search)
                ? search.Bottom+21 : 14;
            Rectangle badge;
            var badgeColumn=new Rectangle(Math.Max(0,left+40),top,30,Math.Max(0,bottom-top));
            if(p.FindRoundBadge(templates["owner"],badgeColumn,out badge)||p.FindRoundBadge(templates["manager"],badgeColumn,out badge)||p.FindRoundBadge(templates["manager2"],badgeColumn,out badge))
                top=AlignRowTop(top,badge.Y-24);
            if(left<0 || left+265>p.Width || bottom<top+40){o.ListFailure="목록 영역 계산 실패 (창 "+p.Width+"×"+p.Height+", 하단 "+o.Footer.X+","+o.Footer.Y+", 좌측 "+rawLeft+"→"+left+", 행 시작 "+top+", 끝 "+bottom+")";return o;}
            // 행이 잘려 읽히지 않아도 스크롤할 목록 영역과 변화 비교값은 먼저 제공합니다.
            o.ListViewportValid=true;
            o.ListHash=p.Fingerprint(new Rectangle(left+15,top,232,bottom-top));
            o.ScrollPoint=new Point(left+235,Math.Max(top+20,bottom-22));
            int rowCount=(bottom-top+4)/52;
            if(rowCount<1 || rowCount>30){o.ListFailure="완전히 표시된 행 없음";return o;}
            for(int i=0;i<rowCount;i++)
            {
                int y=top+i*52;
                // 빈 행/로딩 화면은 일반 참여자로 취급하지 않습니다.
                if(p.Ink(new Rectangle(left+70,y+7,170,28))<6){o.ListFailure=(i+1)+"번째 행 이름 영역 인식 실패";return o;}
                var badgeArea=new Rectangle(left+40,y+18,31,32);
                bool protect=p.FindRoundBadge(templates["owner"],badgeArea,out ignored) || p.FindRoundBadge(templates["manager"],badgeArea,out ignored)
                    || p.FindRoundBadge(templates["manager2"],badgeArea,out ignored) || p.FindRoundBadge(templates["self"],new Rectangle(left+66,y+8,30,28),out ignored);
                // 프로필 사진 색상으로 운영진을 추정하지 않습니다. 원형 배지 안의 왕관 모양까지 일치해야 제외합니다.
                o.Rows.Add(new Row{Click=new Point(left+38,y+20),Protected=protect});
            }
            o.ListValid=true;
            return o;
        }
    }
}
