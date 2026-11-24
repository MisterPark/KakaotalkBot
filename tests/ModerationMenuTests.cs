using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using KakaotalkBot;

internal static class ModerationMenuTests
{
    private const BindingFlags Flags = BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
    private static readonly Assembly App = typeof(Bot).Assembly;
    private static int passed;
    private static Type Type(string name) { return App.GetType(name.Contains(".") ? name : "KakaotalkBot." + name, true); }
    private static object Field(object target, string name) { return target.GetType().GetField(name, Flags).GetValue(target); }
    private static void Set(object target, string name, object value) { target.GetType().GetField(name, Flags).SetValue(target, value); }
    private static object Call(string type, string method, params object[] args) { return Type(type).GetMethod(method, Flags).Invoke(null, args); }
    private static object Invoke(object target, string method, params object[] args) { return target.GetType().GetMethod(method, Flags).Invoke(target, args); }
    private static void Check(bool ok, string message) { if (!ok) throw new Exception(message); passed++; Console.WriteLine("PASS " + message); }
    private static object Window(int handle, int owner, int parent, string kind, int width = 160, int height = 70, uint process = 10, uint thread = 20)
    {
        var window = Activator.CreateInstance(Type("OpenChatRoomCleaner.Window"), true);
        Set(window, "Handle", new IntPtr(handle)); Set(window, "Owner", new IntPtr(owner)); Set(window, "Parent", new IntPtr(parent));
        Set(window, "ClassName", kind); Set(window, "Visible", true); Set(window, "Process", process); Set(window, "Thread", thread);
        Set(window, "Bounds", new Rectangle(100, 200, width, height)); return window;
    }
    private static IList List(string type, params object[] values)
    {
        var list = (IList)Activator.CreateInstance(typeof(System.Collections.Generic.List<>).MakeGenericType(Type(type)));
        foreach (var value in values) list.Add(value); return list;
    }
    private static object Label(string text)
    {
        var label = Activator.CreateInstance(Type("OcrLabel"), true); Set(label, "Text", text); Set(label, "Bounds", new Rectangle(10, 10, 100, 20)); return label;
    }
    private static object Cache(string path) { return Activator.CreateInstance(Type("ModerationMenuImages"), Flags, null, new object[] { path }, null); }
    private static bool Find(object cache, string text, Bitmap image, out Rectangle bounds, string method = "Find")
    {
        var pixels = Activator.CreateInstance(Type("OpenChatRoomCleaner.Pixels"), Flags, null, new object[] { image }, null);
        var args = new object[] { text, pixels, Rectangle.Empty };
        bool found = (bool)Invoke(cache, method, args); bounds = (Rectangle)args[2]; return found;
    }
    private static int Main()
    {
        try
        {
            var room = Window(1, 0, 0, "EVA_Window_Dblclk", 380, 966);
            var small = Window(2, 1, 1, "EVA_Window", 160, 70);
            var native = Window(3, 0, 0, "#32768");
            var otherRoom = Window(4, 0, 0, "EVA_Window_Dblclk", 380, 900);
            var submenu = Window(5, 2, 0, "#32768", 150, 48);
            var unrelated = Window(6, 4, 4, "EVA_Window");
            var all = List("OpenChatRoomCleaner.Window", room, small, native, otherRoom, submenu, unrelated);
            Func<object, int, bool> candidate = (window, menuOwner) => (bool)Call("ModerationMenu", "TextCandidate", window, room, all, new IntPtr(menuOwner));
            Check(candidate(small, 0), "180x100보다 작은 소유 메뉴도 탐색");
            Check(candidate(submenu, 0), "작은 2단 하위 메뉴의 방 소유 관계 확인");
            Check(candidate(native, 1), "Owner가 없는 표준 팝업은 실제 메뉴 소유 방으로 확인");
            Check(!candidate(native, 4) && !candidate(native, 0), "다른 방과 소유 불명 표준 메뉴 제외");
            Check(!candidate(unrelated, 1), "같은 프로세스의 다른 방 메뉴 제외");
            Check(!candidate(room, 1), "채팅 본문에서 메뉴 문구를 찾지 않음");
            Check(!candidate(Window(8, 1, 1, "EVA_VH_ListControl_Dblclk"), 1), "채팅 목록을 메뉴로 오인하지 않음");
            Check(!candidate(Window(9, 1, 1, "RICHEDIT50W"), 1), "입력창을 메뉴로 오인하지 않음");
            Check(!candidate(Window(10, 1, 1, "EVA_Window", process: 99), 1), "다른 프로세스 제외");
            Check(!candidate(Window(11, 0, 0, "#32768", thread: 99), 1), "다른 UI 스레드의 소유 없는 팝업 제외");
            Set(small, "Visible", false); Check(!candidate(small, 0), "숨겨진 메뉴 제외"); Set(small, "Visible", true);
            Set(small, "Owner", new IntPtr(99)); Check(candidate(small, 0), "소유 핸들 외에 실제 부모 경로도 확인");
            var cycle = Window(12, 13, 0, "EVA_Window"); var cycle2 = Window(13, 12, 0, "EVA_Window"); all.Add(cycle); all.Add(cycle2);
            Check(!candidate(cycle, 0), "잘못된 소유 관계 순환에서 종료");

            // 실제 관측한 구조: 제목 없는 422x642 소유 창 → 400x494 내용 → 14픽셀 스크롤바.
            var dialog = Window(21, 1, 1, "EVA_Window_Dblclk", 422, 642);
            Set(dialog, "Root", new IntPtr(21)); Set(dialog, "Title", "");
            var pane = Window(22, 0, 21, "EVA_ChildWindow", 400, 494);
            Set(pane, "Bounds", new Rectangle(121, 269, 400, 494));
            var scrollbar = Window(23, 0, 22, "_EVA_CustomScrollCtrl", 14, 494);
            Set(scrollbar, "Bounds", new Rectangle(507, 269, 14, 494));
            all.Add(dialog); all.Add(pane); all.Add(scrollbar);
            Check(ReferenceEquals(Call("ModerationNavigation", "Dialog", room, all), dialog), "제목 없는 실제 관리 창 구조 식별·이미 열린 창 재사용");
            Check(ReferenceEquals(Call("ModerationNavigation", "Content", dialog, all), pane), "채팅 본문 대신 관리 창 내용 영역 선택");
            Set(dialog, "Owner", new IntPtr(4)); Set(dialog, "Parent", new IntPtr(4));
            Check(Call("ModerationNavigation", "Dialog", room, all) == null, "다른 방의 관리 창은 스크롤하지 않음");
            Set(dialog, "Owner", new IntPtr(1)); Set(dialog, "Parent", new IntPtr(1));
            Set(scrollbar, "Visible", false);
            Check(Call("ModerationNavigation", "Dialog", room, all) == null, "스크롤 영역을 확인하지 못하면 다른 창에서 추측하지 않음");
            Set(scrollbar, "Visible", true);
            var point = (Point)Call("ModerationNavigation", "ScrollPoint", new Size(400, 494));
            Check(new Rectangle(0, 0, 400, 494).Contains(point) && point.X > 350, "휠 입력 지점은 관리 내용 영역 안쪽");

            Check((string)Call("ModerationMenu", "Normalize", "오픈채팅 관리(&O)\tAlt+O") != "오픈채팅관리", "알 수 없는 접미사는 임의로 무시하지 않음");
            Check((string)Call("ModerationMenu", "Normalize", "오픈채팅 &관리\tAlt+O") == "오픈채팅관리", "메뉴 단축키 표기와 공백 처리");
            Check(Call("ModerationMenu", "Exact", List("OcrLabel", Label("오픈채팅 관리 ▶")), "오픈채팅 관리") != null, "OCR 하위 메뉴 화살표 처리");
            // 실제 EVA_Menu(179x335)에서 수집한 OCR 결과: '채'를 '재'로 읽었습니다.
            var observed = List("OcrLabel", Label("69 재팅방 IJRL 복사"), Label("口 재탕방 서람"), Label("口 특기시단"),
                Label("`/ 입장/틔장 표시"), Label("오픈재팅 관리"), Label("오픈재탕 커버 보기"), Label("쇼 것못 (beta)"), Label("대화 내용"), Label("단 재팅방 나가기"));
            var management = Call("ModerationMenu", "Exact", observed, "오픈채팅 관리");
            Check(Call("ModerationMenu", "Exact", List("OcrLabel", Label("오픈재탕 커버 보기")), "오픈채팅 관리") == null, "인접한 오픈채팅 커버 메뉴는 선택하지 않음");
            Check(Call("ModerationMenu", "Exact", List("OcrLabel", Label("내보낸 상대 관리")), "내보내기") == null, "부분 문구로 다른 명령 선택 금지");
            Check(Call("ModerationMenu", "Exact", List("OcrLabel", Label("오픈채팅 관리"), Label("오픈채팅관리")), "오픈채팅 관리") == null, "중복 문구 거부");
            var screen = new Rectangle(-500, 200, 160, 70);
            Check((Rectangle)Call("ModerationMenu", "RelativeBounds", screen, new Rectangle(-490, 210, 130, 24)) == new Rectangle(10, 10, 130, 24), "다중 모니터 메뉴 좌표 변환");
            Check(((Rectangle)Call("ModerationMenu", "RelativeBounds", screen, new Rectangle(0, 0, 130, 24))).IsEmpty, "메뉴 밖 좌표 거부");

            // 실제 Win32 메뉴 API를 검사하되 메뉴를 표시하거나 입력을 보내지 않습니다.
            IntPtr menu = CreatePopupMenu(); if (menu == IntPtr.Zero) throw new Exception("CreatePopupMenu failed");
            try
            {
                if (!AppendMenu(menu, 0, (UIntPtr)1, "오픈채팅 관리") || !AppendMenu(menu, 3, (UIntPtr)2, "내보낸 상대 관리") || !AppendMenu(menu, 0x800, UIntPtr.Zero, null)) throw new Exception("AppendMenu failed");
                var items = ((IEnumerable)Call("ModerationMenu", "ReadItems", menu)).Cast<object>().ToArray();
                Check(items.Length == 2 && (string)Field(items[0], "Text") == "오픈채팅 관리", "Win32에서 한국어 메뉴명 직접 읽기·구분선 제외");
                Check((bool)Field(items[0], "Enabled") && !(bool)Field(items[1], "Enabled"), "활성·비활성 메뉴 구분");
            }
            finally { DestroyMenu(menu); }

            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cases", "menu-cache");
            var cache = Cache(directory); Rectangle found;
            string assets = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ModerationAssets");
            var iconCache = Activator.CreateInstance(Type("ModerationMenuImages"), Flags, null, new object[] { directory, assets }, null);
            using (var icon = new Bitmap(Path.Combine(assets, "openchat-management-icon.png")))
            using (var menuImage = new Bitmap(179, 335))
            {
                using (var graphics = Graphics.FromImage(menuImage))
                { graphics.Clear(Color.White); graphics.DrawImageUnscaled(icon, 8, 135); }
                Check(management == null && Find(iconCache, "오픈채팅 관리", menuImage, out found, "FindIcon") && found.Contains(new Point(8 + icon.Width / 2, 135 + icon.Height / 2)),
                    "실제 OCR이 오픈재팅으로 틀려도 제공된 아이콘으로 처음부터 선택");
                Check(!Find(iconCache, "내보낸 상대 관리", menuImage, out found, "FindIcon"), "제공된 아이콘은 지정한 관리 메뉴에만 사용");
                using (var graphics = Graphics.FromImage(menuImage)) graphics.DrawImageUnscaled(icon, 8, 210);
                Check(!Find(iconCache, "오픈채팅 관리", menuImage, out found, "FindIcon"), "같은 아이콘이 두 행이면 선택하지 않음");
                using (var graphics = Graphics.FromImage(menuImage)) { graphics.Clear(Color.White); graphics.DrawImageUnscaled(icon, 130, 135); }
                Check(!Find(iconCache, "오픈채팅 관리", menuImage, out found, "FindIcon"), "메뉴 오른쪽의 유사 이미지는 제외");
                using (var graphics = Graphics.FromImage(menuImage)) graphics.Clear(Color.White);
                Check(!Find(iconCache, "오픈채팅 관리", menuImage, out found, "FindIcon"), "아이콘 없는 흰 화면은 선택하지 않음");
            }
            using (var button = new Bitmap(Path.Combine(assets, "banned-member-management-button.png")))
            using (var firstPage = new Bitmap(400, 494))
            using (var secondPage = new Bitmap(400, 494))
            using (var lastPage = new Bitmap(400, 494))
            {
                using (var g = Graphics.FromImage(firstPage)) { g.Clear(Color.White); g.FillRectangle(Brushes.Gray, 10, 30, 100, 10); }
                using (var g = Graphics.FromImage(secondPage)) { g.Clear(Color.White); g.DrawImageUnscaled(button, 18, 480); }
                using (var g = Graphics.FromImage(lastPage)) { g.Clear(Color.White); g.DrawImageUnscaled(button, 18, 380); }
                Check(!Find(iconCache, "내보낸 상대 관리", secondPage, out found, "FindButton"), "아래로 잘린 버튼은 클릭하지 않고 더 스크롤");
                Check(Find(iconCache, "내보낸 상대 관리", lastPage, out found, "FindButton") && found.Contains(new Point(18 + button.Width / 2, 380 + button.Height / 2)), "사용자가 제공한 전체 버튼 이미지 탐색");
                var pages = new[] { firstPage, secondPage, lastPage }; int page = 0, scrolls = 0, clicks = 0;
                Action guard = () => { };
                Func<bool> select = () => { if (!Find(iconCache, "내보낸 상대 관리", pages[page], out found, "FindButton")) return false; clicks++; return true; };
                Action scroll = () => { page++; scrolls++; };
                Check((bool)Call("ModerationNavigation", "Search", guard, select, scroll, (Func<bool>)(() => true), 20) && scrolls == 2 && clicks == 1,
                    "위쪽 화면 → 잘린 버튼 → 전체 버튼까지 스크롤 후 정확히 한 번 선택");
                scrolls = clicks = 0;
                Check((bool)Call("ModerationNavigation", "Search", guard, select, scroll, (Func<bool>)(() => true), 20) && scrolls == 0 && clicks == 1,
                    "처음부터 보이는 버튼은 스크롤 없이 선택");
                scrolls = 0; int scans = 0;
                Check(!(bool)Call("ModerationNavigation", "Search", guard, (Func<bool>)(() => { scans++; return false; }), (Action)(() => scrolls++), (Func<bool>)(() => false), 20) && scrolls == 1 && scans == 2,
                    "끝에서 화면이 멈추면 마지막 확인 후 종료");
                scrolls = 0;
                Check(!(bool)Call("ModerationNavigation", "Search", guard, (Func<bool>)(() => false), (Action)(() => scrolls++), (Func<bool>)(() => true), 3) && scrolls == 3,
                    "내용이 계속 변해도 최대 스크롤 횟수 제한");
                scrolls = 0; bool cancelled = false;
                try
                {
                    Call("ModerationNavigation", "Search", (Action)(() => { if (scrolls > 0) throw new OperationCanceledException(); }),
                        (Func<bool>)(() => false), (Action)(() => scrolls++), (Func<bool>)(() => true), 20);
                }
                catch (TargetInvocationException error) { cancelled = error.GetBaseException() is OperationCanceledException; }
                Check(cancelled && scrolls == 1, "스크롤 중 봇 중지·권한 변경 시 다음 동작 중단");
                using (var g = Graphics.FromImage(lastPage)) g.DrawImageUnscaled(button, 18, 250);
                Check(!Find(iconCache, "내보낸 상대 관리", lastPage, out found, "FindButton"), "같은 버튼이 두 개면 임의 선택하지 않음");
            }
            using (var source = new Bitmap(240, 120))
            {
                using (var graphics = Graphics.FromImage(source))
                using (var font = new Font("맑은 고딕", 12))
                { graphics.Clear(Color.White); graphics.DrawString("오픈채팅 관리", font, Brushes.Black, 20, 25); }
                Check(!Find(cache, "오픈채팅 관리", source, out found), "검증되지 않은 이미지를 임의로 사용하지 않음");
                var labelBounds = new Rectangle(20, 25, 150, 28);
                Invoke(cache, "Remember", "오픈채팅 관리", source, labelBounds);
                Check(Find(cache, "오픈채팅 관리", source, out found), "검증한 메뉴 이미지를 OCR 없이 검색");
                Check(Find(Cache(directory), "오픈채팅 관리", source, out found), "재시작 후 저장한 이미지 재사용");
                Check(!Find(cache, "내보낸 상대 관리", source, out found), "메뉴별 이미지 구분");
                labelBounds.Inflate(2, 2);
                using (var crop = source.Clone(labelBounds, PixelFormat.Format32bppArgb))
                using (var moved = new Bitmap(380, 200))
                {
                    using (var graphics = Graphics.FromImage(moved)) { graphics.Clear(Color.White); graphics.DrawImageUnscaled(crop, 70, 85); }
                    Check(Find(cache, "오픈채팅 관리", moved, out found) && found.Contains(new Point(80, 95)), "메뉴 위치가 바뀌어도 이미지 탐색");
                    using (var graphics = Graphics.FromImage(moved)) graphics.DrawImageUnscaled(crop, 70, 140);
                    Check(!Find(cache, "오픈채팅 관리", moved, out found), "동일 이미지가 두 곳이면 선택하지 않음");
                }
                using (var blank = new Bitmap(240, 120))
                    Check(!Find(cache, "오픈채팅 관리", blank, out found), "테마·이미지 불일치 시 재인식 경로로 반환");
            }
            Console.WriteLine("PASS: " + passed + " menu checks (no chat or UI input)"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error.GetBaseException()); return 1; }
    }
    [DllImport("user32.dll")] private static extern IntPtr CreatePopupMenu();
    [DllImport("user32.dll", EntryPoint = "AppendMenuW", CharSet = CharSet.Unicode)] private static extern bool AppendMenu(IntPtr menu, uint flags, UIntPtr id, string text);
    [DllImport("user32.dll")] private static extern bool DestroyMenu(IntPtr menu);
}
