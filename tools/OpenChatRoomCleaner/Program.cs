using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace OpenChatRoomCleaner
{
    static class Program
    {
        [STAThread] static int Main(string[] args)
        {
            if(args.Length>0 && args[0]=="--self-test")return SelfTest(args);
            Desktop.SetProcessDPIAware();Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());return 0;
        }
        internal static string Assets {get{return Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Assets");}}
        static int SelfTest(string[] args)
        {
            try
            {
                Assert(MainForm.EligibleWindow(true,false,false,true,true,Stage.OpenList,"프로필"),"시험 모드에서 기존 전경 프로필 검사");
                Assert(MainForm.EligibleWindow(false,false,false,false,true,Stage.Profile,"프로필"),"새 프로필의 제목이 있어도 검사");
                Assert(!MainForm.EligibleWindow(false,false,false,true,true,Stage.Profile,"다른 방"),"실제 실행에서 기존 다른 방 제외");
                Assert(!MainForm.EligibleWindow(false,false,false,false,true,Stage.Scan,"다른 방"),"목록 검사 중 새 일반 창 제외");
                var testRoom=new Window{Handle=new IntPtr(1),Process=10,ClassName="EVA_Window_Dblclk",Visible=true};
                var testList=new Window{Handle=new IntPtr(2),Owner=testRoom.Handle,Process=10,ClassName="EVA_Window_Dblclk",Visible=true,Bounds=new Rectangle(0,0,270,267)};
                var testProfile=new Window{Handle=new IntPtr(3),Owner=testList.Handle,Process=10,ClassName="EVA_Window_Dblclk",Visible=true,Bounds=new Rectangle(0,0,300,600)};
                var testWindows=new List<Window>{testRoom,testList,testProfile};
                Assert(Desktop.IsProfileWindow(testProfile,testRoom,testWindows),"관측한 프로필 소유 관계 인식");
                testProfile.Owner=testRoom.Handle;Assert(!Desktop.IsProfileWindow(testProfile,testRoom,testWindows),"목록 소유 관계 없는 유사 창 제외");
                testProfile.Owner=testList.Handle;testList.Owner=new IntPtr(99);Assert(!Desktop.IsProfileWindow(testProfile,testRoom,testWindows),"다른 방의 프로필 제외");
                testList.Owner=testRoom.Handle;testProfile.Visible=false;Assert(!Desktop.IsProfileWindow(testProfile,testRoom,testWindows),"숨겨진 프로필 제외");
                testProfile.Visible=true;Assert(Desktop.IsProfileWindow(testProfile,testRoom,testWindows),"재사용 창이 다시 표시되어도 인식");

                Assert(Vision.AlignRowTop(58,60)==60,"검색칸 아래 첫 행 위치 유지");
                Assert(Vision.AlignRowTop(58,112)==60,"두 번째 행 배지로 보정해도 검색칸 높이 유지");
                testRoom.Root=testRoom.Handle;testList.Root=testList.Handle;testProfile.Root=testProfile.Handle;
                var childList=new Window{Handle=new IntPtr(4),Root=testList.Handle,Parent=testList.Handle,Process=10,ClassName="EVA_VH_ListControl_Dblclk",Visible=true};
                testWindows.Add(childList);
                Assert(!Desktop.IsCaptureWindow(childList),"내부 리스트 컨트롤 중복 캡처 차단");
                childList.Root=childList.Handle;childList.Bounds=new Rectangle(0,0,270,375);childList.Owner=testRoom.Handle;
                Assert(!Desktop.IsParticipantPopup(childList,testRoom,testWindows),"리스트 클래스는 루트여도 목록 판독 제외");
                childList.Root=testList.Handle;
                Assert(Desktop.IsCaptureWindow(testList),"소유 팝업은 부모가 있어도 캡처 유지");
                Assert(Desktop.IsParticipantPopup(testList,testRoom,testWindows),"270폭 목록 팝업 전체만 행 판독");
                Assert(!Desktop.IsParticipantPopup(testRoom,testRoom,testWindows),"채팅창 위에 겹친 목록 픽셀을 재판독하지 않음");
                Assert(!Desktop.IsParticipantPopup(testProfile,testRoom,testWindows),"프로필은 목록으로 판독하지 않음");
                // 배지 사각형 모서리의 사진/배경이 달라도 원 안의 왕관만으로 인식합니다.
                using(var originalBadge=new Bitmap(Path.Combine(Assets,"owner.png")))
                using(var changed=(Bitmap)originalBadge.Clone())
                {
                    int cx=changed.Width/2,cy=changed.Height/2,rad=Math.Min(changed.Width,changed.Height)/2-1;
                    for(int y=0;y<changed.Height;y++)for(int x=0;x<changed.Width;x++)
                        if((x-cx)*(x-cx)+(y-cy)*(y-cy)>rad*rad)changed.SetPixel(x,y,Color.DarkGray);
                    Rectangle matched;Assert(new Pixels(changed).FindRoundBadge(new Pixels(originalBadge),new Rectangle(Point.Empty,changed.Size),out matched),"프로필 사진/마우스 배경이 바뀌어도 운영진 배지 인식");
                }
                var vision=new Vision(Assets);
                using(var first=new Bitmap(args[1])) using(var list=first.Clone(new Rectangle(320,82,270,319),first.PixelFormat))
                {
                    Assert(vision.Read(list,false).Footer.IsEmpty,"목록 판독 차단 시 겹친 하단 이미지 무시");
                    var o=vision.Read(list);Assert(o.ListValid,"목록 인식");Assert(o.Rows.Count==5,"5명 행 인식");Assert(o.Rows.All(row=>row.Click.X==38),"닉네임이 아닌 프로필 이미지 중앙 클릭");
                    Assert(o.Rows.Count(r=>r.Protected)==3,"배지가 있는 3명 제외");Assert(o.Rows.Count(r=>!r.Protected)==2,"일반 참여자 2명");
                    for(int border=1;border<=4;border++)
                        using(var trimmed=list.Clone(new Rectangle(border,0,list.Width-border,list.Height),list.PixelFormat))
                        {
                            var borderResult=vision.Read(trimmed);
                            Assert(borderResult.ListValid,"목록 좌측 테두리 "+border+"픽셀 차이 보정: "+borderResult.ListFailure);
                            Assert(borderResult.Rows.Count(r=>r.Protected)==3,"테두리 보정 후 운영진 제외 유지");
                        }
                    using(var padded=new Bitmap(list.Width+16,list.Height+16))
                    {
                        using(var g=Graphics.FromImage(padded)){g.Clear(Color.Black);g.DrawImageUnscaled(list,8,8);}
                        var shifted=vision.Read(padded);Assert(shifted.ListValid&&shifted.Rows.Count(r=>r.Protected)==3,"창 여백 변화에도 운영진 제외");
                    }
                }
                using(var first=new Bitmap(args[1]))Assert(!vision.Read(first).ProfileMenu.IsEmpty,"프로필 메뉴 인식");
                using(var menu=new Bitmap(args[2]))Assert(!vision.Read(menu).Remove.IsEmpty,"신고 없는 내보내기 버튼 인식");
                using(var confirm=new Bitmap(args[3]))Assert(!vision.Read(confirm).Confirm.IsEmpty,"경고 문구와 확인 버튼 인식");
                if(args.Length>=6)
                {
                    using(var center=new Bitmap(args[4])){var at=vision.FindProfileAction(center);Assert(!at.IsEmpty && at.X>130&&at.X<165,"버튼 하나인 빨간 프로필 중앙 아이콘");}
                    using(var right=new Bitmap(args[5])){var at=vision.FindProfileAction(right);Assert(!at.IsEmpty && at.X>220&&at.X<255,"버튼 세 개인 파란 프로필 오른쪽 아이콘");}
                }
                using(var gray=new Bitmap(args[1]))using(var profile=gray.Clone(new Rectangle(13,300,300,600),gray.PixelFormat))
                    Assert(!vision.FindProfileAction(profile).IsEmpty,"회색 프로필 관리 아이콘");
                using(var white=new Bitmap(300,600)){using(var g=Graphics.FromImage(white))g.Clear(Color.White);Assert(vision.FindProfileAction(white).IsEmpty,"흰 배경만으로 아이콘 오인식 금지");}
                if(args.Length>=7)
                {
                    using(var largeSource=new Bitmap(args[6]))using(var large=largeSource.Clone(new Rectangle(7,86,270,375),largeSource.PixelFormat))
                    {
                        var observed=vision.Read(large);
                        Assert(observed.ListValid,"검색칸 있는 큰 목록 인식: "+observed.ListFailure);
                        using(var unreadable=(Bitmap)large.Clone())
                        {
                            using(var g=Graphics.FromImage(unreadable))g.FillRectangle(Brushes.White,70,65,175,35);
                            var middle=vision.Read(unreadable);
                            Assert(!middle.ListValid && middle.ListViewportValid,"행 판독 실패와 스크롤 가능한 영역을 분리");
                            var recover=new Workflow();
                            var move=recover.Decide(middle);
                            Assert(move.Kind==ActionKind.ScrollUp && recover.Stage==Stage.SeekTop,"기존 스크롤 유지 시 인식 대기 대신 위로 이동");
                            for(int retry=0;retry<5;retry++){recover.Applied(move);move=recover.Decide(middle);Assert(move.Kind==ActionKind.ScrollUp,"행이 안 읽힌 채 정지해도 유저 클릭/완료 금지");}
                            recover.Applied(move);move=recover.Decide(observed);
                            for(int retry=0;retry<3 && move.Kind==ActionKind.ScrollUp;retry++){recover.Applied(move);move=recover.Decide(observed);}
                            Assert(move.Kind==ActionKind.Click && move.Point==observed.Rows[1].Click,"최상단 안정화 후 첫 일반 참여자 선택");
                            var reopened=new Workflow{Stage=Stage.ReturnToList};Assert(reopened.Decide(middle).Kind==ActionKind.ScrollUp,"내보내기 시도 후 재개방한 목록도 먼저 맨 위로 이동");
                        }
                        Assert(observed.Rows.Count==5,"큰 목록의 완전한 5개 행만 읽기");
                        Assert(observed.Rows[0].Protected,"방장 행 제외");
                        Assert(observed.Rows[1].Click.Y>=125 && observed.Rows[1].Click.Y<=137,"첫 일반 유저 이미지 중앙 위치");
                        Assert(observed.Rows.Count(row=>!row.Protected)==4,"잘린 여섯 번째 행 제외");
                        var largeFlow=new Workflow{Stage=Stage.Scan};
                        var next=largeFlow.Decide(observed);Assert(next.Kind==ActionKind.Click && next.Point==observed.Rows[1].Click,"큰 목록에서 운영진 다음 일반 참여자 선택");
                    }
                }
                if(args.Length>=8)
                {
                    using(var operators=new Bitmap(args[7]))
                    {
                        var roles=vision.Read(operators);
                        Assert(roles.ListValid,"운영진 목록 인식: "+roles.ListFailure);
                        Assert(roles.Rows.Count==5 && roles.Rows.All(row=>row.Protected),"운영진 5명 제외: "+string.Join("; ",roles.Rows.Select(row=>row.Click+" 제외="+row.Protected+" 불확실="+row.Uncertain)));
                        foreach(int offset in new[]{13,24,37})
                        using(var shifted=(Bitmap)operators.Clone())
                        {
                            for(int y=58;y<331;y++)for(int x=12;x<250;x++)shifted.SetPixel(x,y,operators.GetPixel(x,58+(y-58+offset)%260));
                            var scrolled=vision.Read(shifted);
                            Assert(!scrolled.ListValid || scrolled.Rows.All(row=>row.Protected),"스크롤 "+offset+"픽셀 후 운영진 오선택: "+string.Join(";",scrolled.Rows.Select(row=>row.Click+" 제외="+row.Protected)));
                        }
                    }
                }
                if(args.Length>=9)
                {
                    using(var liveScroll=new Bitmap(args[8]))
                    {
                        var rows=vision.Read(liveScroll);
                        Assert(rows.ListValid && rows.Rows.Count>=3,"실제 스크롤 캡처의 완전한 행 인식: "+rows.ListFailure);
                        Assert(rows.Rows.All(row=>row.Protected),"실제 스크롤 화면 운영진 제외 및 하단 잘린 일반 행 제외: "+string.Join(";",rows.Rows.Select(row=>row.Click+" 제외="+row.Protected)));
                        var scan=new Workflow{Stage=Stage.Scan};Assert(scan.Decide(rows).Kind!=ActionKind.Click,"운영진과 잘린 행을 선택하지 않음");
                    }
                }
                if(args.Length>=10)
                {
                    using(var ordinaryBlue=new Bitmap(args[9]))
                    {
                        var rows=vision.Read(ordinaryBlue);
                        Assert(rows.ListValid && rows.Rows.Count>=3,"파란 일반 프로필 목록 판독: "+rows.ListFailure);
                        Assert(rows.Rows[0].Protected,"파란 왕관 방장 제외 유지");
                        Assert(rows.Rows.Skip(1).All(row=>!row.Protected),"일반 유저 전원 유지");
                        Assert(!rows.Rows[2].Protected,"엉엉 우는 어피치의 파란 사진을 운영진 배지로 오인하지 않음");
                        var focused=vision.Read(ordinaryBlue,true,Stage.Scan);
                        Assert(focused.ListValid && focused.Rows.Count==rows.Rows.Count && !focused.Rows[2].Protected,"단계별 검색에서도 동일한 목록 판정");
                        rows.Rows.RemoveAt(1);
                        var pick=new Workflow{Stage=Stage.Scan}.Decide(rows);
                        Assert(pick.Kind==ActionKind.Click && pick.Point==rows.Rows[1].Click,"앞 유저 제거 후 어피치를 건너뛰지 않고 선택");
                    }
                }
                using(var blank=new Bitmap(400,700)) {var o=vision.Read(blank);Assert(!o.ListValid&&o.Confirm.IsEmpty,"빈 화면에서 클릭/완료 금지");}
                var flow=new Workflow();var empty=new Observation();for(int i=0;i<20;i++)Assert(flow.Decide(empty).Kind==ActionKind.Wait,"예상 이미지 없으면 계속 대기");
                flow.Stage=Stage.Scan;
                var uncertainList=new Observation{ListValid=true,ListHash=5};uncertainList.Rows.Add(new Row{Protected=true,Uncertain=true});
                Assert(flow.Decide(uncertainList).Kind==ActionKind.Wait,"의심 배지를 일반 유저로 선택하거나 완료하지 않음");
                var protectedList=new Observation{ListValid=true,ListHash=99};protectedList.Rows.Add(new Row{Protected=true});
                var d=flow.Decide(protectedList);Assert(d.Kind==ActionKind.ScrollDown,"운영진만 보여도 즉시 완료하지 않음");
                for(int i=0;i<3;i++){flow.Applied(d);d=flow.Decide(protectedList);}
                Assert(d.Kind==ActionKind.Complete,"목록 끝 재확인 후 완료");
                flow=new Workflow{Stage=Stage.Scan};var ordinary=new Observation{ListValid=true,ListHash=12};ordinary.Rows.Add(new Row{Protected=false});
                d=flow.Decide(ordinary);Assert(d.Kind==ActionKind.Click,"일반 참여자 선택");flow.Applied(d);Assert(flow.Stage==Stage.Profile,"입력 접수 후에만 다음 단계");
                Assert(flow.Decide(empty).Kind==ActionKind.Wait && flow.Stage==Stage.Profile,"프로필 로딩 중 단계 유지");
                flow=new Workflow(RunMode.Execute);flow.Stage=Stage.Confirm;d=flow.Decide(new Observation{Confirm=new Rectangle(1,1,20,20)});flow.Applied(d);
                Assert(flow.Attempts==1&&flow.Stage==Stage.ReturnToList,"성공 횟수가 아닌 확인 클릭 시도 횟수 집계");
                var preview=new Workflow{Stage=Stage.Confirm};
                d=preview.Decide(new Observation{Confirm=new Rectangle(1,1,20,20)});
                Assert(d.Kind==ActionKind.PreviewComplete,"자동 시험에서 최종 확인 클릭 금지");preview.Applied(d);Assert(preview.Attempts==0,"자동 시험에서 실제 강퇴 시도 0회");
                Assert(!Desktop.Apply(null,new Decision(ActionKind.ConfirmRemoval,"차단 시험"),()=>false),"입력 계층도 기본값으로 최종 확인 차단");
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test-result.txt"),"PASS: 실제 캡처 "+(args.Length-1)+"장과 좌표/상태 전이 시험 통과. 실제 입력 없음.");return 0;
            }
            catch(Exception e){File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test-result.txt"),e.ToString());return 1;}
        }
        static void Assert(bool value,string text){if(!value)throw new Exception(text);}
    }
    sealed class MainForm:Form
    {
        readonly ComboBox rooms=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Width=320};
        readonly ComboBox mode=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Width=310};
        readonly Button start=new Button{Text="시작",Width=80};
        readonly Button stop=new Button{Text="중지 (F8)",Width=100};
        readonly Label status=new Label{AutoSize=false,Dock=DockStyle.Top,Height=46,Text="방 선택 → 인식 시험 → 실제 실행"};
        readonly TextBox log=new TextBox{Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Vertical,Dock=DockStyle.Fill};
        readonly ManualResetEvent halt=new ManualResetEvent(false);
        Thread worker;volatile bool stopping; string lastStatus;DateTime lastDisplay;
        internal MainForm()
        {
            Text="오픈채팅 참여자 정리 · 목록창-v6 / 확인 직전 시험";Width=760;Height=440;Font=new Font("맑은 고딕",10);
            mode.Items.AddRange(new object[]{"인식만 확인 (클릭 없음)","자동 시험 (최종 확인 직전 중지)","실제 내보내기"});mode.SelectedIndex=1;
            var bar=new FlowLayoutPanel{Dock=DockStyle.Top,Height=85};var refresh=new Button{Text="방 새로고침",AutoSize=true};
            bar.Controls.AddRange(new Control[]{rooms,refresh,mode,start,stop});Controls.Add(log);Controls.Add(status);Controls.Add(bar);
            refresh.Click+=(s,e)=>RefreshRooms();start.Click+=(s,e)=>Start();stop.Click+=(s,e)=>Stop();
            FormClosing+=(s,e)=>{if(worker!=null&&worker.IsAlive){Stop();e.Cancel=true;Report("중지 처리 후 창을 다시 닫아 주세요.");}};
            FormClosed+=(s,e)=>halt.Dispose();RefreshRooms();
            Shown+=(s,e)=>Report("실행 파일: "+Application.ExecutablePath);
            var timer=new System.Windows.Forms.Timer{Interval=100};timer.Tick+=(s,e)=>{if((Desktop.GetAsyncKeyState(0x77)&0x8000)!=0)Stop();};timer.Start();FormClosed+=(s,e)=>timer.Dispose();
        }
        void RefreshRooms()
        {
            if(worker!=null&&worker.IsAlive)return;
            rooms.Items.Clear();foreach(var w in Desktop.Windows().Where(w=>!string.IsNullOrWhiteSpace(w.Title)&&w.Title!="카카오톡"))rooms.Items.Add(w);
            for(int i=0;i<rooms.Items.Count;i++)if(((Window)rooms.Items[i]).Title=="보룸봇 테스트방")rooms.SelectedIndex=i;
            if(rooms.SelectedIndex<0&&rooms.Items.Count>0)rooms.SelectedIndex=0;
        }
        void Start()
        {
            if(worker!=null&&worker.IsAlive)return;var room=rooms.SelectedItem as Window;if(room==null){Report("열린 채팅방을 선택하세요.");return;}
            var runMode=(RunMode)mode.SelectedIndex;
            if(runMode==RunMode.Execute && MessageBox.Show(this,room.Title+"의 운영진/본인을 제외한 모든 참여자에게 내보내기를 시도합니다.\n확인창 안내처럼 재참여가 제한됩니다. 계속할까요?","대상 방 확인",MessageBoxButtons.YesNo,MessageBoxIcon.Warning)!=DialogResult.Yes)return;
            stopping=false;halt.Reset();rooms.Enabled=false;mode.Enabled=false;start.Enabled=false;
            var initial=new HashSet<IntPtr>(Desktop.Windows().Select(w=>w.Handle));
            worker=new Thread(()=>Run(room,runMode,initial)){IsBackground=true,Name="참여자 정리 STA"};worker.SetApartmentState(ApartmentState.STA);worker.Start();
        }
        void Stop(){stopping=true;halt.Set();}
        void Run(Window room,RunMode runMode,HashSet<IntPtr> initial)
        {
            try
            {
                bool test=runMode==RunMode.Observe;
                var vision=new Vision(Program.Assets);var flow=new Workflow(runMode);var permitted=new HashSet<IntPtr>{room.Handle};
                // 시작 버튼을 누른 사용자가 지정한 방만 활성화합니다.
                Desktop.SetForegroundWindow(room.Handle);
                while(!stopping)
                {
                    if(halt.WaitOne(100))break;
                    try
                    {
                        var windows=Desktop.Windows(includeChildren:true);
                        if(!windows.Any(w=>w.Handle==room.Handle&&w.Process==room.Process&&w.Title==room.Title)){Report("선택한 방이 닫혔습니다. 중지합니다.");break;}
                        if(!Desktop.IsForegroundProcess(room.Process)){Report("카카오톡 화면으로 돌아오면 이어서 진행합니다.");continue;}
                        IntPtr foreground=Desktop.GetForegroundWindow();
                        foreach(var w in windows.Where(w=>w.Process==room.Process))
                            if(EligibleWindow(false,Desktop.OwnedBy(w.Handle,room.Handle),permitted.Contains(w.Handle),
                                initial.Contains(w.Handle),w.Handle==foreground,flow.Stage,w.Title))permitted.Add(w.Handle);
                        bool acted=false;string waiting="예상 화면을 찾는 중: "+flow.Stage;
                        var findings=new List<string>();bool invalidListSeen=false;int inspected=0;bool foregroundInspected=false;
                        foreach(var w in windows.Where(w=>w.Process==room.Process && Desktop.IsCaptureWindow(w) &&
                            (permitted.Contains(w.Handle)||(test&&w.Handle==foreground)))
                            .OrderBy(w=>w.Handle==foreground?0:w.Handle==room.Handle?2:1))
                        {
                            using(var frame=Desktop.Capture(w.Handle))
                            {
                                if(frame==null)continue;inspected++;if(w.Handle==foreground)foregroundInspected=true;var observation=vision.Read(frame.Image,Desktop.IsParticipantPopup(w,room,windows),test?(Stage?)null:flow.Stage);
                                observation.NativeProfile=Desktop.IsProfileWindow(w,room,windows);
                                if(observation.NativeProfile && (test || flow.Stage==Stage.Profile || flow.Stage==Stage.ReturnToList) && observation.Remove.IsEmpty && observation.Confirm.IsEmpty)
                                    observation.ProfileMenu=vision.FindProfileAction(frame.Image);
                                if(!observation.Footer.IsEmpty && !observation.ListValid)
                                {
                                    invalidListSeen=true;waiting="목록 재인식: "+observation.ListFailure+" ["+w.ClassName+"]";
                                    if(test)findings.Add(waiting);
                                }
                                if(test)
                                {
                                    string found=!observation.Confirm.IsEmpty?"내보내기 확인창":!observation.Remove.IsEmpty?"내보내기 메뉴":observation.NativeProfile?(observation.ProfileMenu.IsEmpty?"프로필 창(Win32) · 관리 아이콘 인식 대기":"프로필 창(Win32) · 관리 아이콘 위치 "+observation.ProfileMenu.X+","+observation.ProfileMenu.Y):!observation.ProfileMenu.IsEmpty?"프로필 관리 버튼(이미지)":observation.ListValid?"참여자 "+observation.Rows.Count+"명 / 제외 "+observation.Rows.Count(r=>r.Protected)+"명":!observation.Members.IsEmpty?"참여자 목록 열기 버튼":null;
                                    if(found!=null && !findings.Contains(found))findings.Add(found);continue;
                                }
                                // 이미 열린 목록을 읽지 못한 경우 방 헤더를 다시 눌러 목록을 닫지 않습니다.
                                if(w.Handle==room.Handle && invalidListSeen)continue;
                                // 현재 단계와 무관한 창에서는 상태를 변경하지 않습니다.
                                if(!Relevant(flow.Stage,observation))continue;
                                var decision=flow.Decide(observation);waiting=decision.Status;
                                if(observation.NativeProfile && flow.Stage==Stage.Profile)decision.Status="Win32 프로필 확인 → 감지한 관리 아이콘 클릭";
                                if(decision.Kind==ActionKind.PreviewComplete){Report(decision.Status);return;}
                                if(decision.Kind==ActionKind.Complete){Report(decision.Status+" · 확인 클릭 시도 "+flow.Attempts+"회");return;}
                                if(decision.Kind==ActionKind.Wait)continue;
                                if(flow.Stage==Stage.Scan && decision.Kind==ActionKind.Click)
                                    SaveSelection(frame.Image,observation,decision.Point);
                                if(Desktop.Apply(frame,decision,()=>stopping,runMode==RunMode.Execute))
                                {flow.Applied(decision);Report(decision.Status+" · 시도 "+flow.Attempts+"회");acted=true;break;}
                                waiting="창 위치/가림 상태 변경 · 재인식 대기";
                            }
                        }
                        if(test)Report("[시험 · 클릭 없음] 검사 "+inspected+"개 창 · "+
                            (findings.Count>0?string.Join(" / ",findings):"기준 이미지 일치 없음")+
                            (foregroundInspected?" · 현재 창 검사 완료":" · 현재 창 검사 대상 없음"));
                        else if(!acted)Report(waiting);
                    }
                    catch(Exception ex){Report("재시도 대기: "+ex.Message);}
                }
                Report("중지됨");
            }
            catch(Exception ex){Report("시작 실패: "+ex.Message);}
            finally {if(!IsDisposed)BeginInvoke((Action)(()=>{rooms.Enabled=true;mode.Enabled=true;start.Enabled=true;}));}
        }
        // 시험 모드는 현재 카카오톡 창을 읽기만 합니다. 실제 입력은 선택 방의 소유 창과
        // 예상 단계에서 새로 나타난 전경 팝업으로 한정합니다. 제목 유무로 프로필을 배제하지 않습니다.
        internal static bool EligibleWindow(bool test,bool owned,bool known,bool existed,bool foreground,Stage stage,string title)
        {
            if(owned||known)return true;
            if(test&&foreground)return true;
            bool popupExpected=stage==Stage.Profile||stage==Stage.RemoveMenu||stage==Stage.Confirm;
            return !existed && (string.IsNullOrEmpty(title)||(foreground&&popupExpected));
        }
        static bool Relevant(Stage s,Observation o)
        {
            if(s==Stage.Profile)return !o.ProfileMenu.IsEmpty;
            if(s==Stage.RemoveMenu)return !o.Remove.IsEmpty;
            if(s==Stage.Confirm)return !o.Confirm.IsEmpty;
            if(s==Stage.SeekTop)return o.ListViewportValid||o.ListValid;
            if(s==Stage.Scan)return o.ListValid;
            return o.ListViewportValid||o.ListValid||!o.Members.IsEmpty||!o.ProfileMenu.IsEmpty||!o.Remove.IsEmpty;
        }
        void SaveSelection(Bitmap frame,Observation observation,Point selected)
        {
            try
            {
                string directory=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Data");Directory.CreateDirectory(directory);
                using(var marked=(Bitmap)frame.Clone())using(var g=Graphics.FromImage(marked))using(var font=new Font("맑은 고딕",9,FontStyle.Bold))
                {
                    foreach(var row in observation.Rows)
                    {
                        Color color=row.Uncertain?Color.DarkOrange:row.Protected?Color.Green:Color.Blue;
                        using(var pen=new Pen(color,2))g.DrawRectangle(pen,3,row.Click.Y-20,marked.Width-7,49);
                        using(var brush=new SolidBrush(color))g.DrawString(row.Uncertain?"불확실·제외":row.Protected?"운영진/본인 제외":"일반 판정",font,brush,100,row.Click.Y+10);
                    }
                    using(var pen=new Pen(Color.Red,3)){g.DrawEllipse(pen,selected.X-10,selected.Y-10,20,20);g.DrawLine(pen,selected.X-14,selected.Y,selected.X+14,selected.Y);}
                    marked.Save(Path.Combine(directory,"last-selection.png"),System.Drawing.Imaging.ImageFormat.Png);
                }
                File.WriteAllText(Path.Combine(directory,"last-selection.txt"),"목록창-v6 "+DateTime.Now.ToString("O")+Environment.NewLine+
                    "선택="+selected+Environment.NewLine+string.Join(Environment.NewLine,observation.Rows.Select(row=>row.Click+" 제외="+row.Protected+" 불확실="+row.Uncertain)));
            }
            catch(Exception error){if(!(error is IOException || error is UnauthorizedAccessException || error is System.Runtime.InteropServices.ExternalException))throw;Report("선택 진단 저장 실패: "+error.Message);}
        }
        void Report(string text)
        {
            text="[목록창-v6] "+text;
            if(IsDisposed||!IsHandleCreated)return;
            BeginInvoke((Action)(()=>{
                if(text==lastStatus && (DateTime.Now-lastDisplay).TotalSeconds<5)return;
                lastStatus=text;lastDisplay=DateTime.Now;status.Text=text;
                if(log.TextLength>40000)log.Clear();log.AppendText(DateTime.Now.ToString("HH:mm:ss")+" "+text+Environment.NewLine);
            }));
        }
    }
}
