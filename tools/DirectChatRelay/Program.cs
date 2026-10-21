using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using KakaotalkBot;
namespace DirectChatRelay
{
    static class Program
    {
        [STAThread] static int Main(string[] args)
        {
            if(args.Contains("--self-test"))return RelayRules.Test();
            Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);Application.Run(new RelayForm());return 0;
        }
    }
    static class RelayRules
    {
        internal static bool Accept(ChatRoomInfo room,ChatMessage m)
        { return m.Roster==null && m.ChatId==room.ChatId && m.AuthorId==room.DirectUserId && m.AuthorId>0 && !m.IsOwn && m.RoomEvent==ChatRoomEvent.None && !m.RoleChanged && !string.IsNullOrWhiteSpace(m.Message); }
        internal static IEnumerable<string> Parts(string text)
        { for(int i=0;i<text.Length;){int size=Math.Min(2000,text.Length-i);if(i+size<text.Length&&char.IsHighSurrogate(text[i+size-1]))size--;yield return text.Substring(i,size);i+=size;} }
        internal static int Test()
        {
            var room=new ChatRoomInfo{ChatId=10,DirectUserId=20};
            var m=new ChatMessage{ChatId=10,AuthorId=20,Message="텍스트"};
            if(!Accept(room,m))return 1;m.IsOwn=true;if(Accept(room,m))return 2;m.IsOwn=false;m.AuthorId=30;if(Accept(room,m))return 3;
            m.AuthorId=20;m.ChatId=11;if(Accept(room,m))return 4;m.ChatId=10;m.Roster=new ChatRosterSnapshot();if(Accept(room,m))return 5;
            var users=new ChatUserDirectory();users.OwnIds.Add(30);
            object[] row={1L,20L,2,1L,"사진",0L,0L,null};
            if(ChatMessageDecoder.Decode(room,row,users,false).Count!=0)return 7;
            row[2]=1;row[4]="메시지";
            if(!Accept(room,ChatMessageDecoder.Decode(room,row,users,false).Single()))return 8;
            row[1]=30L;row[6]=1L;
            if(Accept(room,ChatMessageDecoder.Decode(room,row,users,false).Single()))return 9;
            string text=new string('가',1999)+"😀"+new string('나',3000);if(string.Concat(Parts(text))!=text||Parts(text).Any(p=>p.Length>2000||char.IsHighSurrogate(p[p.Length-1])))return 6;
            File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test-result.txt"),"PASS: 상대방/방/본인/시스템 필터 및 긴 텍스트 분할. 실제 전송 없음.");return 0;
        }
    }
    sealed class RoomItem
    {
        internal ChatRoomInfo Room;
        public override string ToString(){return Room.Name+" ["+Room.ChatId+"]";}
    }
    sealed class RelayForm:Form
    {
        readonly ComboBox source=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Width=620};
        readonly ComboBox destination=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Width=620};
        readonly ComboBox window=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Width=620};
        readonly Button refresh=new Button{Text="DB 방 목록 읽기",AutoSize=true}, windows=new Button{Text="열린 전송 창 새로고침",AutoSize=true}, start=new Button{Text="전달 시작"},stop=new Button{Text="중지"};
        readonly TextBox log=new TextBox{Multiline=true,ReadOnly=true,Dock=DockStyle.Fill,ScrollBars=ScrollBars.Vertical};
        readonly Label status=new Label{AutoSize=true,Text="상대방의 새 텍스트만 전달합니다."};
        readonly CancellationTokenSource lifetime=new CancellationTokenSource();
        CancellationTokenSource run;ChatDatabaseReceiver receiver;Task worker;bool loading;int sent;
        internal RelayForm()
        {
            Text="1대1 → 오픈채팅 전달";Width=700;Height=560;Font=new Font("맑은 고딕",10);
            var panel=new FlowLayoutPanel{Dock=DockStyle.Top,Height=290,FlowDirection=FlowDirection.TopDown,WrapContents=false,Padding=new Padding(8)};
            panel.Controls.Add(refresh);panel.Controls.Add(new Label{Text="받을 1대1 채팅방",AutoSize=true});panel.Controls.Add(source);
            panel.Controls.Add(new Label{Text="보낼 오픈채팅방",AutoSize=true});panel.Controls.Add(destination);panel.Controls.Add(window);
            var buttons=new FlowLayoutPanel{Width=650,Height=40};buttons.Controls.AddRange(new Control[]{windows,start,stop});panel.Controls.Add(buttons);panel.Controls.Add(status);
            Controls.Add(log);Controls.Add(panel);refresh.Click+=async(s,e)=>await LoadRooms();windows.Click+=(s,e)=>LoadWindows();start.Click+=(s,e)=>Start();stop.Click+=(s,e)=>Stop();
            destination.SelectedIndexChanged+=(s,e)=>LoadWindows();
            var timer=new System.Windows.Forms.Timer{Interval=500};timer.Tick+=(s,e)=>{if(receiver!=null)status.Text=receiver.Status+" · 전송 입력 "+sent+"회";};timer.Start();
            FormClosing+=(s,e)=>{if(loading || worker!=null&&!worker.IsCompleted){Stop();lifetime.Cancel();e.Cancel=true;Report("종료 처리 중입니다. 잠시 후 다시 닫아 주세요.");}};
            FormClosed+=(s,e)=>{timer.Dispose();lifetime.Cancel();lifetime.Dispose();};
        }
        void Report(string text){if(IsDisposed||!IsHandleCreated)return;BeginInvoke((Action)(()=>{if(log.TextLength>30000)log.Clear();log.AppendText(DateTime.Now.ToString("HH:mm:ss")+" "+text+Environment.NewLine);}));}
        void Enable(bool enabled){source.Enabled=destination.Enabled=window.Enabled=refresh.Enabled=windows.Enabled=start.Enabled=enabled;}
        async Task LoadRooms()
        {
            loading=true;Enable(false);Report("카카오톡 DB 목록을 읽는 중");
            try
            {
                var catalog=await Task.Run(()=>ChatCatalog.Load(lifetime.Token));
                source.Items.Clear();destination.Items.Clear();
                foreach(var room in catalog.Rooms.OrderBy(r=>r.Name))
                {if(room.IsDirect&&room.DirectUserId>0)source.Items.Add(new RoomItem{Room=room});else if(room.IsOpenGroup)destination.Items.Add(new RoomItem{Room=room});}
                if(source.Items.Count>0)source.SelectedIndex=0;if(destination.Items.Count>0)destination.SelectedIndex=0;
                Report("1대1 "+source.Items.Count+"개 / 오픈채팅 "+destination.Items.Count+"개. "+string.Join(" / ",catalog.Errors));
            }catch(Exception ex){Report(ex.Message);}finally{loading=false;Enable(true);}
        }
        void LoadWindows()
        {
            window.Items.Clear();var room=destination.SelectedItem as RoomItem;if(room==null)return;
            foreach(var item in Sender.Windows().Where(w=>w.Pid==room.Room.ProcessId&&w.Title==room.Room.Name))window.Items.Add(item);
            if(window.Items.Count==1)window.SelectedIndex=0;
        }
        void Start()
        {
            var from=source.SelectedItem as RoomItem;var to=destination.SelectedItem as RoomItem;var target=window.SelectedItem as Target;
            if(from==null||to==null||target==null){Report("수신방·전송방을 선택하고 카카오톡에서 전송방을 별도 창으로 열어 주세요.");return;}
            if(from.Room.AccountPath!=to.Room.AccountPath){Report("같은 카카오톡 계정의 방을 선택하세요.");return;}
            if(Sender.Windows().Count(w=>w.Pid==target.Pid&&w.Title==target.Title)!=1){Report("같은 이름의 창이 여러 개입니다. 전송방 하나만 열어 주세요.");return;}
            run=new CancellationTokenSource();Enable(false);sent=0;
            receiver=new ChatDatabaseReceiver(from.Room);receiver.Start();
            var done=new TaskCompletionSource<bool>();worker=done.Task;
            var thread=new Thread(()=>
            {
                try
                {
                    string pending=null;bool pasted=false;var parts=new Queue<string>();string lastError=null;
                    while(!run.IsCancellationRequested)
                    {
                        if(pending==null && parts.Count>0)pending=parts.Dequeue();
                        if(pending==null)
                        {
                            ChatMessage message;
                            if(receiver.TryRead(out message))
                            {
                                if(!RelayRules.Accept(from.Room,message))continue;
                                foreach(var part in RelayRules.Parts(message.Message))parts.Enqueue(part);continue;
                            }
                            if(receiver.Completion.IsCompleted){Report("수신기가 종료됐습니다. 목록을 다시 읽고 시작해 주세요.");break;}
                        }
                        else
                        {
                            string result;
                            if(Sender.TrySend(target,pending,()=>run.IsCancellationRequested,ref pasted,out result))
                            {Interlocked.Increment(ref sent);pending=null;lastError=null;Report(result);}
                            else if(lastError!=result){Report(result);lastError=result;}
                        }
                        if(run.Token.WaitHandle.WaitOne(200))break;
                    }
                }
                catch(Exception ex){Report("전달 중지: "+ex.Message+" 재시작 전에 전송방의 입력 내용을 확인하세요.");}
                finally
                {
                    receiver.Dispose();receiver.Completion.ContinueWith(t=>
                    {
                        done.TrySetResult(true);
                        if(!IsDisposed)BeginInvoke((Action)(()=>{receiver=null;run.Dispose();run=null;Enable(true);status.Text="중지됨";}));
                    });
                    Report("중지: 미전송 대기는 폐기합니다. 다시 시작하면 그 이후 새 메시지부터 받습니다.");
                }
            }){IsBackground=true,Name="채팅 전달 STA"};thread.SetApartmentState(ApartmentState.STA);thread.Start();
            Report("수신 준비 후 새 메시지부터 전달합니다. 전송 중에는 클립보드와 전송방 입력창을 사용합니다.");
        }
        void Stop(){if(run!=null)run.Cancel();if(receiver!=null)receiver.Dispose();}
    }
}
