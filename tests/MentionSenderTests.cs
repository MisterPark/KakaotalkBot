using System;
using System.Collections.Generic;
using KakaotalkBot;

internal static class MentionSenderTests
{
    private static int passed;
    private const long Target = 8172472438151128507L;
    private static int Main()
    {
        try
        {
            Run("정확한 64비트 ID·한글·여러 줄 본문·1회 전송", () =>
            {
                var input = new FakeInput(); var sender = Sender(input);
                var ready = sender.Prepare(Target, "동일닉", "첫 줄\r\n둘째 줄\n셋째 줄");
                Check(input.Search == "@동일닉", "검색어");
                Check(ready.Text == "\uFFFC 첫 줄\r둘째 줄\r셋째 줄\r", "본문");
                Check(ready.Mentions[0].UserId == Target && ready.Mentions[0].DisplayText == "@동일닉", "ID/표시 이름");
                Check(input.Retargets == 1 && input.Sends == 0, "준비 중 전송 금지");
                sender.SendPrepared(ready);
                Throws(() => sender.SendPrepared(ready));
                Check(input.Sends == 1, "동일 준비 결과 재전송 금지");
            });
            Run("이미 같은 ID인 후보는 메모리 변경 생략", () =>
            {
                var input = new FakeInput { CandidateId = Target }; var sender = Sender(input);
                var ready = sender.Prepare(Target, "동일닉", "본문");
                Check(input.Retargets == 0, "불필요한 쓰기");
                sender.DiscardPrepared(ready);
                Check(input.Read().IsEmpty && input.Sends == 0, "준비 취소");
            });
            Run("사용자 초안 보존", () =>
            {
                foreach (string draft in new[] { "작성 중\r", "메시지 입력\r" })
                {
                    var input = new FakeInput(); input.Current = Snapshot(draft);
                    Throws(() => Sender(input).Prepare(Target, "닉", "본문"));
                    Check(input.Begins == 0 && input.Clears == 0 && input.Current.Text == draft, "초안 변경");
                }
            });
            Run("안내 문구가 표시된 빈 입력에서 멘션 준비", () =>
            {
                var input = new FakeInput();
                input.Current = Snapshot("메시지 입력\r");
                input.Current.IsPlaceholder = KakaoInputMentions.IsEmptyPlaceholder(input.Current.Text, 0x949494, false);
                var sender = Sender(input);
                var ready = sender.Prepare(Target, "동일닉", "본문");
                Check(input.Begins == 1 && ready.Mentions[0].UserId == Target, "안내 문구를 초안으로 오인");
                sender.DiscardPrepared(ready);
            });
            Run("같은 문구를 직접 입력했거나 서식이 다르면 보존", () =>
            {
                foreach (int color in new[] { 0, 0x949494, -9999999 })
                foreach (bool undo in new[] { false, true })
                {
                    if (color == 0x949494 && !undo) continue;
                    var input = new FakeInput();
                    input.Current = Snapshot("메시지 입력\r");
                    input.Current.IsPlaceholder = KakaoInputMentions.IsEmptyPlaceholder(input.Current.Text, color, undo);
                    Throws(() => Sender(input).Prepare(Target, "닉", "본문"));
                    Check(input.Begins == 0 && input.Clears == 0, "동일 문구의 초안 변경");
                }
                Check(!KakaoInputMentions.IsEmptyPlaceholder("메시지 입력 \r", 0x949494, false), "공백이 포함된 초안 오인");
            });
            Run("후보가 없으면 작성한 검색어만 정리", () =>
            {
                var input = new FakeInput { NoCandidate = true };
                Throws(() => Sender(input).Send(Target, "없는닉", "본문"));
                Check(input.Sends == 0 && input.Clears == 1 && input.Read().IsEmpty, "후보 실패 처리");
            });
            Run("전체 멘션 후보 거부", () =>
            {
                var input = new FakeInput { Display = "@all" };
                Throws(() => Sender(input).Send(Target, "all", "본문"));
                Check(input.Sends == 0 && input.Retargets == 0 && input.Read().IsEmpty, "전체 멘션 오용");
            });
            Run("선택 도중 사용자가 쓴 내용 보존", () =>
            {
                var input = new FakeInput { EditDuringSelection = true };
                Throws(() => Sender(input).Send(Target, "닉", "본문"));
                Check(input.Clears == 0 && input.Sends == 0 && input.Current.Text == "사용자 수정\r", "동시 입력 손실");
            });
            Run("메모리 쓰기 실패 때 보유한 초안만 정리", () =>
            {
                var input = new FakeInput { RetargetFails = true };
                Throws(() => Sender(input).Send(Target, "닉", "본문"));
                Check(input.Sends == 0 && input.Clears == 1 && input.Read().IsEmpty, "쓰기 실패 처리");
            });
            Run("변경 후 잘못된 ID이면 전송 금지", () =>
            {
                var input = new FakeInput { WrongTarget = true };
                Throws(() => Sender(input).Send(Target, "닉", "본문"));
                Check(input.Sends == 0, "잘못된 ID 전송");
            });
            Run("준비 이후 수정된 입력의 전송·정리 금지", () =>
            {
                var input = new FakeInput(); var sender = Sender(input);
                var ready = sender.Prepare(Target, "닉", "본문");
                input.Current = Snapshot("사용자 수정\r");
                Throws(() => sender.SendPrepared(ready));
                Throws(() => sender.DiscardPrepared(ready));
                Check(input.Sends == 0 && input.Current.Text == "사용자 수정\r", "변경한 초안 보존");
            });
            Run("전송 결과 불명확할 때 재전송·자동 정리 금지", () =>
            {
                var input = new FakeInput { Acknowledge = false }; var sender = Sender(input);
                var ready = sender.Prepare(Target, "닉", "본문");
                Throws(() => sender.SendPrepared(ready));
                Throws(() => sender.SendPrepared(ready));
                Check(input.Sends == 1 && input.Clears == 0, "불확실한 전송 중복");
            });
            Run("전송 중 예외가 나도 재시도하지 않음", () =>
            {
                var input = new FakeInput { SendFails = true }; var sender = Sender(input);
                var ready = sender.Prepare(Target, "닉", "본문");
                Throws(() => sender.SendPrepared(ready));
                Throws(() => sender.SendPrepared(ready));
                Check(input.Sends == 1 && input.Clears == 0, "전송 예외 중복");
            });
            Run("입력값 검증은 편집 전에 수행", () =>
            {
                var input = new FakeInput(); var sender = Sender(input);
                Throws(() => sender.Prepare(0, "닉", "본문"));
                Throws(() => sender.Prepare(Target, "", "본문"));
                Throws(() => sender.Prepare(Target, "닉", ""));
                Throws(() => sender.Prepare(Target, "닉", "\uFFFC"));
                Throws(() => sender.Prepare(Target, "닉", new string('가', 3501)));
                Check(input.Begins == 0, "검증 전 입력 발생");
            });
            Run("다른 송신기의 준비 결과 사용 금지", () =>
            {
                var input = new FakeInput(); var sender = Sender(input);
                var ready = sender.Prepare(Target, "닉", "본문");
                Throws(() => Sender(input).SendPrepared(ready));
                Throws(() => sender.Prepare(Target, "닉", "다른 본문"));
                Check(input.Sends == 0, "소유권 없는 전송");
                sender.DiscardPrepared(ready);
            });
            Run("운영진 알림 본문과 여러 줄 검증 후 1회 전송", () =>
            {
                var input = new FakeInput();
                KakaoPlainSender.Send(input, "[운영진]\r\n내용", delay => { });
                Check(input.Sends == 1 && input.Retargets == 0 && input.Search == "[", "일반 알림 송신");
            });
            Run("운영진 알림은 초안을 덮어쓰지 않음", () =>
            {
                var input = new FakeInput(); input.Current = Snapshot("작성 중\r");
                Throws(() => KakaoPlainSender.Send(input, "[알림] 내용", delay => { }));
                Check(input.Sends == 0 && input.Begins == 0, "초안 보존");
            });
            Run("운영진 알림 결과 불명확 시 한 번만 입력", () =>
            {
                var input = new FakeInput { Acknowledge = false };
                Throws(() => KakaoPlainSender.Send(input, "[알림] 내용", delay => { }));
                Check(input.Sends == 1 && input.Clears == 0, "재전송·자동 정리 금지");
            });
            Run("첫 후보가 객체를 만들지 않으면 다른 검색어로 대상 ID 준비", () =>
            {
                var input = new FakeInput { BlockFirstSearch = true };
                var sender = Sender(input);
                var ready = sender.Prepare(Target, "봇", "본문", new[] { "봇", "일반유저" });
                Check(input.Begins == 2 && input.Clears == 1 && input.Search == "@일반유저", "대체 검색");
                Check(ready.Mentions[0].UserId == Target && input.Sends == 0, "최종 대상 검증");
                sender.SendPrepared(ready);
                Check(input.Sends == 1, "단일 전송");
            });
            Run("모든 대체 후보 실패 시 입력 정리와 전송 금지", () =>
            {
                var input = new FakeInput { NoCandidate = true };
                Throws(() => Sender(input).Prepare(Target, "봇", "본문", new[] { "유저1", "유저2" }));
                Check(input.Begins == 3 && input.Read().IsEmpty && input.Sends == 0, "실패 시 정리");
            });
            Run("사용자 입력 변화 시 대체 검색하지 않음", () =>
            {
                var input = new FakeInput { EditDuringSelection = true };
                Throws(() => Sender(input).Prepare(Target, "봇", "본문", new[] { "유저" }));
                Check(input.Begins == 1 && input.Clears == 0 && input.Sends == 0, "동시 입력 보존");
            });
            Console.WriteLine("PASS: " + passed + " cases (no native input or messages sent)");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
    private static KakaoMentionSender Sender(FakeInput input) { return new KakaoMentionSender(input, milliseconds => { }); }
    private static void Run(string name, Action test) { test(); passed++; Console.WriteLine("PASS: " + name); }
    private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
    private static void Throws(Action action)
    {
        try { action(); } catch (ArgumentException) { return; } catch (InvalidOperationException) { return; }
        throw new Exception("예상 오류가 발생하지 않았습니다.");
    }
    private static InputMentionSnapshot Snapshot(string text, long id = 0, string display = "@동일닉")
    {
        return new InputMentionSnapshot { RoomTitle = "시험", Window = (IntPtr)123, ProcessId = 1, ProcessStart = 2, Text = text,
            Mentions = id == 0 ? new List<InputMention>() : new List<InputMention> { new InputMention { Position = 0, UserId = id, DisplayText = display, Address = 1000, Site = 2000 } } };
    }
    private sealed class FakeInput : IMentionInput
    {
        public InputMentionSnapshot Current = Snapshot("\r");
        public bool BlockFirstSearch, NoCandidate, EditDuringSelection, RetargetFails, WrongTarget, SendFails;
        public bool Acknowledge = true;
        public int Begins, Retargets, Sends, Clears;
        public long CandidateId = 6593948531989662102L;
        public string Display = "@동일닉", Search;
        public InputMentionSnapshot Read() { return Current; }
        private void Require(InputMentionSnapshot expected) { if (!KakaoInputMentions.SameSnapshot(Current, expected)) throw new InvalidOperationException("입력 변경"); }
        public void BeginMention(InputMentionSnapshot empty, string search) { Require(empty); Begins++; Search = search; Current = Snapshot(search + "\r"); }
        public void ChooseCandidate(InputMentionSnapshot seed)
        {
            Require(seed);
            if (EditDuringSelection) Current = Snapshot("사용자 수정\r");
            else if (!NoCandidate && !(BlockFirstSearch && Begins == 1)) Current = Snapshot("\uFFFC \r", CandidateId, Display);
        }
        public InputMentionSnapshot Append(InputMentionSnapshot owned, string text)
        { Require(owned); return Current = Snapshot(owned.Text.TrimEnd('\r') + text + "\r", owned.Mentions.Count == 0 ? 0 : owned.Mentions[0].UserId, Display); }
        public InputMentionSnapshot Retarget(InputMentionSnapshot owned, long id)
        {
            Require(owned); Retargets++;
            if (RetargetFails) throw new InvalidOperationException("쓰기 실패");
            return Current = Snapshot(owned.Text, WrongTarget ? 99 : id, Display);
        }
        public void SendOnce(InputMentionSnapshot ready)
        {
            Require(ready); Sends++;
            if (SendFails) throw new InvalidOperationException("키 전송 결과 불명");
            if (Acknowledge) Current = Snapshot("\r");
        }
        public void Clear(InputMentionSnapshot owned) { Require(owned); Clears++; Current = Snapshot("\r"); }
    }
}
