using System;
using System.Linq;
using KakaotalkBot;

internal static class ChatResponseTests
{
    private static int passed;
    private static void Check(bool value, string label)
    { if (!value) throw new Exception(label); passed++; }
    private static int Main()
    {
        try
        {
            foreach (string word in new[] { "씨발", "개새끼", "병신", "존나", "ㅅㅂ", "씨 발", "씨.발", "씨\u200B발", "병\t신", "FUCK", "f.u.c.k", "ｆｕｃｋ", "ShIt", "asshole" })
            {
                string result = ChatProfanityFilter.Mask(word);
                Check(result == new string('*', word.Length), "비속어 또는 우회 표현 누락");
                Check(ChatProfanityFilter.Mask(result) == result, "중복 필터 시 결과 변경");
            }
            string ordinary = "시발점 시발역 시발지 시발차 시바견 새끼손가락 새끼줄 assistant class assignment title Scunthorpe 퀴즈삭제 123 😀\r\n다음 줄";
            Check(ChatProfanityFilter.Mask(ordinary) == ordinary, "정상 문구 오탐 또는 형식 손실");
            Check(ChatProfanityFilter.Mask("[씨발]\n정상\n개새끼님\n씨\n발") == "[**]\n정상\n***님\n씨\n발", "제목/닉네임 필터 또는 줄 경계 오류");
            Check(ChatProfanityFilter.Mask(null) == null && ChatProfanityFilter.Mask("") == "", "빈 응답 처리");
            string fold = new string('\u200B', 500);
            Check(ChatProfanityFilter.Mask("[제목]\n" + fold + "\n씨발") == "[제목]\n" + fold + "\n**", "접기 문자열 손실");
            string large = string.Concat(Enumerable.Repeat("안녕하세요. 퀴즈 고유번호 123 정상 안내\n", 300));
            Check(ChatProfanityFilter.Mask(large) == large, "긴 정상 응답 손실");

            string old = "[명령어 목록]\n" + fold + "\n[사용자 안내]\n/사용자명령 — 직접 관리하는 설명\n\n" +
                "[퀴즈 · 누구나 참여]\n/퀴즈목록 — 가능한 분류\n\n" +
                "[문제 관리 · 방장]\n/퀴즈삭제 분류 | 문제 전체 문장\n예: /퀴즈삭제 과학 | 물의 화학식은?\n\n" +
                "[대법관 선거]\n오래된 선거 안내\n\n[선거 설정 · 방장]\n/선거설정 켜기\n\n" +
                "[다른 안내]\n유지할 내용";
            string help = CommandHelp.Compose(old);
            foreach (string command in new[] { "/선거시작", "/보궐선거", "/선거기록", "/선거기록 2", "/퀴즈삭제 고유번호", "예: /퀴즈삭제 123", "고유번호를 자동 발급", "전체 퀴즈", "[문제 관리 · 방장]", "[선거 설정 · 방장]" })
                Check(help.Contains(command), "새 명령 사용법 누락");
            Check(!help.Contains("가능한 분류") && !help.Contains("/퀴즈삭제 과학 |") && !help.Contains("오래된 선거 안내"), "구형 안내 잔존");
            Check(help.Contains("/사용자명령 — 직접 관리하는 설명") && help.Contains("[다른 안내]\n유지할 내용"), "별도 시트 안내 손실");
            Check(CommandHelp.Compose(help) == help, "반복 조회 시 섹션 중복");
            Check(CommandHelp.Compose(null).StartsWith("[명령어 목록]\n") && CommandHelp.Compose("").Contains("/보궐선거"), "시트 안내 없는 경우 누락");
            Check(!CommandHelp.Compose("[목록]\n예: /퀴즈삭제 과학 | 문제\n/퀴즈목록 — 분류").Contains("/퀴즈삭제 과학 |"), "다른 섹션의 구형 명령 예시 잔존");
            Console.WriteLine("PASS: " + passed + " chat response checks (no network or UI input)");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
