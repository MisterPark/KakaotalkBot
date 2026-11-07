using System;
using System.Reflection;
using KakaotalkBot;

class CommandRoutingTests
{
    static readonly MethodInfo Route = typeof(Bot).GetMethod("ShouldQueueCommand", BindingFlags.Static | BindingFlags.NonPublic);
    static bool Accept(string text, params string[] keywords)
    { return (bool)Route.Invoke(null, new object[] { text, keywords }); }
    static void Check(bool value, string name)
    { if (!value) throw new Exception(name); Console.WriteLine("PASS " + name); }
    static int Main()
    {
        try
        {
            // 시트에는 안내 명령만 있고 실제 관리·선거 키워드 행은 없는 상황입니다.
            foreach (string text in new[] { "/내보낸목록", "/내보내기 @사용자 사유", "/내보내기해제 3 사유",
                "/선거", "/대법관", "/후보", "/선거결과", "/출마 공약", "/출마취소",
                "/투표 1 3 5", "/투표취소", "/선거설정 켜기", "/선거설정 끄기" })
                Check(Accept(text, "/?", "/명령어"), "시트 등록 없이 접수: " + text);
            foreach (string text in new[] { "/내보낸목록추가", "/투표자", "오늘 /내보낸목록", "일반 대화", "", null })
                Check(!Accept(text), "명령이 아닌 입력 제외: " + (text ?? "null"));
            Check(Accept("/내보낸목록 잘못된인수"), "잘못된 인수도 처리기로 전달해 사용법 안내");
            foreach (string text in new[] { "/출석", "/퀴즈", "/퀘스트", "/채팅랭킹", "/네임드", "/메모 @사용자 내용" })
                Check(Accept(text), "기존 내장 명령 유지: " + text);
            Check(Accept("/디스코드", "/디스코드"), "시트 키워드 접수 유지");
            Check(Accept("/조회 @사용자", "/조회"), "기존 시트 키워드 인수 처리 유지");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
