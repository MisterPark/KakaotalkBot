using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace KakaotalkBot
{
    internal static class CommandHelp
    {
        private static readonly string[] Sections = {
            "[퀴즈 · 누구나 참여]\n" +
            "/퀴즈 또는 /상식퀴즈 — 전체 문제 풀기\n" +
            "/분류명퀴즈 — 해당 분류 문제 풀기 (예: /과학퀴즈, /경제퀴즈)\n" +
            "/퀴즈목록 — 전체 퀴즈의 고유번호·분류·난이도·문제 확인\n" +
            "자동출제 모드의 퀴즈 명령은 현재 문제 또는 다음 출제까지 남은 시간을 표시합니다.",

            "[문제 관리 · 방장]\n" +
            "/퀴즈등록 분류 | 난이도 | 문제 | 정답 | 힌트 | 해설\n" +
            "예: /퀴즈등록 과학 | 하 | 물의 화학식은? | H2O | 알파벳과 숫자 | 수소 두 개와 산소 한 개\n" +
            "등록 시 고유번호를 자동 발급하며, 삭제한 번호는 다시 사용하지 않습니다.\n" +
            "/퀴즈삭제 고유번호 — /퀴즈목록에서 확인한 번호로 삭제\n" +
            "예: /퀴즈삭제 123",

            "[대법관 선거]\n" +
            "/선거 — 상태·마감·다음 일정\n/후보 — 후보 번호·공약\n/대법관 — 현재 대법관 목록\n" +
            "/선거결과 — 최근 선거의 후보별 득표·당선 여부\n" +
            "/선거기록 — 전체 회차의 후보별 득표·당선 기록\n/선거기록 2 — 제2회 선거 기록\n" +
            "/출마 공약 — 후보 등록·공약 수정 (300자 이내)\n/출마취소 — 출마 철회\n" +
            "/투표 1 3 5 — 후보 번호를 공백으로 구분\n/투표취소 — 내 투표 철회\n" +
            "다시 투표하면 이전 선택을 교체합니다. 투표 내용은 채팅에 공개됩니다.\n" +
            "시작 당시 참여자만 출마·투표할 수 있습니다.\n" +
            "정기선거: 홀수달 1일 시작, 3일간 투표 (한국 시간), 정원 10명\n" +
            "투표 기간에는 1시간마다 참여 안내를 보냅니다. (보궐·결선 포함)\n" +
            "선거인 10% 이상 득표한 상위 후보 선출, 당선 경계 동률이면 24시간 결선\n" +
            "선택 수: 본선은 선출 인원, 결선은 남은 자리 수까지 (보궐은 공석 수 기준)\n" +
            "회차 번호는 정기·보궐선거를 포함한 개최 순서입니다.",

            "[선거 설정 · 방장]\n" +
            "/선거설정 켜기 — 선거 기능 활성화\n/선거설정 끄기 — 선거 기능 비활성화\n" +
            "/선거시작 — 놓친 현재 정기선거를 지금부터 3일간 진행\n" +
            "/보궐선거 — 기존 대법관을 유지하며 10명 중 빈자리만 3일간 선출\n" +
            "선거 기능이 켜져 있어야 하며, 진행 중인 선거가 있으면 새로 열 수 없습니다.\n" +
            "수동 개시 후에도 다음 정기선거 일정은 유지됩니다. 보궐 당선자는 잔여 임기를 맡습니다."
        };
        private static int SectionIndex(string heading)
        {
            if (heading == "[퀴즈 · 누구나 참여]") return 0;
            if (heading.StartsWith("[문제 관리 ·", StringComparison.Ordinal)) return 1;
            if (heading == "[대법관 선거]") return 2;
            if (heading.StartsWith("[선거 설정 ·", StringComparison.Ordinal)) return 3;
            return -1;
        }
        private static readonly Regex OwnedCommand = new Regex(@"(?<!\S)/(?:퀴즈목록|퀴즈등록|퀴즈삭제|선거|후보|대법관|선거결과|선거기록|출마|출마취소|투표|투표취소|선거설정|선거시작|보궐선거)(?=\s|$)");

        internal static string Compose(string sheetText)
        {
            string text = string.IsNullOrWhiteSpace(sheetText) ? "[명령어 목록]" : sheetText;
            text = text.Replace(new string('\u200B', 500), "").Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = new List<string>();
            var added = new bool[Sections.Length];
            bool skip = false;
            foreach (string line in text.Split('\n'))
            {
                string trimmed = line.Trim();
                if (trimmed.StartsWith("[", StringComparison.Ordinal) && trimmed.EndsWith("]", StringComparison.Ordinal))
                {
                    int section = SectionIndex(trimmed);
                    skip = section >= 0;
                    if (skip && !added[section]) { lines.Add(Sections[section]); lines.Add(""); added[section] = true; }
                }
                if (!skip && !OwnedCommand.IsMatch(line)) lines.Add(line);
            }
            for (int i = 0; i < Sections.Length; i++)
                if (!added[i]) { lines.Add(""); lines.Add(Sections[i]); }
            return Regex.Replace(string.Join("\n", lines).Trim(), @"\n{3,}", "\n\n");
        }
    }
}
