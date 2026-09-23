using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace KakaotalkBot
{
    /// <summary>현재 메시지에서 태그된 사용자와 멘션 위치 정보입니다.</summary>
    public sealed class ChatMention
    {
        internal static readonly IReadOnlyList<ChatMention> Empty = Array.AsReadOnly(new ChatMention[0]);

        public long UserId { get; private set; }
        // 문자열 위치가 아닌, 1부터 시작하는 멘션 순번입니다.
        public IReadOnlyList<int> At { get; private set; }
        // @를 제외한 표시 이름의 길이입니다.
        public int Length { get; private set; }
        // 해당 방의 사용자 정보가 아직 없으면 null입니다. 대상 식별에는 UserId를 사용합니다.
        public string Nickname { get; private set; }

        public ChatMention(long userId, IEnumerable<int> at, int length, string nickname = null)
        {
            if (userId <= 0) throw new ArgumentOutOfRangeException("userId");
            if (at == null) throw new ArgumentNullException("at");
            int[] positions = at.Distinct().OrderBy(index => index).ToArray();
            if (positions.Length == 0 || positions[0] <= 0) throw new ArgumentException("멘션 순번은 1 이상이어야 합니다.", "at");
            if (length <= 0) throw new ArgumentOutOfRangeException("length");
            UserId = userId;
            At = Array.AsReadOnly(positions);
            Length = length;
            Nickname = nickname;
        }
    }

    public struct Command
    {
        public string Nickname;
        public string Keyword;
        public long AuthorId;
        public long LogId;
        public long ChatId;
        public long ReceivedAt;

        /// <summary>실제 멘션 한 명과 뒤따르는 메모 내용을 읽습니다.</summary>
        public bool TryReadMentionText(string verb, out long targetId, out string reason, int limit = 1000)
        {
            targetId = 0; reason = null;
            if (AuthorId <= 0 || Keyword == null || !Keyword.StartsWith(verb + " ", StringComparison.Ordinal) ||
                Mentions.Count != 1 || Mentions[0].At.Count != 1 || Mentions[0].At[0] != 1) return false;
            string argument = Keyword.Substring(verb.Length + 1).TrimStart();
            int length = Mentions[0].Length + 1;
            if (!argument.StartsWith("@", StringComparison.Ordinal) || length > argument.Length) return false;
            string tail = argument.Substring(length);
            if (tail.Length > 0 && !char.IsWhiteSpace(tail[0])) return false;
            tail = tail.Trim();
            if (tail.Length > limit || tail.Any(char.IsControl)) return false;
            targetId = Mentions[0].UserId; reason = tail; return true;
        }

        private IReadOnlyList<ChatMention> mentions;
        public IReadOnlyList<ChatMention> Mentions
        {
            get { return mentions ?? ChatMention.Empty; }
            set { mentions = value == null ? null : Array.AsReadOnly(value.Where(mention => mention != null).ToArray()); }
        }

        /// <summary>등장 순서대로 중복을 제거한 태그 대상 ID입니다. 작성자는 AuthorId로 확인합니다.</summary>
        public IReadOnlyList<long> MentionedUserIds
        {
            get { return Array.AsReadOnly(Mentions.OrderBy(mention => mention.At[0]).Select(mention => mention.UserId).Distinct().ToArray()); }
        }

        /// <summary>실제 멘션 하나의 ID를 읽는다. 이름 문자열이나 가짜 @텍스트로는 대상을 지정할 수 없다.</summary>
        public bool TryReadTarget(string keyword, bool withAmount, out long targetId, out int amount)
        {
            string nickname;
            return TryReadTarget(keyword, withAmount, out targetId, out amount, out nickname);
        }

        /// <summary>대상 ID와 초기 등록용 이름을 함께 읽는다. 프로필이 없으면 멘션 표시 이름을 사용한다.</summary>
        public bool TryReadTarget(string keyword, bool withAmount, out long targetId, out int amount, out string nickname)
        {
            targetId = 0;
            amount = 10;
            nickname = null;
            if (AuthorId <= 0 || Keyword == null || !Keyword.StartsWith(keyword + " ", StringComparison.Ordinal) ||
                Mentions.Count != 1 || Mentions[0].At.Count != 1 || Mentions[0].At[0] != 1) return false;
            string argument = Keyword.Substring(keyword.Length).TrimStart();
            if (!argument.StartsWith("@", StringComparison.Ordinal) || Mentions[0].Length >= argument.Length) return false;
            int mentionLength = Mentions[0].Length + 1;
            string tail = argument.Substring(mentionLength);
            if (tail.Length > 0 && !char.IsWhiteSpace(tail[0])) return false;
            tail = tail.Trim();
            if (tail.Length > 0 && (!withAmount || !int.TryParse(tail, NumberStyles.None, CultureInfo.InvariantCulture, out amount) || amount <= 0)) return false;
            targetId = Mentions[0].UserId;
            nickname = string.IsNullOrWhiteSpace(Mentions[0].Nickname) ? argument.Substring(1, Mentions[0].Length) : Mentions[0].Nickname;
            return true;
        }
    }
}
