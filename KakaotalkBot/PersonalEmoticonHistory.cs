using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace KakaotalkBot
{
    internal sealed class PersonalEmoticonHistory
    {
        private sealed class Turn { internal bool Own; internal string Id; internal long At; }
        private readonly Queue<Turn> recent = new Queue<Turn>();
        internal void Add(ChatMessage message)
        {
            recent.Enqueue(new Turn { Own = message.IsOwn, Id = message.EmoticonId, At = message.SendAt });
            while (recent.Count > 1000) recent.Dequeue();
        }
        internal JObject Snapshot(IEnumerable<PersonalEmoticon> available)
        {
            var allowed = new HashSet<string>(available.Select(e => e.Id));
            var mine = recent.Where(t => t.Own && t.Id != null).GroupBy(t => t.Id).OrderByDescending(g => g.Count()).ThenBy(g => g.Key);
            var theirs = recent.Where(t => !t.Own && t.Id != null).GroupBy(t => t.Id).OrderByDescending(g => g.Count()).ThenBy(g => g.Key);
            var pairs = new Dictionary<string, int>();
            Turn previous = null;
            foreach (var turn in recent)
            {
                // 다른 메시지가 끼면 짝짓지 않습니다. 바로 다음 본인 응답, 15분 이내인 경우만 셉니다.
                if (previous != null && !previous.Own && previous.Id != null && turn.Own && turn.Id != null &&
                    previous.At > 0 && turn.At >= previous.At && turn.At - previous.At <= 900 && allowed.Contains(turn.Id))
                {
                    string key = previous.Id + ">" + turn.Id; int count; pairs.TryGetValue(key, out count); pairs[key] = count + 1;
                }
                previous = turn;
            }
            return new JObject {
                ["sample_messages"] = recent.Count,
                ["owner_favorites"] = new JArray(mine.Take(20).Select(g => new JObject { ["id"] = g.Key, ["count"] = g.Count(), ["available"] = allowed.Contains(g.Key) })),
                ["contact_favorites"] = new JArray(theirs.Take(20).Select(g => new JObject { ["id"] = g.Key, ["count"] = g.Count() })),
                ["reply_pairs"] = new JArray(pairs.OrderByDescending(p => p.Value).ThenBy(p => p.Key).Take(30).Select(p => new JObject {
                    ["contact"] = p.Key.Split('>')[0], ["owner"] = p.Key.Split('>')[1], ["count"] = p.Value }))
            };
        }
    }
}
