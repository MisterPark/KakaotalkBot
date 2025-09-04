using System;
using System.Collections.Generic;

namespace KakaotalkBot
{
    public class Topic
    {
        public string CreatedAt = string.Empty;
        public string Title = string.Empty;

        public List<object> ToRow()
        {
            return new List<object>() { CreatedAt, Title };
        }

        public static Topic ToTopic(List<string> list)
        {
            Topic topic = new Topic();
            topic.CreatedAt = list[0];
            topic.Title = list[1];

            return topic;
        }
    }
}
