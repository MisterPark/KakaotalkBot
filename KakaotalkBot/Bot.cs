﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KakaotalkBot
{
    public class Bot
    {
        public struct QuizAnswer
        {
            public string Nickname;
            public string Answer;
        }

        public static string TargetWindow = string.Empty;

        private Dictionary<string, WindowInfo> windowList = new Dictionary<string, WindowInfo>();

        public static bool IsBotRunning { get; set; }

        private List<string> chatLog = new List<string>();
        private string lastChat = "xx";

        private Queue<Command> commands = new Queue<Command>();
        private Queue<QuizAnswer> quizAnswers = new Queue<QuizAnswer>();

        private Random random;

        CustomTimer soliloquyTimer = new CustomTimer(300000);
        CustomTimer newsTimer = new CustomTimer(3600000);
        CustomTimer dbTimer = new CustomTimer(60000);

        private DateTime lastUpdate = DateTime.MinValue;

        private bool isCorrect = false;

        public Bot()
        {
            UpdateWindowList(); 
            random = new Random(DateTime.Now.Millisecond);
        }

        public void Update()
        {
            if (IsBotRunning == false) return;

            ProcessCopyChat();
            ProcessReset();

            if (soliloquyTimer.Check(Time.DeltaTime))
            {
                //ProcessComonBot();

                ProcessNextCommonSense();
            }

            if (newsTimer.Check(Time.DeltaTime))
            {
                ProcessNews();
            }

            if (dbTimer.Check(Time.DeltaTime))
            {
                ProcessUpdateDB();
            }

            ProcessQuiz();
        }

        private void UpdateWindowList()
        {
            windowList.Clear();
            var list = WindowsMacro.Instance.GetWindowList();
            foreach (var window in list)
            {
                if(windowList.ContainsKey(window.Title) == false)
                {
                    windowList.Add(window.Title, window);
                }
            }
        }

        private void ProcessCopyChat()
        {
            if (string.IsNullOrEmpty(TargetWindow)) return;

            IntPtr handle = windowList[TargetWindow].Handle;

            string chat = WindowsMacro.Instance.CopyChatroomText(handle);
            string[] lines = chat.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            bool isFirstOpen = chatLog.Count == 0;

            chatLog = lines.Where(x => !string.IsNullOrEmpty(x)).ToList();

            if (isFirstOpen && chatLog.Count > 0)
            {
                lastChat = chatLog.Last();
                return;
            }

            int idx = chatLog.Count - 1;
            for (int i = chatLog.Count - 1; i >= 0; i--)
            {
                if (chatLog[i] == lastChat)
                {
                    idx = i;
                    break;
                }
            }

            for (int i = idx + 1; i < chatLog.Count; i++)
            {
                string line = chatLog[i];
                if (line.Contains("님이 들어왔습니다."))
                {
                    int index = line.LastIndexOf("님");
                    string nick = line.Substring(0, index);
                    ProcessKeyword(nick, "/입장");
                    continue;
                }
                else if (line.Contains("님이 나갔습니다."))
                {
                    int index = line.LastIndexOf("님");
                    string nick = line.Substring(0, index);
                    ProcessKeyword(nick, "/퇴장");
                    continue;
                }

                if (!line.StartsWith("[")) continue;

                int firstClose = line.IndexOf(']');
                int secondClose = line.IndexOf(']', firstClose + 1);

                if (firstClose != -1 && secondClose != -1 && secondClose + 2 <= line.Length)
                {
                    string nickname = line.Substring(1, firstClose - 1).Trim();
                    string message = line.Substring(secondClose + 2).Trim();

                    ProcessQuizAnswer(nickname, message);
                    // 키워드 포함 여부 검사
                    if (Database.Instance.Keywords.Any(k => message.StartsWith(k)))
                    {
                        if (string.IsNullOrEmpty(nickname) == false)
                        {
                            ProcessKeyword(nickname, message);
                        }
                    }
                }
            }

            if (chatLog.Count > 0)
            {
                lastChat = chatLog[chatLog.Count - 1];
            }

            ProcessCommand();

            Form1.Instance.UpdateChatLog(string.Join("\n", chatLog));
        }

        private void ProcessKeyword(string nickname, string message)
        {
            //SendTextToChatroom(TargetWindow, message);

            Command command = new Command();
            command.Nickname = nickname;
            command.Keyword = message;
            commands.Enqueue(command);
        }

        private void ProcessQuizAnswer(string nickname, string answer)
        {
            QuizAnswer quizAnswer = new QuizAnswer();
            quizAnswer.Nickname = nickname;
            quizAnswer.Answer = answer;
            quizAnswers.Enqueue(quizAnswer);
        }

        private void ProcessCommand()
        {
            if (commands.Count == 0) return;


            Command command = commands.Dequeue();

            //string answer = GetAnswer(command.Keyword);

            //if (string.IsNullOrEmpty(answer) == false)
            //{
            //    SendTextToChatroom(TargetWindow, $"{answer}");
            //}

            if (command.Keyword == "/?" || command.Keyword == "/명령어" || command.Keyword == "/훈장" || command.Keyword == "/공지사항" || command.Keyword == "/패치노트")
            {
                string answer = Database.Instance.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{answer}");
                }
            }
            else if (command.Keyword == "/입장" || command.Keyword == "/퇴장")
            {
                string answer = Database.Instance.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    SmartString.CurrentNickname = command.Nickname;
                    string parsedAnswer = SmartString.Parse(answer);
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, parsedAnswer);
                }
            }
            else if (command.Keyword == "/출첵")
            {
                string answer = Database.Instance.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    if (Database.Instance.CheckAttendance(command.Nickname))
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"이미 출석한 유저입니다.");
                    }
                    else
                    {
                        if (Database.Instance.FindUser(command.Nickname, out User user))
                        {
                            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"[{command.Nickname}]님이 {answer}\n+10포인트\n(현재 포인트: {user.Point})");
                        }
                        else
                        {
                            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"[{command.Nickname}]님이 {answer}\n+10포인트");
                        }


                    }
                }
            }
            else if (command.Keyword.StartsWith("/조회"))
            {
                if (command.Keyword == "/조회")
                {
                    if (Database.Instance.FindUser(command.Nickname, out User user))
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"=====[유저조회]=====\n닉네임: {user.Nickname}\n포인트: {user.Point}\n인기도: {user.Popularity}\n=================");
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"정보가 없는 유저입니다.");
                    }
                }
                else if (command.Keyword.Length > 4)
                {
                    string param = command.Keyword.Substring(4);
                    param = param.Replace("@", "");
                    if (Database.Instance.FindUser(param, out User user))
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"=====[유저조회]=====\n닉네임: {user.Nickname}\n포인트: {user.Point}\n인기도: {user.Popularity}\n=================");
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"정보가 없는 유저입니다.");
                    }
                }
            }
            else if (command.Keyword.StartsWith("/랭킹"))
            {
                if (command.Keyword == "/랭킹")
                {
                    string answer = Database.Instance.GetAnswer(command.Keyword);

                    if (string.IsNullOrEmpty(answer)) return;

                    StringBuilder sb = new StringBuilder();
                    int beforeRank = 1;
                    int beforePop = 0;
                    List<User> rank = Database.Instance.GetPopularityRank();

                    sb.AppendLine(answer);
                    for (int i = 0; i < rank.Count; i++)
                    {
                        int currentPop = rank[i].Popularity;
                        if (currentPop != beforePop)
                        {
                            beforeRank = i + 1;
                        }

                        string emoji = string.Empty;
                        if (beforeRank == 1)
                        {
                            emoji = "🥇";
                        }
                        else if (beforeRank == 2)
                        {
                            emoji = "🥈";
                        }
                        else if (beforeRank == 3)
                        {
                            emoji = "🥉";
                        }

                        sb.AppendLine($"{emoji}{beforeRank}위 {rank[i].Nickname} {rank[i].Popularity}");

                        beforePop = currentPop;
                    }

                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, sb.ToString());
                }
            }
            else if (command.Keyword.StartsWith("/토론"))
            {
                if (command.Keyword == "/토론")
                {
                    string answer = Database.Instance.GetAnswer(command.Keyword);

                    if (string.IsNullOrEmpty(answer)) return;

                    StringBuilder sb = new StringBuilder();

                    List<Topic> topics = Database.Instance.GetOrderedTopics();
                    string currentCategory = string.Empty;
                    int count = 0;

                    sb.AppendLine(answer);
                    for (int i = 0; i < topics.Count; i++)
                    {
                        if (currentCategory != topics[i].Category)
                        {
                            sb.AppendLine();
                            sb.AppendLine($"{topics[i].Category}");
                            currentCategory = topics[i].Category;
                            count = 0;
                        }
                        sb.AppendLine($"{count + 1}. {topics[i].Title}({topics[i].CreatedAt})");
                        sb.AppendLine($"(토론왕👑: {topics[i].Winner})");
                        count++;
                    }

                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, sb.ToString());
                }
            }
            else if (command.Keyword.StartsWith("/좋아"))
            {
                string[] format = new string[3];
                if (SmartString.ParseCommand(command.Keyword, out format) == ParseResult.Success)
                {
                    string keyword = format[0];
                    string nickname = format[1];
                    string number = format[2];

                    if (nickname.StartsWith("@"))
                    {
                        nickname = nickname.Substring(1);
                    }

                    int point = 10;
                    int.TryParse(number, out point);

                    if (keyword != "/좋아")
                    {
                        return;
                    }

                    if (command.Nickname == nickname)
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"자신에게 할 수 없는 명령입니다.");
                        return;
                    }

                    if (Database.Instance.FindUser(command.Nickname, out User a))
                    {
                        if (Database.Instance.FindUser(nickname, out User b))
                        {
                            if (a.Point < Math.Abs(point))
                            {
                                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"포인트가 부족합니다.\n남은 포인트: {a.Point}");
                            }
                            else
                            {
                                a.Point -= Math.Abs(point);
                                b.Popularity += point;
                                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"[{a.Nickname}]님이 [{b.Nickname}]님에게 👍좋아요.\n인기도 {point}점 상승\n현재 인기도:{b.Popularity}");
                            }
                        }
                        else
                        {
                            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"정보가 없는 유저입니다.");
                        }
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"정보가 없는 유저입니다.");
                    }
                }
                else
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"잘못된 형식입니다.\n형식1: /좋아 [닉네임]\n형식2: /좋아 [닉네임] [숫자]");
                }

            }
            else if (command.Keyword.StartsWith("/싫어"))
            {
                string[] format = new string[3];
                if (SmartString.ParseCommand(command.Keyword, out format) == ParseResult.Success)
                {
                    string keyword = format[0];
                    string nickname = format[1];
                    string number = format[2];

                    if (nickname.StartsWith("@"))
                    {
                        nickname = nickname.Substring(1);
                    }

                    int point = 10;
                    int.TryParse(number, out point);

                    if (keyword != "/싫어")
                    {
                        return;
                    }

                    if (command.Nickname == nickname)
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"자신에게 할 수 없는 명령입니다.");
                        return;
                    }

                    if (Database.Instance.FindUser(command.Nickname, out User a))
                    {
                        if (Database.Instance.FindUser(nickname, out User b))
                        {
                            if (a.Point < Math.Abs(point))
                            {
                                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"포인트가 부족합니다.\n남은 포인트: {a.Point}");
                            }
                            else
                            {
                                a.Point -= Math.Abs(point);
                                b.Popularity -= point;
                                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"[{a.Nickname}]님이 [{b.Nickname}]님에게 👎싫어요.\n인기도 {point}점 하락\n현재 인기도:{b.Popularity}");
                            }
                        }
                        else
                        {
                            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"정보가 없는 유저입니다.");
                        }
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"정보가 없는 유저입니다.");
                    }
                }
                else
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"잘못된 형식입니다.\n형식1: /싫어 [닉네임]\n형식2: /싫어 [닉네임] [숫자]");
                }
            }
            else if (command.Keyword.StartsWith("/정치뉴스"))
            {
                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{News.PoliticsTop6}");
            }
            else if (command.Keyword == "/상식퀴즈")
            {
                ProcessCommonSense();
            }
            else
            {
                string[] answers = Database.Instance.GetAnswers(command.Keyword);
                if (answers == null || answers.Length == 0)
                {
                    return;
                }

                int rand = random.Next(0, answers.Length);
                string answer = answers[rand];

                if (string.IsNullOrEmpty(answer) == false)
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{answer}");
                }
            }

        }

        private void ProcessCommonSense()
        {
            string answer = Database.Instance.GetCommonSenseText();
            if (string.IsNullOrEmpty(answer))
            {
                double left = soliloquyTimer.TimeLeft * 0.001;

                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"다음 퀴즈를 준비하고 있습니다.\n남은 시간: {left}초");
            }
            else
            {
                WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{answer}");
            }
        }

        private void ProcessReset()
        {
            if (lastUpdate.Day != DateTime.Now.Day)
            {
                lastUpdate = DateTime.Now;
                Database.Instance.ResetAttendance();
            }
        }

        private void ProcessNextCommonSense()
        {
            Quiz quiz = Database.Instance.GetCurrentQuiz();
            if (quiz != null)
            {
                if (isCorrect == false)
                {
                    WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"정답자가 없습니다.\n정답: {quiz.Answer}\n해설: {quiz.Explanation}");
                }
            }

            Database.Instance.SetNextCommonSense();
            ProcessCommonSense();
        }

        private void ProcessNews()
        {
            WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"{News.PoliticsTop6}");
        }

        private void ProcessUpdateDB()
        {
            Database.Instance.UpdateCommands();
            Database.Instance.UpdateUserTable();
            Database.Instance.UpdateCommonSenses();
            Database.Instance.UpdateTopic();
            News.Update();
        }

        private void ProcessQuiz()
        {
            Quiz quiz = Database.Instance.GetCurrentQuiz();
            if (quiz == null)
            {
                return;
            }

            while (quizAnswers.Count != 0)
            {
                QuizAnswer quizAnswer = quizAnswers.Dequeue();


                if (quizAnswer.Answer == quiz.Answer)
                {
                    if (Database.Instance.FindUser(quizAnswer.Nickname, out User a))
                    {
                        int point = 0;
                        if (quiz.Difficulty == "최상")
                        {
                            point = 5;
                        }
                        else if (quiz.Difficulty == "상")
                        {
                            point = 4;
                        }
                        else if (quiz.Difficulty == "중")
                        {
                            point = 3;
                        }
                        else if (quiz.Difficulty == "하")
                        {
                            point = 2;
                        }
                        else
                        {
                            point = 1;
                        }

                        a.Point += point;
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"정답자: {quizAnswer.Nickname}\n정답: {quiz.Answer}\n해설: {quiz.Explanation}\n+{point} 포인트 득점!!👍\n 현재 포인트: {a.Point}");
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(TargetWindow, $"정답: {quiz.Answer}\n해설: {quiz.Explanation}");
                    }

                    Database.Instance.CurrentAnswerIndex = -1;
                    break;
                }
            }

            quizAnswers.Clear();

        }
    }
}
