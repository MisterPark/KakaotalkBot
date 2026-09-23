using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KakaotalkBot
{
    /// <summary>일반 텍스트와 사용자 ID 멘션을 구분하는 송신 조각입니다.</summary>
    public sealed class LocoMessagePart
    {
        [JsonProperty("kind")] public string Kind { get; private set; }
        [JsonProperty("text", NullValueHandling = NullValueHandling.Ignore)] public string TextValue { get; private set; }
        [JsonProperty("userId", NullValueHandling = NullValueHandling.Ignore)] public string UserId { get; private set; }
        [JsonProperty("nickname", NullValueHandling = NullValueHandling.Ignore)] public string Nickname { get; private set; }

        private LocoMessagePart() { }
        public static LocoMessagePart Text(string text)
        {
            if (text == null) throw new ArgumentNullException("text");
            return new LocoMessagePart { Kind = "text", TextValue = text };
        }
        public static LocoMessagePart Mention(long userId, string nickname)
        {
            if (userId <= 0) throw new ArgumentOutOfRangeException("userId");
            if (string.IsNullOrWhiteSpace(nickname)) throw new ArgumentException("표시 이름이 필요합니다.", "nickname");
            return new LocoMessagePart { Kind = "mention", UserId = userId.ToString(CultureInfo.InvariantCulture), Nickname = nickname };
        }
    }

    public sealed class LocoSenderException : IOException
    {
        public string Code { get; private set; }
        public int? ServerStatus { get; private set; }
        public LocoSenderException(string code, string message, int? serverStatus = null) : base(message)
        {
            Code = code;
            ServerStatus = serverStatus;
        }
    }

    /// <summary>Node 보조 프로세스를 통한 실험용 LOCO 송신 API입니다. 자동 재시도하지 않습니다.</summary>
    public sealed class LocoMentionSender : IDisposable
    {
        private readonly Process process;
        private readonly SemaphoreSlim gate = new SemaphoreSlim(1, 1);
        private int disposed;

        public LocoMentionSender(string nodePath, string bridgePath, string dataDirectory = null)
        {
            nodePath = Path.GetFullPath(nodePath);
            bridgePath = Path.GetFullPath(bridgePath);
            if (!File.Exists(nodePath) || !File.Exists(bridgePath)) throw new FileNotFoundException("Node 또는 LOCO 보조 프로그램을 찾지 못했습니다.");
            dataDirectory = Path.GetFullPath(dataDirectory ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LocoSenderData"));
            process = new Process { StartInfo = new ProcessStartInfo {
                FileName = nodePath,
                Arguments = Quote(bridgePath) + " " + Quote(dataDirectory),
                WorkingDirectory = Path.GetDirectoryName(bridgePath),
                UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = new UTF8Encoding(false), StandardErrorEncoding = new UTF8Encoding(false)
            } };
            try
            {
                if (!process.Start()) throw new IOException("LOCO 보조 프로세스를 시작하지 못했습니다.");
                // 외부 라이브러리의 진단 문자열은 자격 증명을 포함할 수 있어 로그로 남기지 않습니다.
                process.ErrorDataReceived += (sender, args) => { };
                process.BeginErrorReadLine();
            }
            catch { process.Dispose(); throw; }
        }

        private static string Quote(string path)
        {
            if (path.IndexOf('"') >= 0) throw new ArgumentException("경로에 따옴표를 사용할 수 없습니다.");
            // Windows 인수에서 닫는 따옴표 직전의 역슬래시는 두 번 씁니다.
            return "\"" + path.TrimEnd('\\') + new string('\\', (path.Length - path.TrimEnd('\\').Length) * 2) + "\"";
        }

        public Task<JObject> GetInfoAsync() { return RequestAsync(new JObject { ["op"] = "info" }); }
        public Task<JObject> LoginAsync(string email, string password)
        {
            return RequestAsync(new JObject { ["op"] = "login", ["email"] = email, ["password"] = password });
        }
        public Task<JObject> RequestPasscodeAsync(string email, string password)
        {
            return RequestAsync(new JObject { ["op"] = "requestPasscode", ["email"] = email, ["password"] = password });
        }
        public Task<JObject> RegisterDeviceAsync(string email, string password, string passcode)
        {
            return RequestAsync(new JObject { ["op"] = "registerDevice", ["email"] = email, ["password"] = password, ["passcode"] = passcode });
        }
        public Task<JObject> PreviewAsync(long chatId, IEnumerable<LocoMessagePart> parts)
        {
            return RequestAsync(MessageRequest("preview", chatId, parts));
        }
        public Task<JObject> PrepareAsync(long chatId, IEnumerable<LocoMessagePart> parts)
        {
            return RequestAsync(MessageRequest("prepare", chatId, parts));
        }
        public Task<JObject> SendPreparedAsync(string ticket)
        {
            if (string.IsNullOrWhiteSpace(ticket)) throw new ArgumentException("전송 준비 티켓이 필요합니다.", "ticket");
            return RequestAsync(new JObject { ["op"] = "sendPrepared", ["ticket"] = ticket });
        }
        public async Task<JObject> SendMentionAsync(long chatId, long userId, string nickname, string body)
        {
            var prepared = await PrepareAsync(chatId, new[] { LocoMessagePart.Mention(userId, nickname), LocoMessagePart.Text(" " + body) }).ConfigureAwait(false);
            return await SendPreparedAsync((string)prepared["ticket"]).ConfigureAwait(false);
        }

        private static JObject MessageRequest(string operation, long chatId, IEnumerable<LocoMessagePart> parts)
        {
            if (chatId <= 0) throw new ArgumentOutOfRangeException("chatId");
            if (parts == null) throw new ArgumentNullException("parts");
            var snapshot = parts.ToArray();
            if (snapshot.Any(part => part == null)) throw new ArgumentException("메시지 조각은 null일 수 없습니다.", "parts");
            return new JObject { ["op"] = operation, ["chatId"] = chatId.ToString(CultureInfo.InvariantCulture), ["parts"] = JArray.FromObject(snapshot) };
        }

        private async Task<JObject> RequestAsync(JObject request)
        {
            bool sending = (string)request["op"] == "sendPrepared";
            await gate.WaitAsync().ConfigureAwait(false);
            try
            {
                if (Volatile.Read(ref disposed) != 0) throw new ObjectDisposedException("LocoMentionSender");
                string correlationId = Guid.NewGuid().ToString("N");
                request["id"] = correlationId;
                string line = request.ToString(Formatting.None);
                if (Encoding.UTF8.GetByteCount(line) > 60000) throw new ArgumentException("요청이 너무 큽니다.");
                var exchange = ExchangeAsync(line);
                if (await Task.WhenAny(exchange, Task.Delay(35000)).ConfigureAwait(false) != exchange)
                {
                    Dispose();
                    // 취소된 읽기의 예외가 나중에 발생해도 관찰합니다.
                    Observe(exchange);
                    throw new LocoSenderException(sending ? "SEND_UNKNOWN" : "TIMEOUT", "요청 시간이 초과되었습니다. 전송 중이었다면 채팅방에서 결과를 확인해 주세요.");
                }
                string response = await exchange.ConfigureAwait(false);
                if (response == null) throw new IOException("보조 프로세스가 종료되었습니다.");
                var envelope = JObject.Parse(response);
                if ((string)envelope["id"] != correlationId) throw new IOException("보조 프로세스 응답 순서가 일치하지 않습니다.");
                if ((bool?)envelope["ok"] != true)
                {
                    var error = envelope["error"];
                    var code = (string)error["code"];
                    if (code == "TIMEOUT" || code == "SEND_UNKNOWN") Dispose();
                    throw new LocoSenderException(code, (string)error["message"], (int?)error["status"] ?? (int?)error["httpStatus"]);
                }
                return (JObject)envelope["result"];
            }
            catch (LocoSenderException) { throw; }
            catch (Exception error) when (error is IOException || error is JsonException || error is InvalidOperationException)
            {
                Dispose();
                throw new LocoSenderException(sending ? "SEND_UNKNOWN" : "BRIDGE_CLOSED", sending ?
                    "전송 결과를 확인하지 못했습니다. 채팅방을 확인하기 전에는 다시 보내지 마세요." : "보조 프로세스 연결이 종료되었습니다. 진단 창을 다시 열어 주세요.");
            }
            finally
            {
                request.Remove("password");
                request.Remove("passcode");
                gate.Release();
            }
        }

        private static async void Observe(Task task) { try { await task.ConfigureAwait(false); } catch { } }
        private async Task<string> ExchangeAsync(string line)
        {
            // .NET Framework의 기본 표준 입력 인코딩은 시스템 코드 페이지일 수 있습니다.
            byte[] bytes = Encoding.UTF8.GetBytes(line + "\n");
            await process.StandardInput.BaseStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
            await process.StandardInput.BaseStream.FlushAsync().ConfigureAwait(false);
            return await process.StandardOutput.ReadLineAsync().ConfigureAwait(false);
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref disposed, 1) != 0) return;
            try { if (!process.HasExited) process.Kill(); } catch (InvalidOperationException) { }
            process.Dispose();
        }
    }
}
