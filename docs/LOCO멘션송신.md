# 사용자 ID로 멘션 보내기: LOCO 진단 모듈

## 현재 구현 범위

EDB 수신 및 명령 분석과 별개로 C# → Node 보조 프로세스 → LOCO 송신 경로를 추가했다.
`user_id`를 지정해 멘션을 구성하고, 로그인한 계정의 방과 참가자를 확인한 뒤 전송하도록 구현했다.
현재 신규 기기 등록 API가 HTTP 404로 실패해 로그인과 실제 멘션 송신은 진행할 수 없다.
따라서 `Bot.ProcessCommand()`의 자동 송신 경로에는 연결하지 않았다. 오프라인 멘션 데이터 생성과 프로세스 통신은 검증했다.

사용한 `node-kakao`는 npm의 `4.5.0`을 고정했으며, 상위 저장소에서 유지보수 중단을 표시하고 있다.
라이브러리의 기본 앱 버전은 `3.2.3.2698`이며, 사용자 로그인 시도에서 `WEB_LOGIN(-999)`가 보고되었다.
설치된 `KakaoTalk.exe`의 제품 버전을 읽어 웹 인증과 LOCO 설정에 함께 적용하도록 수정했다.
이번 PC에서는 `26.8.1.5315`를 확인했다. 버전 설정을 맞춘 것만으로 인증 규약 호환성이 보장되는 것은 아니다.

이후 사용자가 **인증번호 요청에서 HTTP 404**를 확인했다.
상위 라이브러리가 사용하는 경로는 `https://katalk.kakao.com/win32/account/request_passcode.json`이다.
현재 이 경로로 신규 기기 등록을 완료할 수 없으며, 사용할 수 있는 대체 인증 절차는 아직 확인하지 못했다.
등록 버튼과 등록 API 호출을 미지원으로 바꾸고, `-100` 이후 같은 보조 프로세스에서 반복 로그인하는 동작도 차단했다.

## 실행

빌드한 실행 파일:

`artifacts\LocoMentionProbe\LocoMentionProbe.exe`

이번 수정 시 기존 exe가 실행 중이어서 수정한 GUI는 같은 폴더의 `LocoMentionProbe.Updated.exe`로 별도 빌드했다.
기존 창에는 새 GUI가 자동 적용되지 않는다. 보조 모듈은 갱신했으며, 다음 정상 빌드는 기본 exe 이름을 사용한다.

다시 빌드하고 열기:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Start-LocoMentionProbe.ps1
```

Node.js 22 이상과 Visual Studio MSBuild가 필요하다. Node가 PATH에 없으면 `-NodePath '절대경로\node.exe'`를 지정한다.
새 환경에서는 먼저 `tools\LocoSender`에서 `npm ci --ignore-scripts`로 잠금 파일에 고정한 의존성을 설치한다.
이번 환경에서는 도구 폴더에 설치를 마쳤고, 실행 파일 폴더로도 복사했다.
실행 파일은 옆의 `LocoSender` 폴더, `Newtonsoft.Json.dll`, `runtime.json`을 함께 사용한다.
Node 실행 파일 경로는 `runtime.json`에 기록된다. 다른 PC에서는 시작 스크립트로 다시 빌드한다.
시작 스크립트는 설치된 카카오톡 제품 버전으로 `LocoSender\client-profile.json`도 만든다.
카카오톡 업데이트 후에는 `Update-LocoClientProfile.ps1 -DestinationDirectory .\artifacts\LocoMentionProbe\LocoSender`를 실행하고 진단 창에서 **연결 초기화**를 누르면 반영된다.

**현재는 인증번호 요청과 기기 등록을 다시 시도하지 않는다.** 새 진단 창에는 두 버튼이 미지원으로 표시된다.
`-100`은 `REGISTRATION_UNSUPPORTED`로 전달하며, 로그인 반복을 중단하고 등록 절차 호환성 문제를 표시한다.
이미 생성한 기기 식별자는 유지한다. 로그인 실패를 이유로 새 식별자를 생성하지 않는다.

**데이터 미리보기**는 로그인 없이 채팅방 ID, 대상 `user_id`, 표시 이름, 본문으로 멘션 JSON을 만든다.
서버 호환성 문제를 해결한 뒤에는 **방·대상 ID 확인** → 결과 검토 → **확인한 내용 1회 전송** 순으로 검증한다.
실제 송신 검증은 반환한 `logId`와 EDB의 `chatLogs.attachement.mentions[].user_id`를 대조해야 끝난다.

비밀번호는 채팅이나 설정 파일에 붙여넣을 필요가 없다. 진단 창에서 입력한 인증 정보는 보조 프로세스의 표준 입력으로만 전달한다.
비밀번호·토큰은 파일에 저장하지 않으며, 로그인 성공 시 비밀번호 입력란을 비운다.
기기 식별자만 **진단 exe 위치의 `LocoSenderData\device.json`**에 자동 저장한다. 저장 폴더 선택은 필요 없다.
현재 `RequestPasscodeAsync`와 `RegisterDeviceAsync`는 서버 요청 없이 `REGISTRATION_UNSUPPORTED` 오류를 반환한다.

## C# 사용 API

아래 송신 예시는 인증 호환성 해결 후 사용할 호출 구조다. 현재 환경에서 동작을 확인한 송신 예제가 아니다.

`KakaotalkBot\LocoMentionSender.cs`는 프로젝트 Compile 항목에 추가했다.
별도 Node 프로세스를 시작하므로 사용하는 동안 인스턴스를 유지하고, 종료할 때 Dispose한다.
`bridgeJsPath`는 진단 빌드 결과의 `LocoSender\bridge.js`를 지정할 수 있다.
소스의 `tools\LocoSender\bridge.js`를 직접 사용한다면 `tools\Update-LocoClientProfile.ps1`로 해당 폴더의 버전 설정을 먼저 만든다.

```csharp
using (var sender = new LocoMentionSender(nodeExePath, bridgeJsPath))
{
    await sender.LoginAsync(email, password);

    // 실제 메시지를 보냅니다. 명령 작성자는 command.AuthorId로 지정할 수 있습니다.
    await sender.SendMentionAsync(
        command.ChatId, command.AuthorId, command.Nickname, "요청을 처리했습니다.");
}
```

여러 사용자 또는 전송 전 검토가 필요한 경우:

```csharp
var parts = new[] {
    LocoMessagePart.Mention(firstUserId, "첫 사용자"),
    LocoMessagePart.Text(" "),
    LocoMessagePart.Mention(secondUserId, "둘째 사용자"),
    LocoMessagePart.Text(" 확인해 주세요.")
};

var localPreview = await sender.PreviewAsync(chatId, parts); // 오프라인
var prepared = await sender.PrepareAsync(chatId, parts);    // 방·사용자 확인
// prepared의 channelName, accountUserId, text, mentions를 표시합니다.
var receipt = await sender.SendPreparedAsync((string)prepared["ticket"]);
```

`SendMentionAsync`는 준비와 전송을 연속 수행하는 API이며 실제 송신 기능이다.
현재 진단 창은 `PrepareAsync`와 `SendPreparedAsync`를 분리해 사용자가 표시된 결과를 확인하게 한다.
사용자 정보 조회에 실패하면 닉네임 검색이나 첫 번째 후보 선택으로 대체하지 않고 오류를 반환한다.

## 데이터와 오류 처리

- C# ↔ Node 경계에서 ID를 십진 문자열로 전달한다. Node에서는 `Long`으로 변환하고 실제 WRITE의 `extra`는 정밀도를 보존한 JSON 숫자로 직렬화한다.
- 멘션은 라이브러리의 `MentionContent`로 구성한다. `at`는 1부터 시작하는 멘션 순번, `len`은 @를 제외한 UTF-16 표시 이름 길이다.
- 일반 @ 문자와 멘션이 섞인 경우의 실제 서버 해석은 검증되지 않아, 현재 도구는 멘션 본문의 일반 @ 및 표시 이름의 @를 거절한다.
- 도구의 입력 제한은 본문 4,000자, 메시지 조각 64개다. 카카오톡 서비스의 공식 제한을 뜻하지 않는다.
- 전송 준비 티켓은 약 2분 유효하며 한 번만 사용할 수 있다. 새 준비, 연결 종료, 전송 시 기존 티켓을 폐기한다.
- `SEND_UNKNOWN`은 메시지가 전송됐을 가능성이 있는 상태다. 실제 방 또는 EDB를 확인한 후 재시도 여부를 결정해야 한다.
- 같은 계정의 동시 PC 접속 제한 여부와 기존 PC 세션에 미치는 영향은 서버 로그인 검증이 필요하다. `forced=false`로만 요청한다.
- 보조 프로세스 요청은 30초, C# 응답 대기는 35초에 종료한다. 통신 오류에는 자격 증명이 포함될 수 있어 원본 HTTP 오류와 토큰을 기록하지 않는다.
- 수신 이벤트에 자동 응답하는 Node 리스너는 등록하지 않았다. 수신·명령 처리는 기존 EDB 경로를 사용한다.

## 2026-09-21 확인 결과

- 봇 프로젝트 및 별도 진단 창 빌드 성공. 기존 다른 파일의 경고는 남아 있다.
- Node 테스트 9개 통과: 큰 ID, 동명이인, 중복 멘션, UTF-16 길이, 실제 라이브러리 WRITE 직렬화, 티켓 중복 방지, 강제 로그인 금지, 표준 입출력 통신, 인증/LOCO 버전 설정 일치, 미지원 등록 요청 차단 등을 검사했다.
- C# 연결 검증 9개 통과: 한글·이모지, 큰 ID, 동시 요청 직렬화, 미로그인 차단, 프로세스 종료 시 전송 결과 불명 처리, 등록 미지원 상태와 오류 전달.
- `booking-loco.kakao.com:443` TCP 연결 성공.
- `https://katalk.kakao.com/` TLS 연결 성공, 루트 경로 HEAD 응답은 404. 인증 경로의 동작이나 로그인이 성공했다는 의미는 아니다.
- 사용자가 초기 기본 버전으로 로그인한 결과는 `WEB_LOGIN(-999)`였다. 설치 버전 `26.8.1.5315` 적용 후 `WEB_LOGIN(-100)`, 인증번호 요청에서 HTTP 404를 확인했다. 이 경로로는 등록과 로그인을 완료할 수 없다.
- 에이전트는 인증번호 요청이나 메시지 전송을 실행하지 않았다. 로그인 성공과 실제 멘션 수신은 검증되지 않았다.
- 설치된 Windows exe의 ASCII/UTF-16 문자열에서 기존 등록 경로와 관련 계정 URL을 제한적으로 검색했으나 유효한 대체 경로는 얻지 못했다. 이것이 대체 절차가 없다는 증거는 아니다.
- 다른 프로젝트에서도 2026-06-29에 macOS의 같은 이름의 인증번호/등록 API가 404여서 등록 구현을 철회한 기록이 있다. Windows의 구체적인 실패 근거는 이번 사용자의 실행 결과이며, macOS 기록만으로 모든 플랫폼의 API 폐지를 단정하지 않는다.

테스트:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-LocoSender.ps1
```

## 참고 소스

- [상위 프로젝트와 유지보수 상태](https://github.com/storycraft/node-kakao)
- [계정 로그인 구현](https://github.com/storycraft/node-kakao/blob/stable/src/api/auth-api-client.ts)
- [멘션 구성 구현](https://github.com/storycraft/node-kakao/blob/stable/src/chat/content/mention.ts)
- [WRITE 송신 구현](https://github.com/storycraft/node-kakao/blob/stable/src/talk/channel/talk-channel-session.ts)
- [macOS 등록 경로 404와 구현 철회 기록](https://github.com/JungHoonGhae/openkakao-cli/blob/main/CHANGELOG.md#133---2026-06-29)
