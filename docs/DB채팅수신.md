# DB 채팅 수신 사용법

## 실행 순서

1. Windows x64에서 봇을 실행한다. 기존 Google Sheets 설정과 보이스룸 리소스 등 앱의 실행 준비는 기존과 같다.
2. 카카오톡에 로그인하고 대상 방을 연다. 초기 키 탐색에는 카카오톡 프로세스에 DB 키가 올라와 있어야 한다.
3. **채팅 DB** 탭의 **DB 목록 읽기**를 누른다. 앱을 처음 표시할 때도 목록을 자동으로 읽는다.
4. 필요하면 계정 필터와 검색창을 사용해 방을 선택한다. 방 선택값은 계정 경로와 `chatId`다.
5. 위쪽 텍스트 상자에는 답변을 보낼 카카오톡 채팅창 이름이 표시된다. 실제 창 제목과 다르면 수정한다.
6. **DB 수신 시작**을 누른다. 초기 복호화가 끝나고 `수신 중`이 표시된 뒤 들어온 메시지가 명령 처리로 전달된다.
7. **DB 수신 중지** 후에는 이전 작업이 끝나야 다른 방을 선택하거나 다시 시작할 수 있다.

현재 로그인 계정이라고 폴더 순서만으로 추정하지 않는다. 실행 중인 카카오톡에서 실제 키를 찾고 복호화할 수 있는 계정의 방 목록을 표시한다.
초기 키 탐색·전체 복호화는 DB 크기에 따라 시간이 걸린다. 이후 기본 확인 간격은 200ms이며, 카카오톡의 DB 저장 시점과 봇의 기존 명령 처리 시간은 추가된다.

## 저장 위치

```text
exe 폴더/DecryptedChat/<계정 폴더명>/
  catalog-rooms.sqlite
  catalog-users.sqlite
  chatListInfo.sqlite
  TalkUserDB.sqlite
  chatLogs_<chatId>.sqlite
```

폴더는 자동 생성된다. 키는 메모리에서 재사용하고 종료 시 해제한다.
이 폴더는 Git 추적에서 제외한다. 사본의 쓰기는 수신 작업자 하나가 담당하며 앱 내부 조회 연결은 갱신 전에 닫는다.
외부 SQLite 편집기로 같은 사본을 열어 둔 경우에는 편집기를 닫고 다시 시도한다.

## 입력 처리

- 창 전체복사와 마지막 문자열 비교 대신 DB의 `logId`로 신규 행을 구분한다.
- 최초 시작은 현재 최대 `logId`를 기준으로 잡아 과거 명령을 실행하지 않는다.
- 일시적인 읽기 오류나 키 재발견 중에는 기존 읽기 기준점을 유지한다.
- 일반 텍스트와 답글을 기존 키워드 검사·퀴즈·기여도 처리로 전달한다. 사진·이모티콘·삭제된 메시지는 명령 입력과 기여도 집계에서 제외한다.
- 입장·초대 및 퇴장·강퇴 이벤트는 해당 멤버의 이름으로 기존 `/입장`, `/퇴장` 명령에 연결한다.
- 본인 메시지는 `나와의 채팅`의 멤버 ID와 이 PC에서 보낸 텍스트의 작성자 ID로 제외한다. 자동 확인이 어려우면 본인 작성자 ID를 직접 입력할 수 있다.
- 오픈채팅 이름은 해당 방의 `linkId`에 맞는 사용자 정보를 사용한다. 작성자 이름이 확인되지 않으면 기준점을 넘기지 않고 사용자 정보를 다시 읽는다.
- 최근 256행을 겹쳐 조회하여 지연 저장을 보완한다. 해당 구간보다 오래된 기록의 뒤늦은 동기화까지 완전하게 재처리하는 기능은 아니다.
- 읽기 상태는 실행 중 메모리에 보관한다. 수동 재시작·앱 재시작은 현재 최대 ID부터 시작한다. 강제 종료까지 포함한 명령의 정확히 한 번 실행은 보장하지 않는다.

수신 큐는 최대 4,096건이며 명령 큐가 밀리면 소비를 잠시 멈춘다. 백그라운드에서는 UI 컨트롤이나 기존 명령 큐에 직접 접근하지 않는다.
`Bot.Update()`가 수신 큐를 비우고 `HandleIncomingMessage()`와 `ProcessCommand()`를 호출한다.

## 멘션 사용자 ID 사용

수신기는 `chatLogs.attachement`의 JSON에서 현재 메시지의 `mentions`를 읽어 `ChatMessage`와 `Command`에 전달한다.
DB 컬럼 이름은 실제 저장 형식에 맞춰 `attachement`를 사용한다.

| 명령 API | 내용 |
| --- | --- |
| `command.AuthorId` | 명령을 보낸 사용자의 ID |
| `command.MentionedUserIds` | 등장 순서대로 중복을 제거한 태그 대상 ID 목록 |
| `command.Mentions` | 사용자 ID, 반복 태그 순번, 표시 이름 길이와 프로필 이름 |
| `mention.UserId` | 태그 대상 ID (`long`) |
| `mention.At` | 1부터 시작하는 멘션 순번 목록. 문자열 글자 위치가 아님 |
| `mention.Length` | `@`를 제외한 표시 이름 길이 |
| `mention.Nickname` | 해당 방의 사용자 정보에서 확인한 이름. 미확인 시 `null` |

`Bot.ProcessCommand()`에서 명령을 꺼낸 뒤 다음과 같이 사용할 수 있다.

```csharp
long senderId = command.AuthorId;
var targetIds = command.MentionedUserIds;
if (targetIds.Count == 1)
{
    long targetUserId = targetIds[0];
    bool targetsSender = targetUserId == senderId;
    // targetUserId를 사용해 명령 대상을 식별합니다.
}

foreach (ChatMention mention in command.Mentions)
{
    long userId = mention.UserId;
    var positions = mention.At;
    string nickname = mention.Nickname;
}
```

동일 사용자를 반복 태그해도 `MentionedUserIds`에는 한 번만 나타나며, 전체 순번은 `Mentions`에 남는다.
두 사용자의 닉네임이 같더라도 서로 다른 ID를 유지한다. 이름이 아직 확인되지 않은 대상도 ID는 전달한다.
목록은 읽기 전용이며 태그가 없으면 빈 목록을 반환한다. `@이름`을 일반 텍스트로만 입력한 경우에는 태그 ID를 추측하지 않는다.
답장 메타데이터의 `src_mentions`는 원본 메시지의 태그이므로 제외하고 현재 메시지의 `mentions`만 전달한다.
첨부 JSON이 잘못된 경우 본문은 계속 수신하고, 해석할 수 없는 멘션 항목은 제외한다.

현재 `/조회`, `/좋아`, `/싫어`와 Google Sheets 사용자 DB는 기존 닉네임 기준을 사용한다.
이번 변경은 ID를 명령 코드에서 사용할 수 있도록 전달하는 범위이며, 기존 사용자 데이터의 ID 전환은 별도 작업이다.

## 1:1 암호 메시지

`/암호검증`, `/암호변경`을 처리하면 해당 작성자의 1:1 방을 5분간 감시한다.
우선 작성자 ID로 방을 찾고, 찾지 못하면 기존 동작에 맞춰 같은 닉네임의 방이 하나뿐인 경우 연결한다.
방이 여러 개라 모호하거나 키가 없으면 상태 표시에서 대기 사유를 알린다. 카카오톡에서 해당 1:1 방을 열어 키를 준비할 수 있다.
요청 시각 이후의 메시지만 기존 암호 처리에 전달하고, 메시지 본문은 채팅 로그 화면에 표시하지 않는다.

## 갱신 구조

| 파일 | 역할 |
| --- | --- |
| `KakaoTalkDecryptor.cs` | 검증된 키 추출·보관과 페이지 복호화 |
| `ChatDatabaseSnapshot.cs` | 일관된 암호화 바이트 수집, WAL 검증, 변경 페이지 반영 |
| `ChatSqlite.cs` | Windows의 `winsqlite3.dll`로 복호화 사본 조회 |
| `ChatDatabaseReceiver.cs` | 계정/방 목록, 사용자 정보, 신규 행 조회, 수신 큐와 작업 수명 |
| `Bot.cs` | 수신 메시지를 기존 명령·퀴즈·기여도·1:1 처리에 연결 |
| `Form1.cs` | DB 방 선택, 상태 표시, 시작·종료 |

최초에는 전체 DB를 복호화한다. 같은 WAL 세대에서는 이전 커밋 이후 프레임만 읽고 검증하며, 변경된 페이지만 복호화한다.
WAL이 초기화되거나 체크포인트 이후 세대가 바뀌면 DB를 다시 읽어 페이지 해시를 대조한다. 내용이 같은 페이지는 복호화를 생략한다.
초기 생성 및 전체 대조 시 SQLite `quick_check`를 수행한다. 갱신 실패 시 이전 사본을 복구하고 읽기 기준점은 진행시키지 않는다.

원본 DB·WAL·SHM은 쓰기와 삭제를 허용하는 공유 읽기로만 연다. 원본에 SQLite 잠금이나 배타 잠금을 걸지 않는다. SHM을 읽을 때도 잠금 바이트 120~127과 버퍼 선행 읽기를 피한다.
읽기 전후 원본 첫 페이지·길이·수정 시각, WAL 헤더·길이, 공유 인덱스 헤더와 체크포인트 위치를 비교한다. 전체 사본 생성 시에는 원본을 다시 읽어 해시도 대조한다. 변경되거나 검증에 실패하면 사본을 반영하지 않고 다음 주기에 재시도한다. 원본 DB 복사에 2초가 넘으면 중단한다.
WAL 헤더/프레임 체크섬과 공유 인덱스에 공개된 마지막 커밋을 확인한다. 활성 롤백 저널이 남아 있으면 카카오톡의 복구를 기다린다.

이 처리는 Windows SQLite WAL 인덱스 형식과 현재 `KakaoTalkDecryptor`의 4,096바이트 SQLCipher 형식을 전제로 한다. 잠금 없는 외부 파일 읽기이므로 SQLite 백업 API의 트랜잭션 보장과 같지는 않다. 변경이 잦으면 수신이 다음 안정된 읽기 시점까지 지연될 수 있다.
답변 전송은 기존 창 이름 기반 방식을 사용하므로 같은 제목의 방이 여러 개면 카카오톡에서 구분 가능한 제목을 지정해야 한다.
DB 수신 중에는 기존의 주기적인 채팅창 닫기/재시작 경로를 사용하지 않는다.

참고 규격: [SQLite WAL 인덱스와 잠금](https://www.sqlite.org/walformat.html), [WAL 파일 형식과 체크섬](https://www.sqlite.org/fileformat2.html#walformat), [Windows VFS 구현](https://github.com/sqlite/sqlite/blob/master/src/os_win.c).
메시지 타입은 실제 DB 구조와 [node-kakao 메시지 타입](https://github.com/storycraft/node-kakao/blob/stable/src/chat/chat-type.ts), [이벤트 타입](https://github.com/storycraft/node-kakao/blob/stable/src/chat/feed/feed-type.ts)을 대조했다.

## 검증

먼저 Debug 구성을 빌드한 뒤 저장소 루트에서 실행한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-KakaoTalkDecryptor.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-ChatDatabaseReceiver.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-BotCommandBridge.ps1
```

합성 SQLite/암호화 DB로 키 재사용, 커밋/미커밋 구분, WAL 세대 변경, 체크섬 오류, 원본 무변경, 카카오톡 쓰기 잠금과의 비간섭, 체크포인트 변경 거부, 취소, 커서 중복 제거와 메시지 분류를 검증한다.
명령 연결 테스트는 메시지 전송이나 Google Sheets 호출 없이 명령 큐·퀴즈·기여도·방 구분만 검증한다.

실제 로그인된 카카오톡의 읽기 전용 점검은 다음 명령으로 실행한다. 임시 폴더에 사본을 만들고 방 목록·채팅 조회·수신 준비·중단을 확인한 뒤 임시 사본을 삭제한다. 명령 실행과 메시지 전송은 하지 않는다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-ChatDatabaseReceiver.ps1 -LiveCatalog
```

## 2026-09-21 접근 충돌 방지 수정

특정 방에서 DB 오류 팝업 후 방을 다시 열면 정상화된 사례를 조사했다. 정확한 오류 문구는 없어 영구 손상이나 원인을 확정할 수 없다. 원본 SHM의 쓰기·체크포인트·복구 잠금을 직접 획득하던 동작이 카카오톡의 DB 작업을 방해할 수 있어 제거했다. 원본 복구·삭제·수정은 수행하지 않는다. 검증은 합성 DB로 수행한다.
