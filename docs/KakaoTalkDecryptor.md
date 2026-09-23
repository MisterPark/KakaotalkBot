# KakaoTalkDecryptor 단계별 사용법

대상 DB를 지정하고 키를 한 번 추출한 다음, 같은 `DecryptionKey` 객체로 반복 복호화한다.
키와 HMAC 키는 객체 안에서 재사용하며 파일에 저장하지 않는다.

## 공개 API

| 단계 | 함수 | 반환값 / 역할 |
| --- | --- | --- |
| 계정 폴더 목록 | `FindChatDataDirectories()` | 로컬에 존재하는 계정들의 `chat_data` 경로. 현재 로그인 계정을 식별하지는 않는다. |
| 프로세스 목록 | `FindKakaoTalkProcessIds()` | 카카오톡 PID 목록. 계정 폴더와의 연결은 호출자가 결정한다. |
| DB 파일 목록 | `FindChatDatabaseFiles(chatDataDirectory)` | 정렬된 `chatLogs_*.edb` 절대 경로 목록 |
| DB 정보 읽기 | `ReadDatabaseInfo(sourcePath)` | `DatabaseInfo`: 절대 경로, 읽은 시점의 페이지 수, salt |
| 키 추출 | `DiscoverKey(processId, database)` | 검증된 `DecryptionKey`. 키를 찾지 못하면 `null` |
| 여러 키 추출 | `DiscoverKeys(processId, databases)` | 한 번의 메모리 탐색으로 찾은 키 사전. 키는 DB 절대 경로이며, 발견되지 않은 DB는 포함되지 않는다. |
| 보유한 키 사용 | `CreateKey(database, rawKey)` | 32바이트 원시 키를 복사·검증하여 재사용 객체 생성 |
| DB 복호화 | `DecryptDatabase(sourcePath, outputPath, key, overwrite)` | 전체 DB와 커밋된 WAL을 반영하여 SQLite 파일 생성. 반환값은 원본 DB 페이지 수이며 WAL 반영 후 크기와 다를 수 있다. |
| 페이지 복호화 | `DecryptPage(encryptedPage, pageNumber, key)` | 암호화된 4,096바이트 페이지 하나를 HMAC 검증 후 복호화. 페이지 번호는 1부터 시작한다. |
| 키 해제 | `key.Dispose()` | 객체가 보관한 원시 키·HMAC 키·salt 배열을 지우고 이후 사용을 차단 |

`DatabaseInfo.Salt`는 복사본이다. `CreateKey()`는 호출자의 원시 키 배열을 변경하거나 소유하지 않는다.
`DiscoverKeys()`로 반환받은 키는 각각 해제해야 한다. 파일 경로가 달라도 같은 salt를 가진 DB를 모두 처리한다.

## 한 채팅방의 키를 유지하며 갱신하기

아래 예시의 `sourcePath`는 대상 방의 `chatLogs_<chatId>.edb` 전체 경로이고,
`processId`는 실행 중인 카카오톡의 PID다. 계정 폴더나 PID가 여러 개면 첫 항목을 자동 선택하지 말고 대상에 맞게 지정한다.

```csharp
using System;
using System.IO;
using KakaotalkBot;

public sealed class ChatDatabaseSession : IDisposable
{
    private readonly KakaoTalkDecryptor decryptor = new KakaoTalkDecryptor();
    private readonly KakaoTalkDecryptor.DecryptionKey key;
    private readonly string sourcePath;

    public string OutputPath { get; private set; }

    public ChatDatabaseSession(string sourcePath, int processId)
    {
        // 1. DB 정보 읽기
        var database = decryptor.ReadDatabaseInfo(sourcePath);
        this.sourcePath = database.SourcePath;
        OutputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "DecryptedChat", Path.GetFileNameWithoutExtension(sourcePath) + ".sqlite");

        // 2. 최초 한 번만 키 추출
        key = decryptor.DiscoverKey(processId, database);
        if (key == null)
            throw new InvalidOperationException("카카오톡에서 대상 채팅방을 연 뒤 다시 시도하세요.");
    }

    public void Refresh()
    {
        // 3. 반복 호출해도 프로세스 탐색 및 HMAC 키 파생을 다시 하지 않는다.
        // 출력 폴더는 DecryptDatabase가 자동으로 생성한다.
        decryptor.DecryptDatabase(sourcePath, OutputPath, key, overwrite: true);
    }

    public void Dispose()
    {
        // 4. 채팅방 감시 종료 시 호출
        key.Dispose();
    }
}
```

세션 객체는 방을 감시하는 동안 필드 등에 보관하고, 갱신할 때마다 `Refresh()`를 호출한다.
매 갱신마다 생성하면 메모리 탐색을 반복하게 된다. 키 탐색과 전체 DB 복호화는 동기 함수이므로
WinForms에서는 백그라운드 작업에서 호출하고, 컨트롤 갱신은 UI 스레드에서 수행한다.
출력 파일을 교체하기 전에는 해당 SQLite 파일을 읽는 연결을 닫아야 한다.

## 키 수명과 오류 처리

- 프로세스 메모리에 대상 DB의 키가 있어야 `DiscoverKey()`가 성공한다. 키 탐색은 Windows x64 프로세스에서 실행해야 한다.
- 추출한 키로 복호화할 때는 카카오톡 PID나 프로세스 핸들을 사용하지 않는다.
- 같은 키 객체를 사용하는 작업과 `Dispose()`는 직렬화된다. 해제 후 사용하면 `ObjectDisposedException`이 발생한다.
- DB의 salt가 바뀌면 `CryptographicException`이 발생한다. 기존 키를 해제하고 `ReadDatabaseInfo()`부터 다시 수행한다.
- HMAC 오류는 잘못된 키나 손상되었거나 읽는 중 변경된 페이지에서도 발생할 수 있다.
- 실패한 복호화는 임시 파일을 정리하고 기존 출력 파일을 유지한다. 원본 DB 및 해당 `-wal`, `-shm` 경로를 출력으로 지정할 수 없다.

## 기존 API와 현재 범위

`DecryptAvailable()`과 `DecryptFile(..., byte[] key, ...)`는 기존 형태로 계속 호출할 수 있다.
`DecryptAvailable()`은 일괄 작업을 끝낸 뒤 내부 키를 해제한다. 반복 갱신에는 단계별 API를 사용한다.

`DecryptDatabase()`는 호출할 때마다 전체 DB를 복호화하는 기존 API다.
실제 봇 수신에서는 `ChatDatabaseSnapshot`이 `DecryptPage()`를 이용해 변경 페이지를 반영하고, 원본을 잠그지 않고 읽기 전후 상태·WAL 체크섬·커밋 경계를 확인한다.
`ChatDatabaseReceiver`가 신규 행을 조회하여 기존 명령 처리에 전달한다. 자세한 동작과 실행 방법은 `DB채팅수신.md`를 참고한다.

## 검증 실행

Windows의 .NET Framework 4.8 환경에서 저장소 루트를 기준으로 실행한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-KakaoTalkDecryptor.ps1
```

테스트는 임시 폴더의 합성 암호화 DB와 테스트 프로세스 자체의 메모리를 사용한다.
키 추출·재사용·해제, salt 변경, HMAC 오류, 커밋된 WAL 반영, 기존 API 호환성과 실패 시 출력 보존을 검증한다.
