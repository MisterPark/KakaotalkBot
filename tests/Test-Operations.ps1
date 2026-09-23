$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$build = Join-Path $root 'artifacts\OperationsTests'
New-Item -ItemType Directory -Path $build -Force | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$json = Join-Path $root 'packages\Newtonsoft.Json.13.0.3\lib\net45\Newtonsoft.Json.dll'
Copy-Item -LiteralPath $json -Destination $build
$sources = @('User.cs','UserActivity.cs','OperationsStore.cs','Command.cs','KakaoTalkDecryptor.cs','ChatSqlite.cs','ChatDatabaseSnapshot.cs','ChatDatabaseReceiver.cs') | ForEach-Object { Join-Path $root ('KakaotalkBot\' + $_) }
& $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll "/r:$json" "/out:$build\OperationsTests.exe" $sources (Join-Path $PSScriptRoot 'OperationsTests.cs')
if ($LASTEXITCODE -ne 0) { throw '운영 기능 테스트 컴파일 실패' }
& "$build\OperationsTests.exe"
if ($LASTEXITCODE -ne 0) { throw '운영 기능 테스트 실패' }
