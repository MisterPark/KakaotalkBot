$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repository 'artifacts\MentionDeliveryProbe'
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$compiler = 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\Roslyn\csc.exe'
$executable = Join-Path $outputDirectory 'MentionDeliveryProbe.exe'
$sources = @('tools\MentionDeliveryProbe.cs','KakaotalkBot\KakaoInputMentions.cs','KakaotalkBot\KakaoTalkDecryptor.cs','KakaotalkBot\ChatDatabaseSnapshot.cs','KakaotalkBot\ChatSqlite.cs') | ForEach-Object { Join-Path $repository $_ }
& $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll /r:System.Web.Extensions.dll "/out:$executable" $sources
if ($LASTEXITCODE -ne 0) { throw '멘션 송신 검증 도구 빌드 실패' }
Write-Output $executable
