$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repository 'artifacts\MentionSenderTests'
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $outputDirectory 'MentionSenderTests.exe'
& $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll "/out:$executable" `
    (Join-Path $repository 'KakaotalkBot\KakaoInputMentions.cs') `
    (Join-Path $repository 'KakaotalkBot\KakaoMentionSender.cs') `
    (Join-Path $PSScriptRoot 'MentionSenderTests.cs')
if ($LASTEXITCODE -ne 0) { throw '멘션 송신 테스트 컴파일 실패' }
& $executable
if ($LASTEXITCODE -ne 0) { throw '멘션 송신 테스트 실패' }
