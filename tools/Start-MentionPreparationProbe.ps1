param(
    [Parameter(Mandatory = $true)][string]$Room,
    [Parameter(Mandatory = $true)][long]$UserId,
    [Parameter(Mandatory = $true)][string]$CandidateNickname
)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repository 'artifacts\MentionPreparationProbe'
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $outputDirectory 'MentionPreparationProbe.exe'
& $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll "/out:$executable" `
    (Join-Path $repository 'KakaotalkBot\KakaoInputMentions.cs') `
    (Join-Path $repository 'KakaotalkBot\KakaoMentionSender.cs') `
    (Join-Path $PSScriptRoot 'MentionPreparationProbe.cs')
if ($LASTEXITCODE -ne 0) { throw '멘션 준비 진단 도구 컴파일 실패' }
# 준비된 시험 입력만 확인 후 정리하며, 전송 함수를 호출하지 않는다.
& $executable $Room $UserId $CandidateNickname
if ($LASTEXITCODE -ne 0) { throw '멘션 준비 진단 실패' }
