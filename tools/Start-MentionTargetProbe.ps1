param([string]$Room = '보룸봇 테스트방', [switch]$BuildOnly)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$output = Join-Path $repository 'artifacts\MentionTargetProbe'
New-Item -ItemType Directory -Path $output -Force | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $output 'MentionTargetProbe.exe'
& $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll "/out:$executable" (Join-Path $PSScriptRoot 'MentionTargetProbe.cs') (Join-Path $repository 'KakaotalkBot\KakaoInputMentions.cs')
if ($LASTEXITCODE -ne 0) { throw '멘션 대상 검사 도구 빌드 실패' }
Write-Output $executable
if (-not $BuildOnly) {
    & $executable $Room
    if ($LASTEXITCODE -ne 0) { throw '멘션 대상 검사 실패' }
}
