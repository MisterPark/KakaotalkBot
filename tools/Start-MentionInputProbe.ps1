param([string]$Room = '보룸봇 테스트방', [string]$UserId, [string]$Nickname, [switch]$BuildOnly)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repository 'artifacts\MentionInputProbe'
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $outputDirectory 'MentionInputProbe.exe'
& $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll /r:Accessibility.dll "/out:$executable" (Join-Path $PSScriptRoot 'MentionInputProbe.cs') (Join-Path $PSScriptRoot 'MentionMemoryProbe.cs')
if ($LASTEXITCODE -ne 0) { throw '멘션 입력창 진단 도구 빌드 실패' }
Write-Output $executable
if (-not $BuildOnly) {
    if ($UserId) { & $executable $Room $UserId $Nickname }
    else { & $executable $Room }
    if ($LASTEXITCODE -ne 0) { throw '멘션 입력창 진단 실패' }
}
