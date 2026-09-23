param([switch]$BuildOnly)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repository 'artifacts\MentionClipboardProbe'
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $outputDirectory 'MentionClipboardProbe.exe'
& $compiler /nologo /target:winexe /platform:x64 /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll "/out:$executable" (Join-Path $PSScriptRoot 'MentionClipboardProbe.cs')
if ($LASTEXITCODE -ne 0) { throw '멘션 진단 도구 빌드 실패' }
Write-Output $executable
if (-not $BuildOnly) {
    # 사용자가 복사 상태별로 버튼을 누르는 대화형 진단 창입니다.
    & $executable
}
