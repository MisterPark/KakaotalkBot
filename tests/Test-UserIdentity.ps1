param([string]$BinaryDirectory)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
if (-not $BinaryDirectory) { $BinaryDirectory = Join-Path $repository 'artifacts\IdentityBuild' }
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $BinaryDirectory 'UserIdentityTests.exe'
& $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll "/r:$BinaryDirectory\KakaotalkBot.exe" "/out:$executable" (Join-Path $PSScriptRoot 'UserIdentityTests.cs')
if ($LASTEXITCODE -ne 0) { throw '사용자 ID 테스트 컴파일 실패' }
& $executable
if ($LASTEXITCODE -ne 0) { throw '사용자 ID 테스트 실패' }
