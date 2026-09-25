param([string]$BinaryDirectory)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
if (-not $BinaryDirectory) { $BinaryDirectory = Join-Path $repository 'KakaotalkBot\bin\Debug' }
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $BinaryDirectory 'BotIdentityTests.exe'
& $compiler /nologo /r:System.Core.dll "/r:$BinaryDirectory\Newtonsoft.Json.dll" "/out:$executable" (Join-Path $PSScriptRoot 'BotIdentityTests.cs') (Join-Path $repository 'KakaotalkBot\BotIdentity.cs')
if ($LASTEXITCODE -ne 0) { throw '고정 봇 ID 테스트 컴파일 실패' }
& $executable
if ($LASTEXITCODE -ne 0) { throw '고정 봇 ID 테스트 실패' }
