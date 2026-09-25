$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$output = Join-Path $root 'artifacts\DailyQuestTests'
New-Item -ItemType Directory -Force -Path $output | Out-Null
Copy-Item -LiteralPath (Join-Path $root 'KakaotalkBot\bin\Debug\Newtonsoft.Json.dll') -Destination $output -Force
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$exe = Join-Path $output 'Tests.exe'
& $compiler /nologo /r:System.Core.dll "/r:$output\Newtonsoft.Json.dll" "/out:$exe" (Join-Path $root 'KakaotalkBot\DailyQuests.cs') (Join-Path $PSScriptRoot 'DailyQuestTests.cs')
if ($LASTEXITCODE -ne 0) { throw '퀘스트 테스트 빌드 실패' }
& $exe
if ($LASTEXITCODE -ne 0) { throw '퀘스트 테스트 실패' }
