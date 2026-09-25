$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$build = Join-Path $root 'artifacts\QuizSyncTests'
New-Item -ItemType Directory -Path $build -Force | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$json = Join-Path $root 'packages\Newtonsoft.Json.13.0.3\lib\net45\Newtonsoft.Json.dll'
Copy-Item -LiteralPath $json -Destination $build
& $compiler /nologo /target:exe /r:System.Core.dll "/r:$json" "/out:$build\QuizSyncTests.exe" (Join-Path $root 'KakaotalkBot\Quiz.cs') (Join-Path $PSScriptRoot 'QuizSyncTests.cs')
if ($LASTEXITCODE -ne 0) { throw '퀴즈 동기화 테스트 컴파일 실패' }
& "$build\QuizSyncTests.exe"
if ($LASTEXITCODE -ne 0) { throw '퀴즈 동기화 테스트 실패' }
