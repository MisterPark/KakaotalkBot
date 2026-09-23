$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$output = Join-Path $repository 'artifacts\StaInputWorkerTests'
New-Item -ItemType Directory -Force -Path $output | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $output 'Tests.exe'
& $compiler /nologo /r:System.Windows.Forms.dll "/out:$executable" (Join-Path $repository 'KakaotalkBot\StaInputWorker.cs') (Join-Path $PSScriptRoot 'StaInputWorkerTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'STA 입력 테스트 빌드 실패' }
& $executable
if ($LASTEXITCODE -ne 0) { throw 'STA 입력 테스트 실패' }
