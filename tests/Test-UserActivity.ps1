param()
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$build = Join-Path $root 'artifacts\ActivityTests'
New-Item -ItemType Directory -Path $build -Force | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /target:exe /r:System.Core.dll "/out:$build\UserActivityTests.exe" (Join-Path $root 'KakaotalkBot\User.cs') (Join-Path $root 'KakaotalkBot\UserActivity.cs') (Join-Path $PSScriptRoot 'UserActivityTests.cs')
if ($LASTEXITCODE -ne 0) { throw '활동 테스트 컴파일 실패' }
& "$build\UserActivityTests.exe"
if ($LASTEXITCODE -ne 0) { throw '활동 테스트 실패' }
