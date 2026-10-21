$ErrorActionPreference = 'Stop'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$msbuild = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe | Select-Object -First 1
if (-not $msbuild) { throw 'Visual Studio MSBuild를 찾지 못했습니다.' }
& $msbuild (Join-Path $PSScriptRoot 'DirectChatRelay.csproj') /t:Build /v:quiet /nologo /p:CopyRetryCount=0
if ($LASTEXITCODE -ne 0) { throw '빌드 실패. 실행 중인 전달 프로그램을 종료해 주세요.' }
