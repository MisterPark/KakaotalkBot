$ErrorActionPreference = 'Stop'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) { throw 'Visual Studio Build Tools와 .NET Framework 4.8 개발 도구가 필요합니다.' }
$msbuild = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
if (-not $msbuild) { throw 'MSBuild를 찾지 못했습니다.' }
& $msbuild (Join-Path $PSScriptRoot 'OpenChatRoomCleaner.csproj') /t:Build /v:minimal /nologo
if ($LASTEXITCODE -ne 0) { throw '빌드 실패' }
Write-Output (Join-Path $PSScriptRoot 'bin\Debug\OpenChatRoomCleaner.exe')
