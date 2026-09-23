param([string]$NodePath)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
if (-not $NodePath) { $NodePath = (Get-Command node -ErrorAction Stop).Source }
& $NodePath --test (Join-Path $repository 'tools\LocoSender\test\sender.test.js')
if ($LASTEXITCODE -ne 0) { throw 'Node LOCO 테스트 실패' }
$outputDirectory = Join-Path $repository 'artifacts\LocoSenderTests'
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$json = Join-Path $repository 'packages\Newtonsoft.Json.13.0.3\lib\net45\Newtonsoft.Json.dll'
Copy-Item -LiteralPath $json -Destination $outputDirectory -Force
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$installation = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -property installationPath
if (-not $installation) { throw 'Visual Studio MSBuild를 찾지 못했습니다.' }
$compiler = Join-Path $installation 'MSBuild\Current\Bin\Roslyn\csc.exe'
$executable = Join-Path $outputDirectory 'LocoSenderBridgeTests.exe'
& $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll "/r:$json" "/out:$executable" (Join-Path $repository 'KakaotalkBot\LocoMentionSender.cs') (Join-Path $PSScriptRoot 'LocoSenderBridgeTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'C# LOCO 테스트 빌드 실패' }
& $executable $NodePath (Join-Path $repository 'tools\LocoSender\bridge.js')
if ($LASTEXITCODE -ne 0) { throw 'C# LOCO 테스트 실패' }
