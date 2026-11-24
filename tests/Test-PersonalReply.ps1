param([string]$BinaryDirectory)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
if (-not $BinaryDirectory) { $BinaryDirectory = Join-Path $repository 'artifacts\PersonalAiBuild' }
$build = (Resolve-Path -LiteralPath $BinaryDirectory).Path
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $build 'PersonalReplyTests.exe'
& $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll /r:System.Net.Http.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll `
    "/r:$build\Newtonsoft.Json.dll" "/r:$build\KakaotalkBot.exe" "/out:$executable" (Join-Path $PSScriptRoot 'PersonalReplyTests.cs') (Join-Path $PSScriptRoot 'PersonalActionTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Personal AI tests compilation failed' }
& $executable
if ($LASTEXITCODE -ne 0) { throw 'Personal AI tests failed' }
