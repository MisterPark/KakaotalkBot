$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repository 'artifacts\ChatResponseTests'
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $outputDirectory 'ChatResponseTests.exe'
& $compiler /nologo /target:exe /r:System.Core.dll "/out:$executable" `
    (Join-Path $repository 'KakaotalkBot\ChatProfanityFilter.cs') `
    (Join-Path $repository 'KakaotalkBot\CommandHelp.cs') `
    (Join-Path $PSScriptRoot 'ChatResponseTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Chat response tests compilation failed' }
& $executable
if ($LASTEXITCODE -ne 0) { throw 'Chat response tests failed' }
