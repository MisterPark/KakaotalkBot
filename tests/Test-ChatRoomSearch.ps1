$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repository 'artifacts\ChatRoomSearchTests'
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $outputDirectory 'ChatRoomSearchTests.exe'
& $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll /r:System.Windows.Forms.dll "/out:$executable" `
    (Join-Path $repository 'KakaotalkBot\ChatRoomSearch.cs') `
    (Join-Path $PSScriptRoot 'ChatRoomSearchTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Chat room search tests compilation failed' }
& $executable
if ($LASTEXITCODE -ne 0) { throw 'Chat room search tests failed' }
