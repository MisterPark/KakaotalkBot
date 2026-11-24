$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$output = Join-Path $root 'artifacts\GovernanceReminderTests'
New-Item -ItemType Directory -Force -Path $output | Out-Null
$json = Join-Path $root 'packages\Newtonsoft.Json.13.0.3\lib\net45\Newtonsoft.Json.dll'
Copy-Item -LiteralPath $json -Destination $output
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /r:System.Core.dll "/r:$json" "/out:$output\GovernanceReminderTests.exe" `
    (Join-Path $root 'KakaotalkBot\GovernanceStore.cs') (Join-Path $PSScriptRoot 'GovernanceReminderTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Election reminder tests compilation failed' }
& "$output\GovernanceReminderTests.exe"
if ($LASTEXITCODE -ne 0) { throw 'Election reminder tests failed' }
