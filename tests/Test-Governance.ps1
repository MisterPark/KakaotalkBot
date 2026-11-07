$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$output = Join-Path $root 'artifacts\GovernanceTests'
New-Item -ItemType Directory -Force -Path $output | Out-Null
$json = Join-Path $root 'packages\Newtonsoft.Json.13.0.3\lib\net45\Newtonsoft.Json.dll'
Copy-Item -LiteralPath $json -Destination $output
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /r:System.Core.dll "/r:$json" "/out:$output\GovernanceTests.exe" (Join-Path $root 'KakaotalkBot\GovernanceStore.cs') (Join-Path $PSScriptRoot 'GovernanceTests.cs')
if ($LASTEXITCODE -ne 0) { throw '선거 테스트 빌드 실패' }
& "$output\GovernanceTests.exe"
if ($LASTEXITCODE -ne 0) { throw '선거 테스트 실패' }
