param([switch]$BuildOnly, [string]$NodePath)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
if (-not $NodePath) { $NodePath = (Get-Command node -ErrorAction Stop).Source }
$source = Join-Path $PSScriptRoot 'LocoSender'
if (-not (Test-Path -LiteralPath (Join-Path $source 'node_modules\node-kakao\dist\index.js'))) {
    throw 'tools\LocoSender에서 npm ci --ignore-scripts를 먼저 실행해 주세요.'
}
$outputDirectory = Join-Path $repository 'artifacts\LocoMentionProbe'
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$destination = Join-Path $outputDirectory 'LocoSender'
New-Item -ItemType Directory -Path $destination -Force | Out-Null
foreach ($file in @('bridge.js', 'message.js', 'sender.js', 'profile.js', 'package.json', 'package-lock.json', 'node_modules')) {
    Copy-Item -LiteralPath (Join-Path $source $file) -Destination $destination -Recurse -Force
}
& (Join-Path $PSScriptRoot 'Update-LocoClientProfile.ps1') -DestinationDirectory $destination
$json = Join-Path $repository 'packages\Newtonsoft.Json.13.0.3\lib\net45\Newtonsoft.Json.dll'
Copy-Item -LiteralPath $json -Destination $outputDirectory -Force
# VS의 컴파일러는 프로젝트와 동일한 C# 문법을 지원합니다.
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$installation = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -property installationPath
if (-not $installation) { throw 'Visual Studio MSBuild를 찾지 못했습니다.' }
$compiler = Join-Path $installation 'MSBuild\Current\Bin\Roslyn\csc.exe'
$executable = Join-Path $outputDirectory 'LocoMentionProbe.exe'
& $compiler /nologo /target:winexe /platform:x64 /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll "/r:$json" "/out:$executable" (Join-Path $repository 'KakaotalkBot\LocoMentionSender.cs') (Join-Path $PSScriptRoot 'LocoMentionProbe.cs')
if ($LASTEXITCODE -ne 0) { throw 'LOCO 진단 창 빌드 실패' }
$runtime = @{ nodePath = [IO.Path]::GetFullPath($NodePath) } | ConvertTo-Json
[IO.File]::WriteAllText((Join-Path $outputDirectory 'runtime.json'), $runtime + "`r`n", (New-Object Text.UTF8Encoding($false)))
Write-Output $executable
if (-not $BuildOnly) { Start-Process -FilePath $executable -WorkingDirectory $outputDirectory }
