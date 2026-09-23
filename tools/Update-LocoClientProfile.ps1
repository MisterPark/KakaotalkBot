param([string]$DestinationDirectory)
$ErrorActionPreference = 'Stop'
if (-not $DestinationDirectory) { $DestinationDirectory = Join-Path $PSScriptRoot 'LocoSender' }
$candidates = @()
$candidates += Get-Process KakaoTalk -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path -Unique
$candidates += Join-Path $env:ProgramFiles 'Kakao\KakaoTalk\KakaoTalk.exe'
if (${env:ProgramFiles(x86)}) { $candidates += Join-Path ${env:ProgramFiles(x86)} 'Kakao\KakaoTalk\KakaoTalk.exe' }
$installedPath = $candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
if (-not $installedPath) { throw '설치된 KakaoTalk.exe를 찾지 못했습니다. 카카오톡을 실행한 뒤 다시 시도해 주세요.' }
$version = (Get-Item -LiteralPath $installedPath).VersionInfo.ProductVersion
if ($version -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,6}$') { throw '카카오톡 제품 버전 형식을 인식하지 못했습니다.' }
New-Item -ItemType Directory -Path $DestinationDirectory -Force | Out-Null
$profile = @{ appVersion = $version } | ConvertTo-Json
[IO.File]::WriteAllText((Join-Path $DestinationDirectory 'client-profile.json'), ($profile -replace '\r?\n', "`r`n") + "`r`n", (New-Object Text.UTF8Encoding($false)))
Write-Output ('설치된 카카오톡 버전 적용: ' + $version)
