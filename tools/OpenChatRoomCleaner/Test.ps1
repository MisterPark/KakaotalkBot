param(
    [Parameter(Mandatory=$true)][string]$ParticipantScreenshot,
    [Parameter(Mandatory=$true)][string]$RemoveMenuScreenshot,
    [Parameter(Mandatory=$true)][string]$ConfirmScreenshot,
    [string]$SingleButtonScreenshot,
    [string]$ThreeButtonScreenshot,
    [string]$LargeListScreenshot,
    [string]$OperatorsScreenshot
)
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'Build.ps1')
$executable = Join-Path $PSScriptRoot 'bin\Debug\OpenChatRoomCleaner.exe'
if ([bool]$SingleButtonScreenshot -ne [bool]$ThreeButtonScreenshot) { throw '프로필 버튼 캡처 2개를 함께 전달하세요.' }
$profileScreenshots = @()
if ($SingleButtonScreenshot) { $profileScreenshots = @($SingleButtonScreenshot,$ThreeButtonScreenshot) }
$arguments = @('--self-test') + @($ParticipantScreenshot,$RemoveMenuScreenshot,$ConfirmScreenshot | ForEach-Object { '"' + (Resolve-Path -LiteralPath $_).Path + '"' })
foreach ($profileScreenshot in $profileScreenshots) { $arguments += '"' + (Resolve-Path -LiteralPath $profileScreenshot).Path + '"' }
if ($LargeListScreenshot) {
    if (-not $SingleButtonScreenshot) { throw '큰 목록 시험에는 프로필 캡처 2개도 함께 전달하세요.' }
    $arguments += '"' + (Resolve-Path -LiteralPath $LargeListScreenshot).Path + '"'
}
if ($OperatorsScreenshot) {
    if (-not $LargeListScreenshot) { throw '운영진 목록 시험에는 큰 목록 캡처도 함께 전달하세요.' }
    $arguments += '"' + (Resolve-Path -LiteralPath $OperatorsScreenshot).Path + '"'
}
$process = Start-Process -FilePath $executable -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
Get-Content -LiteralPath (Join-Path $PSScriptRoot 'bin\Debug\test-result.txt')
if ($process.ExitCode -ne 0) { throw '이미지/상태 전이 테스트 실패' }
