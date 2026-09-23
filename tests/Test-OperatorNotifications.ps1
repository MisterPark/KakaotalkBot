param([string]$BinaryDirectory)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
if (-not $BinaryDirectory) { $BinaryDirectory = Join-Path $root 'artifacts\IdentityBuild' }
$build = Join-Path ([IO.Path]::GetTempPath()) ('OperatorNotices-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $build | Out-Null
try {
    Copy-Item -LiteralPath (Join-Path $BinaryDirectory 'KakaotalkBot.exe') -Destination $build
    Get-ChildItem -LiteralPath $BinaryDirectory -Filter '*.dll' | Copy-Item -Destination $build
    $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
    & $compiler /nologo /target:exe /r:System.Core.dll "/r:$build\KakaotalkBot.exe" "/out:$build\OperatorNotificationTests.exe" (Join-Path $PSScriptRoot 'OperatorNotificationTests.cs')
    if ($LASTEXITCODE -ne 0) { throw '알림 테스트 컴파일 실패' }
    & "$build\OperatorNotificationTests.exe" $build
    if ($LASTEXITCODE -ne 0) { throw '알림 테스트 실패' }
} finally {
    $resolved = [IO.Path]::GetFullPath($build)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) { throw '임시 경로 검증 실패' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
