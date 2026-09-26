param([string]$BinaryDirectory)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
if (-not $BinaryDirectory) { $BinaryDirectory = Join-Path $repository 'KakaotalkBot\bin\Debug' }
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$testDirectory = Join-Path ([IO.Path]::GetTempPath()) ('VoiceWorkerBuild-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testDirectory | Out-Null
try {
    Copy-Item -LiteralPath (Join-Path $binaryDirectory 'KakaotalkBot.exe') -Destination $testDirectory
    Get-ChildItem -LiteralPath $binaryDirectory -Filter '*.dll' | Copy-Item -Destination $testDirectory
    $executable = Join-Path $testDirectory 'VoiceRoomWorkerTests.exe'
    & $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll "/r:$testDirectory\KakaotalkBot.exe" "/out:$executable" (Join-Path $PSScriptRoot 'VoiceRoomWorkerTests.cs')
    if ($LASTEXITCODE -ne 0) { throw '보이스룸 작업자 테스트 컴파일 실패' }
    & $executable
    if ($LASTEXITCODE -ne 0) { throw '보이스룸 작업자 테스트 실패' }
}
finally {
    $resolved = [IO.Path]::GetFullPath($testDirectory)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) { throw '임시 폴더 밖의 경로입니다.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
