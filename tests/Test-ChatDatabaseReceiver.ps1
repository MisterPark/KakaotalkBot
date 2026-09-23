param([switch]$LiveCatalog)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$testDirectory = Join-Path ([IO.Path]::GetTempPath()) ('ChatReceiverBuild-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testDirectory | Out-Null
try {
    $executable = Join-Path $testDirectory 'ChatDatabaseReceiverTests.exe'
    $json = Join-Path $repository 'packages\Newtonsoft.Json.13.0.3\lib\net45\Newtonsoft.Json.dll'
    Copy-Item -LiteralPath $json -Destination $testDirectory
    $sources = @('Command.cs','KakaoTalkDecryptor.cs','ChatSqlite.cs','ChatDatabaseSnapshot.cs','ChatDatabaseReceiver.cs') | ForEach-Object { Join-Path $repository ('KakaotalkBot\' + $_) }
    & $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll "/r:$json" "/out:$executable" $sources (Join-Path $PSScriptRoot 'ChatDatabaseReceiverTests.cs')
    if ($LASTEXITCODE -ne 0) { throw '수신 테스트 컴파일 실패' }
    if ($LiveCatalog) { & $executable --live-catalog } else { & $executable }
    if ($LASTEXITCODE -ne 0) { throw '수신 테스트 실패' }
}
finally {
    $resolved = [IO.Path]::GetFullPath($testDirectory)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) { throw '임시 폴더 밖의 경로입니다.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
