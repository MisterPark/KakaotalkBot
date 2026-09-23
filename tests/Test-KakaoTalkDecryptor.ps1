$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$testDirectory = Join-Path ([IO.Path]::GetTempPath()) ('KakaoDecryptorBuild-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testDirectory | Out-Null
try {
    $executable = Join-Path $testDirectory 'KakaoTalkDecryptorTests.exe'
    & $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll "/out:$executable" `
        (Join-Path $repository 'KakaotalkBot\KakaoTalkDecryptor.cs') `
        (Join-Path $PSScriptRoot 'KakaoTalkDecryptorTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
    & $executable
    if ($LASTEXITCODE -ne 0) { throw 'Decryptor tests failed.' }
}
finally {
    # This directory is created above with a unique name and contains only test build output.
    $resolved = [IO.Path]::GetFullPath($testDirectory)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Test cleanup path is outside the temporary directory.'
    }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
