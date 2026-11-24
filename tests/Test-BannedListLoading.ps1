$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$testDirectory = Join-Path ([IO.Path]::GetTempPath()) ('BannedListLoadingBuild-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testDirectory | Out-Null
try {
    $executable = Join-Path $testDirectory 'BannedListLoadingTests.exe'
    & $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll /r:System.Drawing.dll "/out:$executable" (Join-Path $repository 'KakaotalkBot\BannedMemberMemory.cs') (Join-Path $repository 'KakaotalkBot\BannedListLoading.cs') (Join-Path $PSScriptRoot 'BannedListLoadingTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Banned-list loading test compilation failed.' }
    & $executable
    if ($LASTEXITCODE -ne 0) { throw 'Banned-list loading tests failed.' }
}
finally {
    $resolved = [IO.Path]::GetFullPath($testDirectory)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Test directory is outside the temporary root.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
