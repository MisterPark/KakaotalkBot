param([string]$BinaryDirectory)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
if (-not $BinaryDirectory) { $BinaryDirectory = Join-Path $repository 'KakaotalkBot\bin\Debug' }
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$testDirectory = Join-Path ([IO.Path]::GetTempPath()) ('ModerationMenuBuild-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testDirectory | Out-Null
try {
    Copy-Item -LiteralPath (Join-Path $BinaryDirectory 'KakaotalkBot.exe') -Destination $testDirectory
    Get-ChildItem -LiteralPath $BinaryDirectory -Filter '*.dll' | Copy-Item -Destination $testDirectory
    $assets = Join-Path $testDirectory 'ModerationAssets'
    New-Item -ItemType Directory -Path $assets | Out-Null
    Copy-Item -LiteralPath (Join-Path $BinaryDirectory 'ModerationAssets\openchat-management-icon.png') -Destination $assets
    Copy-Item -LiteralPath (Join-Path $BinaryDirectory 'ModerationAssets\banned-member-management-button.png') -Destination $assets
    $executable = Join-Path $testDirectory 'ModerationMenuTests.exe'
    & $compiler /nologo /target:exe /platform:x64 /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll "/r:$testDirectory\KakaotalkBot.exe" "/out:$executable" (Join-Path $PSScriptRoot 'ModerationMenuTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Menu test compilation failed.' }
    & $executable
    if ($LASTEXITCODE -ne 0) { throw 'Menu tests failed.' }
}
finally {
    $resolved = [IO.Path]::GetFullPath($testDirectory)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Test directory is outside the temporary root.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
