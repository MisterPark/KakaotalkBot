$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$build = Join-Path $root 'artifacts\PixelDetectorTests'
New-Item -ItemType Directory -Path $build -Force | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /target:exe /r:System.Core.dll /r:System.Drawing.dll "/out:$build\PixelDetectorTests.exe" (Join-Path $root 'KakaotalkBot\ScreenPixelDetector.cs') (Join-Path $PSScriptRoot 'PixelDetectorTests.cs')
if ($LASTEXITCODE -ne 0) { throw '픽셀 감지 테스트 컴파일 실패' }
& "$build\PixelDetectorTests.exe"
if ($LASTEXITCODE -ne 0) { throw '픽셀 감지 테스트 실패' }
