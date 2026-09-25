$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$build = Join-Path $root 'artifacts\NewsTests'
New-Item -ItemType Directory -Path $build -Force | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$html = Join-Path $root 'packages\HtmlAgilityPack.1.12.2\lib\Net40\HtmlAgilityPack.dll'
Copy-Item -LiteralPath $html -Destination $build
& $compiler /nologo /target:exe /r:System.Core.dll /r:System.Net.Http.dll "/r:$html" "/out:$build\NewsTests.exe" (Join-Path $root 'KakaotalkBot\News.cs') (Join-Path $root 'KakaotalkBot\Article.cs') (Join-Path $PSScriptRoot 'NewsTests.cs')
if ($LASTEXITCODE -ne 0) { throw '뉴스 테스트 컴파일 실패' }
& "$build\NewsTests.exe"
if ($LASTEXITCODE -ne 0) { throw '뉴스 테스트 실패' }
