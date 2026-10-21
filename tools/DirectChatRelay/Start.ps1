& (Join-Path $PSScriptRoot 'Build.ps1')
Start-Process -FilePath (Join-Path $PSScriptRoot 'bin\Debug\DirectChatRelay.exe')
