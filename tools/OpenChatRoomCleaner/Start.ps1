$ErrorActionPreference = 'Stop'
# 항상 현재 소스에서 같은 경로에 빌드한 실행 파일을 엽니다.
& (Join-Path $PSScriptRoot 'Build.ps1')
$executable = Join-Path $PSScriptRoot 'bin\Debug\OpenChatRoomCleaner.exe'
Start-Process -FilePath $executable -WorkingDirectory (Split-Path -Parent $executable) -WindowStyle Normal
