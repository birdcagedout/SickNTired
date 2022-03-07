net user administrator /active:yes
taskkill /IM explorer.exe /F
DEL /A /Q "%localappdata%\IconCache*.db"
DEL /A /F /Q "%localappdata%\Microsoft\Windows\Explorer\iconcache*.*"
START explorer.exe
net user administrator /active:no

