@echo off
:: Y�u c?u quy?n admin
:: N?u kh�ng ch?y v?i quy?n admin th� s? t? n�ng quy?n
:: (nh? d�ng n�y m� kh�ng c?n click chu?t ph?i)

:: Ki?m tra quy?n admin
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Yeu cau quyen quan tri... Dang khoi dong lai voi quyen admin.
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Ch?y PowerShell v?i l?nh c?n thi?t
powershell -NoExit -Command ^
  "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; irm https://get.activated.win | iex"
