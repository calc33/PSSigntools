@ECHO OFF
REM バッチファイルと同名の.ps1を実行する
PowerShell -File %~dpn0.ps1 %*
