@echo off
@REM get current time


echo Set the Path to point to the v2.0 and v2.0 Framework runtimes
PATH=%PATH%;%windir%\Microsoft.Net\Framework\v2.0.50727;


regasm /unregister "D:\BSCEFSuperTrumpCom\BSCEFSuperTrumpCom\bin\BSCEFSuperTrumpCom.dll"


echo Register the new COM+ Components to Component Services

regsvcs /exapp  /appname:BSCEFSuperTrumpCom "D:\BSCEFSuperTrumpCom\BSCEFSuperTrumpCom\bin\BSCEFSuperTrumpCom.dll"

echo ST COM+ Server Deploy Complete

echo Deploy complete
pause


