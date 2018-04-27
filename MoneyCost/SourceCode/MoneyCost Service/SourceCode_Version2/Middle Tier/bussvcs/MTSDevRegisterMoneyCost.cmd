@echo off
@REM get current time


echo Set the Path to point to the v2.0 and v2.0 Framework runtimes
PATH=%PATH%;%windir%\Microsoft.Net\Framework\v2.0.50727;

echo Unregister the old COM+ Components from the GAC
"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\gacutil"  /u BSMoneyCostBL
"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\gacutil"  /u BSMoneyCostDL
"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\gacutil"  /u BSMoneyCostAuto

								  
"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\regasm" /unregister "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSDBAdapter.dll"

"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\regasm" /unregister "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostAuto.dll"

"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\regasm" /unregister "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostBL.dll"

"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\regasm" /unregister "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostDL.dll"

"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\regasm" /unregister "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostEntity.dll"

"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\regasm" /unregister "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\Interop.BSLDAP.dll"


echo Register the new COM+ Components to the GAC
"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\gacutil"  /i "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSDBAdapter.dll"

"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\gacutil"  /i "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostAuto.dll"

"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\gacutil"  /i "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostBL.dll"

"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\gacutil"  /i "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostDL.dll"

"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\gacutil"  /i "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostEntity.dll"

"C:\WINNT\Microsoft.NET\Framework\v2.0.50727\gacutil"  /i "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\Interop.BSLDAP.dll"

echo Register the new COM+ Components to Component Services



regsvcs /exapp  /appname:MoneyCost "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostDL.dll"

regsvcs /exapp  /appname:MoneyCost "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostBL.dll"

regsvcs /exapp  /appname:MoneyCost "D:\MoneyCostSite\SourceCode_Version2\Middle Tier\bussvcs\Bin\BSMoneyCostAuto.dll"


echo MoneyCost COM+ Server Deploy Complete

echo Deploy complete
pause


