<%@ Language="VBScript" TRANSACTION ="NOT_SUPPORTED"%>
<% Option Explicit %>
<%
'This function sends the error email to developer
Sub SendEmail()
Dim lobjNewMail
Dim lobjRegistry
Dim lstrTo
Dim lstrSubject
Dim lstrFrom
Dim lstrDomain
Dim lstrBody
Dim lstrSendErrorEmail

	Set lobjNewMail = Server.CreateObject("CDONTS.NewMail")
	set lobjRegistry = Server.CreateObject("WScript.Shell.1")

    lstrTo = lobjRegistry.RegRead("HKEY_LOCAL_MACHINE\Software\FacilitySettings\MoneyCost\Email\Developer")
    lstrSendErrorEmail = lobjRegistry.RegRead("HKEY_LOCAL_MACHINE\Software\FacilitySettings\MoneyCost\Debug\SendErrorEmail")
    lstrDomain = lobjRegistry.RegRead("HKEY_LOCAL_MACHINE\Software\FacilitySettings\MoneyCost\URLs\DomainName")
    lstrFrom = lobjRegistry.RegRead("HKEY_LOCAL_MACHINE\Software\FacilitySettings\MoneyCost\Email\AdminEmailID")

	If lstrSendErrorEmail = "1" Then
		'if reporting error by email is ON, send the error email to developer
		lobjNewMail.From = lstrFrom
		lobjNewMail.To   = lstrTo
		lstrBody = "Error details are as follows:"  & vbCrLf
		lstrBody = lstrBody & "Error Number : " & mstrErrorNum & vbCrLf
		lstrBody = lstrBody & "Error Source : " & mstrErrorSource & vbCrLf
		lstrBody = lstrBody & "Error Description : " & mstrErrorDesc & vbCrLf
		lstrBody = lstrBody & "Referer : " & Request.ServerVariables("HTTP_REFERER") & vbCrLf

		lobjNewMail.Subject = "Error encountered in " & lstrDomain  & " domain"

		lobjNewMail.Body    = lstrBody

		lobjNewMail.Send
	End If
	' relase the memory

	Set lobjNewMail = Nothing
End Sub

'This Sub is for handling the error and storing the error into Log file
Sub logasperror(ByVal astrNum, ByVal astrSrc, ByVal astrDesc)
Dim lstrLogFileName
Dim lobjRegistry
Dim lstrLogData
Dim liFileSizeLimit
Dim liLogLevel

On Error Resume Next

	Set lobjRegistry = Server.CreateObject("WScript.Shell.1")

	liLogLevel = lobjRegistry.RegRead("HKEY_LOCAL_MACHINE\Software\FacilitySettings\MoneyCost\Debug\DebugLevel")

	If liLogLevel = "" Or liLogLevel = 0 Then
		Exit Sub
	End If

	lstrLogFileName = lobjRegistry.RegRead("HKEY_LOCAL_MACHINE\Software\FacilitySettings\MoneyCost\Debug\AspLogFile")
	liFileSizeLimit = lobjRegistry.RegRead("HKEY_LOCAL_MACHINE\Software\FacilitySettings\MoneyCost\Debug\MaxDebugFileSize")
	lstrLogData = "Time: " & Now() & vbTab

	' write the ASP Error number using hex on the number property
	lstrLogData = lstrLogData & "Num: " & "0x" & Hex(astrNum) & vbTab

	' returns a description of why the code failed with various error related data
	lstrLogData = lstrLogData & "Desc: "  & astrDesc & vbTab

	lstrLogData = lstrLogData & "Source: " & astrSrc & vbTab

	lstrLogData = lstrLogData & "Referer: " & Request.ServerVariables("HTTP_REFERER") & vbTab

	WriteTofile lstrLogFileName, lstrLogData, liFileSizeLimit, false

	On Error GoTo 0
End Sub

Function WriteTofile(ByVal astrFileName, ByVal astrData, ByVal aiMaxFileZise, ByVal abFileOverwrite)

On Error Resume Next

Dim lobjFileSystem
Dim lobjFile
Dim lobjTxtStream
Dim liIOMode

    Set lobjFileSystem = createobject("scripting.filesystemobject")

    'Check if debug file exists
    If lobjFileSystem.FileExists(astrFileName) Then
        'Retrieve the debug file
        Set lobjFile = lobjFileSystem.GetFile(astrFileName)

        'If flag set to overwrite file
        If abFileOverwrite Then
            'Set file mode to overwrite
            liIOMode = 2

        'Else if the overwrite flag is not set
        Else
            'If debug file size exceeds the max. defined size
            If lobjFile.Size > aiMaxFileZise Then
                'Copy the existing file as a .bak file and set file mode to overwrite
                liIOMode = 2
                lobjFile.Copy Replace(astrFileName, ".txt", ".bak"), True

            'Else If debug file size doesn't exceeds the max. defined size
            Else

                'Set file mode to append
                liIOMode = 8
            End If
        End If

    'Else if debug file doesn't exists
    Else
        'Create and open the file
       lobjFileSystem.CreateTextFile(astrFileName)
       Set lobjFile = lobjFileSystem.GetFile(astrFileName)
       liIOMode = 2
    End If

    Err.Clear

    'Open the file in the appropriate mode
    Set lobjTxtStream = lobjFile.OpenAsTextStream(liIOMode, -2)

    'Write data to the debug file
    lobjTxtStream.Writeline astrData

    'Close the file
    lobjTxtStream.Close

    Set lobjFile = Nothing
    Set lobjTxtStream = Nothing
    Set lobjFileSystem = Nothing
    WriteTofile = True
    Exit Function
End Function

' Begin Page execution
Dim mstrErrorNum
Dim mstrErrorDesc
Dim mstrErrorSource
Dim mstrEmail
Dim mbMaintainacePeriod
Dim mstrUserError
Dim mdtStartTime
Dim lobjRegistry
Dim mdtEndTime

Set lobjRegistry = Server.CreateObject("WScript.Shell.1")
mdtStartTime = lobjRegistry.RegRead("HKEY_LOCAL_MACHINE\Software\FacilitySettings\MoneyCost\Maintainance\StartTime")
mdtEndTime = lobjRegistry.RegRead("HKEY_LOCAL_MACHINE\Software\FacilitySettings\MoneyCost\Maintainance\EndTime")
mstrErrorNum = Session("Number")
mstrErrorDesc = Session("Description")
mstrErrorSource = Session("Source")
mstrEmail = Session("Email")
mstrUserError = Session("UserFriendlyErrorDescription")

mbMaintainacePeriod = false

If mdtStartTime <> "" And mdtEndTime <> "" Then
	If CDate(mdtStartTime) < Now() And CDate(mdtEndTime) > Now() Then
		'Allow Admin in during the maintenance period
		mbMaintainacePeriod = True
	End If
End If

'if not mbMaintainacePeriod then
'	call SendEmail()
	Call logasperror(mstrErrorNum, mstrErrorSource, mstrErrorDesc)
'end if
' Sub which gather details about the error from the ASPError object
%>
<HTML><HEAD><TITLE>GE Update Money Cost</TITLE>
<script language="Javascript">
<!--
function form_submit(url)
{
	document.frmError.action= url;
	document.frmError.submit();
}
function form_Retry()
{window.history.back(-1);}
//-->
</script></HEAD>
<BODY bgColor=#efefef leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<form name="frmError" method="post">
<input type="hidden" name="HID_Desc" value="<%=mstrErrorDesc%>">
<input type="hidden" name="HID_Num" value="<%=mstrErrorNum%>">
<input type="hidden" name="HID_Source" value="<%=mstrErrorSource%>">
<input type="hidden" name="HID_Email" value="<%=mstrEmail%>">
<input type="hidden" name="HID_Time" value="<%=Now()%>">
<input type="hidden" name="HID_Referer" value="<%=Request.ServerVariables("HTTP_REFERER")%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	  <td width="24%" align="right" bgcolor="396797"><img src="images/top_lft.jpg" width="279" height="98"></td>
      <td width="76%" bgcolor="396797"><img src="images/top_bar.jpg" height="97"></td>
</tr>
<tr bgcolor="04204E">
	<td colspan="2" height=5></td>
</tr>
</table>
<br>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
<TR><TD>&nbsp;</TD>
<TD height=150 align=center>
<font face="Arial, Helvetica, sans-serif"><b><font size="4" color="#CE0000">
<%	If Not (mbMaintainacePeriod) Then%>
Internal Error
<%	Else%>
Maintainance Period
<%	End If%>
</font></b></font></TD>
<TD>&nbsp;</TD></TR>
<!-- add the code here to dynamically show the row-->
<TR>
<TD width="5%">&nbsp;</TD>
<%	If Not (mbMaintainacePeriod) Then%>
<TD width="91%" align=center><FONT face="Arial, Helvetica, sans-serif" size=3>
<b> <% If mstrUserError <> "" Then%>
<% = mstrUserError %>
<%Else%> There was an internal error while performing the requested operation.<%End If%><br>
Please click <a href="PricingAnalyst/UpdateDetails.asp">here</a> to continue.<br></b></FONT></TD>
<%Else%>
<TD width="91%" align=center><FONT face="Arial, Helvetica, sans-serif" size=3><b> 
The Money Cost application is currently under scheduled/emergency maintainance<br><br>
Please check back after <% = CDate(mdtEndTime)%><br>
We appreciate your co-operation during this period. Thank you.</b></FONT><br>
</TD>
<%End If%>
<TD width="4%">&nbsp;</TD></TR>
<!-- add the code here to dynamically show the row-->
<% If Not(mbMaintainacePeriod) Then%>
<TR>
<TD width="5%">&nbsp;</TD>
<TD width="91%"><br><br><b><font size="3" face="Arial, Helvetica, sans-serif">Detailed Technical Description:<br><br></font></b>
<font size="2" face="Arial, Helvetica, sans-serif"><b>Error source:</b><%= mstrErrorSource%></font><br>
<font size="2" face="Arial, Helvetica, sans-serif"><b>Error number:</b><%= mstrErrorNum%></font><br>
<font size="2" face="Arial, Helvetica, sans-serif"><b>Error Description:</b><%= mstrErrorDesc%></font></TD>
<TD width="4%">&nbsp;</TD></TR>
<% End If%>
</TABLE></form></BODY></HTML>