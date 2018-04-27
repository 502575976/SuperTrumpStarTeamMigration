Attribute VB_Name = "modBSMoneyCostAuto"
'================================================================
' GE Capital Proprietary and Confidential
' Copyright (c) 2001-2002 by GE Capital - All rights reserved.
'
' This code may not be reproduced in any way without express
' permission from GE Capital.
'================================================================

'================================================================
'MODULE  : modBSMoneyCostAuto
'PURPOSE : This module will contain all common functions, public
'          variables & constants specific to this component.
'================================================================

Option Explicit

'=== Constant for module name ===================================
Private Const cMODULE_NAME      As String = "modBSMoneyCostAuto"

Public Const cCOMPONENT_NAME    As String = "BSMoneyCostAuto"
'================================================================

'=== Registry Constants =========================================
Public Const cFACILITY_CONFIG_REG_PATH      As String = "HKEY_LOCAL_MACHINE\SOFTWARE\FacilitySettings\"
Public Const cFACILITY_ID                   As String = "MoneyCostAuto"
Public Const cCONN_STRINGS_REG_PATH         As String = "\ConnectStrings\"
Public Const cCONN_STRING_KEY               As String = "MCDataWarehouse"
Public Const cHIERARCH_CONN_STRING_KEY      As String = "MoneyCostHierarch"
Public Const cErrorMailBoxKey               As String = "Clarify\ErrorMailBox"
Public Const cEmailOverrideKey              As String = "Clarify\EmailOverride"
Public Const cDeveloperEmailKey             As String = "Clarify\DeveloperEmail"
Public Const cEmailFromKey                  As String = "Clarify\EmailFrom"
Public Const cClarifyPriorityKey            As String = "Clarify\ClarifyPriority"
Public Const cClarifyContactFNameKey        As String = "Clarify\ClarifyContact_Fname"
Public Const cClarifyContactLNameKey        As String = "Clarify\ClarifyContact_Lname"
Public Const cClarifyContactPhoneKey        As String = "Clarify\ClarifyContact_Phone"
Public Const cClarifyQNameKey               As String = "Clarify\ClarifyQueueName"
Public Const cClarifySiteIdKey              As String = "Clarify\SiteID"
Public Const cClarifyEmailSubject           As String = "Clarify\EmailSubject"
'================================================================

Public Const cDEBUG_REG_PATH                    As String = "\Debug"
Public Const cDEBUG_LEVEL_COMPONENT_REG_PATH    As String = ""
Public Const cDEBUG_LEVEL_SIZE_KEY              As String = "DebugLevel"
Public Const cDEBUG_LOG_FILE_PATH_NAME_KEY      As String = "DebugFile"
Public Const cDEBUG_ERROR_FILE_PATH_NAME_KEY    As String = "ErrorLogFile"
Public Const cDEBUG_MAX_FILE_SIZE_KEY           As String = "MaxDebugFileSize"
Public Const cWORKING_DIRECTORY_PATH            As String = "WorkingDirectory"
Public Const cBACKUP_LOCATION                   As String = "Backup_Location"
Public Const cNETWORK_LOCATION                  As String = "Network_Location"
Public Const cFTP_LOCATION                      As String = "FTP_Location"
Public Const cFTP_DIRECTORY                     As String = "FTP_Directory"
Public Const cFTP_USER                          As String = "FTP_User"
Public Const cFTP_PASSWORD                      As String = "FTP_Password"

Public Const cFTP_LOCATION_NEWDATEFORMAT        As String = "FTP_LocationNewDateFormat"
Public Const cFTP_USER_NEWDATEFORMAT            As String = "FTP_UserNewDateFormat"
Public Const cFTP_PASSWORD_NEWDATEFORMAT        As String = "FTP_PasswordNewDateFormat"
Public Const cFTP_DIRECTORY_NEWDATEFORMAT       As String = "FTP_DirectoryNewDateFormat"

'=== Error Constants ============================================
Public Const cHIGHEST_ERROR     As Variant = vbObjectError + 256
Public Const cINVALID_PARMS     As Variant = cHIGHEST_ERROR + 1000
Public Const cINVALID_SQL_ID    As Variant = cHIGHEST_ERROR + 1010
'================================================================

'==== Declaring Global Variables =================================
Public gstrErrorLogFile         As String

'=== Debug Constants & variables ================================
Public Const cAPPEND_IO_MODE    As Integer = 8
Public Const cWRITE_IO_MODE     As Integer = 2

Public Enum eDebugLevel
    ecDebugDoNotLog = 0
    ecDebugFatalError = 3
    ecDebugCriticalError = 6
    ecDebugWarning = 9
    ecDebugInputTrace = 12
    ecDebugOutputTrace = 15
    ecDebugLogData = 18
    ecDebugLoglargeData = 21
End Enum

''MapNetwork path
Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Type NETRESOURCE
        dwScope As Long
        dwType As Long
        dwDisplayType As Long
        dwUsage As Long
        lpLocalName As String
        lpRemoteName As String
        lpComment As String
        lpProvider As String
End Type
Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_PUBLICNET = &H2
Public Const RESOURCE_REMEMBERED = &H3

Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_UNKNOWN = &HFFFF

Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
Public Const RESOURCEUSAGE_RESERVED = &H80000000

Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEDISPLAYTYPE_FILE = &H4

'============================================================
'METHOD  : CopyFiles
'PURPOSE : Copy/Move files to location specified
'PARMS   : [argSource] From where the files have to be moved
'          [argDestination] Location where the files have to be moved or copied.
'          [argMove] Boolean argument specifying copy or move
'RETURN  : XML Error String if any
'============================================================
Public Function CopyFiles(ByVal argSource As String, ByVal argDestination As String, Optional argMove As Boolean = False) As String
    On Error GoTo CopyFiles_ErrorHandler

    Dim lobjFSO             As New FileSystemObject
    Dim lstrReturnString    As String

    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): In CopyFiles() method"

    ' -------------------------------------------
    ' Check move(true) or copy argument
    ' -------------------------------------------
    If argMove Then
        If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): Deleting Source file - " & argSource

        lobjFSO.DeleteFile argSource, True

        If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): Move Source file - " & argSource & " to destination - " & argDestination

        lobjFSO.MoveFile argSource, argDestination
    Else
        If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): Copying Source file - " & argSource & " to destination - " & argDestination

        lobjFSO.CopyFile argSource, argDestination, True
    End If

CopyFiles_CleanMemory:
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): Exit CopyFiles() method"

    Set lobjFSO = Nothing
    CopyFiles = lstrReturnString
    Exit Function

CopyFiles_ErrorHandler:
    'gstrErrDesc = "MCUSD file not available for updation at the following path " & vbCrLf & _
                   argSource & vbCrLf & _
                   "Date : " & Date & _
                   " Time : " & Time
    'Call SendNotificationToBusssiness(lstrErrDesc)

    lstrReturnString = "<ERROR_DETAILS>" & _
                            "<ERROR_NUMBER>" & Err.Number & "</ERROR_NUMBER>" & _
                            "<ERROR_DESCRIPTION>" & Err.Description & "</ERROR_DESCRIPTION>" & _
                            "<ERROR_SOURCE>" & Err.Source & "::CopyFiles()</ERROR_SOURCE>" & _
                        "</ERROR_DETAILS>"

    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): Error occured - " & Err.Description

    Resume CopyFiles_CleanMemory
End Function

'================================================================
'METHOD  : GetDateFormat
'PURPOSE : To get particular date in mmddyy format
'PARMS   : astrDate [String] = date, which needs to convert in mmddyy format
'RETURN  : [String] = date formatted into mmddyy string
'================================================================
Public Function GetDateFormat(ByVal astrDate As Date) As String
    Dim lstrDay As String
    Dim lstrMonth As String
    Dim lstrYear As String

    lstrDay = Day(astrDate)
    If Len(lstrDay) = 1 Then
        lstrDay = "0" & lstrDay
    End If

    lstrMonth = Month(astrDate)
    If Len(lstrMonth) = 1 Then
        lstrMonth = "0" & lstrMonth
    End If

    lstrYear = Year(astrDate)
    If Len(lstrYear) > 2 Then
        lstrYear = Right(lstrYear, 2)
    End If

    GetDateFormat = lstrMonth & lstrDay & lstrYear
End Function

'================================================================
'METHOD  : SendErrNotification
'PURPOSE : To send error notification to the designated error
'          mail box.
'PARMS   :
'          astrBody [String] = Error Message.
'RETURN  : None.
'================================================================
Public Function SendErrNotification(ByVal astrQueueName As String, _
                                    ByVal astrBusinessContact As String, _
                                    ByVal astrBody As String, _
                                    ByVal abCutTicket As Boolean, _
                                    ByVal abSendNotification As Boolean, _
                                    ByVal astrMCCode, _
                                    ByVal astrProcessDate) As String

On Error GoTo SendErrNotification_ErrHandler

Dim lstrMailBody        As String
Dim lstrTo              As String
Dim lstrFrom            As String
Dim lstrReturnString    As String
Dim lPos                As Integer

    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): In SendErrNotification() method"

    
        Dim lobjCDONTS As New CDO.Message
    
        Dim objCDOSYSCon As New CDO.Configuration
        'Out going SMTP server
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.ad.ge.com"
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
        objCDOSYSCon.Fields.Update

Set lobjCDONTS.Configuration = objCDOSYSCon

    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): CDONTS instance created"

    If abCutTicket Then
        If gstrEmailOverride <> "" Then
            lstrTo = gstrEmailOverride
        Else
            lstrTo = gstrErrMailBox
        End If

        'Create the email body as per the format required to open a clarify case.
        lstrMailBody = "CASE_START" & vbCrLf & _
                        "CALL_TYPE: Problem" & vbCrLf & _
                        "SEVERITY: High" & vbCrLf & _
                        "PRIORITY: " & gstrClarifyPriority & vbCrLf & _
                        "SITE_ID: " & gstrClarifySiteId & vbCrLf & _
                        "CONTACT_LAST_NAME: " & gstrClarifyContactLname & vbCrLf & _
                        "CONTACT_FIRST_NAME: " & astrQueueName & vbCrLf & _
                        "CONTACT_PHONE: " & gstrClarifyContactPhone & vbCrLf & _
                        "CASE_SUMMARY: " & gstrClarifyEmailSub & vbCrLf & _
                        "CASE_DESCRIPTION:" & vbCrLf & _
                            astrBody & vbCrLf & _
                        "CASE_END"

        'Set the properties of the CDONTS object
        lobjCDONTS.From = gstrDeveloperEmail
        lobjCDONTS.To = lstrTo
        If gstrDeveloperEmail <> "" Then lobjCDONTS.Cc = gstrDeveloperEmail
        lobjCDONTS.Subject = "Case Request"
        lobjCDONTS.HTMLBody = "<pre>" & lstrMailBody

        If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): Sending email to - " & lstrTo & " and cc - " & gstrDeveloperEmail

        'Send email
        Call lobjCDONTS.send
        Set lobjCDONTS = Nothing
    End If

    If abSendNotification Then
       
        lPos = InStr(1, astrBody, "<", vbBinaryCompare)

        lstrMailBody = "Dear Money Cost User," & vbCrLf & vbCrLf & _
                       "An error was reported by service while updating the money cost file." & IIf(abCutTicket, " A clarify case has been created and dispatched to IT support team." & vbCrLf, vbCrLf) & vbCrLf & _
                       "Money Cost Support Team." & vbCrLf & vbCrLf

        If lPos <= 0 Then
            lstrMailBody = lstrMailBody & "Error Reported:" & vbCrLf & astrBody & vbCrLf & vbCrLf
        Else
            lstrMailBody = lstrMailBody & "Error Reported:" & vbCrLf & Left(astrBody, lPos - 1) & vbCrLf & vbCrLf
        End If

        '************************
        'Changes made on 08 Nov 2006 by Nizar
        'Changes made to put an message as auto generated email.
        lstrMailBody = lstrMailBody & "Note: This is an auto-generated email from Money Cost Service. Please do not reply to this email."
        
        'Changes made to get From Email address from EmailFrom (in registry)
        lobjCDONTS.From = gstrFrom
        'lobjCDONTS.From = gstrDeveloperEmail
        '***********************
        
        lobjCDONTS.To = astrBusinessContact
        If gstrDeveloperEmail <> "" Then lobjCDONTS.Cc = gstrDeveloperEmail
        lobjCDONTS.Subject = "MoneyCostService: Error reported while updating " & astrMCCode & " for " & astrProcessDate
        lobjCDONTS.HTMLBody = "<pre>" & lstrMailBody

        If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): Sending email to - " & lstrTo & " and cc - " & gstrDeveloperEmail

        'Send email
        Call lobjCDONTS.send
    End If

SendErrNotification_CleanMemory:
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): Exit SendErrNotification() method"

    Set lobjCDONTS = Nothing
    SendErrNotification = lstrReturnString
    Set objCDOSYSCon = Nothing
    Exit Function

SendErrNotification_ErrHandler:
    lstrReturnString = "<ERROR_DETAILS><ERROR_NUMBER>" & Err.Number & "</ERROR_NUMBER>" & _
                        "<ERROR_DESCRIPTION>" & Err.Description & "</ERROR_DESCRIPTION>" & _
                        "<ERROR_SOURCE>" & Err.Source & "::SendErrNotification()</ERROR_SOURCE></ERROR_DETAILS>"

    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): Error occured - " & Err.Description & vbCrLf & lstrReturnString

    Resume SendErrNotification_CleanMemory
End Function

'============================================================================================
'METHOD :   fnSortXmlData
'PURPOSE:   This Function will use to sort the XML records depending upon the column.
'PARMS  :   astrXmlDoc      [String] = XML string.
'           astrXslFileName [String] = Path of XSL file.
'RETURN :   Sorted HTML output.
'============================================================================================
Public Function fnSortXmlData(astrXmlDoc, astrXslFilePath)
On Error GoTo fnSortXmlData_ErrHandler

Dim lstrErrSrc     As String    'to store error source
Dim lstrMethodName As String    'to store method name
Dim lstrErrDesc    As String    'to store error description
Dim llErrNbr       As Long      'to store error number

    lstrMethodName = "fnSortXmlData"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

    Dim lobjXMLDoc          As New MSXML2.FreeThreadedDOMDocument40     'to load XML Document
    Dim lobjXSLDoc          As New MSXML2.FreeThreadedDOMDocument40     'to load XSL Document
    Dim lobjTemplate        As New MSXML2.XSLTemplate40                 'to define sytlesheet (xsl)
    Dim lobjProcessor       As IXSLProcessor                            'to set xsl processor
    Dim lbTransformResult   As Boolean                                  'to get result output

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrXmlDoc, ecDebugInputTrace
    End If

    lobjXMLDoc.async = False
    lobjXMLDoc.loadXML astrXmlDoc

    lobjXSLDoc.async = False
    lobjXSLDoc.Load (astrXslFilePath)

    'Create template and pass parameter to the XSL file for soting

    Set lobjTemplate.stylesheet = lobjXSLDoc
    Set lobjProcessor = lobjTemplate.createProcessor()
    lobjProcessor.input = lobjXMLDoc

    lbTransformResult = lobjProcessor.Transform()

    fnSortXmlData = lobjProcessor.output

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, fnSortXmlData, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

fnSortXmlData_CleanMemory:
    'clear all local object variables from memory
    Set lobjXMLDoc = Nothing
    Set lobjXSLDoc = Nothing
    Set lobjTemplate = Nothing

    Exit Function

fnSortXmlData_ErrHandler:
    lstrErrSrc = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description

    fnSortXmlData = vbNullString

    'write error message to log file
    WriteToTextDebugFile cMODULE_NAME & lstrMethodName, BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    Err.Raise llErrNbr, lstrErrSrc, lstrErrDesc

    Resume fnSortXmlData_CleanMemory
End Function

