VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "IBSMoneyCostAutoService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'================================================================
' GE Capital Proprietary and Confidential
' Copyright (c) 2001-2002 by GE Capital - All rights reserved.
'
' This code may not be reproduced in any way without express
' permission from GE Capital.
'================================================================

'================================================================
'MODULE  : IBSMoneyCostAutoService
'PURPOSE : This will contain non-transactional wrapper methods
'          to avoid multiple invocation of business component
'          methods by UI tier.
'================================================================

Option Explicit

'Constant for module name =======================================
Private Const cMODULE_NAME As String = "IBSMoneyCostAutoService"
'================================================================

'================================================================
'METHOD  :  Ping
'PURPOSE :  Allows component to be pinged to verify it can be
'           instantiated
'PARMS   :  none
'RETURN  :  String with date and time
'================================================================
Public Function Ping() As String
    Ping = "Ping request to " & cCOMPONENT_NAME & "." & cMODULE_NAME & " returned at " & Format(Now, "mm/dd/yyyy Hh:Nn:Ss AM/PM") & " server time."
End Function

'================================================================
'METHOD  : Test
'PURPOSE : Returns a string that indicates that the component
'          can connect to the database and the registry.
'PARMS   : NONE
'RETURN  : String
'================================================================
Public Function Test() As String
On Error GoTo Test_ErrHandler

Dim lobjcDataClass  As New cDataClass
Dim lrsTest         As New ADODB.Recordset

    'Execute the Test SQL statement which returns a count of the records
    Set lrsTest = lobjcDataClass.Execute(ecTestSQL, ecRSExecuteRecords, "")

    'Return the total records
    Test = "Retrieved " & lrsTest.Fields(0).Value & " records."
    lrsTest.Close

Test_CleanMemory:
    Set lobjcDataClass = Nothing
    Set lrsTest = Nothing
    Exit Function

Test_ErrHandler:
    Test = vbNullString

    Err.Raise Err.Number, cCOMPONENT_NAME & "." & cMODULE_NAME & ":Test()", Err.Description

    Resume Test_CleanMemory
End Function

Private Sub Class_Initialize()
On Error GoTo Class_ErrHandler

    'Get Debugging parameters from registry for this Component
    giDebugLevel = Val(ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & cDEBUG_REG_PATH & cDEBUG_LEVEL_COMPONENT_REG_PATH & "\" & cDEBUG_LEVEL_SIZE_KEY))
    'Read the Log File Name with location
    gstrDebugFile = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & cDEBUG_REG_PATH & cDEBUG_LEVEL_COMPONENT_REG_PATH & "\" & cDEBUG_LOG_FILE_PATH_NAME_KEY)
    'Read the Error File Name with location
    gstrErrorLogFile = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & cDEBUG_REG_PATH & cDEBUG_LEVEL_COMPONENT_REG_PATH & "\" & cDEBUG_ERROR_FILE_PATH_NAME_KEY)
    'Read the Max size of the file
    glMaxDebugFileSize = Val(ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & cDEBUG_REG_PATH & cDEBUG_LEVEL_COMPONENT_REG_PATH & "\" & cDEBUG_MAX_FILE_SIZE_KEY))

    gstrErrMailBox = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cErrorMailBoxKey)
    gstrEmailOverride = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cEmailOverrideKey)
    gstrDeveloperEmail = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cDeveloperEmailKey)
    gstrClarifySiteId = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cClarifySiteIdKey)
    gstrClarifyEmailSub = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cClarifyEmailSubject)
    gstrClarifyPriority = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cClarifyPriorityKey)
    gstrClarifyContactFname = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cClarifyContactFNameKey)
    gstrClarifyContactLname = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cClarifyContactLNameKey)
    gstrClarifyContactPhone = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cClarifyContactPhoneKey)
    gstrFrom = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cEmailFromKey)

    Exit Sub

Class_ErrHandler:
     If giDebugLevel > 1 Then
            WriteToTextDebugFile "Class_Initialize()", BuildErrXML(Err.Number, Err.Source & cMODULE_NAME & ".Class_Initialize()", Err.Description), ecDebugCriticalError
    End If

    Resume Next
End Sub

'================================================================
'METHOD  : ExecuteServiceFlow
'PURPOSE : Main Controller procedure for the service flow
'PARMS   : NONE
'RETURN  : NONE
'================================================================
Public Sub ExecuteServiceFlow()
On Error GoTo ExecuteServiceFlow_ErrHandler

Dim lstrMethodName  As String   'to store method name
Dim lstrErrSrc      As String   'to store error source
Dim lstrErrDesc     As String   'to store error description
Dim llErrNbr        As Long     'to store error number

If giDebugLevel > 1 Then
        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow:  method Start"
End If

    lstrMethodName = "ExecuteServiceFlow"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lcnCSVConnection                    As New ADODB.Connection
Dim lrsCSVRecordset                     As New ADODB.Recordset
Dim lobjBSMCIMoneyCostService           As New BSMoneyCost.IMoneyCostService
Dim lobjBSMCIMoneyCostMgr               As New BSMoneyCost.IMoneyCostMgr
Dim lobjcMoneyCostAutoService           As New cBSMoneyCostAutoService
Dim lobjcFTP                            As New cFTP
Dim lobjAllMCFileDOM                    As New DOMDocument40
Dim lobjIndexRateDOM                    As New DOMDocument40
Dim lobjIndexRateDOM_PrevDay            As New DOMDocument40
Dim lobjIndexDataDOMXml                 As New DOMDocument40
Dim lobjCSVInsertDOMXml                 As New DOMDocument40
Dim lobjMCFileNodeList                  As IXMLDOMNodeList
Dim lobjIndexRateNodeList               As IXMLDOMNodeList
Dim lobjIndexRateNodeList_PrevDay       As IXMLDOMNodeList
Dim lobjIndexDataNodeList               As IXMLDOMNodeList
Dim lobjIndexDataNodeList_PrevDay       As IXMLDOMNodeList
Dim lobjCSVInsertNodeList               As IXMLDOMNodeList
Dim lobjMCFileNode                      As IXMLDOMNode
Dim lobjIndexRateNode                   As IXMLDOMNode
Dim lobjIndexRateNode_PrevDay           As IXMLDOMNode
Dim lobjIndexDataNode                   As IXMLDOMNode
Dim lobjCSVInsertNode                   As IXMLDOMNode
Dim liCounter1                          As Integer
Dim liCounter2                          As Integer
Dim liCounter3                          As Integer
Dim liSQ_MC_ID                          As Integer
Dim liDAYS_TO_SKIP                      As Integer
Dim lbMarketOpenFlag                    As Boolean
Dim lstrMC_CODE                         As String
Dim lstrDESCRIPTION                     As String
Dim lstrSTART_TIME                      As String
Dim lstrEND_TIME                        As String
Dim lstrWorkingDirectory                As String
Dim lstrCopyFileResponse                As String
Dim lstrProcessDate                     As String
Dim lstrProcessDate_PrevDay             As String
Dim lstrIndexRateReqXml                 As String
Dim lstrIndexRateRespXml                As String
Dim lstrIndexRateRespXml_PrevDay        As String
Dim lstrINDEX_CODEList                  As String
Dim lstrINDEX_CODEList_PrevDay          As String
Dim lstrINDEX_TERMList                  As String
Dim lstrINDEX_TERMList_PrevDay          As String
Dim lstrProcessDateFieldName            As String
Dim lstrCURRENCY_CODE                   As String
Dim lstrIndexDataReqXml                 As String
Dim lstrIndexDataRespXml                As String
Dim lstrGetMCFilesRespXml               As String
Dim lstrINTEREST_RATE_DWH               As String
Dim lstrMissingYieldCurve               As String
Dim lstrMissingInterestRate             As String
Dim lstrCSV_INSERT_RECORDXml            As String
Dim lstrCSVInsertSQL                    As String
Dim lstrClarifyQName                    As String
Dim lstrBusinessContact                 As String
Dim lstrSendErrNotiResult               As String
Dim lstrSqlQry                          As String
Dim lstrCSVColumnHeaders                As String
Dim lstrMCDBAllIndexCode                As String
Dim lstrMCDBAllIndexTerm                As String
Dim lstrMCFileStartExec                 As String
Dim lstrMCFileEndExec                   As String
Dim lstrUpdateMCLogReqXml               As String
Dim lstrUpdateMCLogRespXml              As String
Dim lstrCommonErrorDetails              As String
Dim lstrErrorDetails                    As String
Dim larrMissingIndexCode()              As String
Dim larrMissingIntRate()                As String
Dim lstrBackup_Location                 As String
Dim lstrNetwork_Location                As String
Dim lstrFTP_Location                    As String
Dim lstrFTP_Directory                   As String
Dim lstrFTP_User                        As String
Dim lstrFTP_Password                    As String
Dim lstrBlankRegistryKey                As String
Dim lbCopyFileResponse                  As Boolean
Dim lbDataDeletionFlag                  As Boolean

Dim lbServiceRunFlag                    As Boolean
Dim lstrFREQUENCY                       As String
Dim liFREQUENCY_COUNT                   As Integer
Dim lstrLAST_SCHEDULE_PROCESS_DATE      As String
Dim liMARKET_CLOSED_DWH_CHECK_COUNTER   As Integer
Dim lstrScheduleProcessDate             As String
Dim lstrUpdateMCFileReqXml              As String
Dim lstrUpdateMCFileRespXml             As String
Dim lrsCSVDeleteRecordset               As New ADODB.Recordset

Dim lstrLAST_UPDATED_IND                As Boolean
Dim lstrFTP_LocationForNewDateFormat    As String
Dim lstrFTP_DirectoryForNewDateFormat   As String
Dim lstrFTP_UserForNewDateFormat        As String
Dim lstrFTP_PasswordForNewDateFormat    As String
Dim lstrDateFormat                      As String
Dim lstrDateFormatRequired              As Boolean
Dim lstrCSVUpdateDateSQL                As String
'added
Dim lstrMissingYieldCurve_NotCopied     As String
Dim larrMissingIndexCode_NotCopied()    As String

Dim lobjFSOExecute                      As New FileSystemObject

Dim lstrProcessDate_NewDateFormat       As String
Dim lstrDay_NewDateFormat               As String
Dim lstrMonth_NewDateFormat             As String

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "", ecDebugInputTrace
    End If

If giDebugLevel > 1 Then
    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():  Check necessary registry keys"
End If
    ' --------------------------------------------
    ' Check if necessary registry keys are defined
    ' --------------------------------------------
    lstrWorkingDirectory = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cWORKING_DIRECTORY_PATH)
    lstrClarifyQName = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cClarifyQNameKey)
    lstrBackup_Location = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cBACKUP_LOCATION)
    lstrNetwork_Location = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cNETWORK_LOCATION)
    lstrFTP_Location = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cFTP_LOCATION)
    lstrFTP_Directory = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cFTP_DIRECTORY)
    lstrFTP_User = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cFTP_USER)
    lstrFTP_Password = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cFTP_PASSWORD)
    
    lstrFTP_LocationForNewDateFormat = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cFTP_LOCATION_NEWDATEFORMAT)
    lstrFTP_DirectoryForNewDateFormat = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cFTP_DIRECTORY_NEWDATEFORMAT)
    lstrFTP_UserForNewDateFormat = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cFTP_USER_NEWDATEFORMAT)
    lstrFTP_PasswordForNewDateFormat = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & "\" & cFTP_PASSWORD_NEWDATEFORMAT)

    If giDebugLevel > 1 Then
        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():  Check if necessary registry entries are blank"
    End If
    ' ---------------------------------------------
    ' Check if necessary registry entries are blank
    ' ---------------------------------------------
    If (lstrWorkingDirectory = "" Or lstrClarifyQName = "" Or lstrBackup_Location = "" Or lstrFTP_Location = "") Then
        If lstrWorkingDirectory = "" Then lstrBlankRegistryKey = lstrBlankRegistryKey & cWORKING_DIRECTORY_PATH & ","
        If lstrClarifyQName = "" Then lstrBlankRegistryKey = lstrBlankRegistryKey & cClarifyQNameKey & ","
        If lstrBackup_Location = "" Then lstrBlankRegistryKey = lstrBlankRegistryKey & cBACKUP_LOCATION & ","
        If lstrFTP_Location = "" Then lstrBlankRegistryKey = lstrBlankRegistryKey & cFTP_LOCATION & ","

        If Right(lstrBlankRegistryKey, 1) = "," Then lstrBlankRegistryKey = Mid(lstrBlankRegistryKey, 1, Len(lstrBlankRegistryKey) - 1)

        lstrCommonErrorDetails = "The updation failed because the following registry keys were not defined." & vbCrLf & _
                                 "Keys List: " & lstrBlankRegistryKey
        If giDebugLevel > 1 Then
            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrCommonErrorDetails
        End If
    End If

    If lstrCommonErrorDetails = "" Then
        If Right(lstrWorkingDirectory, 1) <> "\" Then lstrWorkingDirectory = lstrWorkingDirectory & "\"
        If Right(lstrBackup_Location, 1) <> "\" Then lstrBackup_Location = lstrBackup_Location & "\"
        If Right(lstrNetwork_Location, 1) <> "\" Then lstrNetwork_Location = lstrNetwork_Location & "\"

        'Changed to use Microsoft Jet OLEDB driver that allows where clause and order by actions on the recordset
        lcnCSVConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & lstrWorkingDirectory & ";" & _
                                "Extended Properties=""text;HDR=YES;FMT=Delimited"""

        If giDebugLevel > 1 Then
            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get List of MC ALL Files Start"
        End If
        
        ' -------------------------------------------
        ' Get List of MC ALL Files
        ' -------------------------------------------
        ' Invoke BSMoneyCost.GetAllMCFiles()
        lstrGetMCFilesRespXml = lobjBSMCIMoneyCostService.GetAllMCFiles()

        'validate, if response xml is well-formed
        If Not lobjAllMCFileDOM.loadXML(lstrGetMCFilesRespXml) Then
            lstrCommonErrorDetails = "Error No. : " & lobjAllMCFileDOM.parseError.errorCode & vbCrLf & "Error Description : " & _
                                "Error loading response XML from BSMOneyCost.IMoneyCostService.GetAllMCFiles(). " & _
                                lobjAllMCFileDOM.parseError.reason
                                
        If giDebugLevel > 1 Then
            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get List of MC ALL Files Error" & lstrCommonErrorDetails
        End If
        
            GoTo ExecuteServiceFlow_ErrHandler
        End If

        If giDebugLevel > 1 Then
            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get List of MC ALL Files Complete"
        End If
        
        Set lobjMCFileNodeList = lobjAllMCFileDOM.selectNodes("/MC_FILE_RESPONSE/MC_FILESet/MC_FILE")

        If lobjMCFileNodeList.length <= 0 Then
            lstrCommonErrorDetails = "No Money Cost file found to process."
            
            If giDebugLevel > 1 Then
                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): No Money Cost file found to process."
            End If
            
            GoTo ExecuteServiceFlow_ErrHandler
        End If

        lstrUpdateMCLogReqXml = "<UPDATE_MC_FILE_LOGS_REQUEST><MC_FILE_LOGSet>"

        lstrUpdateMCFileReqXml = lstrUpdateMCFileReqXml & "<UPDATE_MC_FILE_REQUEST><MC_FILESet>"

        If giDebugLevel > 1 Then
            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Iterate through the Recordset with all MC Files list START:-"
        End If
        ' ----------------------------------------------------
        ' Iterate through the Recordset with all MC Files list
        ' ----------------------------------------------------
        For liCounter1 = 0 To lobjMCFileNodeList.length - 1
            On Error GoTo ResumeNextMCFile

            lstrMCFileStartExec = Now()

            ' ----------------------------------------------------
            ' Iterate through the Recordset with all MC Files list
            ' ----------------------------------------------------
            If giDebugLevel > 1 Then
                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Iterate through the Recordset with all MC Files list"
            End If
            ' --------------------------------------------------------------
            ' For each MC File record in the recordset process the following
            ' --------------------------------------------------------------
            Set lobjMCFileNode = lobjMCFileNodeList.Item(liCounter1)

            liSQ_MC_ID = 0
            lstrMC_CODE = ""
            lstrDESCRIPTION = ""
            lstrCURRENCY_CODE = ""
            lstrSTART_TIME = ""
            lstrEND_TIME = ""
            liDAYS_TO_SKIP = 0
            lstrClarifyQName = ""
            lstrBusinessContact = ""
            lstrMissingYieldCurve = ""
            lstrMissingInterestRate = ""
            lstrSendErrNotiResult = ""
            lstrCSV_INSERT_RECORDXml = "<CSV_INSERT_DATA_REQUEST><CSV_INSERT_DATASet>"
            lstrMCDBAllIndexCode = ""
            lstrMCDBAllIndexTerm = ""
            lstrErrorDetails = ""
            lbMarketOpenFlag = True
            lbServiceRunFlag = False
            lstrFREQUENCY = ""
            liFREQUENCY_COUNT = 0
            lstrLAST_SCHEDULE_PROCESS_DATE = ""
            liMARKET_CLOSED_DWH_CHECK_COUNTER = 0
            lstrINDEX_CODEList = ""
            lstrINDEX_CODEList_PrevDay = ""
            lstrINDEX_TERMList = ""
            lstrLAST_UPDATED_IND = True
            lstrDateFormat = ""
            lstrDateFormatRequired = False
            lstrMissingYieldCurve_NotCopied = ""
            Erase larrMissingIndexCode_NotCopied
            
            lstrProcessDate_NewDateFormat = ""
            lstrDay_NewDateFormat = ""
            lstrMonth_NewDateFormat = ""

            'fetch MC file details into local variables, from XML
            If IsXMLElementPresent(lobjMCFileNode, "SQ_MC_ID") Then
                liSQ_MC_ID = Val(lobjMCFileNode.selectNodes("SQ_MC_ID").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "MC_CODE") Then
                lstrMC_CODE = Trim(lobjMCFileNode.selectNodes("MC_CODE").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "DESCRIPTION") Then
                lstrDESCRIPTION = Trim(lobjMCFileNode.selectNodes("DESCRIPTION").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "CURRENCY_CODE") Then
                lstrCURRENCY_CODE = Trim(lobjMCFileNode.selectNodes("CURRENCY_CODE").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "START_TIME") Then
                lstrSTART_TIME = Trim(lobjMCFileNode.selectNodes("START_TIME").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "END_TIME") Then
                lstrEND_TIME = Trim(lobjMCFileNode.selectNodes("END_TIME").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "DAYS_TO_SKIP") Then
                liDAYS_TO_SKIP = Val(lobjMCFileNode.selectNodes("DAYS_TO_SKIP").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "CLARIFY_QUEUE") Then
                lstrClarifyQName = Trim(lobjMCFileNode.selectNodes("CLARIFY_QUEUE").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "BUSINESS_CONTACT") Then
                lstrBusinessContact = Trim(lobjMCFileNode.selectNodes("BUSINESS_CONTACT").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "FREQUENCY") Then
                lstrFREQUENCY = Trim(lobjMCFileNode.selectNodes("FREQUENCY").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "FREQUENCY_COUNT") Then
                liFREQUENCY_COUNT = Trim(lobjMCFileNode.selectNodes("FREQUENCY_COUNT").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "LAST_SCHEDULE_PROCESS_DATE") Then
                lstrLAST_SCHEDULE_PROCESS_DATE = Trim(lobjMCFileNode.selectNodes("LAST_SCHEDULE_PROCESS_DATE").Item(0).Text)
            End If

            If IsXMLElementPresent(lobjMCFileNode, "MARKET_CLOSED_DWH_CHECK_COUNTER") Then
                liMARKET_CLOSED_DWH_CHECK_COUNTER = Trim(lobjMCFileNode.selectNodes("MARKET_CLOSED_DWH_CHECK_COUNTER").Item(0).Text)
            End If
            
            If IsXMLElementPresent(lobjMCFileNode, "LAST_UPDATED_IND") Then
                lstrLAST_UPDATED_IND = Trim(lobjMCFileNode.selectNodes("LAST_UPDATED_IND").Item(0).Text)
            End If
            
            If IsXMLElementPresent(lobjMCFileNode, "DATE_FORMAT") Then
                lstrDateFormat = Trim(lobjMCFileNode.selectNodes("DATE_FORMAT").Item(0).Text)
            End If
            
            If IsXMLElementPresent(lobjMCFileNode, "DATE_FORMAT_REQUIRED") Then
                lstrDateFormatRequired = Trim(lobjMCFileNode.selectNodes("DATE_FORMAT_REQUIRED").Item(0).Text)
            End If
            

            If lstrSTART_TIME <> "" And lstrEND_TIME <> "" Then
                'if MC file process time is out of specified time, skip to next MC file
                If DateDiff("s", Format(Time(), "hh:mm:ss"), lstrSTART_TIME) < 0 Or DateDiff("s", Format(Time(), "hh:mm:ss"), lstrEND_TIME) > 0 Then
                
                If giDebugLevel > 1 Then
                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): if MC file process time is out of specified time, skip to next MC file"
                End If
                    GoTo ResumeNextMCFile
                End If
            End If

            If giDebugLevel > 1 Then
                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Process date  START"
            End If
            ' -------------------------------------------------
            ' Process date = System Date - MC_FILE.Days_TO_SKIP
            ' -------------------------------------------------
            lstrProcessDate = DateAdd("d", -liDAYS_TO_SKIP, Date)
            
            'get the previous processdate as well to get the previous day's adder rates.
            lstrProcessDate_PrevDay = DateAdd("d", -(liDAYS_TO_SKIP + 1), Date)
            
            'The week end check of the previous day is based on the current process date.
            'dont change the order of the cases.
            Select Case Weekday(lstrProcessDate)
                Case 7, 1 'Saturday , Sunday
                    lstrProcessDate_PrevDay = DateAdd("d", -(liDAYS_TO_SKIP + 1 + 2), Date)
            End Select
            
            Select Case Weekday(lstrProcessDate)
                Case 7, 1 'Saturday , Sunday
                    lstrProcessDate = DateAdd("d", -(liDAYS_TO_SKIP + 2), Date)
            End Select
            

            lstrScheduleProcessDate = DateAdd(lstrFREQUENCY, liFREQUENCY_COUNT, lstrLAST_SCHEDULE_PROCESS_DATE)
            Select Case Weekday(lstrScheduleProcessDate)
                Case 7, 1 'Saturday , Sunday
                    lstrScheduleProcessDate = DateAdd(lstrFREQUENCY, liFREQUENCY_COUNT + 2, lstrLAST_SCHEDULE_PROCESS_DATE)
            End Select

            If DateDiff("d", lstrProcessDate, lstrScheduleProcessDate) = 0 And Weekday(Date) <> 1 And Weekday(Date) <> 7 Then
                lbServiceRunFlag = True
            End If

            If giDebugLevel > 1 Then
            
                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Process date = System Date - MC_FILE.Days_TO_SKIP lbServiceRunFlag=" & lbServiceRunFlag
            End If
            If lbServiceRunFlag = True Then
                
                If giDebugLevel > 1 Then
                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Process date = FTP Location Mapped start"
                End If
                
                'FTP Location Mapped start
                If IsFTPLocationMapped(lstrFTP_Location, lstrFTP_User, lstrFTP_Password) Then
                    
                    If Right(lstrFTP_Location, 1) <> "\" Then lstrFTP_Location = lstrFTP_Location & "\"
                    If Right(lstrFTP_Directory, 1) <> "\" Then lstrFTP_Directory = lstrFTP_Directory & "\"
                    ' -----------------------------------------------------------------
                    ' Create a Backup Copy of existing MCXXX File as BKMMDDYY.MCUSD.csv
                    ' on Backup location, from FTP location
                    ' -----------------------------------------------------------------
                    lstrCopyFileResponse = CopyFiles(lstrFTP_Location & lstrFTP_Directory & lstrMC_CODE & ".csv", lstrBackup_Location & "BK" & GetDateFormat(Date) & "." & lstrMC_CODE & ".csv")
                    
'                    If lobjcFTP.Connected Then lobjcFTP.Disconnect
'                    lobjcFTP.Connect lstrFTP_Location, lstrFTP_User, lstrFTP_Password
'                    lobjcFTP.Directory = lstrFTP_Directory
'                    lstrCopyFileResponse = lobjcFTP.GetFile(lstrMC_CODE & ".csv", lstrBackup_Location & "BK" & GetDateFormat(Date) & "." & lstrMC_CODE & ".csv")
'                    lobjcFTP.Disconnect
    
                    'if any error occurred, raise clarify case
                    If lstrCopyFileResponse <> "" Then
                        lstrErrorDetails = "Error while creating backup copy of existing " & lstrMC_CODE & ".csv file as " & _
                                            lstrBackup_Location & "BK" & GetDateFormat(Date) & "." & lstrMC_CODE & ".csv on " & _
                                            lstrBackup_Location & " : " & lstrCopyFileResponse
                                            
                        If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
    
                    If giDebugLevel > 1 Then
                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                    End If
    
                        GoTo ResumeNextMCFile
                    End If
    
                    ' Invoke Common Method.CopyFiles() to copy latest MCXXX file from Backup Location to working Direcotry.
                    ' Working directory value should be defined in Registry
                    
                    If giDebugLevel > 1 Then
                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get Latest MCXXX file from Network Path Start"
                    End If
                    ' ---------------------------------------
                    ' Get Latest MCXXX file from Network Path
                    ' ---------------------------------------
                    ' Invoke Common Method.CopyFiles() to copy latest MCXXX file from Backup Location to working Direcotry.
                    ' Working directory value should be defined in Registry
                    lstrCopyFileResponse = CopyFiles(lstrBackup_Location & "BK" & GetDateFormat(Date) & "." & lstrMC_CODE & ".csv", lstrWorkingDirectory & lstrMC_CODE & ".csv")
    
                    'if error occurred, raise clarify case
                    If lstrCopyFileResponse <> "" Then
                        lstrErrorDetails = "Error while getting latest " & lstrMC_CODE & ".csv to working " & "directory " & _
                                            lstrWorkingDirectory & " : " & lstrCopyFileResponse
    
                    If giDebugLevel > 1 Then
                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                    End If
                        GoTo ResumeNextMCFile
                    End If
    
                    If giDebugLevel > 1 Then
                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get Latest MCXXX file from Network Path Start END"
                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Copy MCXXX.ini file as schema.ini file in working direcotry Start"
                    End If
                    ' -----------------------------------------------------------
                    ' Copy MCXXX.ini file as schema.ini file in working direcotry
                    ' -----------------------------------------------------------
                    ' Invoke Common Method.CopyFiles()
                    lstrCopyFileResponse = CopyFiles(lstrWorkingDirectory & lstrMC_CODE & ".ini", lstrWorkingDirectory & "schema.ini")
    
                    'if any error occurred, raise clarify case
                    If lstrCopyFileResponse <> "" Then
                        'raise clarify case
                        'schema file not present for MCXXX.csv file.
                        'Send notification to business MCXXX file not updated for Process Date.
                        'Details: Schema file not defined for MCXXX.csv
    
                        lstrErrorDetails = "Schema file not present for " & lstrMC_CODE & ".csv" & vbCrLf & _
                                            "Error while defining schema file : " & lstrCopyFileResponse
                        If giDebugLevel > 1 Then
                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                        End If
                        
                        GoTo ResumeNextMCFile
                    End If
                    
                    If giDebugLevel > 1 Then
                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Copy MCXXX.ini file as schema.ini file in working direcotry END"
                    End If
    
                    If lcnCSVConnection.State = adStateOpen Then
                        If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
    
                        'get first data row from MC file, to get date field name
                        lrsCSVRecordset.Open "SELECT TOP 1 * FROM " & lstrMC_CODE & ".csv", lcnCSVConnection, adOpenKeyset, adLockReadOnly
    
                        If lrsCSVRecordset.EOF = False Then lstrProcessDateFieldName = Trim(lrsCSVRecordset.Fields(0).Name)
    
                        lrsCSVRecordset.Close
    
                        If giDebugLevel > 1 Then
                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Delete any record for the Process date if already exist in the Recordset START:-"
                        End If
                        
                        ' ------------------------------------------------------------------------
                        ' Delete any record for the Process date if already exist in the Recordset
                        ' ------------------------------------------------------------------------
                        If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                        lrsCSVRecordset.Open "SELECT TOP 1 * FROM " & lstrMC_CODE & ".csv WHERE " & lstrProcessDateFieldName & " = #" & CDate(lstrProcessDate) & "#", lcnCSVConnection
    
                        If lrsCSVRecordset.EOF = False Then
                            If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
    
                            lbDataDeletionFlag = UpdateFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrProcessDate)
    
                            If lbDataDeletionFlag = False Then
                                lstrErrorDetails = "Error while deleting records from " & lstrMC_CODE & ".csv for Process Date : " & lstrProcessDate
                                If giDebugLevel > 1 Then
                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                End If
                                GoTo ResumeNextMCFile
                            End If
                        End If
    
                        If giDebugLevel > 1 Then
                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Delete any record for the Process date if already exist in the Recordset END"
                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Load the CSV file from Working DIR into an ADO object. STRAT:-"
                        End If
                        ' ------------------------------------------------------
                        ' Load the CSV file from Working DIR into an ADO object.
                        ' ------------------------------------------------------
                        If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
    
                        lstrSqlQry = "SELECT * FROM " & lstrMC_CODE & ".csv Order By " & lstrProcessDateFieldName
    
                        lrsCSVRecordset.Open lstrSqlQry, lcnCSVConnection, adOpenKeyset, adLockReadOnly
    
                        ' ----------------------------------------------------
                        ' Check if Process date = 01/01 and Rs.Recordcount > 0
                        ' ----------------------------------------------------
                        If DatePart("d", lstrProcessDate) = 1 And DatePart("m", lstrProcessDate) = 1 And lrsCSVRecordset.EOF = False Then
                            'Yes
                            ' -------------------------------------------------------
                            ' Save MCXXX.csv file to FTP location as MCXXX & YYYY.csv
                            ' -------------------------------------------------------
                            If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                            
                            lstrCopyFileResponse = CopyFiles(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrFTP_Location & lstrFTP_Directory & lstrMC_CODE & CStr(CInt(DatePart("yyyy", lstrProcessDate)) - 1) & ".csv")
                            
'                            If lobjcFTP.Connected Then lobjcFTP.Disconnect
'                            lobjcFTP.Connect lstrFTP_Location, lstrFTP_User, lstrFTP_Password
'                            lobjcFTP.Directory = lstrFTP_Directory
'                            lstrCopyFileResponse = lobjcFTP.PutFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrMC_CODE & CStr(CInt(DatePart("yyyy", lstrProcessDate)) - 1) & ".csv")
'                            lobjcFTP.Disconnect
    
                            'if any error occurred, raise clarify case
                            If lstrCopyFileResponse <> "" Then
                                lstrErrorDetails = "Error while saving " & lstrMC_CODE & ".csv file as " & lstrMC_CODE & _
                                                    CStr(CInt(DatePart("yyyy", lstrProcessDate)) - 1) & ".csv on FTP location : " & lstrCopyFileResponse
    
                                If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                If giDebugLevel > 1 Then
                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                End If
                                GoTo ResumeNextMCFile
                            End If
    
                            ' -----------------------------------
                            ' Delete all the data from ADO object
                            ' -----------------------------------
                            lbDataDeletionFlag = UpdateFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", "")
    
                            If lbDataDeletionFlag = False Then
                                lstrErrorDetails = "Error creating blank " & lstrMC_CODE & ".csv for current year"
                                If giDebugLevel > 1 Then
                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                End If
                                GoTo ResumeNextMCFile
                            End If
                        'End
                        End If
                        
                        If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                        lrsCSVRecordset.Open "SELECT TOP 1 * FROM " & lstrMC_CODE & ".csv WHERE " & lstrProcessDateFieldName & " = #" & CDate(lstrProcessDate) & "#", lcnCSVConnection
    
                        If giDebugLevel > 1 Then
                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get Index Rate information for the MC File being processed Start:-"
                        End If
                        
                        
                        ' ----------------------------------------------------------
                        ' Get Index Rate information for the MC File being processed
                        ' ----------------------------------------------------------
                        ' Invoke BSMoneyCost.GetIndexRates(SQ_MC_ID, Process Date)
                        lstrIndexRateReqXml = "<INDEX_RATE_REQUEST>" & _
                                                    "<SQ_MC_ID>" & liSQ_MC_ID & "</SQ_MC_ID>" & _
                                                    "<PROCESS_DATE>" & lstrProcessDate & "</PROCESS_DATE>" & _
                                                "</INDEX_RATE_REQUEST>"
    
                        lstrIndexRateRespXml = lobjBSMCIMoneyCostService.GetIndexRates(lstrIndexRateReqXml)
    
                        'validate, if Index Rate Response XML is well-formed
                        If Not lobjIndexRateDOM.loadXML(lstrIndexRateRespXml) Then
                            'if error, raise clarify case
                            lstrErrorDetails = "Error while loading Response XML from lobjBSMCIMoneyCostService.GetIndexRates " & _
                                                "method." & vbCrLf & "Error no. : " & lobjIndexRateDOM.parseError.errorCode & _
                                                vbCrLf & "Error description : " & lobjIndexRateDOM.parseError.reason
    
                            If giDebugLevel > 1 Then
                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                            End If
                            GoTo ResumeNextMCFile
                        End If
    
                        If giDebugLevel > 1 Then
                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get Index Rate information for the MC File being processed End:-"
                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Check whether no. of columns in CSV and Money Cost DB match Start:-"
                        End If
                        '-------------------------------------------------------------
                        ' Check whether no. of columns in CSV and Money Cost DB match.
                        ' if not, raise error
                        '-------------------------------------------------------------
                        Set lobjIndexRateNodeList = lobjIndexRateDOM.selectNodes("/INDEX_RATE_RESPONSE/INDEX_RATESet/INDEX_RATE")
    
                        If Val(lobjIndexRateNodeList.length) <> lrsCSVRecordset.Fields.Count - 1 Then
                            'get list of all Index Codes and Index Terms, available in Money Cost DB
                            For liCounter2 = 0 To lobjIndexRateNodeList.length - 1
                                Set lobjIndexRateNode = lobjIndexRateNodeList.Item(liCounter2)
    
                                lstrMCDBAllIndexCode = lstrMCDBAllIndexCode & "," & Trim(lobjIndexRateNode.selectNodes("INDEX_CODE").Item(0).Text)
                                lstrMCDBAllIndexTerm = lstrMCDBAllIndexTerm & "," & Trim(lobjIndexRateNode.selectNodes("INDEX_TERM").Item(0).Text)
                            Next
    
                            'remove first comma
                            If Left(lstrMCDBAllIndexCode, 1) = "," Then lstrMCDBAllIndexCode = Mid(lstrMCDBAllIndexCode, 2, Len(lstrMCDBAllIndexCode))
                            If Left(lstrMCDBAllIndexTerm, 1) = "," Then lstrMCDBAllIndexTerm = Mid(lstrMCDBAllIndexTerm, 2, Len(lstrMCDBAllIndexTerm))
    
                            lstrCSVColumnHeaders = ""
    
                            'get comma separated list of Money Cost file header
                            For liCounter2 = 0 To lrsCSVRecordset.Fields.Count - 1
                                lstrCSVColumnHeaders = lstrCSVColumnHeaders & Trim(lrsCSVRecordset.Fields(Val(liCounter2)).Name) & ","
                            Next
    
                            If Right(lstrCSVColumnHeaders, 1) = "," Then lstrCSVColumnHeaders = Left(lstrCSVColumnHeaders, Len(lstrCSVColumnHeaders) - 1)
    
                            'raise clarify case
                            lstrErrorDetails = "Columns defined in schema file not matching to what are defined in Money Cost DB" & _
                                                vbCrLf & "CSV column headers : " & lstrCSVColumnHeaders & vbCrLf & "Index Codes " & _
                                                "in Money Cost DB: " & lstrMCDBAllIndexCode & vbCrLf & "Index Terms in Money Cost " & _
                                                "DB : " & lstrMCDBAllIndexTerm
                            If giDebugLevel > 1 Then
                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                            End If
                            GoTo ResumeNextMCFile
                        End If
    
                        If giDebugLevel > 1 Then
                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Check whether no. of columns in CSV and Money Cost DB match END:-"
                        End If
                        
                        
                        'get comma separated list of all unique Index Codes and Index Terms from MoneyCost DB
                        Set lobjIndexRateNodeList = lobjIndexRateDOM.selectNodes("/INDEX_RATE_RESPONSE/INDEX_RATESet/INDEX_RATE[IND_QUERYDB = 'True']")
    
                        For liCounter2 = 0 To lobjIndexRateNodeList.length - 1
                            Set lobjIndexRateNode = lobjIndexRateNodeList.Item(liCounter2)
    
                            If InStr("," + lstrINDEX_CODEList + ",", ",'" + UCase(Trim(lobjIndexRateNode.selectNodes("INDEX_CODE").Item(0).Text)) + "',") <= 0 Then
                                lstrINDEX_CODEList = lstrINDEX_CODEList + ",'" + UCase(Trim(lobjIndexRateNode.selectNodes("INDEX_CODE").Item(0).Text)) + "'"
                            End If
    
                            If InStr("," + lstrINDEX_TERMList + ",", ",'" + UCase(Trim(lobjIndexRateNode.selectNodes("INDEX_TERM").Item(0).Text)) + "',") <= 0 Then
                                lstrINDEX_TERMList = lstrINDEX_TERMList + ",'" + UCase(Trim(lobjIndexRateNode.selectNodes("INDEX_TERM").Item(0).Text)) + "'"
                            End If
                        Next
    
                        'remove first comma
                        If Left(lstrINDEX_CODEList, 1) = "," Then lstrINDEX_CODEList = Mid(lstrINDEX_CODEList, 2, Len(lstrINDEX_CODEList))
                        If Left(lstrINDEX_TERMList, 1) = "," Then lstrINDEX_TERMList = Mid(lstrINDEX_TERMList, 2, Len(lstrINDEX_TERMList))
    
                        Do While True
                            If giDebugLevel > 1 Then
                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Generate CurveTypes list to pull Data from DW only with DB query flag = 1 START:-"
                            End If
                            ' -------------------------------------------------------------------------
                            ' Generate CurveTypes list to pull Data from DW only with DB query flag = 1
                            ' -------------------------------------------------------------------------
                            lstrIndexDataReqXml = "<INDEX_DATA_REQUEST>" & _
                                                    "<PROCESSING_DATE>" & lstrProcessDate & "</PROCESSING_DATE>" & _
                                                    "<CURRENCY_CODE>" & lstrCURRENCY_CODE & "</CURRENCY_CODE>" & _
                                                    "<YIELD_CURVE_TYPE_LIST>" & lstrINDEX_CODEList & "</YIELD_CURVE_TYPE_LIST>" & _
                                                    "<DWC_TERM_PERIOD_LIST>" & lstrINDEX_TERMList & "</DWC_TERM_PERIOD_LIST>" & _
                                                    "<LAST_UPDATED_IND>" & lstrLAST_UPDATED_IND & "</LAST_UPDATED_IND>" & _
                                                  "</INDEX_DATA_REQUEST>"
    
                            ' Invoke cMoneyCostAutoService.GetIndexData(CurveTypes)
                            lstrIndexDataRespXml = UCase(lobjcMoneyCostAutoService.GetIndexData(lstrIndexDataReqXml))
    
                            'validate, if response xml is well-formed
                            If Not lobjIndexDataDOMXml.loadXML(lstrIndexDataRespXml) Then
    
                                'if error, raise clarify case
                                lstrErrorDetails = "Error, while loading Response XML from lobjcMoneyCostAutoService.GetIndexData " & _
                                                    "method. " & vbCrLf & "Error No. : " & lobjIndexDataDOMXml.parseError.errorCode & _
                                                    vbCrLf & "Error Description : " & lobjIndexDataDOMXml.parseError.reason
                                If giDebugLevel > 1 Then
                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                End If
                                GoTo ResumeNextMCFile
                            End If
    
                            If giDebugLevel > 1 Then
                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): loading Response XML from lobjcMoneyCostAutoService.GetIndexData END"
                            End If
                            
                            Set lobjIndexDataNodeList = lobjIndexDataDOMXml.selectNodes("/INDEX_DATA_RESPONSE/INDEX_DATASET/INDEX_DATA")
    
                            If giDebugLevel > 1 Then
                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Check if data was fetched from DW START:-"
                            End If
                            
                            'Check if data was fetched from DW
                            If lobjIndexDataNodeList.length > 0 Then
                            'True
    
                                If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
    
                                ' ------------------------------------------------------------------------
                                ' Delete any record for the Process date if already exist in the Recordset
                                ' ------------------------------------------------------------------------
                                If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close
                                lrsCSVDeleteRecordset.Open "SELECT TOP 1 * FROM " & lstrMC_CODE & ".csv WHERE " & lstrProcessDateFieldName & " = #" & CDate(lstrProcessDate) & "#", lcnCSVConnection
    
                                If lrsCSVDeleteRecordset.EOF = False Then
                                    If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close
    
                                    lbDataDeletionFlag = UpdateFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", Format(lstrProcessDate, "mm/dd/yyyy"))
    
                                    If lbDataDeletionFlag = False Then
                                        lstrErrorDetails = "Error while deleting records from " & lstrMC_CODE & ".csv for Process Date : " & lstrProcessDate
                                        If giDebugLevel > 1 Then
                                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                        End If
                                        GoTo ResumeNextMCFile
                                    End If
                                Else
                                    If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close
                                End If

                                If giDebugLevel > 1 Then
                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Load the CSV file from Working DIR into an ADO object. START:-"
                                End If
                                
                                ' ------------------------------------------------------
                                ' Load the CSV file from Working DIR into an ADO object.
                                ' ------------------------------------------------------
    
                                lstrSqlQry = "SELECT * FROM " & lstrMC_CODE & ".csv Order By " & lstrProcessDateFieldName
    
                                lrsCSVRecordset.Open lstrSqlQry, lcnCSVConnection, adOpenKeyset, adLockReadOnly
    
                                'Reach the Last record in the existing MC recordsset. Value from this record would be picked
                                'to fill missing data in new row being inserted.
                                lrsCSVRecordset.MoveLast
    
                                'For each Index Rate from MC_FILE with DB_Query = 1
                                For liCounter2 = 0 To lobjIndexRateNodeList.length - 1
                                    Set lobjIndexRateNode = lobjIndexRateNodeList.Item(liCounter2)
    
                                    ' -----------------------------------
                                    ' If Index Rate found in DW resultset
                                    ' -----------------------------------
                                    Set lobjIndexDataNodeList = lobjIndexDataDOMXml.selectNodes("/INDEX_DATA_RESPONSE/INDEX_DATASET/INDEX_DATA[YIELD_CURVE_TYPE = '" & UCase(Trim(lobjIndexRateNode.selectNodes("INDEX_CODE").Item(0).Text)) & "' and TERM_PERIOD = '" & UCase(Trim(lobjIndexRateNode.selectNodes("INDEX_TERM").Item(0).Text)) & "']")
    
                                    If lobjIndexDataNodeList.length > 0 Then
                                    'True
                                        Set lobjIndexDataNode = lobjIndexDataNodeList.Item(0)
    
                                        ' Get rate fetched from DW.
                                        lstrINTEREST_RATE_DWH = lobjIndexDataNode.selectNodes("INTEREST_RATE").Item(0).Text
    
                                        'check if the interest rate returned from DWH is not blank...
                                        '---------------------------------------
                                        If IsNumeric(lstrINTEREST_RATE_DWH) Then
                                            If UCase(Trim(lobjIndexRateNode.selectNodes("IND_PERCENTILE").Item(0).Text)) = "TRUE" Then
                                                ' Apply Percentile for the Index Rates requiring the same.
                                                lstrINTEREST_RATE_DWH = Val(lstrINTEREST_RATE_DWH) / 100
                                            End If
        
                                            ' Add Adder value defined for each Index Rate to the processed value above
                                            lstrINTEREST_RATE_DWH = Val(lstrINTEREST_RATE_DWH) + Val(lobjIndexRateNode.selectNodes("AMT_ADDER").Item(0).Text)
                                        Else
                                            'if the interest rate returned from the DWH is blank or some non-numeric value , then check for previous day's index data required or not
                                            'if IND_PREV_INDEXRATES_REQ = 0 then not required to copy previous day's index data
                                            If Trim(lobjIndexRateNode.selectNodes("IND_PREV_INDEXRATES_REQ").Item(0).Text) = "0" Then
                                                lstrMissingYieldCurve_NotCopied = lstrMissingYieldCurve_NotCopied & "," & Trim(lobjIndexRateNode.selectNodes("DESCRIPTION").Item(0).Text)
                                            Else
                                        'if the interest rate returned from the DWH is blank or some non-numeric value , then copy the previous day's index data.
                                        '---------------------------------------
                                            ' Get the value for Index Rate from Previous days record in MC File Recordset
                                                If lrsCSVRecordset.EOF = False Then
                                                    If IsNull(lrsCSVRecordset.Fields(Val(lobjIndexRateNode.selectNodes("MC_FILE_COL_POSITION").Item(0).Text)).Value) = False Then
                                                        lstrINTEREST_RATE_DWH = lrsCSVRecordset.Fields(Val(lobjIndexRateNode.selectNodes("MC_FILE_COL_POSITION").Item(0).Text)).Value
                                                    Else
                                                        lstrINTEREST_RATE_DWH = ""
                                                    End If
                                                End If
                                            '------------------------------------------------------------------
                                            'This is the case where the tag is formed, but has a blank or a non numeric value
                                            'On this index rate, update the adder amount with the latest adder amount.
                                            'Temporarily substracting the last adder amount and then adding the current adder amount.
                                            'this logic will be permanently fixed in the .net version.
                                            
                                            'Hit SQL Server again to get the adder amount for the previous day.
                                            'Invoke BSMoneyCost.GetIndexRates(SQ_MC_ID, Process Date)
                                                lstrIndexRateReqXml = "<INDEX_RATE_REQUEST>" & _
                                                                            "<SQ_MC_ID>" & liSQ_MC_ID & "</SQ_MC_ID>" & _
                                                                            "<PROCESS_DATE>" & lstrProcessDate_PrevDay & "</PROCESS_DATE>" & _
                                                                        "</INDEX_RATE_REQUEST>"
                            
                                                lstrIndexRateRespXml_PrevDay = lobjBSMCIMoneyCostService.GetIndexRates(lstrIndexRateReqXml)
                                                'validate, if Index Rate Response XML is well-formed
                                                If Not lobjIndexRateDOM_PrevDay.loadXML(lstrIndexRateRespXml_PrevDay) Then
                                                    'if error, raise clarify case
                                                    lstrErrorDetails = "Error while loading Response XML from lobjBSMCIMoneyCostService.GetIndexRates " & _
                                                                        "method for Previous Day." & vbCrLf & "Error no. : " & lobjIndexRateDOM.parseError.errorCode & _
                                                                        vbCrLf & "Error description : " & lobjIndexRateDOM_PrevDay.parseError.reason
                                                If giDebugLevel > 1 Then
                                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                                End If
                                                    GoTo ResumeNextMCFile
                                                End If
                                                   
                                                'Not Checking whether no. of columns in CSV and Money Cost DB match as it has already been checked in the first call.
                                            
                                                'get comma separated list of all unique Index Codes and Index Terms from MoneyCost DB
                                                Set lobjIndexRateNodeList_PrevDay = lobjIndexRateDOM_PrevDay.selectNodes("/INDEX_RATE_RESPONSE/INDEX_RATESet/INDEX_RATE[IND_QUERYDB = 'True']")
                            
                                                For liCounter3 = 0 To lobjIndexRateNodeList_PrevDay.length - 1
                                                    Set lobjIndexRateNode_PrevDay = lobjIndexRateNodeList_PrevDay.Item(liCounter3)
                                                      If liCounter2 = liCounter3 Then
                                                       lstrINTEREST_RATE_DWH = Val(lstrINTEREST_RATE_DWH) - Val(lobjIndexRateNode_PrevDay.selectNodes("AMT_ADDER").Item(0).Text) + Val(lobjIndexRateNode.selectNodes("AMT_ADDER").Item(0).Text)
                                                      End If
                                                Next
                                            
                                            '------------------------------------------------------------------
                                                ' Maintain a list of all such Index Rates and the values copied.
                                                lstrMissingYieldCurve = lstrMissingYieldCurve & "," & Trim(lobjIndexRateNode.selectNodes("DESCRIPTION").Item(0).Text)
                                                lstrMissingInterestRate = lstrMissingInterestRate & "," & lstrINTEREST_RATE_DWH
                                            End If
                                        '---------------------------------------
                                        End If
                                    'Else
                                    Else
                                    If giDebugLevel > 1 Then
                                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): if the interest rate returned from the DWH is blank or some non-numeric value , then check for previous day's index data required or not"
                                    End If
                                        'if the interest rate returned from the DWH is blank or some non-numeric value , then check for previous day's index data required or not
                                            'if IND_PREV_INDEXRATES_REQ = 0 then not required to copy previous day's index data
                                            If Trim(lobjIndexRateNode.selectNodes("IND_PREV_INDEXRATES_REQ").Item(0).Text) = "0" Then
                                                lstrMissingYieldCurve_NotCopied = lstrMissingYieldCurve_NotCopied & "," & Trim(lobjIndexRateNode.selectNodes("DESCRIPTION").Item(0).Text)
                                            Else
                                        'if the interest rate returned from the DWH is blank or some non-numeric value , then copy the previous day's index data.
                                        '---------------------------------------
                                            ' Get the value for Index Rate from Previous days record in MC File Recordset
                                                If lrsCSVRecordset.EOF = False Then
                                                    If IsNull(lrsCSVRecordset.Fields(Val(lobjIndexRateNode.selectNodes("MC_FILE_COL_POSITION").Item(0).Text)).Value) = False Then
                                                        lstrINTEREST_RATE_DWH = lrsCSVRecordset.Fields(Val(lobjIndexRateNode.selectNodes("MC_FILE_COL_POSITION").Item(0).Text)).Value
                                                    Else
                                                        lstrINTEREST_RATE_DWH = ""
                                                    End If
                                                End If
            
                                                '------------------------------------------------------------------
                                            'This is the case where no tag is formed for the specific yield curve as the data is missing. This is the ideal condition.
                                            'On this index rate, update the adder amount with the latest adder amount.
                                            'Temporarily substracting the last adder amount and then adding the current adder amount.
                                            'this logic will be permanently fixed in the .net version.
                                            
                                            'Hit SQL Server again to get the adder amount for the previous day.
                                            'Invoke BSMoneyCost.GetIndexRates(SQ_MC_ID, Process Date)
                                                lstrIndexRateReqXml = "<INDEX_RATE_REQUEST>" & _
                                                                            "<SQ_MC_ID>" & liSQ_MC_ID & "</SQ_MC_ID>" & _
                                                                            "<PROCESS_DATE>" & lstrProcessDate_PrevDay & "</PROCESS_DATE>" & _
                                                                        "</INDEX_RATE_REQUEST>"
                            
                                                lstrIndexRateRespXml_PrevDay = lobjBSMCIMoneyCostService.GetIndexRates(lstrIndexRateReqXml)
                                                'validate, if Index Rate Response XML is well-formed
                                                If Not lobjIndexRateDOM_PrevDay.loadXML(lstrIndexRateRespXml_PrevDay) Then
                                                    'if error, raise clarify case
                                                    lstrErrorDetails = "Error while loading Response XML from lobjBSMCIMoneyCostService.GetIndexRates " & _
                                                                        "method for Previous Day." & vbCrLf & "Error no. : " & lobjIndexRateDOM.parseError.errorCode & _
                                                                        vbCrLf & "Error description : " & lobjIndexRateDOM_PrevDay.parseError.reason
                                                 If giDebugLevel > 1 Then
                                                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                                 End If
                                                    GoTo ResumeNextMCFile
                                                End If
                                                   
                                                'Not Checking whether no. of columns in CSV and Money Cost DB match as it has already been checked in the first call.
                                            
                                                'get comma separated list of all unique Index Codes and Index Terms from MoneyCost DB
                                                Set lobjIndexRateNodeList_PrevDay = lobjIndexRateDOM_PrevDay.selectNodes("/INDEX_RATE_RESPONSE/INDEX_RATESet/INDEX_RATE[IND_QUERYDB = 'True']")
                            
                                                For liCounter3 = 0 To lobjIndexRateNodeList_PrevDay.length - 1
                                                    Set lobjIndexRateNode_PrevDay = lobjIndexRateNodeList_PrevDay.Item(liCounter3)
                                                      If liCounter2 = liCounter3 Then
                                                       lstrINTEREST_RATE_DWH = Val(lstrINTEREST_RATE_DWH) - Val(lobjIndexRateNode_PrevDay.selectNodes("AMT_ADDER").Item(0).Text) + Val(lobjIndexRateNode.selectNodes("AMT_ADDER").Item(0).Text)
                                                      End If
                                                Next
                                            '------------------------------------------------------------------
                                                ' Maintain a list of all such Index Rates and the values copied.
                                                lstrMissingYieldCurve = lstrMissingYieldCurve & "," & Trim(lobjIndexRateNode.selectNodes("DESCRIPTION").Item(0).Text)
                                                lstrMissingInterestRate = lstrMissingInterestRate & "," & lstrINTEREST_RATE_DWH
                                            End If
                                    'Endif
                                    End If
                                    'If NOT lstrMissingYieldCurve_NotCopied = "" Then
                                        lstrCSV_INSERT_RECORDXml = lstrCSV_INSERT_RECORDXml & _
                                                                    "<CSV_INSERT_DATA>" & _
                                                                        "<COL_POSITION>" & Trim(lobjIndexRateNode.selectNodes("MC_FILE_COL_POSITION").Item(0).Text) & "</COL_POSITION>" & _
                                                                        "<INTEREST_RATE>" & lstrINTEREST_RATE_DWH & "</INTEREST_RATE>" & _
                                                                    "</CSV_INSERT_DATA>"
                                    'End If
                                'Next
                                Next
                            If giDebugLevel > 1 Then
                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): remove first comma START:-"
                            End If
                            'remove first comma
                            If Not lstrMissingYieldCurve_NotCopied = "" Then
                            
                            lstrMissingYieldCurve_NotCopied = lstrMissingYieldCurve_NotCopied & lstrMissingYieldCurve
                            
                                If Left(lstrMissingYieldCurve_NotCopied, 1) = "," Then lstrMissingYieldCurve_NotCopied = Mid(lstrMissingYieldCurve_NotCopied, 2, Len(lstrMissingYieldCurve_NotCopied))
      
                                'If data was found missing for some Index Rates, Create a p3 case for the respective queue and include the list of Index Rates missing.
                                    lstrErrorDetails = "Errors were reported while updating " & lstrMC_CODE & " file for " & Format(lstrProcessDate, "mm/dd/yyyy") & ". " & vbCrLf & _
                                                        "Interest rates for the following were missing. So " & lstrMC_CODE & " file for " & Format(lstrProcessDate, "mm/dd/yyyy") & " is not updated. " & vbCrLf & vbCrLf & _
                                                        "S.No." & vbTab & "Index Name" & vbCrLf & vbCrLf
    
                                    larrMissingIndexCode_NotCopied = Split(lstrMissingYieldCurve_NotCopied, ",")
                                    If giDebugLevel > 1 Then
                                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                    End If
                                    
                                    For liCounter2 = 0 To UBound(larrMissingIndexCode_NotCopied, 1)
                                        lstrErrorDetails = lstrErrorDetails & CStr(liCounter2 + 1) & ")" & vbTab & _
                                                            larrMissingIndexCode_NotCopied(liCounter2) & vbCrLf
                                    Next
                                    
                                    SendErrNotification lstrClarifyQName, lstrBusinessContact, lstrErrorDetails, False, True, lstrMC_CODE, lstrProcessDate
                                    'lstrErrorDetails = ""
                                    '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
                                    'close the connection, if not closed
                                    If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                    If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close
        
                                    'Copy Working File back to the FTP Location.
                                    lstrCopyFileResponse = CopyFiles(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrFTP_Location & lstrFTP_Directory & lstrMC_CODE & ".csv")
                                    
'                                    If lobjcFTP.Connected Then lobjcFTP.Disconnect
'                                    lobjcFTP.Connect lstrFTP_Location, lstrFTP_User, lstrFTP_Password
'                                    lobjcFTP.Directory = lstrFTP_Directory
'                                    lstrCopyFileResponse = lobjcFTP.PutFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrMC_CODE & ".csv")
'                                    lobjcFTP.Disconnect
        
                                    'if any error occurred, raise clarify case
                                    If lstrCopyFileResponse <> "" Then
                                        lstrErrorDetails = lstrErrorDetails & "Error while updating " & lstrMC_CODE & ".csv file as on the FTP location : " & lstrFTP_Location & " : " & lstrCopyFileResponse
                                        If giDebugLevel > 1 Then
                                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                        End If
                                        If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                        GoTo ResumeNextMCFile
                                    End If
        
                                    'Copy working File over network location, if Network Location is defined in Registry
                                    If lstrNetwork_Location <> "" Then
                                        lstrCopyFileResponse = CopyFiles(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrNetwork_Location & lstrMC_CODE & ".csv")
        
                                        'if any error occurred, raise clarify case
                                        If lstrCopyFileResponse <> "" Then
                                            lstrErrorDetails = lstrErrorDetails & "Error while saving " & lstrMC_CODE & ".csv to network location : " & _
                                                                lstrNetwork_Location & " as " & lstrMC_CODE & ".csv" & vbCrLf & _
                                                                lstrCopyFileResponse
                                        If giDebugLevel > 1 Then
                                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                        End If
                                            GoTo ResumeNextMCFile
                                        End If
                                    End If
                                    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
                                    
                                    GoTo ResumeNextMCFile
                            End If
                            
                               
                                'remove first comma
                                If Left(lstrMissingYieldCurve, 1) = "," Then lstrMissingYieldCurve = Mid(lstrMissingYieldCurve, 2, Len(lstrMissingYieldCurve))
                                If Left(lstrMissingInterestRate, 1) = "," Then lstrMissingInterestRate = Mid(lstrMissingInterestRate, 2, Len(lstrMissingInterestRate))
                                
                                 If giDebugLevel > 1 Then
                                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): remove first comma END"
                                 End If
                                'If data was found missing for some Index Rates, Create a p3 case for the respective queue and include the list of Index Rates missing.
                                If lstrMissingYieldCurve <> "" And lstrMissingInterestRate <> "" Then
    
                                    lstrErrorDetails = "Errors were reported while updating " & lstrMC_CODE & " file for " & Format(lstrProcessDate, "mm/dd/yyyy") & ". " & vbCrLf & _
                                                        "Interest rates for following were copied with prior days values. " & vbCrLf & vbCrLf & _
                                                        "S.No." & vbTab & "Index Name & Value" & vbCrLf & vbCrLf
                                    If giDebugLevel > 1 Then
                                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                    End If
                                    larrMissingIndexCode = Split(lstrMissingYieldCurve, ",")
                                    larrMissingIntRate = Split(lstrMissingInterestRate, ",")
    
                                    For liCounter2 = 0 To UBound(larrMissingIndexCode, 1)
                                        lstrErrorDetails = lstrErrorDetails & CStr(liCounter2 + 1) & ")" & vbTab & _
                                                            larrMissingIndexCode(liCounter2) & " = " & _
                                                            larrMissingIntRate(liCounter2) & vbCrLf
                                    Next
                                End If
                                
                                Set lobjIndexRateNodeList = lobjIndexRateDOM.selectNodes("/INDEX_RATE_RESPONSE/INDEX_RATESet/INDEX_RATE[IND_QUERYDB = 'False']")
    
                                If giDebugLevel > 1 Then
                                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): For each Index Rate from MC_FILE with DB_Query = 0"
                                End If
                                'For each Index Rate from MC_FILE with DB_Query = 0
                                For liCounter2 = 0 To lobjIndexRateNodeList.length - 1
                                    Set lobjIndexRateNode = lobjIndexRateNodeList.Item(liCounter2)
    
                                    ' Add Adder value defined for each Index Rate to the processed value above
                                    lstrINTEREST_RATE_DWH = Trim(lobjIndexRateNode.selectNodes("AMT_ADDER").Item(0).Text)
    
                                    lstrCSV_INSERT_RECORDXml = lstrCSV_INSERT_RECORDXml & _
                                                                "<CSV_INSERT_DATA>" & _
                                                                    "<COL_POSITION>" & Trim(lobjIndexRateNode.selectNodes("MC_FILE_COL_POSITION").Item(0).Text) & "</COL_POSITION>" & _
                                                                    "<INTEREST_RATE>" & lstrINTEREST_RATE_DWH & "</INTEREST_RATE>" & _
                                                                "</CSV_INSERT_DATA>"
                                'Next
                                Next
    
                                lstrCSV_INSERT_RECORDXml = lstrCSV_INSERT_RECORDXml & "</CSV_INSERT_DATASet></CSV_INSERT_DATA_REQUEST>"
    
                                'Arrange all Index Rates based upon the Column Position and prepare an Insert Statement
                                lstrCSV_INSERT_RECORDXml = fnSortXmlData(lstrCSV_INSERT_RECORDXml, lstrWorkingDirectory & "CSVInsertXMLSorting.xsl")
    
                                lstrCSV_INSERT_RECORDXml = Replace(lstrCSV_INSERT_RECORDXml, "<?xml version=""1.0""?>", "")
    
                                'validate, if dynamically build XMl for CSV insert is well-formed
                                If Not lobjCSVInsertDOMXml.loadXML(lstrCSV_INSERT_RECORDXml) Then
                                    lstrErrorDetails = "Error while loading request XML for CSV data insert. " & vbCrLf & _
                                                        "Error No. : " & lobjCSVInsertDOMXml.parseError.errorCode & vbCrLf & _
                                                        "Error Description : " & lobjCSVInsertDOMXml.parseError.reason
                                      If giDebugLevel > 1 Then
                                            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails
                                    End If
                                
                                    GoTo ResumeNextMCFile
                                End If
    
                                lstrCSVInsertSQL = "INSERT INTO " & lstrMC_CODE & ".csv VALUES('" & lstrProcessDate & "', "
    
                                'Insert new row to MC File Recordset
                                Set lobjCSVInsertNodeList = lobjCSVInsertDOMXml.selectNodes("/CSV_INSERT_DATA_REQUEST/CSV_INSERT_DATASet/CSV_INSERT_DATA")
    
                                For liCounter2 = 0 To lobjCSVInsertNodeList.length - 1
                                    Set lobjCSVInsertNode = lobjCSVInsertNodeList.Item(liCounter2)
    
                                    lstrCSVInsertSQL = lstrCSVInsertSQL & "'" & Trim(lobjCSVInsertNode.selectNodes("INTEREST_RATE").Item(0).Text) & "', "
                                Next
    
                                lstrCSVInsertSQL = Mid(lstrCSVInsertSQL, 1, Len(lstrCSVInsertSQL) - 2)
    
                                lstrCSVInsertSQL = lstrCSVInsertSQL & ");"
    
                                'Save ADO recordset back to the working file as MCXXX.csv.
                                lcnCSVConnection.Execute lstrCSVInsertSQL
    
                                'close the connection, if closed
                                If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close
    
                                'Copy Working File back to the FTP Location.
                                lstrCopyFileResponse = CopyFiles(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrFTP_Location & lstrFTP_Directory & lstrMC_CODE & ".csv")
                                
'                                If lobjcFTP.Connected Then lobjcFTP.Disconnect
'                                lobjcFTP.Connect lstrFTP_Location, lstrFTP_User, lstrFTP_Password
'                                lobjcFTP.Directory = lstrFTP_Directory
'                                lstrCopyFileResponse = lobjcFTP.PutFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrMC_CODE & ".csv")
'                                lobjcFTP.Disconnect
    
                                'if any error occurred, raise clarify case
                                If lstrCopyFileResponse <> "" Then
                                    lstrErrorDetails = "Error while updating " & lstrMC_CODE & ".csv file as on the FTP location : " & lstrFTP_Location & " : " & lstrCopyFileResponse
                                If giDebugLevel > 1 Then
                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails
                                End If
                                    If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                    GoTo ResumeNextMCFile
                                End If
    
                                If giDebugLevel > 1 Then
                                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Copy working File over network location, if Network Location is defined in Registry"
                                End If
                                'Copy working File over network location, if Network Location is defined in Registry
                                If lstrNetwork_Location <> "" Then
                                    lstrCopyFileResponse = CopyFiles(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrNetwork_Location & lstrMC_CODE & ".csv")
    
                                    'if any error occurred, raise clarify case
                                    If lstrCopyFileResponse <> "" Then
                                        lstrErrorDetails = "Error while saving " & lstrMC_CODE & ".csv to network location : " & _
                                                            lstrNetwork_Location & " as " & lstrMC_CODE & ".csv" & vbCrLf & _
                                                            lstrCopyFileResponse
                                      If giDebugLevel > 1 Then
                                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails
                                      End If
    
                                        GoTo ResumeNextMCFile
                                    End If
                                End If
    
                                If giDebugLevel > 1 Then
                                        WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Take the generated files , change the date format and move it to the specified location as defined in registry."
                                End If
                                
                                'Take the generated files , change the date format and move it to the specified location as defined in registry.
                                If (lstrFTP_LocationForNewDateFormat <> "") Then
                                
                                    If Right(lstrFTP_LocationForNewDateFormat, 1) <> "\" Then lstrFTP_LocationForNewDateFormat = lstrFTP_LocationForNewDateFormat & "\"
                                    If Right(lstrFTP_DirectoryForNewDateFormat, 1) <> "\" Then lstrFTP_DirectoryForNewDateFormat = lstrFTP_DirectoryForNewDateFormat & "\"
                                    
                                    'check if Date Formatting and copying is required for this particular Currency
                                    If lstrDateFormatRequired = True Then
                                    
                                        'Copy the new ini file with the extn of MCCode_AUD.ini which stores the date in dd/mm/yyyy format.
                                        lstrCopyFileResponse = CopyFiles(lstrWorkingDirectory & lstrMC_CODE & "_AUD.ini", lstrWorkingDirectory & "schema.ini")
                                        
                                        'if any error occurred, raise clarify case
                                        If lstrCopyFileResponse <> "" Then
                                            'raise clarify case
                                            'schema file not present for MCXXX.csv file.
                                            'Send notification to business MCXXX file not updated for Process Date.
                                            'Details: Schema file not defined for MCXXX.csv
                        
                                            lstrErrorDetails = "Schema file with new Date Format not present for " & lstrMC_CODE & ".csv" & vbCrLf & _
                                                                "Error while defining schema file : " & lstrCopyFileResponse
                                            If giDebugLevel > 1 Then
                                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails
                                            End If
                                            GoTo ResumeNextMCFile
                                        End If
                                    
                                    
                                        'copy from new ftp location to working directory
                                        lstrCopyFileResponse = CopyFiles(lstrFTP_LocationForNewDateFormat & lstrFTP_DirectoryForNewDateFormat & lstrMC_CODE & ".csv", lstrWorkingDirectory & lstrMC_CODE & ".csv")
                                        
'                                        If lobjcFTP.Connected Then lobjcFTP.Disconnect
'                                        lobjcFTP.Connect lstrFTP_LocationForNewDateFormat, lstrFTP_UserForNewDateFormat, lstrFTP_PasswordForNewDateFormat
'                                        lobjcFTP.Directory = lstrFTP_DirectoryForNewDateFormat
'                                        lstrCopyFileResponse = lobjcFTP.GetFile(lstrMC_CODE & ".csv", lstrWorkingDirectory & lstrMC_CODE & ".csv")
'                                        lobjcFTP.Disconnect
                        
                                        'if any error occurred, raise clarify case
                                        If lstrCopyFileResponse <> "" Then
                                            lstrErrorDetails = "Error while creating working copy of existing " & lstrMC_CODE & ".csv file as " & _
                                                                lstrMC_CODE & ".csv from New FTP on " & lstrWorkingDirectory & " : " & lstrCopyFileResponse
                                                                
                                            If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                            If giDebugLevel > 1 Then
                                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails
                                            End If
                                            GoTo ResumeNextMCFile
                                        End If
                                        If giDebugLevel > 1 Then
                                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Delete any record for the Process date if already exist in the Recordset START:-"
                                        End If
                                        
                                        ' ------------------------------------------------------------------------
                                        ' Delete any record for the Process date if already exist in the Recordset
                                        ' ------------------------------------------------------------------------
                                        If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                        lrsCSVRecordset.Open "SELECT TOP 1 * FROM " & lstrMC_CODE & ".csv WHERE " & lstrProcessDateFieldName & " = #" & Format(CDate(Day(lstrProcessDate) & "/" & Month(lstrProcessDate) & "/" & Year(lstrProcessDate)), "DD/MM/YYYY") & "#", lcnCSVConnection
                                        
                                        lstrDay_NewDateFormat = Day(lstrProcessDate)
                                        If Len(lstrDay_NewDateFormat) = 1 Then lstrDay_NewDateFormat = "0" & lstrDay_NewDateFormat
                                        lstrMonth_NewDateFormat = Month(lstrProcessDate)
                                        If Len(lstrMonth_NewDateFormat) = 1 Then lstrMonth_NewDateFormat = "0" & lstrMonth_NewDateFormat
                                        
                                        lstrProcessDate_NewDateFormat = lstrDay_NewDateFormat & "/" & lstrMonth_NewDateFormat & "/" & Year(lstrProcessDate)
                                        
                                        If lrsCSVRecordset.EOF = False Then
                                            If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                    
                                            lbDataDeletionFlag = UpdateFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrProcessDate_NewDateFormat)
                    
                                            If lbDataDeletionFlag = False Then
                                                lstrErrorDetails = "Error while deleting records from " & lstrMC_CODE & ".csv for Process Date : " & lstrProcessDate_NewDateFormat
                                                
                                                If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                            If giDebugLevel > 1 Then
                                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails
                                            End If
                                                GoTo ResumeNextMCFile
                                            End If
                                        End If
                                                       
                                        'Insert  with the new formatted date....
                                        lstrCSVInsertSQL = "INSERT INTO " & lstrMC_CODE & ".csv VALUES('" & Format(lstrProcessDate, lstrDateFormat) & "', "
                                        
                                        'Insert new row to MC File Recordset with the new formatted date.
                                        Set lobjCSVInsertNodeList = lobjCSVInsertDOMXml.selectNodes("/CSV_INSERT_DATA_REQUEST/CSV_INSERT_DATASet/CSV_INSERT_DATA")
                                        
                                        For liCounter2 = 0 To lobjCSVInsertNodeList.length - 1
                                        Set lobjCSVInsertNode = lobjCSVInsertNodeList.Item(liCounter2)
                                        
                                        lstrCSVInsertSQL = lstrCSVInsertSQL & "'" & Trim(lobjCSVInsertNode.selectNodes("INTEREST_RATE").Item(0).Text) & "', "
                                        Next
                                        
                                        lstrCSVInsertSQL = Mid(lstrCSVInsertSQL, 1, Len(lstrCSVInsertSQL) - 2)
                                        
                                        lstrCSVInsertSQL = lstrCSVInsertSQL & ");"
                                        
                                        lcnCSVConnection.Execute lstrCSVInsertSQL
                                        
                                        'close the connection, if closed
                                        If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                        If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close
                                        If giDebugLevel > 1 Then
                                                WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): copy this version of the file in the new FTP location. START:-"
                                        End If
                                        'copy this version of the file in the new FTP location.
                                        lstrCopyFileResponse = CopyFiles(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrFTP_LocationForNewDateFormat & lstrFTP_DirectoryForNewDateFormat & lstrMC_CODE & ".csv")
                                        
'                                        If lobjcFTP.Connected Then lobjcFTP.Disconnect
'                                        lobjcFTP.Connect lstrFTP_LocationForNewDateFormat, lstrFTP_UserForNewDateFormat, lstrFTP_PasswordForNewDateFormat
'                                        lobjcFTP.Directory = lstrFTP_DirectoryForNewDateFormat
'                                        lstrCopyFileResponse = lobjcFTP.PutFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrMC_CODE & ".csv")
'                                        lobjcFTP.Disconnect
                                        
                                        'if any error occurred, raise clarify case
                                        If lstrCopyFileResponse <> "" Then
                                            lstrErrorDetails = "Error while updating " & lstrMC_CODE & ".csv file as on the New FTP location : " & lstrFTP_LocationForNewDateFormat & " : " & lstrCopyFileResponse
                                            If giDebugLevel > 1 Then
                                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails
                                            End If
                                            
                                            If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                            GoTo ResumeNextMCFile
                                        End If
                                        
                                    End If
                                    
                                End If
    
    
                                'If data was found missing for some Index Rates, Create a p3 case for the respective queue
                                'and include the list of Index Rates missing.
                                If lstrMissingYieldCurve <> "" And lstrMissingInterestRate <> "" Then
                                    SendErrNotification lstrClarifyQName, lstrBusinessContact, lstrErrorDetails, False, True, lstrMC_CODE, lstrProcessDate
                                    lstrErrorDetails = ""
                                If giDebugLevel > 1 Then
                                    WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): If data was found missing for some Index Rates, Create a p3 case for the respective queue and include the list of Index Rates missing."
                                End If
                                    GoTo ResumeNextMCFile
                                End If
    
                                lstrErrorDetails = ""
    
                                Exit Do
    
                            Else
                                If liMARKET_CLOSED_DWH_CHECK_COUNTER = 0 Then
                                    If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                                    ' Assumption that Markets were closed for Process date. Stop the process for current file and skip to next file.
                                    lbMarketOpenFlag = False
    
                                    Exit Do
                                End If
    
                                lstrProcessDate = DateAdd("d", -1, lstrProcessDate)
                                Select Case Weekday(lstrProcessDate)
                                Case 7, 1 'Saturday , Sunday
                                    lstrProcessDate = DateAdd("d", -3, lstrProcessDate)
                                End Select
    
                                liMARKET_CLOSED_DWH_CHECK_COUNTER = liMARKET_CLOSED_DWH_CHECK_COUNTER - 1
                            End If
                        Loop
                    End If
    
                    'In case of any errors encountered during MC File processing. Skip to the next File
ResumeNextMCFile:
                    
                    If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
                    
                    lstrMCFileEndExec = Now()
                    
                    If lobjFSOExecute.FileExists(lstrWorkingDirectory & lstrMC_CODE & ".csv") Then
                        Kill lstrWorkingDirectory & lstrMC_CODE & ".csv"
                    End If
                    
                    If lobjFSOExecute.FileExists(lstrWorkingDirectory & lstrMC_CODE & ".csv") Then
                        Kill lstrWorkingDirectory & "schema.ini"
                    End If
    
                    'if some error found during processing of Money Cost file
                    If lstrErrorDetails <> "" Or Err.Description <> "" Then
                        'raise clarify ticket
                        lstrSendErrNotiResult = SendErrNotification(lstrClarifyQName, lstrBusinessContact, _
                                                IIf(lstrErrorDetails = "", "Error No. : " & Err.Number & vbCrLf & _
                                                "Error Description : " & Err.Description, lstrErrorDetails), True, _
                                                True, lstrMC_CODE, lstrProcessDate)
    
                        'update MC_LOGS table as per approriate status/ details
                        lstrUpdateMCLogReqXml = lstrUpdateMCLogReqXml & _
                                                    "<MC_FILE_LOG>" & _
                                                        "<SQ_MC_ID>" & liSQ_MC_ID & "</SQ_MC_ID>" & _
                                                        "<DATE_START>" & lstrMCFileStartExec & "</DATE_START>" & _
                                                        "<DATE_END>" & lstrMCFileEndExec & "</DATE_END>" & _
                                                        "<STATUS>1</STATUS>" & _
                                                        "<DETAILS>" & IIf(lstrErrorDetails = "", "Error No. : " & Err.Number & vbCrLf & "Error Description : " & Err.Description, lstrErrorDetails) & "</DETAILS>" & _
                                                    "</MC_FILE_LOG>"
    
                    'if no error occurred
                    Else
                        'if data fetched from DWH, update MC_LOGS table with success status
                        If lbMarketOpenFlag = True Then
                            lstrUpdateMCLogReqXml = lstrUpdateMCLogReqXml & _
                                                    "<MC_FILE_LOG>" & _
                                                        "<SQ_MC_ID>" & liSQ_MC_ID & "</SQ_MC_ID>" & _
                                                        "<DATE_START>" & lstrMCFileStartExec & "</DATE_START>" & _
                                                        "<DATE_END>" & lstrMCFileEndExec & "</DATE_END>" & _
                                                        "<STATUS>0</STATUS>" & _
                                                        "<DETAILS>Money Cost file processed successfully.</DETAILS>" & _
                                                    "</MC_FILE_LOG>"
    
                        'if no data found from DWH for the processing date, assume market was closed for the day,
                        'hence update MC_LOGS table with success status and details regarding non-availability of data
                        Else
                            lstrUpdateMCLogReqXml = lstrUpdateMCLogReqXml & _
                                                    "<MC_FILE_LOG>" & _
                                                        "<SQ_MC_ID>" & liSQ_MC_ID & "</SQ_MC_ID>" & _
                                                        "<DATE_START>" & lstrMCFileStartExec & "</DATE_START>" & _
                                                        "<DATE_END>" & lstrMCFileEndExec & "</DATE_END>" & _
                                                        "<STATUS>0</STATUS>" & _
                                                        "<DETAILS>No data found in the Data warehouse for " & lstrMC_CODE & _
                                                            ".csv file on Proessing Date : " & Format(lstrProcessDate, "mm/dd/yyyy") & "." & _
                                                        "</DETAILS>" & _
                                                    "</MC_FILE_LOG>"
    
                            'send notification email to business contacts, if no data found in DWH,
                            'provided execution day is not Saturday or Sunday
                            Select Case Weekday(Date)
                            Case 2, 3, 4, 5, 6  'all weekday Monday thru Friday
                                SendErrNotification lstrClarifyQName, lstrBusinessContact, "No data found in the Data warehouse for " & lstrMC_CODE & ".csv file on Proessing Date : " & Format(lstrProcessDate, "mm/dd/yyyy") & ".", False, True, lstrMC_CODE, lstrProcessDate
                            End Select
                        End If
                    End If
    
                    'append MC_FILE update request XML, with last schedules process date
                    lstrUpdateMCFileReqXml = lstrUpdateMCFileReqXml & _
                                                "<MC_FILE>" & _
                                                    "<SQ_MC_ID>" & liSQ_MC_ID & "</SQ_MC_ID>" & _
                                                    "<LAST_SCHEDULE_PROCESS_DATE>" & lstrScheduleProcessDate & "</LAST_SCHEDULE_PROCESS_DATE>" & _
                                                "</MC_FILE>"
                End If
                'FTP Location Mapped End
            End If

        'Next MC File
        Next

        lstrUpdateMCLogReqXml = lstrUpdateMCLogReqXml & "</MC_FILE_LOGSet></UPDATE_MC_FILE_LOGS_REQUEST>"
        lstrUpdateMCFileReqXml = lstrUpdateMCFileReqXml & "</MC_FILESet></UPDATE_MC_FILE_REQUEST>"

        'call BSMoneyCost component for updation of MC_LOGS table
        lstrUpdateMCLogRespXml = lobjBSMCIMoneyCostMgr.UpdateMCLogs(lstrUpdateMCLogReqXml)

        'call BSMoneyCost component for updation of MC_FILE table, for last schedule process date
        lstrUpdateMCFileRespXml = lobjBSMCIMoneyCostMgr.UpdateMCFile(lstrUpdateMCFileReqXml)
    Else
        GoTo ExecuteServiceFlow_ErrHandler
    End If

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "", ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

ExecuteServiceFlow_CleanMemory:
    'clear object variables and recordset variables from memory
    Set lobjAllMCFileDOM = Nothing
    Set lobjBSMCIMoneyCostService = Nothing
    Set lobjcMoneyCostAutoService = Nothing
    Set lobjCSVInsertDOMXml = Nothing
    Set lobjCSVInsertNode = Nothing
    Set lobjCSVInsertNodeList = Nothing
    Set lobjIndexDataDOMXml = Nothing
    Set lobjIndexDataNode = Nothing
    Set lobjIndexDataNodeList = Nothing
    Set lobjIndexRateDOM = Nothing
    Set lobjIndexRateNode = Nothing
    Set lobjIndexRateNodeList = Nothing
    Set lobjMCFileNode = Nothing
    Set lobjMCFileNodeList = Nothing
    
    If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close
    Set lrsCSVRecordset = Nothing

    If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close
    Set lrsCSVDeleteRecordset = Nothing

    If lcnCSVConnection.State = adStateOpen Then lcnCSVConnection.Close
    Set lcnCSVConnection = Nothing

    If lobjcFTP.Connected Then lobjcFTP.Disconnect
    Set lobjcFTP = Nothing
    
    Set lobjFSOExecute = Nothing

    Exit Sub

ExecuteServiceFlow_ErrHandler:
    If InStr(lstrUpdateMCLogReqXml, "</MC_FILE_LOGSet></UPDATE_MC_FILE_LOGS_REQUEST>") <= 0 Then
        lstrUpdateMCLogReqXml = lstrUpdateMCLogReqXml & "</MC_FILE_LOGSet></UPDATE_MC_FILE_LOGS_REQUEST>"
    End If

    'update MC_LOGS table in MoneyCost DB, using BSMoneyCost component
    'lstrUpdateMCLogRespXml = lobjBSMCIMoneyCostMgr.UpdateMCLogs(lstrUpdateMCLogReqXml)

    llErrNbr = Err.Number
    lstrErrSrc = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
    lstrErrDesc = Err.Description

    'raise clarify ticket
    lstrSendErrNotiResult = SendErrNotification(lstrClarifyQName, lstrBusinessContact, _
                            IIf(lstrCommonErrorDetails = "", llErrNbr & " : " & lstrErrDesc & " Source: " & lstrErrSrc, lstrCommonErrorDetails), True, False, lstrMC_CODE, lstrProcessDate)

    'write error to log file
     If giDebugLevel > 1 Then
            WriteToTextDebugFile cCOMPONENT_NAME & "." & cMODULE_NAME & ".ExecuteServiceFlow ", _
                        BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError
    End If
    If giDebugLevel > 1 Then
            WriteToTextDebugFile gstrDebugFile, "BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & llErrNbr & " " & lstrErrSrc & " " & lstrErrDesc
    End If
    
    Resume ExecuteServiceFlow_CleanMemory
End Sub

Private Function UpdateFile(ByVal astrFileName As String, ByVal astrProcessDate As String) As Boolean
On Error GoTo UpdateFile_ErrHandler

    Dim lFSOFileSystem      As New FileSystemObject
    Dim lTSFileStreamHandle As TextStream
    Dim lvarData()          As Variant
    Dim lstrSearch          As String
    Dim lstrText            As String
    Dim llLine              As Long
    Dim liCount             As Long

    lstrSearch = astrProcessDate
    Set lTSFileStreamHandle = lFSOFileSystem.OpenTextFile(astrFileName)

    llLine = 0

    lstrText = lTSFileStreamHandle.ReadLine
    ReDim Preserve lvarData(llLine)
    lvarData(llLine) = lstrText
    llLine = llLine + 1

    If lstrSearch <> "" Then
        While lTSFileStreamHandle.AtEndOfStream = False
            lstrText = lTSFileStreamHandle.ReadLine

            If InStr(1, lstrText, lstrSearch, vbBinaryCompare) = 0 Then
                ReDim Preserve lvarData(llLine)

                lvarData(llLine) = lstrText
                llLine = llLine + 1
            End If
        Wend
    End If

    lTSFileStreamHandle.Close

    If IsEmpty(lvarData) = True Then
        UpdateFile = True
        Exit Function
    End If

    Set lTSFileStreamHandle = lFSOFileSystem.OpenTextFile(astrFileName, ForWriting)

    For liCount = 0 To UBound(lvarData)
        lTSFileStreamHandle.WriteLine lvarData(liCount)
    Next

    lTSFileStreamHandle.Close

    UpdateFile = True
    Exit Function

UpdateFile_ErrHandler:
    lTSFileStreamHandle.Close
    UpdateFile = False
End Function

    Private Function IsFTPLocationMapped(ByVal astrFTP_Location As String, ByVal astrFTP_User As String, ByVal astrFTP_Password As String) As Boolean
        Dim lstrUNCPath As String
        Dim fso As New FileSystemObject
                
        lstrUNCPath = astrFTP_Location
        If fso.FolderExists(lstrUNCPath) Then
           IsFTPLocationMapped = True
        Else
            If MapDrive(lstrUNCPath, astrFTP_User, astrFTP_Password) Then
                IsFTPLocationMapped = True
            Else
                IsFTPLocationMapped = False
            End If
        End If
            
    End Function

    Public Function MapDrive(ByVal UNCPath As String, ByVal astrFTP_User As String, ByVal astrFTP_Password As String) As Boolean
    
        Dim N As NETRESOURCE
        Dim strUsername As String
        Dim strPassword As String
        Dim l As Long
    
        N.dwScope = RESOURCE_PUBLICNET
        N.dwType = RESOURCETYPE_PRINT
        N.dwDisplayType = RESOURCEDISPLAYTYPE_GENERIC
        N.dwUsage = RESOURCEUSAGE_CONNECTABLE
        N.lpLocalName = ""
        strUsername = astrFTP_User
        strPassword = astrFTP_Password
        N.lpRemoteName = UNCPath + Chr$(0)
        N.lpComment = Chr$(0)
        N.lpProvider = Chr$(0)
       
        l = WNetAddConnection2(N, strPassword, strUsername, 0)
        
            If l = 0 Then
                MapDrive = True
            Else
                MapDrive = False
            End If
    End Function

