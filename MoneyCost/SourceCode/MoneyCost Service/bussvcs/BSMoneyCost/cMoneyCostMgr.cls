VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMoneyCostMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================
' GE Capital Proprietary and Confidential
' Copyright (c) 2001-2002 by GE Capital - All rights reserved.
'
' This code may not be reproduced in any way without express
' permission from GE Capital.
'================================================================

'================================================================
'MODULE  : cMoneyCostMgr
'PURPOSE : This will contain all Money Cost related transactional
'          methods.
'================================================================

Option Explicit

'Constant for module name =======================================
Private Const cMODULE_NAME As String = "cMoneyCostMgr"
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

Dim lobjDataClass   As New cDataClass
Dim lrsTest         As New ADODB.Recordset

    'Execute the Test SQL statement which returns a count of the records
    Set lrsTest = lobjDataClass.Execute(ecTestSQL, ecRSExecuteRecords, "")

    'Return the total records
    Test = "Retrieved " & lrsTest.Fields(0).Value & " records."
    lrsTest.Close

Test_CleanMemory:
    Set lrsTest = Nothing
    Set lobjDataClass = Nothing
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

    Exit Sub

Class_ErrHandler:
    WriteToTextDebugFile "Class_Initialize()", BuildErrXML(Err.Number, Err.Source & cMODULE_NAME & ".Class_Initialize()", Err.Description), ecDebugCriticalError

    Resume Next
End Sub

'================================================================
'METHOD  : UpdateMCDetails
'PURPOSE : Wrapper method to update the details of selected Money
'          Cost file.
'PARMS   :
'          astrUpdateMCDetailsXML [String] = Input paramters
'          to update the selected Money Cost File details.
'          Sample XML structure:
'           <UPDATE_MC_FILE_DETAIL_REQUEST>
'               <USER_SSOID>500674629</USER_SSOID>
'               <MC_FILE_DETAILSet>
'                   <MC_FILE_DETAIL>
'                       <SQ_INDEX_ID>1</SQ_INDEX_ID>
'                       <AMT_ADDER>0.0814</AMT_ADDER>
'                       <DATE_EFFECTIVE>05/04/2005</DATE_EFFECTIVE>
'                   </MC_FILE_DETAIL>
'                   ....
'               </MC_FILE_DETAILSet>
'           </UPDATE_MC_FILE_DETAIL_REQUEST>
'RETURN  : String = status of MC File details updation, as XML string
'          Sample XML structure:
'           <UPDATE_MC_FILE_DETAIL_RESPONSE>
'               <STATUS>SUCCESS</STATUS>
'           </UPDATE_MC_FILE_DETAIL_RESPONSE>
'================================================================
Public Function UpdateMCDetails(ByVal astrUpdateMCDetailsXML As String) As String
On Error GoTo UpdateMCDetails_ErrHandler

Dim lstrMethodName  As String   'to store method name
Dim lstrErrSrc      As String   'to store error source
Dim lstrErrDesc     As String   'to store error description
Dim llErrNbr        As Long     'to store error number

    lstrMethodName = "UpdateMCDetails"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjcMoneyCostMgr       As New cMoneyCostMgr    'object of cMoneyCostMgr class module, to access its methods
Dim lobjcDataClass          As New cDataClass       'object of cDataClass to access its methods
Dim lobjRequestXmlDOM       As New DOMDocument40    'object of XML DOM to load input Request XML
Dim lobjMCFileDetailsNdLst  As IXMLDOMNodeList      'to store node list of MC_FILE_DETAIL node set
Dim lobjMCFileDetailElem    As IXMLDOMElement       'to store single node of MC_FILE_DETAIL node set
Dim llSQ_INDEX_ID           As Long                 'to store SQ_INDEX_ID for each Index row
Dim ldAMT_ADDER             As Double               'to store Adder Rate, needs to be updated
Dim lstrUSER_SSOID          As String               'to store logged-in user's SSO ID
Dim lstrBatchUpdateQry      As String               'to store dynamically batch update query
Dim lstrDATE_EFFECTIVE      As String               'to store Effective Date for Adder Rate, needs to be updated
Dim liCounter               As Integer              'to store loop counter

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrUpdateMCDetailsXML, ecDebugInputTrace
    End If

    'Validate if the Input Request XML is well-formed
    If Not lobjRequestXmlDOM.loadXML(astrUpdateMCDetailsXML) Then
        'Raise Error
        Err.Raise lobjRequestXmlDOM.parseError.errorCode, _
                    , "Error loading Input Request XML to UpdateMCDetails(). " & lobjRequestXmlDOM.parseError.reason
    End If

    'check, if SSO ID found in Input XML, store in local variable
    If lobjRequestXmlDOM.getElementsByTagName("USER_SSOID").length > 0 Then
        If lobjRequestXmlDOM.getElementsByTagName("USER_SSOID").Item(0).Text <> "" Then
            lstrUSER_SSOID = Trim(lobjRequestXmlDOM.getElementsByTagName("USER_SSOID").Item(0).Text)
        End If
    End If

    'get all node list of MC_FILE_DETAIL node set
    Set lobjMCFileDetailsNdLst = lobjRequestXmlDOM.selectNodes("/UPDATE_MC_FILE_DETAIL_REQUEST/MC_FILE_DETAILSet/MC_FILE_DETAIL")

    lstrBatchUpdateQry = ""

    'loop for each individual MC_FILE_DETAIL node set
    For liCounter = 0 To lobjMCFileDetailsNdLst.length - 1
        'get single node set of MC_FILE_DETAIL
        Set lobjMCFileDetailElem = lobjMCFileDetailsNdLst.Item(liCounter)

        're-initialize local variables
        llSQ_INDEX_ID = 0
        ldAMT_ADDER = 0
        lstrDATE_EFFECTIVE = ""

        'check, if SQ_INDEX_ID found in node set, store in local variable
        If lobjMCFileDetailElem.getElementsByTagName("SQ_INDEX_ID").length > 0 Then
            If lobjMCFileDetailElem.getElementsByTagName("SQ_INDEX_ID").Item(0).Text <> "" Then
                llSQ_INDEX_ID = lobjMCFileDetailElem.getElementsByTagName("SQ_INDEX_ID").Item(0).Text
            End If
        End If

        'check, if AMT_ADDER found in node set, store in local variable
        If lobjMCFileDetailElem.getElementsByTagName("AMT_ADDER").length > 0 Then
            If lobjMCFileDetailElem.getElementsByTagName("AMT_ADDER").Item(0).Text <> "" Then
                ldAMT_ADDER = lobjMCFileDetailElem.getElementsByTagName("AMT_ADDER").Item(0).Text
            Else
                Err.Raise cINVALID_PARMS, , "Adder Rate is not specified."
            End If
        Else
            Err.Raise cINVALID_PARMS, , "Adder Rate is not specified."
        End If

        'check, if DATE_EFFECTIVE found in node set, store in local variable
        If lobjMCFileDetailElem.getElementsByTagName("DATE_EFFECTIVE").length > 0 Then
            If lobjMCFileDetailElem.getElementsByTagName("DATE_EFFECTIVE").Item(0).Text <> "" Then
                lstrDATE_EFFECTIVE = Trim(lobjMCFileDetailElem.getElementsByTagName("DATE_EFFECTIVE").Item(0).Text)
            Else
                Err.Raise cINVALID_PARMS, , "Effective Date is not specified."
            End If
        Else
            Err.Raise cINVALID_PARMS, , "Effective Date is not specified."
        End If

        'dynamically build update statement
        lstrBatchUpdateQry = lstrBatchUpdateQry & "UPDATE INDEX_RATES SET AMT_ADDER = " & ldAMT_ADDER & ", " & _
                                                "SSO_UPDATE = '" & lstrUSER_SSOID & "' , " & _
                                                "DATE_EFFECTIVE = '" & lstrDATE_EFFECTIVE & "', DATE_UPDATE = " & _
                                                "GETDATE() WHERE SQ_INDEX_ID = " & llSQ_INDEX_ID & "; " & vbCrLf
    Next

    'if some data found, which needs to be updated/ inserted
    If lstrBatchUpdateQry <> "" Then
        'execute batch update SQL
        Call lobjcDataClass.Execute(ecExecuteSQL, ecRSExecuteNoRecords, lstrBatchUpdateQry)

        'return Output XML with status as SUCCESS
        UpdateMCDetails = "<UPDATE_MC_FILE_DETAIL_RESPONSE><STATUS>SUCCESS</STATUS></UPDATE_MC_FILE_DETAIL_RESPONSE>"
    Else
        'raise error, if no data found to update
        Err.Raise -1, , "No data found to update in UpdateMCDetails() method. "
    End If

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, UpdateMCDetails, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

UpdateMCDetails_CleanMemory:
    'clear object variables from memory
    Set lobjcDataClass = Nothing
    Set lobjcMoneyCostMgr = Nothing
    Set lobjMCFileDetailElem = Nothing
    Set lobjMCFileDetailsNdLst = Nothing
    Set lobjRequestXmlDOM = Nothing
    Exit Function

UpdateMCDetails_ErrHandler:
    llErrNbr = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    UpdateMCDetails = vbNullString

    'write error message to log file
    WriteToTextDebugFile cMODULE_NAME & lstrMethodName, BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    'raise error
    Err.Raise llErrNbr, lstrErrSrc, lstrErrDesc

    Resume UpdateMCDetails_CleanMemory
End Function

'================================================================
'METHOD  : UpdateMCLogs
'PURPOSE : Wrapper method to update the MC Logs, on execution of
'          BSSTMoneyCostAuto service
'PARMS   :
'          astrUpdateMCLogsXML [String] = Input paramters
'          to update the MC_Logs table
'          Sample XML structure:
'           <UPDATE_MC_FILE_LOGS_REQUEST>
'               <MC_FILE_LOGSet>
'                   <MC_FILE_LOG>
'                       <SQ_MC_ID>1</SQ_MC_ID>
'                       <DATE_START>4/15/2005 6:14:13 PM</DATE_START>
'                       <DATE_END>4/15/2005 6:15:04 PM</DATE_END>
'                       <STATUS>1</STATUS>
'                       <DETAILS>Error description</DETAILS>
'                   </MC_FILE_LOG>
'                   ....
'               </MC_FILE_LOGSet>
'           </UPDATE_MC_FILE_LOGS_REQUEST>
'RETURN  : String = status of MC File details updation, as XML string
'          Sample XML structure:
'           <UPDATE_MC_FILE_LOGS_RESPONSE>
'               <STATUS>SUCCESS</STATUS>
'           </UPDATE_MC_FILE_LOGS_RESPONSE>
'================================================================
Public Function UpdateMCLogs(ByVal astrUpdateMCLogsXML As String) As String
On Error GoTo UpdateMCLogs_ErrHandler

Dim lstrMethodName  As String   'to store method name
Dim lstrErrSrc      As String   'to store error source
Dim lstrErrDesc     As String   'to store error description
Dim llErrNbr        As Long     'to store error number

    lstrMethodName = "UpdateMCLogs"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjcDataClass          As New cDataClass       'object of cDataClass to access its methods
Dim lobjRequestXmlDOM       As New DOMDocument40    'object of XML DOM to load input Request XML
Dim lobjMCFileLogsNdLst     As IXMLDOMNodeList      'to store node list of MC_FILE_LOG node set
Dim lobjMCFileLogsElem      As IXMLDOMElement       'to store single node of MC_FILE_LOG node set
Dim lstrBatchUpdateQry      As String               'to store dynamically batch update query
Dim liCounter               As Integer              'to store loop counter
Dim liSQ_MC_ID              As Integer              'to store Money Cost file Id
Dim lstrDATE_START          As String               'to store start date/time of money cost file processing
Dim lstrDATE_END            As String               'to store end date/time of money cost file processed
Dim lbySTATUS               As Byte                 'to store status of money cost file processing (0: Success; 1: Failure)
Dim lstrDETAILS             As String               'to store details of money cost file processing result

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrUpdateMCLogsXML, ecDebugInputTrace
    End If

    'Validate if the Input Request XML is well-formed
    If Not lobjRequestXmlDOM.loadXML(astrUpdateMCLogsXML) Then
        'Raise Error
        Err.Raise lobjRequestXmlDOM.parseError.errorCode, _
                    , "Error loading Input Request XML to UpdateMCLogs(). " & lobjRequestXmlDOM.parseError.reason
    End If

    'get all node list of MC_FILE_DETAIL node set
    Set lobjMCFileLogsNdLst = lobjRequestXmlDOM.selectNodes("/UPDATE_MC_FILE_LOGS_REQUEST/MC_FILE_LOGSet/MC_FILE_LOG")

    lstrBatchUpdateQry = ""

    'loop for each individual MC_FILE_DETAIL node set
    For liCounter = 0 To lobjMCFileLogsNdLst.length - 1

        'get single node set of MC_FILE_DETAIL
        Set lobjMCFileLogsElem = lobjMCFileLogsNdLst.Item(liCounter)

        're-initialize local variables
        liSQ_MC_ID = 0
        lstrDATE_START = ""
        lstrDATE_END = ""
        lbySTATUS = 1
        lstrDETAILS = ""

        'check, if SQ_MC_ID found in Input XML, store in local variable
        If IsXMLElementPresent(lobjMCFileLogsElem, "SQ_MC_ID") Then
            If Trim(lobjMCFileLogsElem.selectNodes("SQ_MC_ID").Item(0).Text) <> "" Then
                liSQ_MC_ID = Val(lobjMCFileLogsElem.selectNodes("SQ_MC_ID").Item(0).Text)
            Else
                Err.Raise cINVALID_PARMS, , "SQ_MC_ID is not specified."
            End If
        Else
            Err.Raise cINVALID_PARMS, , "SQ_MC_ID is not specified."
        End If

        'check, if DATE_START found in Input XML, store in local variable
        If IsXMLElementPresent(lobjMCFileLogsElem, "DATE_START") Then
            If Trim(lobjMCFileLogsElem.selectNodes("DATE_START").Item(0).Text) <> "" Then
                lstrDATE_START = Trim(lobjMCFileLogsElem.selectNodes("DATE_START").Item(0).Text)
            Else
                Err.Raise cINVALID_PARMS, , "Date Start is not specified."
            End If
        Else
            Err.Raise cINVALID_PARMS, , "Date Start is not specified."
        End If

        'check, if DATE_START found in Input XML, store in local variable
        If IsXMLElementPresent(lobjMCFileLogsElem, "DATE_END") Then
            If Trim(lobjMCFileLogsElem.selectNodes("DATE_END").Item(0).Text) <> "" Then
                lstrDATE_END = Trim(lobjMCFileLogsElem.selectNodes("DATE_END").Item(0).Text)
            Else
                Err.Raise cINVALID_PARMS, , "Date End is not specified."
            End If
        Else
            Err.Raise cINVALID_PARMS, , "Date End is not specified."
        End If

        'check, if STATUS found in Input XML, store in local variable
        If IsXMLElementPresent(lobjMCFileLogsElem, "STATUS") Then
            If Trim(lobjMCFileLogsElem.selectNodes("STATUS").Item(0).Text) <> "" Then
                lbySTATUS = Trim(lobjMCFileLogsElem.selectNodes("STATUS").Item(0).Text)
            Else
                Err.Raise cINVALID_PARMS, , "Status is not specified."
            End If
        Else
            Err.Raise cINVALID_PARMS, , "Status is not specified."
        End If

        'check, if STATUS found in Input XML, store in local variable
        If IsXMLElementPresent(lobjMCFileLogsElem, "DETAILS") Then
            If Trim(lobjMCFileLogsElem.selectNodes("DETAILS").Item(0).Text) <> "" Then
                lstrDETAILS = Trim(lobjMCFileLogsElem.selectNodes("DETAILS").Item(0).Text)

                'replace one Single Quote to two Single Quotes
                lstrDETAILS = Replace(lstrDETAILS, "'", "''")
            Else
                Err.Raise cINVALID_PARMS, , "Details is not specified."
            End If
        Else
            Err.Raise cINVALID_PARMS, , "Details is not specified."
        End If

        'dynamically build insert statement for MC_LOGS table
        lstrBatchUpdateQry = lstrBatchUpdateQry & _
            "INSERT INTO MC_LOGS (SQ_MC_ID, DATE_START, DATE_END, STATUS, DETAILS) VALUES(" & liSQ_MC_ID & ", '" & _
            lstrDATE_START & "', '" & lstrDATE_END & "', " & lbySTATUS & ", '" & lstrDETAILS & "'); " & vbCrLf
    Next

    'if some data found, which needs to be updated/ inserted
    If lstrBatchUpdateQry <> "" Then
        'execute batch update SQL
        Call lobjcDataClass.Execute(ecExecuteSQL, ecRSExecuteNoRecords, lstrBatchUpdateQry)

        'return Output XML with status as SUCCESS
        UpdateMCLogs = "<UPDATE_MC_FILE_DETAIL_RESPONSE><STATUS>SUCCESS</STATUS></UPDATE_MC_FILE_DETAIL_RESPONSE>"
'    Else
'        'raise error, if no data found to update
'        Err.Raise -1, , "No data found to update MC_LOGS table in UpdateMCLogs() method. "
    End If

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, UpdateMCLogs, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

UpdateMCLogs_CleanMemory:
    'clear object variables from memory
    Set lobjcDataClass = Nothing
    Set lobjMCFileLogsElem = Nothing
    Set lobjMCFileLogsNdLst = Nothing
    Set lobjRequestXmlDOM = Nothing
    Exit Function

UpdateMCLogs_ErrHandler:
    llErrNbr = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    UpdateMCLogs = vbNullString

    'write error message to log file
    WriteToTextDebugFile cMODULE_NAME & lstrMethodName, BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    'raise error
    Err.Raise llErrNbr, lstrErrSrc, lstrErrDesc

    Resume UpdateMCLogs_CleanMemory
End Function

'================================================================
'METHOD  : UpdateMCFile
'PURPOSE : Wrapper method to update the MC File, on execution of
'          BSSTMoneyCostAuto service
'PARMS   :
'          astrUpdateMCFileXML [String] = Input paramters
'          to update the MC_File table
'          Sample XML structure:
'           <UPDATE_MC_FILE_REQUEST>
'               <MC_FILESet>
'                   <MC_FILE>
'                       <SQ_MC_ID>1</SQ_MC_ID>
'                       <LAST_SCHEDULE_PROCESS_DATE>09/30/2005</LAST_SCHEDULE_PROCESS_DATE>
'                   </MC_FILE>
'                   ....
'               </MC_FILESet>
'           </UPDATE_MC_FILE_REQUEST>
'RETURN  : String = status of MC File details updation, as XML string
'          Sample XML structure:
'           <UPDATE_MC_FILE_RESPONSE>
'               <STATUS>SUCCESS</STATUS>
'           </UPDATE_MC_FILE_RESPONSE>
'================================================================
Public Function UpdateMCFile(ByVal astrUpdateMCFileXML As String) As String
On Error GoTo UpdateMCFile_ErrHandler

Dim lstrMethodName  As String   'to store method name
Dim lstrErrSrc      As String   'to store error source
Dim lstrErrDesc     As String   'to store error description
Dim llErrNbr        As Long     'to store error number

    lstrMethodName = "UpdateMCFile"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjcDataClass          As New cDataClass       'object of cDataClass to access its methods
Dim lobjRequestXmlDOM       As New DOMDocument40    'object of XML DOM to load input Request XML
Dim lobjMCFileNdLst         As IXMLDOMNodeList      'to store node list of MC_FILE node set
Dim lobjMCFileElem          As IXMLDOMElement       'to store single node of MC_FILE node set
Dim lstrBatchUpdateQry      As String               'to store dynamically batch update query
Dim liCounter               As Integer              'to store loop counter
Dim liSQ_MC_ID              As Integer              'to store Money Cost file Id
Dim lstrLastScheduleDate    As String               'to store last schedule process date for a particular MC File

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrUpdateMCFileXML, ecDebugInputTrace
    End If

    'Validate if the Input Request XML is well-formed
    If Not lobjRequestXmlDOM.loadXML(astrUpdateMCFileXML) Then
        'Raise Error
        Err.Raise lobjRequestXmlDOM.parseError.errorCode, _
                    , "Error loading Input Request XML to UpdateMCFile(). " & lobjRequestXmlDOM.parseError.reason
    End If

    'get all node list of MC_FILE_DETAIL node set
    Set lobjMCFileNdLst = lobjRequestXmlDOM.selectNodes("/UPDATE_MC_FILE_REQUEST/MC_FILESet/MC_FILE")

    lstrBatchUpdateQry = ""

    'loop for each individual MC_FILE_DETAIL node set
    For liCounter = 0 To lobjMCFileNdLst.length - 1

        'get single node set of MC_FILE
        Set lobjMCFileElem = lobjMCFileNdLst.Item(liCounter)

        're-initialize local variables
        liSQ_MC_ID = 0
        lstrLastScheduleDate = ""

        'check, if SQ_MC_ID found in Input XML, store in local variable
        If IsXMLElementPresent(lobjMCFileElem, "SQ_MC_ID") Then
            If Trim(lobjMCFileElem.selectNodes("SQ_MC_ID").Item(0).Text) <> "" Then
                liSQ_MC_ID = Val(lobjMCFileElem.selectNodes("SQ_MC_ID").Item(0).Text)
            Else
                Err.Raise cINVALID_PARMS, , "SQ_MC_ID is not specified."
            End If
        Else
            Err.Raise cINVALID_PARMS, , "SQ_MC_ID is not specified."
        End If

        'check, if DATE_START found in Input XML, store in local variable
        If IsXMLElementPresent(lobjMCFileElem, "LAST_SCHEDULE_PROCESS_DATE") Then
            If Trim(lobjMCFileElem.selectNodes("LAST_SCHEDULE_PROCESS_DATE").Item(0).Text) <> "" Then
                lstrLastScheduleDate = Trim(lobjMCFileElem.selectNodes("LAST_SCHEDULE_PROCESS_DATE").Item(0).Text)
            Else
                Err.Raise cINVALID_PARMS, , "Last Schedule Process Date is not specified."
            End If
        Else
            Err.Raise cINVALID_PARMS, , "Last Schedule Process Date is not specified."
        End If

        'dynamically build insert statement for MC_LOGS table
        lstrBatchUpdateQry = lstrBatchUpdateQry & "UPDATE MC_FILE SET LAST_SCHEDULE_PROCESS_DATE = '" & _
                                            lstrLastScheduleDate & "' WHERE SQ_MC_ID = " & liSQ_MC_ID & ";" & vbCrLf
    Next

    'if some data found, which needs to be updated/ inserted
    If lstrBatchUpdateQry <> "" Then
        'execute batch update SQL
        Call lobjcDataClass.Execute(ecExecuteSQL, ecRSExecuteNoRecords, lstrBatchUpdateQry)

        'return Output XML with status as SUCCESS
        UpdateMCFile = "<UPDATE_MC_FILE_RESPONSE><STATUS>SUCCESS</STATUS></UPDATE_MC_FILE_RESPONSE>"
'    Else
'        'raise error, if no data found to update
'        Err.Raise -1, , "No data found to update MC_File table in UpdateMCFile() method. "
    End If

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, UpdateMCFile, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

UpdateMCFile_CleanMemory:
    'clear object variables from memory
    Set lobjcDataClass = Nothing
    Set lobjMCFileElem = Nothing
    Set lobjMCFileNdLst = Nothing
    Set lobjRequestXmlDOM = Nothing
    Exit Function

UpdateMCFile_ErrHandler:
    llErrNbr = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    UpdateMCFile = vbNullString

    'write error message to log file
    WriteToTextDebugFile cMODULE_NAME & lstrMethodName, BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    'raise error
    Err.Raise llErrNbr, lstrErrSrc, lstrErrDesc

    Resume UpdateMCFile_CleanMemory
End Function
