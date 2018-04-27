VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMoneyCostService"
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
'MODULE  : cMoneyCostService
'PURPOSE : This will contain all money cost related non-transactional
'          methods.
'================================================================

Option Explicit

'Constant for module name =======================================
Private Const cMODULE_NAME As String = "cMoneyCostService"
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
'METHOD  : GetMCFiles
'PURPOSE : Returns a list of Money Cost Files, associated with
'          logged-in user, in ascending order of Money Cost File Code
'PARMS   :
'          astrGetMCFilesXML [String] = Filter criteria
'          Sample XML structure:
'           <MC_FILES_REQUEST>
'               <USER_SSOID>500975793</USER_SSOID>
'           </MC_FILES_REQUEST>
'RETURN  : String = list of MC Files records as XML string
'          Sample XML structure:
'           <MC_FILE_RESPONSE>
'               <MC_FILESet>
'                   <MC_FILE>
'                       <!-- Details from MC_SECURITY & MC_FILE tables -->
'                       <SQ_MC_ID>1</SQ_MC_ID>
'                       <MONEY_COST_FILE>MCUSD-USD Money Cost File</MONEY_COST_FILE>
'                   </MC_FILE>
'                   ...
'               </MC_FILESet>
'           </MC_FILE_RESPONSE>
'================================================================
Public Function GetMCFiles(ByVal astrGetMCFilesXML As String) As String
On Error GoTo GetMCFiles_ErrHandler

Dim lstrErrSrc     As String    'to store error source
Dim lstrMethodName As String    'to store method name
Dim lstrErrDesc    As String    'to store error description
Dim llErrNbr       As Long      'to store error number

    lstrMethodName = "GetMCFiles"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjRSUtils         As New RecordsetUtils.IRecordsetUtils   'object variable for Recordset Utilities Component
Dim lobjDataClass       As New cDataClass                       'object variable to access cDataClass method(s)
Dim lobjRequestXmlDOM   As New DOMDocument40                    'object variable for DOM, to load Request XML
Dim lstrUSER_SSOID      As String                               'to store logged-in User Gessouid, from Request XML
Dim lstrResult          As String                               'to store final Response XML

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrGetMCFilesXML, ecDebugInputTrace
    End If

    'Validate if the Request XML is well-formed
    If Not lobjRequestXmlDOM.loadXML(astrGetMCFilesXML) Then
        'Raise Error
        Err.Raise lobjRequestXmlDOM.parseError.errorCode, _
                    , "Error loading Request XML to GetMCFiles(). " & lobjRequestXmlDOM.parseError.reason
    End If

    'If logged-in User SSO ID found in Input XML, fetch the value in local variable
    If IsXMLElementPresent(lobjRequestXmlDOM.documentElement, "USER_SSOID") Then
        If Trim(lobjRequestXmlDOM.getElementsByTagName("USER_SSOID").Item(0).Text) <> "" Then
            lstrUSER_SSOID = lobjRequestXmlDOM.getElementsByTagName("USER_SSOID").Item(0).Text
        Else
            Err.Raise cINVALID_PARMS, , "User's SSO ID not specified."
        End If
    Else
        Err.Raise cINVALID_PARMS, , "User's SSO ID not specified."
    End If

    'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
    'method of Recordset Utilities component to form the Output XML, in local variable
    lstrResult = lobjRSUtils.RSToXML("MC_FILE", lobjDataClass.Execute(ecGetMCFiles, ecRSExecuteRecords, lstrUSER_SSOID))

    'Return the XML as output
    GetMCFiles = "<MC_FILE_RESPONSE>" & lstrResult & "</MC_FILE_RESPONSE>"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, GetMCFiles, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

GetMCFiles_CleanMemory:
    'clear all local object variables from memory
    Set lobjDataClass = Nothing
    Set lobjRequestXmlDOM = Nothing
    Set lobjRSUtils = Nothing
    Exit Function

GetMCFiles_ErrHandler:
    lstrErrSrc = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description

    GetMCFiles = vbNullString

    'write error message to log file
    WriteToTextDebugFile cMODULE_NAME & lstrMethodName, BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    Err.Raise llErrNbr, lstrErrSrc, lstrErrDesc

    Resume GetMCFiles_CleanMemory
End Function

'================================================================
'METHOD  : GetMCFileDetails
'PURPOSE : Returns details of selected Money Cost file as per given
'          filter criteria. Filter Criteria will be Money Cost File
'          Id. Data will be in ascending order of column position
'PARMS   :
'          astrGetMCFileDetailsXML [String] = Filter criteria
'          Sample XML structure:
'           <MC_FILE_DETAIL_REQUEST>
'               <SQ_MC_ID>1</SQ_MC_ID>
'           </MC_FILE_DETAIL_REQUEST>
'RETURN  : String = details of selected MC File as XML string
'          Sample XML structure:
'           <MC_FILE_DETAIL_RESPONSE>
'               <MC_FILE_DETAILSet>
'                   <MC_FILE_DETAIL>
'                       <!-- Details from INDEX_RATES tables -->
'                       <SQ_INDEX_ID>1</SQ_INDEX_ID>
'                       <MC_FILE_COL_POSITION>1</MC_FILE_COL_POSITION>
'                       <INDEX_CODE>US TREASURY</INDEX_CODE>
'                       <DESCRIPTION>30yr-Swap</DESCRIPTION>
'                       <AMT_ADDER>0.081</AMT_ADDER>
'                       <DATE_EFFECTIVE>04/03/2005</DATE_EFFECTIVE>
'                   </MC_FILE_DETAIL>
'                   ...
'               </MC_FILE_DETAILSet>
'           </MC_FILE_DETAIL_RESPONSE>
'================================================================
Public Function GetMCFileDetails(ByVal astrGetMCFileDetailsXML As String) As String
On Error GoTo GetMCFileDetails_ErrHandler

Dim lstrErrSrc     As String    'to store error source
Dim lstrMethodName As String    'to store method name
Dim lstrErrDesc    As String    'to store error description
Dim llErrNbr       As Long      'to store error number

    lstrMethodName = "GetMCFileDetails"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjRSUtils         As New RecordsetUtils.IRecordsetUtils   'object variable for Recordset Utilities Component
Dim lobjDataClass       As New cDataClass                       'object variable to access cDataClass method(s)
Dim lobjRequestXmlDOM   As New DOMDocument40                    'object variable for DOM, to load Request XML
Dim llSQ_MC_ID          As Long                                 'to store selected Money Cost File ID, from Request XML
Dim lstrResult          As String                               'to store final Response XML

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrGetMCFileDetailsXML, ecDebugInputTrace
    End If

    'Validate if the Input Request XML is well-formed
    If Not lobjRequestXmlDOM.loadXML(astrGetMCFileDetailsXML) Then
        'Raise Error
        Err.Raise lobjRequestXmlDOM.parseError.errorCode, _
                    , "Error loading Input Request XML to GetMCFileDetails(). " & lobjRequestXmlDOM.parseError.reason
    End If

    'If selected Money Cost File ID found in Input Request XML, fetch the value in local variable
    If IsXMLElementPresent(lobjRequestXmlDOM.documentElement, "SQ_MC_ID") Then
        If Trim(lobjRequestXmlDOM.getElementsByTagName("SQ_MC_ID").Item(0).Text) <> "" Then
            llSQ_MC_ID = lobjRequestXmlDOM.getElementsByTagName("SQ_MC_ID").Item(0).Text
        Else
            Err.Raise cINVALID_PARMS, , "Money Cost File ID not specified."
        End If
    Else
        Err.Raise cINVALID_PARMS, , "Money Cost File ID not specified."
    End If

    'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
    'method of Recordset Utilities component to form the Output XML, in local variable
    lstrResult = lobjRSUtils.RSToXML("MC_FILE_DETAIL", lobjDataClass.Execute(ecGetMCFileDetails, ecRSExecuteRecords, llSQ_MC_ID))

    'Return the XML as output
    GetMCFileDetails = "<MC_FILE_DETAIL_RESPONSE>" & lstrResult & "</MC_FILE_DETAIL_RESPONSE>"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, GetMCFileDetails, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

GetMCFileDetails_CleanMemory:
    'clear all local object variables from memory
    Set lobjDataClass = Nothing
    Set lobjRequestXmlDOM = Nothing
    Set lobjRSUtils = Nothing
    Exit Function

GetMCFileDetails_ErrHandler:
    lstrErrSrc = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description

    GetMCFileDetails = vbNullString

    'write error message to log file
    WriteToTextDebugFile cMODULE_NAME & lstrMethodName, BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    Err.Raise llErrNbr, lstrErrSrc, lstrErrDesc

    Resume GetMCFileDetails_CleanMemory
End Function

'================================================================
'METHOD  : GetAllMCFiles
'PURPOSE : Returns a list of all Money Cost Files in ascending
'           order of Money Cost File Code
'PARMS   : None
'RETURN  : String = list of all MC Files records as XML string
'          Sample XML structure:
'           <MC_FILE_RESPONSE>
'               <MC_FILESet>
'                   <MC_FILE>
'                       <!-- Details from MC_FILE table -->
'                       <SQ_MC_ID>1</SQ_MC_ID>
'                       <MC_CODE>MCUSD</MC_CODE>
'                       <DESCRIPTION>USD Money Cost File</DESCRIPTION>
'                       <CURRENCY_CODE>USD</CURRENCY_CODE>
'                       <START_TIME>09:00</START_TIME>
'                       <END_TIME>11:00</END_TIME>
'                       <FREQUENCY>d</FREQUENCY>
'                       <FREQUENCY_COUNT>1</FREQUENCY_COUNT>
'                       <LAST_SCHEDULE_PROCESS_DATE>09/30/2005</LAST_SCHEDULE_PROCESS_DATE>
'                       <MARKET_CLOSED_DWH_CHECK_COUNTER>7</MARKET_CLOSED_DWH_CHECK_COUNTER>
'                       <CLARIFY_QUEUE>test</CLARIFY_QUEUE>
'                       <BUSINESS_CONTACT>singh.manpreet@ge.com</BUSINESS_CONTACT>
'                   </MC_FILE>
'                   ...
'               </MC_FILESet>
'           </MC_FILE_RESPONSE>
'================================================================
Public Function GetAllMCFiles() As String
On Error GoTo GetAllMCFiles_ErrHandler

Dim lstrErrSrc     As String    'to store error source
Dim lstrMethodName As String    'to store method name
Dim lstrErrDesc    As String    'to store error description
Dim llErrNbr       As Long      'to store error number

    lstrMethodName = "GetAllMCFiles"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjRSUtils         As New RecordsetUtils.IRecordsetUtils   'object variable for Recordset Utilities Component
Dim lobjDataClass       As New cDataClass                       'object variable to access cDataClass method(s)
Dim lstrResult          As String                               'to store final Response XML

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "", ecDebugInputTrace
    End If

    'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
    'method of Recordset Utilities component to form the Output XML, in local variable
    lstrResult = lobjRSUtils.RSToXML("MC_FILE", lobjDataClass.Execute(ecGetAllMCFiles, ecRSExecuteRecords))

    'Return the XML as output
    GetAllMCFiles = "<MC_FILE_RESPONSE>" & lstrResult & "</MC_FILE_RESPONSE>"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, GetAllMCFiles, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

GetAllMCFiles_CleanMemory:
    'clear all local object variables from memory
    Set lobjDataClass = Nothing
    Set lobjRSUtils = Nothing
    Exit Function

GetAllMCFiles_ErrHandler:
    lstrErrSrc = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description

    GetAllMCFiles = vbNullString

    'write error message to log file
    WriteToTextDebugFile cMODULE_NAME & lstrMethodName, BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    Err.Raise llErrNbr, lstrErrSrc, lstrErrDesc

    Resume GetAllMCFiles_CleanMemory
End Function

'================================================================
'METHOD  : GetIndexRates
'PURPOSE : Returns a list of all Money Cost Files in ascending
'           order of Money Cost File Code
'PARMS   :
'          astrGetIndexRatesXML [String] = Input paramters
'          to retrieve the information for the MC Files list.
'          Sample XML structure:
'           <INDEX_RATE_REQUEST>
'               <SQ_MC_ID>2</SQ_MC_ID>
'               <PROCESS_DATE>04/15/2005</PROCESS_DATE>
'           </INDEX_RATE_REQUEST>
'RETURN  : String = information containing index rates of selected
'          Money Cost file as XML string
'          Sample XML structure:
'           <INDEX_RATE_RESPONSE>
'               <INDEX_RATESet>
'                   <INDEX_RATE>
'                       <!-- Details from INDEX_RATES, INDEX_AUDIT tables -->
'                       <SQ_INDEX_ID>1</SQ_INDEX_ID>
'                       <INDEX_CODE>INTEREST RATE SWAP</INDEX_CODE>
'                       <INDEX_TERM>24</INDEX_TERM>
'                       <AMT_ADDER>0.0450</AMT_ADDER>
'                       <DATE_EFFECTIVE>04/07/2005</DATE_EFFECTIVE>
'                       <IND_PERCENTILE>0</IND_PERCENTILE>
'                       <MC_FILE_COL_POSITION>1</MC_FILE_COL_POSITION>
'                       <IND_QUERYDB>1</IND_QUERYDB>
'                       <DESCRIPTION>GE MC</DESCRIPTION>
'                   </INDEX_RATE>
'                   ...
'               </INDEX_RATESet>
'           </INDEX_RATE_RESPONSE>
'================================================================
Public Function GetIndexRates(ByVal astrGetIndexRatesXML As String) As String
On Error GoTo GetIndexRates_ErrHandler

Dim lstrErrSrc     As String    'to store error source
Dim lstrMethodName As String    'to store method name
Dim lstrErrDesc    As String    'to store error description
Dim llErrNbr       As Long      'to store error number

    lstrMethodName = "GetIndexRates"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjRSUtils         As New RecordsetUtils.IRecordsetUtils   'object variable for Recordset Utilities Component
Dim lobjDataClass       As New cDataClass                       'object variable to access cDataClass method(s)
Dim lobjRequestXmlDOM   As New DOMDocument40                    'object variable for DOM, to load Request XML
Dim liSQ_MC_ID          As Integer                              'to store Money Cost file ID, from Request XML
Dim lstrPROCESS_DATE    As String                               'to store MC File Process Date, from Request XML
Dim lstrResult          As String                               'to store Response XML

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrGetIndexRatesXML, ecDebugInputTrace
    End If

    'validate, if Request xml is well-formed
    If Not lobjRequestXmlDOM.loadXML(astrGetIndexRatesXML) Then
        'Raise Error
        Err.Raise lobjRequestXmlDOM.parseError.errorCode, _
                    , "Error loading Input Request XML to GetIndexRates(). " & lobjRequestXmlDOM.parseError.reason
    End If

    'check, whether MC File Id found, in Request XML
    If IsXMLElementPresent(lobjRequestXmlDOM.documentElement, "SQ_MC_ID") Then
        If lobjRequestXmlDOM.getElementsByTagName("SQ_MC_ID").length > 0 Then
            liSQ_MC_ID = Val(lobjRequestXmlDOM.getElementsByTagName("SQ_MC_ID").Item(0).Text)
        Else
            Err.Raise cINVALID_PARMS, , "Money Cost File ID not specified."
        End If
    Else
        Err.Raise cINVALID_PARMS, , "Money Cost File ID not specified."
    End If

    'check, whether MC File Process Date found, in Request XML
    If IsXMLElementPresent(lobjRequestXmlDOM.documentElement, "PROCESS_DATE") Then
        If lobjRequestXmlDOM.getElementsByTagName("PROCESS_DATE").length > 0 Then
            lstrPROCESS_DATE = Trim(lobjRequestXmlDOM.getElementsByTagName("PROCESS_DATE").Item(0).Text)
        Else
            Err.Raise cINVALID_PARMS, , "Money Cost File Process Date not specified."
        End If
    Else
        Err.Raise cINVALID_PARMS, , "Money Cost File Process Date not specified."
    End If

    'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
    'method of Recordset Utilities component to form the Output XML, in local variable
    lstrResult = lobjRSUtils.RSToXML("INDEX_RATE", lobjDataClass.Execute(ecGetIndexRates, ecRSExecuteRecords, liSQ_MC_ID, lstrPROCESS_DATE))

    'Return the XML as output
    GetIndexRates = "<INDEX_RATE_RESPONSE>" & lstrResult & "</INDEX_RATE_RESPONSE>"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, GetIndexRates, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

GetIndexRates_CleanMemory:
    'clear all local object variables from memory
    Set lobjDataClass = Nothing
    Set lobjRequestXmlDOM = Nothing
    Set lobjRSUtils = Nothing
    Exit Function

GetIndexRates_ErrHandler:
    lstrErrSrc = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description

    GetIndexRates = vbNullString

    'write error message to log file
    WriteToTextDebugFile cMODULE_NAME & lstrMethodName, BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    Err.Raise llErrNbr, lstrErrSrc, lstrErrDesc

    Resume GetIndexRates_CleanMemory
End Function
