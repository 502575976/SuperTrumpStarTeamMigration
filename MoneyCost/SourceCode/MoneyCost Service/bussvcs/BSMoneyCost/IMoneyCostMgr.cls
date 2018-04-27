VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "IMoneyCostMgr"
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
'MODULE  : IMoneyCostMgr
'PURPOSE : This will contain transactional wrapper methods to
'          avoid multiple invocation of business component
'          methods by UI tier.
'================================================================

Option Explicit

'Constant for module name =======================================
Private Const cMODULE_NAME  As String = "IMoneyCostMgr"
'================================================================
Private Const cError_Msg    As String = ""

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
'                       <AMT_ADDER>0.081</AMT_ADDER>
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
Dim lobjMTSContext          As ObjectContext        'object for MTS transaction

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrUpdateMCDetailsXML, ecDebugInputTrace
    End If

    'start MTS transaction
    Set lobjMTSContext = GetObjectContext()

    'call UpdateMCDetails() method of cMoneyCostMgr class module to update
    'selected Money Cost file details, and return final output XML
    UpdateMCDetails = lobjcMoneyCostMgr.UpdateMCDetails(astrUpdateMCDetailsXML)

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, UpdateMCDetails, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

    'complete MTS transaction
    lobjMTSContext.SetComplete

UpdateMCDetails_CleanMemory:
    Set lobjcMoneyCostMgr = Nothing
    Set lobjMTSContext = Nothing
    Exit Function

UpdateMCDetails_ErrHandler:
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    'return error XML
    UpdateMCDetails = BuildErrorXML(llErrNbr, lstrErrSrc, lstrErrDesc, cError_Msg)

    'write error to log file
    WriteToTextDebugFile cCOMPONENT_NAME & "." & cMODULE_NAME & ".UpdateMCDetails ", _
                        BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    'Rollback all MTS transaction
    lobjMTSContext.SetAbort

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

Dim lobjcMoneyCostMgr       As New cMoneyCostMgr    'object of cMoneyCostMgr class module, to access its methods
Dim lobjMTSContext          As ObjectContext        'object for MTS transaction

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrUpdateMCLogsXML, ecDebugInputTrace
    End If

    'start MTS transaction
    Set lobjMTSContext = GetObjectContext()

    'call UpdateMCLogs() method of cMoneyCostMgr class module to update MC_Logs table,
    'on execution of BSSTMoneyCost service, and return final output XML
    UpdateMCLogs = lobjcMoneyCostMgr.UpdateMCLogs(astrUpdateMCLogsXML)

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, UpdateMCLogs, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

    'complete MTS transaction
    lobjMTSContext.SetComplete

UpdateMCLogs_CleanMemory:
    Set lobjcMoneyCostMgr = Nothing
    Set lobjMTSContext = Nothing
    Exit Function

UpdateMCLogs_ErrHandler:
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    'return error XML
    UpdateMCLogs = BuildErrorXML(llErrNbr, lstrErrSrc, lstrErrDesc, cError_Msg)

    'write error to log file
    WriteToTextDebugFile cCOMPONENT_NAME & "." & cMODULE_NAME & ".UpdateMCLogs ", _
                        BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    'Rollback all MTS transaction
    lobjMTSContext.SetAbort

    Resume UpdateMCLogs_CleanMemory
End Function

'================================================================
'METHOD  : UpdateMCFile
'PURPOSE : Wrapper method to update the MC File for last schedule
'          process date, on execution of BSSTMoneyCostAuto service
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

Dim lobjcMoneyCostMgr       As New cMoneyCostMgr    'object of cMoneyCostMgr class module, to access its methods
Dim lobjMTSContext          As ObjectContext        'object for MTS transaction

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrUpdateMCFileXML, ecDebugInputTrace
    End If

    'start MTS transaction
    Set lobjMTSContext = GetObjectContext()

    'call UpdateMCFile() method of cMoneyCostMgr class module to update MC_File table,
    'on execution of BSSTMoneyCost service, and return final output XML
    UpdateMCFile = lobjcMoneyCostMgr.UpdateMCFile(astrUpdateMCFileXML)

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, UpdateMCFile, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

    'complete MTS transaction
    lobjMTSContext.SetComplete

UpdateMCFile_CleanMemory:
    Set lobjcMoneyCostMgr = Nothing
    Set lobjMTSContext = Nothing
    Exit Function

UpdateMCFile_ErrHandler:
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    'return error XML
    UpdateMCFile = BuildErrorXML(llErrNbr, lstrErrSrc, lstrErrDesc, cError_Msg)

    'write error to log file
    WriteToTextDebugFile cCOMPONENT_NAME & "." & cMODULE_NAME & ".UpdateMCFile ", _
                        BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    'Rollback all MTS transaction
    lobjMTSContext.SetAbort

    Resume UpdateMCFile_CleanMemory
End Function
