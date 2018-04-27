VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDataClass"
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
'MODULE  : cDataClass
'PURPOSE : Handles all data access.
'================================================================

Option Explicit

'Constants ======================================================
Private Const cDELIMITER As String = "|~^"
Private Const cMODULE_NAME As String = "cDataClass"
'================================================================

'================================================================
' Bitwise Enum for Execution Options.
'================================================================
Public Enum eDCExecOptions
    ecRSExecuteRecords = 1
    ecRSExecuteNoRecords = 2
    ecRSExecuteHierarch = 4
    ecQuoteReplacement = 8
End Enum
'================================================================

'================================================================
' Enum for SQL Statements. Insert all entries in between the
' lower bound and upper bound. DO NOT put a number on upper bound.
' It will automatically default to the next highest number than
' the one you assign. The eLowerBound and eUpperBound Enum
' members will not appear in intellisense! Since it must be a
' public enum, use your Business Server Name to make it unique.
'================================================================
Public Enum eDCActions
    [_eLowerBound] = -1
    ecRSUpdateBatch
    ecRSUpdateBatchCurrentOnly
    ecTestSQL
    'to get Money Cost File List
    ecGetMCFiles
    'to get details of Money Cost File
    ecGetMCFileDetails
    'to perform single parameter, dynamically build query
    ecExecuteSQL
    'to get all Money Cost files list
    ecGetAllMCFiles
    'to get all index rates for a Money Cost file
    ecGetIndexRates
    '---All Enums go here--

    [_eUpperBound]
End Enum
'================================================================

'================================================================
'  SQL Statements go here. The statement should be implemented as
'  a constant with a meaningful name. The same name(with
'  alternate "e" prefix, of course) should be placed in the enum
'  above and assigned the next sequential number. An Example:
'
'  Private CONST cGetCustomer as String = "Select * From CUSTOMER"
'  and in the Enum above:    eGetCustomer = 1
'
'  The final step is to place the enum/constant within the Select
'  Case statement in the GetSQLStatement function. Obviously the
'  Enum will take care of assigning the value.
'  See the method AddParm for additional info about substitution
'  parameters.
'================================================================
'test SQL
Private Const cTestSQL          As String = "SELECT COUNT(*) FROM MC_FILE"

'to get list of Money Cost Files, associated with logged-in user
Private Const cGetMCFiles       As String = "SELECT MC_FILE.SQ_MC_ID, MC_FILE.MC_CODE + (CASE WHEN MC_FILE.[DESCRIPTION] " & _
                                            "IS NULL THEN '' ELSE '-' + MC_FILE.[DESCRIPTION] END) AS 'MONEY_COST_FILE' " & _
                                            "FROM MC_FILE INNER JOIN MC_SECURITY ON MC_SECURITY.SQ_MC_ID = " & _
                                            "MC_FILE.SQ_MC_ID WHERE MC_FILE.IND_ACTIVE = 1 AND MC_SECURITY.SSO_ID = |~^1 ORDER BY MONEY_COST_FILE"

'to get details of selected Money Cost File
Private Const cGetMCFileDetails As String = "SELECT SQ_INDEX_ID, MC_FILE_COL_POSITION, INDEX_CODE, ISNULL([DESCRIPTION], " & _
                                            "'') AS 'DESCRIPTION', AMT_ADDER, CONVERT(VARCHAR(10), DATE_EFFECTIVE, 101) " & _
                                            "AS DATE_EFFECTIVE FROM INDEX_RATES WHERE SQ_MC_ID = |~^1 ORDER BY " & _
                                            "MC_FILE_COL_POSITION"

'to get list of all Money Cost files
Private Const cGetAllMCFiles    As String = "SELECT SQ_MC_ID, MC_CODE, ISNULL([DESCRIPTION], '') AS 'DESCRIPTION', " & _
                                            "CURRENCY_CODE, CONVERT(VARCHAR(8), START_TIME, 108) AS START_TIME, " & _
                                            "CONVERT(VARCHAR(8), END_TIME, 108) AS END_TIME, DAYS_TO_SKIP, " & _
                                            "FREQUENCY, FREQUENCY_COUNT, LAST_SCHEDULE_PROCESS_DATE, " & _
                                            "MARKET_CLOSED_DWH_CHECK_COUNTER, CLARIFY_QUEUE, ISNULL([BUSINESS_CONTACT], '') " & _
                                            "AS BUSINESS_CONTACT , LAST_UPDATED_IND , DATE_FORMAT , DATE_FORMAT_REQUIRED  FROM MC_FILE WHERE IND_ACTIVE = 1 ORDER BY MC_CODE"

'to get all index rate data for particulat money Cost file
Private Const cGetIndexRates    As String = "SELECT INDEX_RATES.SQ_INDEX_ID, INDEX_RATES.INDEX_CODE, " & _
                                            "INDEX_RATES.INDEX_TERM, X.AMT_ADDER, X.DATE_EFFECTIVE, " & _
                                            "INDEX_RATES.IND_PERCENTILE, INDEX_RATES.MC_FILE_COL_POSITION, " & _
                                            "INDEX_RATES.IND_QUERYDB, INDEX_RATES.[DESCRIPTION],INDEX_RATES.IND_PREV_INDEXRATES_REQ FROM INDEX_RATES " & _
                                            "INNER JOIN INDEX_AUDIT X ON X.SQ_INDEX_ID = INDEX_RATES.SQ_INDEX_ID " & _
                                            "WHERE INDEX_RATES.SQ_MC_ID = |~^1 AND X.DATE_UPDATE = (SELECT " & _
                                            "MAX(A.DATE_UPDATE) FROM INDEX_AUDIT A WHERE A.SQ_INDEX_ID = " & _
                                            "INDEX_RATES.SQ_INDEX_ID AND A.DATE_EFFECTIVE = X.DATE_EFFECTIVE) AND " & _
                                            "X.DATE_EFFECTIVE = (SELECT MAX(B.DATE_EFFECTIVE) FROM INDEX_AUDIT B " & _
                                            "WHERE DATEDIFF(d, '|~^2', B.DATE_EFFECTIVE) <= 0 AND B.SQ_INDEX_ID = " & _
                                            "INDEX_RATES.SQ_INDEX_ID) ORDER BY INDEX_RATES.MC_FILE_COL_POSITION"

'to perform single parameter, dynamically build query
Private Const cExecuteSQL       As String = "|~^1"
'---All SQL constants go here--

'================================================================
'*                      END  of SQL Statements
'================================================================

'Public Variables ===============================================
Private gcnDataClass    As New ADODB.Connection
Private glMaxRecs       As Long
'================================================================

'================================================================
'METHOD  : Connect
'PURPOSE : Connect to some database using a Trusted Connection.
'          The connect string for the appropriate DB will be
'          returned by GetConnectString.
'PARMS   : astrConnectStringKey [String] = connection string key
'          name in the registry. If no value is supplied the
'          value in cCONN_STRING_KEY constant will be used
'RETURN  : NONE
'================================================================
Public Sub Connect(Optional ByVal astrConnectStringKey As String)
On Error GoTo Connect_ErrHandler

    If astrConnectStringKey = "" Then astrConnectStringKey = cCONN_STRING_KEY

    gcnDataClass.Open GetConnectString(astrConnectStringKey)
    Exit Sub

Connect_ErrHandler:
    Err.Raise Err.Number, cMODULE_NAME & ":Connect()", Err.Description
End Sub

'================================================================
'METHOD  : GetConnectString
'PURPOSE : This method will get the DB connection string from the
'          registry.
'PARMS   : astrKey [String] = Connection string registry Key name
'RETURN  : String = The connection string to connect to the
'          database with
'================================================================
Private Function GetConnectString(ByVal astrKey As String) As String
    GetConnectString = ReadRegistry(cFACILITY_CONFIG_REG_PATH & cFACILITY_ID & cCONN_STRINGS_REG_PATH & astrKey)
End Function

'================================================================
'METHOD  : Disconnect
'PURPOSE : Disconnects the current connection from the datasource
'PARMS   : NONE
'RETURN  : NONE
'================================================================
Public Sub Disconnect()
On Error GoTo Disconnect_ErrHandler

    gcnDataClass.Close
    Exit Sub

Disconnect_ErrHandler:
    Err.Raise Err.Number, cMODULE_NAME & ":Disconnect()", Err.Description
End Sub

'================================================================
'METHOD  : Execute
'PURPOSE : This method performs an action based on the Action ID
'          provided.
'          If alActionID = eIService_RSUpdateBatch then it is
'           assumed the avParms() is an array of ADOR.Recordset
'           objects. It is possible to pass in more than one
'           Recordset object to be updated.
'          Elseif alActionID = anything else then Execute
'           performs the action based on one of the
'           predefined SQL statements above.
'PARMS   : alActionID [Long] = Action ID of the action to take
'          like Execute a SQL statement, Update a RS, etc.
'          alOptions [Long] = Options for this execution like
'          Return an RS, Don't return an RS etc.
'          avParms [Variant]  = Paramaters for the action to
'          execute.
'RETURN  : ADODB.Recordset = Recorset containing the requested
'          data
'================================================================
Public Function Execute(ByVal alActionID As eDCActions, _
                        ByVal alOptions As eDCExecOptions, _
                        ParamArray avParms() As Variant) _
                        As ADODB.Recordset
On Error GoTo Execute_ErrHandler

Dim llCounter As Long
Dim lbQuotes As Boolean
Dim lrstemp As New ADODB.Recordset
Dim lstrSQL As String
Dim lvParmsTemp As Variant

    'Bitwise the options to see if there is a No Quote Replacement flag
    lbQuotes = alOptions And ecQuoteReplacement

    'Check if non hierarchical record needs to be executed
    If Not CBool(alOptions And ecRSExecuteHierarch) Then

        'Get a DB connection
        If Len(gcnDataClass.ConnectionString) = 0 Then Connect

        'Check if batch update needs to be done
        If alActionID = ecRSUpdateBatch Then

            'If no Recordsets were passed as parameters then Raise Error
            If UBound(avParms()) = -1 Then Err.Raise cINVALID_PARMS

            'Loop through each parameters
            For llCounter = 0 To UBound(avParms())

                'Make sure it's a RS before performing UpdateBatch!
                If TypeName(avParms(llCounter)) = "Recordset" Then

                    'Do Batch update
                    Set avParms(llCounter).ActiveConnection = gcnDataClass
                    avParms(llCounter).UpdateBatch adAffectAllChapters
                    Set avParms(llCounter).ActiveConnection = Nothing
                End If
            Next llCounter

        'Otherwise check if Batch update needs to be done
        ElseIf alActionID = ecRSUpdateBatchCurrentOnly Then

            'If no Recordsets were passed as parameters then Raise Error
            If UBound(avParms()) = -1 Then Err.Raise cINVALID_PARMS

            'Loop through each parameters
            For llCounter = 0 To UBound(avParms())

                'Make sure it's a RS before performing UpdateBatch!
                If TypeName(avParms(llCounter)) = "Recordset" Then

                    'Do Batch update
                    Set avParms(llCounter).ActiveConnection = gcnDataClass
                    avParms(llCounter).UpdateBatch adAffectCurrent
                    Set avParms(llCounter).ActiveConnection = Nothing
                End If
            Next llCounter

        'Otherwise if not batch update
        Else

            'Put ParamArray in Local Var since a ParamArray can't be passed around
            lvParmsTemp = avParms

            'Get completed SQL statement to execute based on alActionID and supplied parameters
            lstrSQL = GetSQLStatement(alActionID, lvParmsTemp, lbQuotes)

            'Check if records needs to be returned
            If alOptions And ecRSExecuteRecords Then

                lrstemp.MaxRecords = glMaxRecs
                lrstemp.CursorLocation = adUseClient

                'Execute SQL statement and get the recordset
                lrstemp.Open lstrSQL, gcnDataClass, adOpenStatic, adLockBatchOptimistic

                Set lrstemp.ActiveConnection = Nothing

                'Return the recorset
                Set Execute = lrstemp

            'Otherwise if no records to be returned
            ElseIf alOptions And ecRSExecuteNoRecords Then

                'Execute SQL statement
                gcnDataClass.Execute lstrSQL, , adExecuteNoRecords
                Set Execute = Nothing

            'Otherwise Raise error
            Else
                Err.Raise cINVALID_PARMS
            End If
        End If

    'Otherwise, if hierarchical record needs to be executed
    Else

        'Get hierarchical connection string
        If Len(gcnDataClass.ConnectionString) = 0 Then Connect (cHIERARCH_CONN_STRING_KEY)

        'Put ParamArray in Local Var since a ParamArray can't be passed around
        lvParmsTemp = avParms

        'Get completed SQL statement to execute based on alActionID and supplied parameters
        lstrSQL = GetSQLStatement(alActionID, lvParmsTemp, lbQuotes)

        Set lrstemp = New ADODB.Recordset
        lrstemp.MaxRecords = glMaxRecs
        lrstemp.CursorLocation = adUseClient

        'Execute SQL statement and get the recordset
        lrstemp.Open lstrSQL, gcnDataClass, adOpenStatic, adLockBatchOptimistic
        Set lrstemp.ActiveConnection = Nothing

        'Return the recorset
        Set Execute = lrstemp

    End If
    Exit Function

Execute_ErrHandler:
    Err.Raise Err.Number, cMODULE_NAME & ":Execute()", Err.Description
End Function

'================================================================
'METHOD  : GetSQLStatement
'PURPOSE : Based on an Action ID and an array of parms
'          GetSQLStatement returns the SQL Statement in
'          "executeable" form i.e. all parms are replaced in the
'          string with the parms provided in Parms()
'PARMS   : alActionID [eDCActions] = SQL Statement to Execute
'          avParms [Variant] = Parms to Populate SQL Statement
'          with
'RETURN  : String = Completed SQL statement
'================================================================
Private Function GetSQLStatement(alActionID As eDCActions, _
                                    avParms As Variant, _
                                    Optional ByVal abQuoteReplacement As Boolean = False) As String
On Error GoTo GetSQLStatement_ErrHandler

Dim lstrSQL As String
Dim liPos As Integer
Dim liLoopCtr As Integer
Dim liUBound As Integer

    Select Case alActionID
        '---All Enum to SQL constant mapping goes here---
        Case ecTestSQL
            lstrSQL = cTestSQL
        Case ecGetMCFiles
            lstrSQL = cGetMCFiles
        Case ecGetMCFileDetails
            lstrSQL = cGetMCFileDetails
        Case ecExecuteSQL
            lstrSQL = cExecuteSQL
        Case ecGetAllMCFiles
            lstrSQL = cGetAllMCFiles
        Case ecGetIndexRates
            lstrSQL = cGetIndexRates

        Case Else
            Err.Raise cINVALID_SQL_ID
    End Select

    'If there are no substitution parms, we're done
    liPos = InStr(1, lstrSQL, cDELIMITER)

    If liPos = 0 Then
        GetSQLStatement = lstrSQL
        Exit Function
    End If

    If IsArray(avParms) Then
        If UBound(avParms) = -1 Then
            Err.Raise cINVALID_PARMS
        End If
    End If

    'Loop through each of the param
    liUBound = UBound(avParms)
    For liLoopCtr = liUBound + 1 To 1 Step -1

        'First, replace any apostrophes (') in the parm with two ('') for SQL Server.
        If abQuoteReplacement Then avParms(liLoopCtr - 1) = Replace(avParms(liLoopCtr - 1), "'", "''")

        'Then Insert the parm into the SQL statement.
        lstrSQL = Replace(lstrSQL, cDELIMITER & CStr(liLoopCtr), avParms(liLoopCtr - 1))
    Next liLoopCtr

    'REPLACE ANY STRINGS THAT CAME UP 'NULL' WITH NULL
    lstrSQL = Replace(lstrSQL, "'NULL'", "NULL")

    'Make sure all substitution parms have been swapped out
    liPos = InStr(1, lstrSQL, cDELIMITER)
    If liPos > 0 Then Err.Raise cINVALID_PARMS

    'Return Completed SQL statement
    GetSQLStatement = lstrSQL
    Exit Function

GetSQLStatement_ErrHandler:
    Err.Raise Err.Number, cMODULE_NAME & ":GetSQLStatement()", Err.Description
End Function

Private Sub Class_Terminate()
    If gcnDataClass Is Nothing Then Exit Sub
    If gcnDataClass.State = adStateOpen Then gcnDataClass.Close
    Set gcnDataClass = Nothing
End Sub

Public Property Get MaxRecords() As Long
    MaxRecords = glMaxRecs
End Property

Public Property Let MaxRecords(ByVal alValue As Long)
    glMaxRecs = alValue
End Property
