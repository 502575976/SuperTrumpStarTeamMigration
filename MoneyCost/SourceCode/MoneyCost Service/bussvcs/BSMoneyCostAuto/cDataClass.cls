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

    'to use Test SQL
    ecTestSQL

    'to fetch index rates from DWH
    ecGetIndexData
    
    'To fetch data from DWH without the last_updated_ind flag
    ecGetIndexDataWithoutCondition
    
    

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
Private Const cTestSQL      As String = "SELECT COUNT(*) FROM DWC_OWNER.DWC_INTEREST_RATE"

'to fetch index rates from DWH for given Index Rate configuration, given currency, and specific processing date
Private Const cGetIndexData As String = "SELECT YIELD_CURVE_TYPE, TERM_PERIOD, INTEREST_RATE FROM " & _
                                        "DWC_OWNER.DWC_INTEREST_RATE, DWC_OWNER.DWC_DATE, DWC_TERM_PERIOD, " & _
                                        "DWC_OWNER.DWC_YIELD_CURVE_TYPE, DWC_OWNER.DWC_CURRENCY WHERE " & _
                                        "((DWC_INTEREST_RATE.DATE_KEY = DWC_DATE.DATE_KEY) AND " & _
                                        "(DWC_INTEREST_RATE.CURRENCY_KEY = DWC_CURRENCY.CURRENCY_KEY) AND " & _
                                        "(DWC_INTEREST_RATE.YIELD_CURVE_KEY = DWC_YIELD_CURVE_TYPE.YIELD_CURVE_KEY) " & _
                                        "AND (DWC_INTEREST_RATE.TERM_PERIOD_KEY = DWC_TERM_PERIOD.TERM_PERIOD_KEY)) " & _
                                        "AND CALENDAR_DATE = TO_DATE('|~^1', 'MM/DD/YYYY') AND CURRENCY_CODE = '|~^2' " & _
                                        "AND NVL(DWC_INTEREST_RATE.LAST_UPDATED_IND,'N') <> 'Y' " & _
                                        "AND UPPER(YIELD_CURVE_TYPE) IN (|~^3) AND TERM_PERIOD IN (|~^4) ORDER BY " & _
                                        "YIELD_CURVE_TYPE, TERM_PERIOD"


'to fetch index rates from DWH for given Index Rate configuration, given currency, and specific processing date
'where we dont want to fetch the data with the last_updated condition.
Private Const cGetIndexDataWithoutCondition As String = "SELECT YIELD_CURVE_TYPE, TERM_PERIOD, INTEREST_RATE FROM " & _
                                        "DWC_OWNER.DWC_INTEREST_RATE, DWC_OWNER.DWC_DATE, DWC_TERM_PERIOD, " & _
                                        "DWC_OWNER.DWC_YIELD_CURVE_TYPE, DWC_OWNER.DWC_CURRENCY WHERE " & _
                                        "((DWC_INTEREST_RATE.DATE_KEY = DWC_DATE.DATE_KEY) AND " & _
                                        "(DWC_INTEREST_RATE.CURRENCY_KEY = DWC_CURRENCY.CURRENCY_KEY) AND " & _
                                        "(DWC_INTEREST_RATE.YIELD_CURVE_KEY = DWC_YIELD_CURVE_TYPE.YIELD_CURVE_KEY) " & _
                                        "AND (DWC_INTEREST_RATE.TERM_PERIOD_KEY = DWC_TERM_PERIOD.TERM_PERIOD_KEY)) " & _
                                        "AND CALENDAR_DATE = TO_DATE('|~^1', 'MM/DD/YYYY') AND CURRENCY_CODE = '|~^2' " & _
                                        "AND UPPER(YIELD_CURVE_TYPE) IN (|~^3) AND TERM_PERIOD IN (|~^4) ORDER BY " & _
                                        "YIELD_CURVE_TYPE, TERM_PERIOD"



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
On Error GoTo Errhandler

    If astrConnectStringKey = "" Then astrConnectStringKey = cCONN_STRING_KEY

    gcnDataClass.Open GetConnectString(astrConnectStringKey)

    Exit Sub

Errhandler:
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
Attribute Disconnect.VB_Description = "Disconnects the current connection from the datasource."
On Error GoTo Errhandler

    gcnDataClass.Close
    Exit Sub

Errhandler:
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
On Error GoTo Errhandler

Dim llCounter   As Long
Dim lbQuotes    As Boolean
Dim lrstemp     As New ADODB.Recordset
Dim lstrSQL     As String
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

Errhandler:
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
On Error GoTo Errhandler

Dim lstrSQL     As String
Dim liPos       As Integer
Dim liLoopCtr   As Integer
Dim liUBound    As Integer

    Select Case alActionID
        Case ecTestSQL
            lstrSQL = cTestSQL
        Case ecGetIndexData
            lstrSQL = cGetIndexData
        Case ecGetIndexDataWithoutCondition
            lstrSQL = cGetIndexDataWithoutCondition

        '---All Enum to SQL constant mapping goes here---

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

Errhandler:
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
