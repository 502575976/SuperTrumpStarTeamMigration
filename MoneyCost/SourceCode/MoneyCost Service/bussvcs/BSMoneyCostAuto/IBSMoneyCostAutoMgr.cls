VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "IBSMoneyCostAutoMgr"
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
'MODULE  : IBSMoneyCostAutoMgr
'PURPOSE : This will contain transactional wrapper methods to
'          avoid multiple invocation of business component
'          methods by UI tier.
'================================================================

Option Explicit

'Constant for module name =======================================
Private Const cMODULE_NAME As String = "IBSMoneyCostAutoMgr"
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
    Set lrsTest = lobjDataClass.Execute(ecTestSQL, _
                                        ecRSExecuteRecords, _
                                        "")

    'Return the total records
    Test = "Retrieved " & lrsTest.Fields(0).Value & " records."
    lrsTest.Close

Test_CleanMemory:
    Set lrsTest = Nothing
    Set lobjDataClass = Nothing
    Exit Function

Test_ErrHandler:
    Test = vbNullString

    Err.Raise Err.Number, _
                cCOMPONENT_NAME & "." & cMODULE_NAME & ":Test()", _
                Err.Description

    Resume Test_CleanMemory
End Function
