VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBSMoneyCostAutoService"
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
'MODULE  : cBSMoneyCostAutoService
'PURPOSE : This will contain all DW Data related non-transactional
'          methods.
'================================================================

Option Explicit

'Constant for module name =======================================
Private Const cMODULE_NAME As String = "cBSMoneyCostAutoService"
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

'================================================================
'METHOD  : GetIndexData
'PURPOSE : Returns a string that contains all index rates from DWH
'PARMS   :
'          astrGetIndexDataXML [String] = Filter criteria
'          Sample XML structure:
'           <INDEX_DATA_REQUEST>
'               <PROCESSING_DATE>04/15/2005</PROCESSING_DATE>
'               <CURRENCY_CODE>USD</CURRENCY_CODE>
'               <YIELD_CURVE_TYPE_LIST>'CP','EURODOLLAR DEPOSITS','INTEREST RATE SWAP','LIBOR','PRIME','US TREASURY'</YIELD_CURVE_TYPE_LIST>
'               <DWC_TERM_PERIOD_LIST>'0','1','3','6','12','24','36','48','60','84','120','240','360'</DWC_TERM_PERIOD_LIST>
'           </INDEX_DATA_REQUEST>
'RETURN  : String = list of index rates records as XML string
'          Sample XML structure:
'           <INDEX_DATA_RESPONSE>
'               <INDEX_DATASet>
'                   <INDEX_DATA>
'                       <!-- Details from index rates data from DWH -->
'                       <YIELD_CURVE_TYPE>INTEREST RATE SWAP</YIELD_CURVE_TYPE>
'                       <TERM_PERIOD>12</TERM_PERIOD>
'                       <INTEREST_RATE>1.2454</INTEREST_RATE>
'                   </INDEX_DATA>
'                   ...
'               </INDEX_DATASet>
'           </INDEX_DATA_RESPONSE>
'================================================================
Public Function GetIndexData(ByVal astrGetIndexDataXML As String) As String
On Error GoTo GetIndexData_ErrHandler

Dim lstrErrSrc     As String    'to store error source
Dim lstrMethodName As String    'to store method name
Dim lstrErrDesc    As String    'to store error description
Dim llErrNbr       As Long      'to store error number

    lstrMethodName = "GetIndexData"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjRSUtils             As New RecordsetUtils.IRecordsetUtils   'to access Recordset Utilities Component
Dim lobjRequestXmlDOM       As New DOMDocument40                    'to load Request XML
Dim lobjcDataClass          As New cDataClass                       'to access cDataClass method(s)
Dim lstrProcessingDate      As String                               'to store Processing Date, from Request XML
Dim lstrCurrencyCode        As String                               'to store Curreny Code of MC under processing, from Request XML
Dim lstrYieldCurveType      As String                               'to store comma separated Yield Curve Types/ Index Codes available in MoneyCost DB, from Request XML
Dim lstrTermPeriod          As String                               'to store comma separated Term Periods/ Index Terms available in MoneyCost DB, from Request XML
Dim lstrResult              As String                               'to store final Response XML
Dim lstrInterestRateType    As String                               'to store comma separated Interest Rate Types available in MoneyCost DB, from Request XML
Dim lstrSourceSystemName    As String                               'to store comma separated Source System Names available in MoneyCost DB, from Request XML
Dim lstrLastUpdatedFlag     As Boolean                               'to store whether the last_updated_flag is to be used in the dwh query or not. CAD should not used this flag , rest all currencies should use this.

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrGetIndexDataXML, ecDebugInputTrace
    End If

    'Validate if the Request XML is well-formed
    If Not lobjRequestXmlDOM.loadXML(astrGetIndexDataXML) Then
        'Raise Error
        Err.Raise lobjRequestXmlDOM.parseError.errorCode, _
                    , "Error loading Request XML to GetIndexData(). " & lobjRequestXmlDOM.parseError.reason
    End If

    'If Processing Date found in Request XML, fetch the value in local variable
    If IsXMLElementPresent(lobjRequestXmlDOM.documentElement, "PROCESSING_DATE") Then
        If Trim(lobjRequestXmlDOM.getElementsByTagName("PROCESSING_DATE").Item(0).Text) <> "" Then
            lstrProcessingDate = lobjRequestXmlDOM.getElementsByTagName("PROCESSING_DATE").Item(0).Text
        Else
            Err.Raise cINVALID_PARMS, , "Processing date not specified."
        End If
    Else
        Err.Raise cINVALID_PARMS, , "Processing date not specified."
    End If

    'If Currency Code found in Request XML, fetch the value in local variable
    If IsXMLElementPresent(lobjRequestXmlDOM.documentElement, "CURRENCY_CODE") Then
        If Trim(lobjRequestXmlDOM.getElementsByTagName("CURRENCY_CODE").Item(0).Text) <> "" Then
            lstrCurrencyCode = lobjRequestXmlDOM.getElementsByTagName("CURRENCY_CODE").Item(0).Text
        Else
            Err.Raise cINVALID_PARMS, , "Currency code not specified."
        End If
    Else
        Err.Raise cINVALID_PARMS, , "Currency code not specified."
    End If

    'If Yield Curve Type list found in Request XML, fetch the value in local variable
    If IsXMLElementPresent(lobjRequestXmlDOM.documentElement, "YIELD_CURVE_TYPE_LIST") Then
        If Trim(lobjRequestXmlDOM.getElementsByTagName("YIELD_CURVE_TYPE_LIST").Item(0).Text) <> "" Then
            lstrYieldCurveType = lobjRequestXmlDOM.getElementsByTagName("YIELD_CURVE_TYPE_LIST").Item(0).Text
        Else
            Err.Raise cINVALID_PARMS, , "Yield curve type list not specified."
        End If
    Else
        Err.Raise cINVALID_PARMS, , "Yield curve type list not specified."
    End If

    'If Term Period list found in Request XML, fetch the value in local variable
    If IsXMLElementPresent(lobjRequestXmlDOM.documentElement, "DWC_TERM_PERIOD_LIST") Then
        If Trim(lobjRequestXmlDOM.getElementsByTagName("DWC_TERM_PERIOD_LIST").Item(0).Text) <> "" Then
            lstrTermPeriod = lobjRequestXmlDOM.getElementsByTagName("DWC_TERM_PERIOD_LIST").Item(0).Text
        Else
            Err.Raise cINVALID_PARMS, , "Term period list not specified."
        End If
    Else
        Err.Raise cINVALID_PARMS, , "Term period list not specified."
    End If

    'If LAST_UPDATED_IND is found in Request XML , fetch the value in local variable.
    If IsXMLElementPresent(lobjRequestXmlDOM.documentElement, "LAST_UPDATED_IND") Then
        If Trim(lobjRequestXmlDOM.getElementsByTagName("LAST_UPDATED_IND").Item(0).Text) <> "" Then
            lstrLastUpdatedFlag = lobjRequestXmlDOM.getElementsByTagName("LAST_UPDATED_IND").Item(0).Text
        Else
            Err.Raise cINVALID_PARMS, , "The Last Updated Flag is to be used not is not specified."
        End If
    Else
        Err.Raise cINVALID_PARMS, , "The Last Updated Flag is to be used not is not specified."
    End If



    'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
    'method of Recordset Utilities component to form the Output XML, in local variable
    'Note : Currencies where lstrLastUpdatedFlag = 0 , the last_updated condition shouldnt be used.
    ' to fetch the data from the warehouse. If it is 1 , then this condition should be there in the query.

    If lstrLastUpdatedFlag = False Then
        lstrResult = lobjRSUtils.RSToXML("INDEX_DATA", lobjcDataClass.Execute(ecGetIndexDataWithoutCondition, ecRSExecuteRecords, lstrProcessingDate, lstrCurrencyCode, lstrYieldCurveType, lstrTermPeriod))
    Else
        lstrResult = lobjRSUtils.RSToXML("INDEX_DATA", lobjcDataClass.Execute(ecGetIndexData, ecRSExecuteRecords, lstrProcessingDate, lstrCurrencyCode, lstrYieldCurveType, lstrTermPeriod))
    End If

    'Return the XML as output
    GetIndexData = "<INDEX_DATA_RESPONSE>" & lstrResult & "</INDEX_DATA_RESPONSE>"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, GetIndexData, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

GetIndexData_CleanMemory:
    'clear all local object variables from memory
    Set lobjcDataClass = Nothing
    Set lobjRequestXmlDOM = Nothing
    Set lobjRSUtils = Nothing
    Exit Function

GetIndexData_ErrHandler:
    lstrErrSrc = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description

    GetIndexData = vbNullString

    'write error message to log file
    WriteToTextDebugFile cMODULE_NAME & lstrMethodName, BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    Err.Raise llErrNbr, lstrErrSrc, lstrErrDesc

    Resume GetIndexData_CleanMemory
End Function
