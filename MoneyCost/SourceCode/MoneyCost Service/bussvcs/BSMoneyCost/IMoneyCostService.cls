VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "IMoneyCostService"
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
'MODULE  : IMoneyCostService
'PURPOSE : This will contain non-transactional wrapper methods
'          to avoid multiple invocation of business component
'          methods by UI tier.
'================================================================

Option Explicit

'Constant for module name =======================================
Private Const cMODULE_NAME As String = "IMoneyCostService"

'================================================================
Private Const cError_Msg As String = "There was an error while retrieving data."

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
'PURPOSE : Wrapper method to fetch the list of all the Money Cost
'          files that the logged-in user is associated with
'PARMS   :
'          astrGetMCFilesXML [String] = Input paramters
'          to retrieve the information for the MC Files list.
'          Sample XML structure:
'           <MC_FILES_REQUEST>
'               <USER_GESSOUID>2ee87ed4-95af-11d6-a612-00d0b785330f</USER_GESSOUID>
'           </MC_FILES_REQUEST>
'RETURN  : String = information required to display list of MC Files
'          as XML string
'          Sample XML structure:
'           <USER_MC_FILE_RESPONSE>
'               <LDAPSEARCH>
'                   <RECORDS>
'                       <RECORD>
'                           <uid>500674626</uid>
'                           <givenname>Manish</givenname>
'                           <sn>Dhingra</sn>
'                           <objectClass>mdsMVLink</objectClass>
'                           <objectClass>geworkerperson</objectClass>
'                           <objectClass>geworkerassignment</objectClass>
'                           <objectClass>gessosecurity</objectClass>
'                           <objectClass>gemmfworkerperson</objectClass>
'                       </RECORD>
'                   </RECORDS>
'               </LDAPSEARCH>
'               <MC_FILE_RESPONSE>
'                   <MC_FILESet>
'                       <MC_FILE>
'                           <!-- Details from MC_SECURITY & MC_FILE tables -->
'                           <SQ_MC_ID>1</SQ_MC_ID>
'                           <MONEY_COST_FILE>MCUSD-US Money Cost</MONEY_COST_FILE>
'                       </MC_FILE>
'                       ...
'                   </MC_FILESet>
'               </MC_FILE_RESPONSE>
'           </USER_MC_FILE_RESPONSE>
'================================================================
Public Function GetMCFiles(ByVal astrGetMCFilesXML As String) As String
On Error GoTo GetMCFiles_ErrHandler

Dim lstrMethodName  As String   'to store method name
Dim lstrErrSrc      As String   'to store error source
Dim lstrErrDesc     As String   'to store error description
Dim llErrNbr        As Long     'to store error number

    lstrMethodName = "GetMCFiles"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjcMoneyCostService   As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods
Dim lobjBSLDAPIService2     As New BSLDAP.Iservice2     'object of BSLDAP component for getting logged-in user's details
Dim lobjRequestXmlDOM       As New DOMDocument40        'DOM object variable for loading Request XML
Dim lobjUserDetailXmlDOM    As New DOMDocument40        'DOM object variable for loading Response XML from LDAP
Dim lstrUSER_GESSOUID       As String                   'to store logged-in user's GESSOUID from Request XML
Dim lstrUSER_SSOID          As String                   'to store logged-in user's SSO ID, retrieved from LDAP
Dim lstrBSLDAPRequestXML    As String                   'to store Request XML for LDAP component
Dim lstrBSLDAPResponseXML   As String                   'to store Response XML from LDAP component
Dim lstrMCFileRequestXML    As String                   'to store dynamically build Request XML for cMoneyCostService.GetMCFile() method
Dim lstrMCFilesResponseXML  As String                   'to store Response XML from cMoneyCostService.GetMCFiles() method

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrGetMCFilesXML, ecDebugInputTrace
    End If

    'validate, if the Input Request XML is well-formed
    If Not lobjRequestXmlDOM.loadXML(astrGetMCFilesXML) Then
        'Raise Error
        Err.Raise lobjRequestXmlDOM.parseError.errorCode, _
                    , "Error loading Input XML to GetMCFiles(). " & lobjRequestXmlDOM.parseError.reason
    End If

    'If logged-in User SSO ID found in Input XML, fetch the value in local variable
    If IsXMLElementPresent(lobjRequestXmlDOM.documentElement, "USER_GESSOUID") Then
        If Trim(lobjRequestXmlDOM.getElementsByTagName("USER_GESSOUID").Item(0).Text) <> "" Then
            lstrUSER_GESSOUID = lobjRequestXmlDOM.getElementsByTagName("USER_GESSOUID").Item(0).Text
        End If
    End If

    'build Request XML for BSLDAP component for fetching logged-in User's details
    lstrBSLDAPRequestXML = "<LDAPSearch FETCH='uid,givenname,sn' ou='geworker'>" & _
                                "<LDAP_ATTRIB NAME='gessouid' OPERATOR='EQ'>" & UCase(lstrUSER_GESSOUID) & "</LDAP_ATTRIB>" & _
                            "</LDAPSearch>"

    'call BSLDAP component to fetch logged-in user's details
    lstrBSLDAPResponseXML = lobjBSLDAPIService2.GetUserDetailsByAttributes(lstrBSLDAPRequestXML)

    'validate, if the Input Request XML is well-formed
    If Not lobjUserDetailXmlDOM.loadXML(lstrBSLDAPResponseXML) Then
        'Raise Error
        Err.Raise lobjUserDetailXmlDOM.parseError.errorCode, _
                    , "Error loading Response XML from LDAP. " & lobjUserDetailXmlDOM.parseError.reason
    End If

    If IsXMLElementPresent(lobjUserDetailXmlDOM.documentElement, "//uid") Then
        If lobjUserDetailXmlDOM.getElementsByTagName("uid").length > 0 Then
            lstrUSER_SSOID = Trim(lobjUserDetailXmlDOM.getElementsByTagName("uid").Item(0).Text)
        Else
            Err.Raise cINVALID_PARMS, , "User's SSO ID could not found from LDAP."
        End If
    Else
        Err.Raise cINVALID_PARMS, , "User's SSO ID could not found from LDAP."
    End If

    lstrMCFileRequestXML = "<MC_FILES_REQUEST><USER_SSOID>" & lstrUSER_SSOID & "</USER_SSOID></MC_FILES_REQUEST>"

    'call GetMCFiles() method of cMoneyCostService class module to get logged-in user's Money Cost File List
    lstrMCFilesResponseXML = lobjcMoneyCostService.GetMCFiles(lstrMCFileRequestXML)

    'return final output XML
    GetMCFiles = "<USER_MC_FILE_RESPONSE>" & lstrBSLDAPResponseXML & lstrMCFilesResponseXML & "</USER_MC_FILE_RESPONSE>"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, GetMCFiles, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

GetMCFiles_CleanMemory:
    Set lobjBSLDAPIService2 = Nothing
    Set lobjcMoneyCostService = Nothing
    Set lobjRequestXmlDOM = Nothing
    Exit Function

GetMCFiles_ErrHandler:
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    'return error XML
    GetMCFiles = BuildErrorXML(llErrNbr, lstrErrSrc, lstrErrDesc, cError_Msg)

    'write error to log file
    WriteToTextDebugFile cCOMPONENT_NAME & "." & cMODULE_NAME & ".GetMCFiles ", _
                        BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    Resume GetMCFiles_CleanMemory
End Function

'================================================================
'METHOD  : GetMCFileDetails
'PURPOSE : Wrapper method to fetch the details of the selected
'          Money Cost file, selected by the logged-in user
'PARMS   :
'          astrGetMCFileDetailsXML [String] = Input paramters
'          to retrieve the information for the selected MC File
'          details.
'          Sample XML structure:
'           <MC_FILE_DETAIL_REQUEST>
'               <SQ_MC_ID>1</SQ_MC_ID>
'           </MC_FILE_DETAIL_REQUEST>
'RETURN  : String = information required to display selected MC File
'          details as XML string
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

Dim lstrMethodName  As String   'to store method name
Dim lstrErrSrc      As String   'to store error source
Dim lstrErrDesc     As String   'to store error description
Dim llErrNbr        As Long     'to store error number

    lstrMethodName = "GetMCFileDetails"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjcMoneyCostService   As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrGetMCFileDetailsXML, ecDebugInputTrace
    End If

    'call GetMCFileDetails() method of cMoneyCostService class module to get selected Money Cost File details,
    'and return final outpur Response XML
    GetMCFileDetails = lobjcMoneyCostService.GetMCFileDetails(astrGetMCFileDetailsXML)

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, GetMCFileDetails, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

GetMCFileDetails_CleanMemory:
    Set lobjcMoneyCostService = Nothing
    Exit Function

GetMCFileDetails_ErrHandler:
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    'return error XML
    GetMCFileDetails = BuildErrorXML(llErrNbr, lstrErrSrc, lstrErrDesc, cError_Msg)

    'write error to log file
    WriteToTextDebugFile cCOMPONENT_NAME & "." & cMODULE_NAME & ".GetMCFileDetails ", _
                        BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    Resume GetMCFileDetails_CleanMemory
End Function

'================================================================
'METHOD  : GetAllMCFiles
'PURPOSE : Wrapper method to fetch the list of all the Money Cost
'          files that the logged-in user is associated with
'PARMS   : None
'RETURN  : String = information required to display list of MC Files
'          as XML string
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
'                       <DAYS_TO_SKIP>2</DAYS_TO_SKIP>
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

Dim lstrMethodName  As String   'to store method name
Dim lstrErrSrc      As String   'to store error source
Dim lstrErrDesc     As String   'to store error description
Dim llErrNbr        As Long     'to store error number

    lstrMethodName = "GetAllMCFiles"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjcMoneyCostService   As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "", ecDebugInputTrace
    End If

    'call GetAllMCFiles() method of cMoneyCostService class module
    'to get all Money Cost File List, to return response XML
    GetAllMCFiles = lobjcMoneyCostService.GetAllMCFiles()

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, GetAllMCFiles, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

GetAllMCFiles_CleanMemory:
    Set lobjcMoneyCostService = Nothing
    Exit Function

GetAllMCFiles_ErrHandler:
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    'return error XML
    GetAllMCFiles = BuildErrorXML(llErrNbr, lstrErrSrc, lstrErrDesc, cError_Msg)

    'write error to log file
    WriteToTextDebugFile cCOMPONENT_NAME & "." & cMODULE_NAME & ".GetAllMCFiles ", _
                        BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    Resume GetAllMCFiles_CleanMemory
End Function

'================================================================
'METHOD  : GetIndexRates
'PURPOSE : Wrapper method to fetch the Index Rates, associated
'          with particular Money Cost file
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

Dim lstrMethodName  As String   'to store method name
Dim lstrErrSrc      As String   'to store error source
Dim lstrErrDesc     As String   'to store error description
Dim llErrNbr        As Long     'to store error number

    lstrMethodName = "GetIndexRates"

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "In " & lstrMethodName & "() method", ecDebugLogData
    End If

Dim lobjcMoneyCostService   As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, astrGetIndexRatesXML, ecDebugInputTrace
    End If

    'call GetIndexRates() method of cMoneyCostService class module
    'to get Index Rates for particular Money Cost File, to return response XML
    GetIndexRates = lobjcMoneyCostService.GetIndexRates(astrGetIndexRatesXML)

    If giDebugLevel >= ecDebugInputTrace Then
        'write details in log file
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, GetIndexRates, ecDebugOutputTrace
        WriteToTextDebugFile cMODULE_NAME & lstrMethodName, "Exit " & lstrMethodName & "() Method", ecDebugLogData
    End If

GetIndexRates_CleanMemory:
    Set lobjcMoneyCostService = Nothing
    Exit Function

GetIndexRates_ErrHandler:
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    'return error XML
    GetIndexRates = BuildErrorXML(llErrNbr, lstrErrSrc, lstrErrDesc, cError_Msg)

    'write error to log file
    WriteToTextDebugFile cCOMPONENT_NAME & "." & cMODULE_NAME & ".GetIndexRates ", _
                        BuildErrXML(llErrNbr, lstrErrSrc, lstrErrDesc), ecDebugCriticalError

    Resume GetIndexRates_CleanMemory
End Function
