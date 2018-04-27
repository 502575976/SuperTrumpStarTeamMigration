Imports BSMoneyCostDL
Imports System.Reflection
Imports System.IO
Imports BSMoneyCostEntity
Imports System.Data.SqlClient
Imports BSMoneyCostDL.MoneyCostAutoDataClass
Imports BSMoneyCostBL
Imports BSMoneyCostAuto.modCEFCommon
Imports System.EnterpriseServices
Imports System.Runtime.InteropServices

Public Interface IMoneyCostAutoSvc
    Sub ExecuteServiceFlow()
    Function ping() As String
    Function Test() As String
End Interface
<JustInTimeActivation(), _
EventTrackingEnabled(), _
ClassInterface(ClassInterfaceType.None), _
Transaction(TransactionOption.NotSupported, isolation:=TransactionIsolationLevel.Serializable, timeout:=120), _
ComponentAccessControl(True)> _
Public Class cMoneyCostAutoSvc
    Inherits ServicedComponent
    Implements IMoneyCostAutoSvc
    Dim STLogger As log4net.ILog

    Dim lstrMethodName As String = ""  'to store method name
    Dim lstrErrSrc As String = ""  'to store error source
    Dim lstrErrDesc As String = ""  'to store error description
    Dim llErrNbr As Long     'to store error number        
    Dim Count As Integer
    Dim lrsCSVRecordset As DataTable
    Dim lobjAllMCFileDOM As New Xml.XmlDocument
    Dim lobjIndexRateDOM As New Xml.XmlDocument
    Dim lobjIndexRateDOM_PrevDay As New Xml.XmlDocument
    Dim lobjIndexDataDOMXml As New Xml.XmlDocument
    Dim lobjCSVInsertDOMXml As New Xml.XmlDocument
    Dim lobjMCFileNodeList As Xml.XmlNodeList
    Dim lobjIndexRateNodeList As Xml.XmlNodeList
    Dim lobjIndexRateNodeList_PrevDay As Xml.XmlNodeList
    Dim lobjIndexDataNodeList As Xml.XmlNodeList
    Dim lobjCSVInsertNodeList As Xml.XmlNodeList
    Dim lobjMCFileNode As Xml.XmlNode
    Dim lobjIndexRateNode As Xml.XmlNode
    Dim lobjIndexRateNode_PrevDay As Xml.XmlNode
    Dim lobjIndexDataNode As Xml.XmlNode
    Dim lobjCSVInsertNode As Xml.XmlNode
    Dim liCounter1 As Integer
    Dim liCounter2 As Integer
    Dim liCounter3 As Integer
    Dim liSQ_MC_ID As Integer
    Dim liDAYS_TO_SKIP As Integer
    Dim lbMarketOpenFlag As Boolean
    Dim lstrMC_CODE As String = ""
    Dim lstrDESCRIPTION As String = ""
    Dim lstrSTART_TIME As String = ""
    Dim lstrEND_TIME As String = ""
    Dim lstrWorkingDirectory As String = ""
    Dim lstrCopyFileResponse As String = ""
    Dim lstrProcessDate As String = ""
    Dim lstrProcessDate_PrevDay As String = ""
    Dim lstrIndexRateReqXml As String = ""
    Dim lstrIndexRateRespXml As String = ""
    Dim lstrIndexRateRespXml_PrevDay As String = ""
    Dim lstrINDEX_CODEList As String = ""
    Dim lstrINDEX_CODEList_PrevDay As String
    Dim lstrINDEX_TERMList As String
    Dim lstrProcessDateFieldName As String = ""
    Dim lstrCURRENCY_CODE As String = ""
    Dim lstrIndexDataReqXml As String = ""
    Dim lstrIndexDataRespXml As String = ""
    Dim lstrGetMCFilesRespXml As String = ""
    Dim lstrINTEREST_RATE_DWH As String = ""
    Dim lstrMissingYieldCurve As String = ""
    Dim lstrMissingInterestRate As String = ""
    Dim lstrCSV_INSERT_RECORDXml As String = ""
    Dim lstrCSVInsertSQL As String = ""
    Dim lstrClarifyQName As String = ""
    Dim lstrBusinessContact As String = ""
    Dim lstrSendErrNotiResult As String = ""
    Dim lstrSqlQry As String = ""
    Dim lstrCSVColumnHeaders As String = ""
    Dim lstrMCDBAllIndexCode As String = ""
    Dim lstrMCDBAllIndexTerm As String = ""
    Dim lstrMCFileStartExec As String = ""
    Dim lstrMCFileEndExec As String = ""
    Dim lstrUpdateMCLogReqXml As String = ""
    Dim lstrUpdateMCLogRespXml As String = ""
    Dim lstrCommonErrorDetails As String = ""
    Dim lstrErrorDetails As String = ""
    Dim larrMissingIndexCode() As String
    Dim larrMissingIntRate() As String
    Dim lstrBackup_Location As String = ""
    Dim lstrNetwork_Location As String = ""
    Dim lstrFTP_Location As String = ""
    Dim lstrFTP_Directory As String = ""
    Dim lstrFTP_User As String = ""
    Dim lstrFTP_Password As String = ""
    Dim lstrBlankRegistryKey As String = ""
    Dim lbDataDeletionFlag As Boolean
    Dim lbServiceRunFlag As Boolean
    Dim lstrFREQUENCY As String = ""
    Dim liFREQUENCY_COUNT As Integer
    Dim lstrLAST_SCHEDULE_PROCESS_DATE As String = ""
    Dim liMARKET_CLOSED_DWH_CHECK_COUNTER As Integer
    Dim lstrScheduleProcessDate As String = ""
    Dim lstrUpdateMCFileReqXml As String = ""
    Dim lstrUpdateMCFileRespXml As String = ""
    Dim lrsCSVDeleteRecordset As DataTable
    Dim lstrLAST_UPDATED_IND As Boolean
    Dim lstrFTP_LocationForNewDateFormat As String = ""
    Dim lstrFTP_DirectoryForNewDateFormat As String = ""
    Dim lstrFTP_UserForNewDateFormat As String = ""
    Dim lstrFTP_PasswordForNewDateFormat As String = ""
    Dim lstrDateFormat As String = ""
    Dim lstrDateFormatRequired As Boolean
    Dim lstrMissingYieldCurve_NotCopied As String = ""
    Dim larrMissingIndexCode_NotCopied() As String
    Dim lstrProcessDate_NewDateFormat As String = ""
    Dim lstrDay_NewDateFormat As String = ""
    Dim lstrMonth_NewDateFormat As String = ""
    Dim objcDataEntity As New cDataEntity
    Dim objFileEntity As New FileInfoEntity
    Dim objxmlErrEntity As New XmlErrEntity
    Dim EncrptFileFlag As String
    Public Sub SetLog4Net()
        Try
            If log4net.LogManager.GetRepository.Configured = False Then
                log4net.Config.XmlConfigurator.ConfigureAndWatch(New System.IO.FileInfo(GetConfigurationKey("MoneyCostLog4Net")))
            End If
            STLogger = log4net.LogManager.GetLogger("MoneyCost")
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Public Sub SetConfigSettings()
        Try
            'gstrErrorLogFile = GetConfigurationKey(cDEBUG_ERROR_FILE_PATH_NAME_KEY)
            gstrErrMailBox = GetConfigurationKey(cErrorMailBoxKey)
            gstrEmailOverride = GetConfigurationKey(cEmailOverrideKey)
            gstrDeveloperEmail = GetConfigurationKey(cDeveloperEmailKey)
            gstrClarifySiteId = GetConfigurationKey(cClarifySiteIdKey)
            gstrClarifyEmailSub = GetConfigurationKey(cClarifyEmailSubject)
            gstrClarifyPriority = GetConfigurationKey(cClarifyPriorityKey)
            gstrClarifyContactFname = GetConfigurationKey(cClarifyContactFNameKey)
            gstrClarifyContactLname = GetConfigurationKey(cClarifyContactLNameKey)
            gstrClarifyContactPhone = GetConfigurationKey(cClarifyContactPhoneKey)
            gstrFrom = GetConfigurationKey(cEmailFromKey)
            EncrptFileFlag = GetConfigurationKey(cEncrptFileFlag)
        Catch ex As Exception
            Throw
        End Try
    End Sub
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
    <AutoComplete()> _
    Public Function Ping() As String Implements IMoneyCostAutoSvc.ping
        Return "Ping request to " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & "." & cMODULE_NAME & " returned at " & Format(Now, "mm/dd/yyyy Hh:Nn:Ss AM/PM") & " server time."
    End Function

    '================================================================
    'METHOD  : Test
    'PURPOSE : Returns a string that indicates that the component
    '          can connect to the database and the registry.
    'PARMS   : NONE
    'RETURN  : String
    '================================================================
    <AutoComplete()> _
    Public Function Test() As String Implements IMoneyCostAutoSvc.Test
        Dim lobjDataClass As New cDataClass
        Dim lrsTest As New DataTable
        Try
            'Execute the Test SQL statement which returns a count of the records
            lrsTest = lobjDataClass.Test
            'Return the total records
            Return "Retrieved " & lrsTest.Rows(0)(0).ToString & " records."
        Catch ex As Exception
            Return vbNullString
            Throw
        Finally
            If Not IsNothing(lrsTest) Then
                lrsTest = Nothing
            End If
            If Not IsNothing(lobjDataClass) Then
                lobjDataClass = Nothing
            End If
        End Try
    End Function

    '================================================================
    'METHOD  : ExecuteServiceFlow
    'PURPOSE : Main Controller procedure for the service flow
    'PARMS   : NONE
    'RETURN  : NONE
    '================================================================

    <AutoComplete()> _
    Public Function UpdateFile(ByVal lobjEntity As cDataEntity) As cDataEntity

        Dim lTSFileStreamHandle As StreamReader = Nothing
        Dim lTSFileWriteHandle As StreamWriter = Nothing
        Dim lvarData() As Object
        Dim lstrSearch As String
        Dim lstrText As String
        Dim llLine As Long
        Dim liCount As Long
        SetLog4Net()
        Try
            lstrSearch = lobjEntity.ProcessDate.GetDateTimeFormats.GetValue(3).ToString
            lTSFileStreamHandle = File.OpenText(lobjEntity.astrFileName)
            llLine = 0

            lstrText = lTSFileStreamHandle.ReadLine
            ReDim Preserve lvarData(llLine)
            lvarData(llLine) = lstrText
            llLine = llLine + 1

            If lstrSearch <> "" Then
                While lTSFileStreamHandle.EndOfStream = False
                    lstrText = lTSFileStreamHandle.ReadLine

                    If InStr(1, lstrText, lstrSearch, vbBinaryCompare) = 0 Then
                        ReDim Preserve lvarData(llLine)

                        lvarData(llLine) = lstrText
                        llLine = llLine + 1
                    End If
                End While
            End If

            lTSFileStreamHandle.Close()

            If lvarData.Length = 0 Then
                lobjEntity.OutputString = "True"
                Return lobjEntity
                Exit Function
            End If

            'lTSFileWriteHandle = File.OpenWrite(lobjEntity.astrFileName)
            lTSFileWriteHandle = New StreamWriter(lobjEntity.astrFileName)

            For liCount = 0 To UBound(lvarData)
                lTSFileWriteHandle.WriteLine(lvarData(liCount))
            Next
            lTSFileWriteHandle.Close()
            lobjEntity.OutputString = "True"
            Return lobjEntity
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            lobjEntity.OutputString = "False"
            Return lobjEntity
            Throw
        Finally
            lTSFileStreamHandle.Close()
            lTSFileWriteHandle.Close()
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
            If Not IsNothing(lTSFileStreamHandle) Then
                lTSFileStreamHandle = Nothing
            End If
            If Not IsNothing(lTSFileWriteHandle) Then
                lTSFileWriteHandle = Nothing
            End If
        End Try

    End Function
    <AutoComplete()> _
    Private Function IsFTPLocationMapped(ByVal objFtpEntity As FTPEntity) As cDataEntity
        Dim lobjEntity As New cDataEntity
        Dim lstrUNCPath As String
        Try
            lstrUNCPath = objFtpEntity.FTPLocation
            If Directory.Exists(lstrUNCPath) Then
                lobjEntity.OutputString = "True"
                Return lobjEntity
            Else
                If Convert.ToBoolean(MapDrive(objFtpEntity).OutputString) Then
                    lobjEntity.OutputString = "True"
                    Return lobjEntity
                Else
                    lobjEntity.OutputString = "False"
                    Return lobjEntity
                End If
            End If
        Catch ex As Exception
            lobjEntity.OutputString = "False"
            Return lobjEntity
            Throw
        Finally
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
            If Not IsNothing(objFtpEntity) Then
                objFtpEntity = Nothing
            End If
        End Try

    End Function

#Region " MapDrive Code Block "
    <AutoComplete()> _
    Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" _
    (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, _
    ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer

    Public Declare Function WNetCancelConnection2 Lib "mpr" Alias "WNetCancelConnection2A" _
  (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer

    Public Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String
    End Structure

    Public Const ForceDisconnect As Integer = 1
    Public Const RESOURCETYPE_DISK As Long = &H1
    <AutoComplete()> _
    Public Function MapDrive(ByVal objFtpEntity As FTPEntity) As cDataEntity
        Dim cdataEntity As New cDataEntity
        Dim nr As NETRESOURCE
        Dim strUsername As String
        Dim strPassword As String
        Try
            nr = New NETRESOURCE
            nr.lpRemoteName = objFtpEntity.FTPLocation
            nr.lpLocalName = ""     'DriveLetter & ":"
            strUsername = objFtpEntity.FTPUser '(add parameters to pass this if necessary)
            strPassword = objFtpEntity.FTPPassword '(add parameters to pass this if necessary)
            nr.dwType = RESOURCETYPE_DISK

            Dim result As Integer
            result = WNetAddConnection2(nr, strPassword, strUsername, 0)

            If result = 0 Then
                cdataEntity.OutputString = "True"
                MapDrive = cdataEntity
            Else
                cdataEntity.OutputString = "False"
                MapDrive = cdataEntity
            End If
        Catch ex As Exception
            cdataEntity.OutputString = "False"
            MapDrive = cdataEntity
            Throw
        Finally
            cdataEntity = Nothing
            objFtpEntity = Nothing
        End Try

    End Function

    'Public Function UnMapDrive(ByVal DriveLetter As String) As Boolean
    '    Dim rc As Integer
    '    rc = WNetCancelConnection2(DriveLetter & ":", 0, ForceDisconnect)

    '    If rc = 0 Then
    '        Return True
    '    Else
    '        Return False
    '    End If

    'End Function
#End Region
    <AutoComplete()> _
    Public Function GetIndexData(ByVal astrGetIndexDataXML As String) As String


        Dim lstrErrSrc As String = ""   'to store error source
        Dim lstrMethodName As String    'to store method name
        Dim lstrErrDesc As String = ""   'to store error description
        Dim llErrNbr As Long = 0    'to store error number
        Dim lobjRequestXmlDOM As New Xml.XmlDocument                    'to load Request XML
        Dim lobjcDataClass As New MoneyCostAutoDataClass                   'to access cDataClass method(s)
        Dim lstrProcessingDate As String = ""                             'to store Processing Date, from Request XML
        Dim lstrCurrencyCode As String = ""                             'to store Curreny Code of MC under processing, from Request XML
        Dim lstrYieldCurveType As String = ""                             'to store comma separated Yield Curve Types/ Index Codes available in MoneyCost DB, from Request XML
        Dim lstrTermPeriod As String = ""                          'to store comma separated Term Periods/ Index Terms available in MoneyCost DB, from Request XML
        Dim lstrResult As String                               'to store final Response XML
        Dim lstrInterestRateType As String = ""                              'to store comma separated Interest Rate Types available in MoneyCost DB, from Request XML
        Dim lstrSourceSystemName As String = ""                             'to store comma separated Source System Names available in MoneyCost DB, from Request XML
        Dim lstrLastUpdatedFlag As Boolean                               'to store whether the last_updated_flag is to be used in the dwh query or not. CAD should not used this flag , rest all currencies should use this.
        Dim objcdataEntity As New cDataEntity
        Dim objxmlErrEntity As New XmlErrEntity
        SetLog4Net()
        Try
            lstrMethodName = "GetIndexData"
            'write details in log file
            STLogger.Debug(cMODULE_NAME & lstrMethodName & "In " & lstrMethodName & "() method")
            'write details in log file
            STLogger.Debug(cMODULE_NAME & lstrMethodName & astrGetIndexDataXML)

            'Validate if the Request XML is well-formed
            Try
                lobjRequestXmlDOM.LoadXml(astrGetIndexDataXML)
            Catch ex As Exception
                'Raise Error            
                STLogger.Error(Err.Number & "Error loading Request XML to GetIndexData(). " & Err.Description)
            End Try


            'If Processing Date found in Request XML, fetch the value in local variable

            objxmlErrEntity.IXMLDOMNode = lobjRequestXmlDOM.DocumentElement
            objxmlErrEntity.ElementXPath = "PROCESSING_DATE"
            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                If Trim(lobjRequestXmlDOM.GetElementsByTagName("PROCESSING_DATE").Item(0).InnerText) <> "" Then
                    lstrProcessingDate = lobjRequestXmlDOM.GetElementsByTagName("PROCESSING_DATE").Item(0).InnerText
                Else
                    STLogger.Error(cINVALID_PARMS & "Processing date not specified.")
                End If
            Else
                STLogger.Error(cINVALID_PARMS & "Processing date not specified.")
            End If

            'If Currency Code found in Request XML, fetch the value in local variable
            objxmlErrEntity.IXMLDOMNode = lobjRequestXmlDOM.DocumentElement
            objxmlErrEntity.ElementXPath = "CURRENCY_CODE"
            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                If Trim(lobjRequestXmlDOM.GetElementsByTagName("CURRENCY_CODE").Item(0).InnerText) <> "" Then
                    lstrCurrencyCode = lobjRequestXmlDOM.GetElementsByTagName("CURRENCY_CODE").Item(0).InnerText
                Else
                    STLogger.Error(cINVALID_PARMS & "Currency code not specified.")
                End If
            Else
                STLogger.Error(cINVALID_PARMS & "Currency code not specified.")
            End If

            'If Yield Curve Type list found in Request XML, fetch the value in local variable
            objxmlErrEntity.IXMLDOMNode = lobjRequestXmlDOM.DocumentElement
            objxmlErrEntity.ElementXPath = "YIELD_CURVE_TYPE_LIST"
            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                If Trim(lobjRequestXmlDOM.GetElementsByTagName("YIELD_CURVE_TYPE_LIST").Item(0).InnerText) <> "" Then
                    lstrYieldCurveType = lobjRequestXmlDOM.GetElementsByTagName("YIELD_CURVE_TYPE_LIST").Item(0).InnerText
                Else
                    STLogger.Error(cINVALID_PARMS & "Yield curve type list not specified.")
                End If
            Else
                STLogger.Error(cINVALID_PARMS & "Yield curve type list not specified.")
            End If

            'If Term Period list found in Request XML, fetch the value in local variable
            objxmlErrEntity.IXMLDOMNode = lobjRequestXmlDOM.DocumentElement
            objxmlErrEntity.ElementXPath = "DWC_TERM_PERIOD_LIST"
            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                If Trim(lobjRequestXmlDOM.GetElementsByTagName("DWC_TERM_PERIOD_LIST").Item(0).InnerText) <> "" Then
                    lstrTermPeriod = lobjRequestXmlDOM.GetElementsByTagName("DWC_TERM_PERIOD_LIST").Item(0).InnerText
                Else
                    STLogger.Error(cINVALID_PARMS & "Term period list not specified.")
                End If
            Else
                STLogger.Error(cINVALID_PARMS & "Term period list not specified.")
            End If

            'If LAST_UPDATED_IND is found in Request XML , fetch the value in local variable.
            objxmlErrEntity.IXMLDOMNode = lobjRequestXmlDOM.DocumentElement
            objxmlErrEntity.ElementXPath = "LAST_UPDATED_IND"
            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                If Trim(lobjRequestXmlDOM.GetElementsByTagName("LAST_UPDATED_IND").Item(0).InnerText) <> "" Then
                    lstrLastUpdatedFlag = lobjRequestXmlDOM.GetElementsByTagName("LAST_UPDATED_IND").Item(0).InnerText
                Else
                    STLogger.Error(cINVALID_PARMS & "The Last Updated Flag is to be used not is not specified.")
                End If
            Else
                STLogger.Error(cINVALID_PARMS & "The Last Updated Flag is to be used not is not specified.")
            End If



            'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
            'method of Recordset Utilities component to form the Output XML, in local variable
            'Note : Currencies where lstrLastUpdatedFlag = 0 , the last_updated condition shouldnt be used.
            ' to fetch the data from the warehouse. If it is 1 , then this condition should be there in the query.

            lstrProcessingDate = lstrProcessingDate
            objcdataEntity.ProcessDate = lstrProcessingDate
            objcdataEntity.CurrencyCode = lstrCurrencyCode
            objcdataEntity.YieldCurveType = lstrYieldCurveType
            objcdataEntity.TermPeriod = lstrTermPeriod
            objcdataEntity.QuoteReplacement = "False"

            If lstrLastUpdatedFlag = False Then
                objcdataEntity.ActionID = eDCActions.ecGetIndexDataWithoutCondition
                lstrResult = lobjcDataClass.GetIndexData(objcdataEntity).OutputString
            Else
                objcdataEntity.ActionID = eDCActions.ecGetIndexData
                lstrResult = lobjcDataClass.GetIndexData(objcdataEntity).OutputString
            End If

            'Return the XML as output
            Return "<INDEX_DATA_RESPONSE>" & lstrResult & "</INDEX_DATA_RESPONSE>"
            'write details in log file
            STLogger.Debug(cMODULE_NAME & lstrMethodName & GetIndexData & eDebugLevel.ecDebugOutputTrace)
            STLogger.Debug(cMODULE_NAME & lstrMethodName & "Exit " & lstrMethodName & "() Method")
        Catch ex As Exception
            lstrErrSrc = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
            llErrNbr = Err.Number
            lstrErrDesc = Err.Description
            Return vbNullString
            'write error message to log file
            objxmlErrEntity.ErrNbr = Err.Number
            objxmlErrEntity.ErrSource = Err.Source
            objxmlErrEntity.ErrDesc = Err.Description
            STLogger.Error(cMODULE_NAME & lstrMethodName & BuildErrXMLauto(objxmlErrEntity).OutputString & eDebugLevel.ecDebugCriticalError)
            Throw
        Finally
            'clear all local object variables from memory
            If Not IsNothing(lobjcDataClass) Then
                lobjcDataClass.Dispose()
                lobjcDataClass = Nothing
            End If
            If Not IsNothing(lobjRequestXmlDOM) Then
                lobjRequestXmlDOM = Nothing
            End If
            If Not IsNothing(objxmlErrEntity) Then
                objxmlErrEntity = Nothing
            End If
        End Try

    End Function

#Region "Process Treasury Assessment Block"
    'Treasury Assessment
    <AutoComplete()> _
    Public Sub ProcessTreasuryAssessment()
        Dim lobjMoneyCostDataClass As New cDataClass
        Dim lobjBSMCIMoneyCostService As New BSMoneyCostBL.cMoneyCostUISvc
        Dim lobjcDataClass As New MoneyCostAutoDataClass
        Dim lobjcMoneyCostUISvc As New cMoneyCostUISvc
        Dim dsDWHData As DataSet, dtCSVColNames As DataTable
        Dim drDwhData As DataRow, drFilteredData As DataRow
        Dim bDatasetComparison As Boolean = True
        Dim strMailBody As String = ""
        'Dim dsMoneyCostData As DataSet
        Dim objMoneyCostAutoDataClass As New MoneyCostAutoDataClass
        Dim ProcessDate As String, strLastProcessDateInCSV As String = ""
        Dim strProcessStartDate As String, strProcessEndDate As String, strLogMessage As String, strCostTypes As String
        Dim intStatus As Int32 = 0
        Dim strQuater As String = ""

        If IsNothing(objcDataEntity) Then
            objcDataEntity = New cDataEntity
        End If

        If IsNothing(objFileEntity) Then
            objFileEntity = New FileInfoEntity
        End If

        strProcessStartDate = ""
        strProcessEndDate = ""
        strLogMessage = ""
        strCostTypes = ""

        ' Set Global Variables.
        SetConfigSettings()
        SetConfigValue()


        Try
            ' Set ProcessDate
            ' -------------------------------------------
            ' Get Details of TreasuryAssessment Run
            ' -------------------------------------------                
            lstrGetMCFilesRespXml = lobjBSMCIMoneyCostService.GetTreasuryDetails().OutputString
            'validate, if response xml is well-formed
            Try
                If IsNothing(lobjAllMCFileDOM) Then lobjAllMCFileDOM = New Xml.XmlDocument

                lobjAllMCFileDOM.LoadXml(lstrGetMCFilesRespXml)
            Catch ex As Exception
                lstrCommonErrorDetails = "Error No. : " & Err.Number & vbCrLf & "Error Description : " & _
                                    "Error loading response XML from BSMOneyCost.IMoneyCostService.GetTreasuryDetails(). " & _
                                    Err.Description
                STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Get Treasury Detail Error" & lstrCommonErrorDetails)
                Throw
            End Try
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Get Treasury Detail Complete")
            lobjMCFileNodeList = lobjAllMCFileDOM.SelectNodes("/MC_FILE_RESPONSE/TREASURY_DETAIL/TREASURY_DETAIL")

            If lobjMCFileNodeList.Count <= 0 Then
                lstrCommonErrorDetails = "No Treasury file found to process."
                STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): No Treasury file found to process.")
                Throw New Exception(lstrCommonErrorDetails)
                strLogMessage = "No Treasury Data to Update."
            End If

            ''''Set Process Flag lbServiceRunFlag True Or False
            Call SetXmlValue(lstrFTP_DirectoryForNewDateFormat, lstrFTP_Directory)

            lstrMC_CODE = "TreasuryAssessment.csv"

            Call SetDateProcessFlag()

            ProcessDate = lstrScheduleProcessDate
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Process date  END ; Process Flag=" & lbServiceRunFlag.ToString)

            If lbServiceRunFlag = True Then

                strProcessStartDate = Now()

                '*****************************************************************
                'get first data row from Treasury file, to get date field name
                ' Fetch Adder Data from MoneyCost DB
                'objcDataEntity.ProcessDate = "02/18/2012"
                'dsMoneyCostData = lobjcMoneyCostUISvc.GetTreasuryAssessmentData(objcDataEntity)

                'Dim flag As Boolean = dsMoneyCostData.Equals(dsMoneyCostData)

                ' Loop thru Moneycost DB and add adder to DWH entries -- [Resultant Dataset]
                ' Adder part is on hold as of now -- 03/16/2012                

                Dim objFtp As New FTPEntity
                objFtp.FTPLocation = lstrFTP_Location
                objFtp.FTPUser = lstrFTP_User
                objFtp.FTPPassword = lstrFTP_Password
                If Convert.ToBoolean(IsFTPLocationMapped(objFtp).OutputString) Then
                    NotificationMailFlag = False

                    If Right(lstrWorkingDirectory, 1) <> "\" Then lstrWorkingDirectory = lstrWorkingDirectory & "\"
                    If Right(lstrFTP_Location, 1) <> "\" Then lstrFTP_Location = lstrFTP_Location & "\"
                    If Right(lstrFTP_Directory, 1) <> "\" Then lstrFTP_Directory = lstrFTP_Directory & "\"
                    ' -----------------------------------------------------------------
                    ' Create a Backup Copy of existing TreasuryAssessment File as BKMMDDYY.TreasuryAssessment.csv
                    ' on Backup location, from FTP location
                    ' -----------------------------------------------------------------
                    objcDataEntity.ProcessDate = Now.Date
                    objFileEntity.Source = lstrFTP_Location & lstrFTP_Directory & lstrMC_CODE
                    objFileEntity.Destination = lstrBackup_Location & "\BK" & GetDateFormat(objcDataEntity).OutputString & lstrMC_CODE
                    lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString

                    'if any error occurred
                    If lstrCopyFileResponse <> "" Then
                        objcDataEntity.ProcessDate = Now.Date
                        lstrErrorDetails = "Error while creating backup copy of existing " & lstrMC_CODE & " file as " & _
                                            lstrBackup_Location & "BK" & GetDateFormat(objcDataEntity).OutputString & "." & lstrMC_CODE & " on " & _
                                            lstrBackup_Location & " : " & lstrCopyFileResponse


                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment():" & lstrErrorDetails)
                        lstrCommonErrorDetails = lstrErrorDetails
                        Throw New Exception(lstrErrorDetails)
                    End If

                    ' Invoke Common Method.CopyFiles() to copy latest TreasuryAssessment file from Backup Location to working Direcotry.
                    ' Working directory value should be defined in Registry
                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Get Latest " & lstrMC_CODE & " file from Network Path Start")

                    ' ---------------------------------------
                    ' Get Latest TreasuryAssessment file from Network Path
                    ' ---------------------------------------
                    ' Invoke Common Method.CopyFiles() to copy latest TreasuryAssessment file from Backup Location to working Direcotry.
                    ' Working directory value should be defined in Registry 
                    objcDataEntity.ProcessDate = Format(Convert.ToDateTime(Now.Date), "MM/dd/yyyy")
                    objFileEntity.Source = lstrBackup_Location & "\BK" & GetDateFormat(objcDataEntity).OutputString & lstrMC_CODE
                    objFileEntity.Destination = lstrWorkingDirectory & lstrMC_CODE
                    lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString

                    'Add these lines to decrypt the files #sumit
                    If EncrptFileFlag = "True" Then
                        If EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE).ToString.ToUpper = "ENCRYPTED" Then
                            Dim strtest As String = EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE)
                        End If
                    End If

                    'if error occurred
                    If lstrCopyFileResponse <> "" Then
                        lstrErrorDetails = "Error while getting latest " & lstrMC_CODE & " to working " & "directory " & _
                                            lstrWorkingDirectory & " : " & lstrCopyFileResponse
                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment():" & lstrErrorDetails)
                        lstrCommonErrorDetails = lstrErrorDetails
                        Throw New Exception(lstrErrorDetails)
                    End If

                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Get Latest " & lstrMC_CODE & " file from Network Path Start END")
                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Copy TreasuryAssessment.ini file as schema.ini file in working direcotry Start")


                    '''''' Fetch CSV Columns. This should be executed before copying INI file.
                    '''''objcDataEntity.CommonSQL = "SELECT Top 1 * FROM " & lstrMC_CODE
                    '''''objcDataEntity.WrkDirectory = lstrWorkingDirectory
                    '''''Dim dtColNamesFromCSV As DataTable = lobjcDataClass.GetCsvRecords(objcDataEntity).CsvOutput

                    ' -----------------------------------------------------------
                    ' Copy TreasuryAssessment.ini file as schema.ini file in working direcotry
                    ' -----------------------------------------------------------
                    ' Invoke Common Method.CopyFiles()
                    objFileEntity.Source = lstrWorkingDirectory & "\TreasuryAssessment.ini"
                    objFileEntity.Destination = lstrWorkingDirectory & "\Schema.ini"
                    lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString

                    'if any error occurred, raise clarify case
                    If lstrCopyFileResponse <> "" Then

                        'schema file not present for TreasuryAssessment.csv file.
                        'Details: Schema file not defined for TreasuryAssessment.csv

                        lstrErrorDetails = "Schema file not present for " & lstrMC_CODE & vbCrLf & _
                                            "Error while defining schema file : " & lstrCopyFileResponse
                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment():" & lstrErrorDetails)
                        lstrCommonErrorDetails = lstrErrorDetails
                        Throw New Exception(lstrErrorDetails)
                    End If
                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Copy TreasuryAssessment.ini file as schema.ini file in working direcotry END")


                    ' Fetch Column Name.
                    objcDataEntity.CommonSQL = "SELECT TOP 1 * FROM " & lstrMC_CODE
                    objcDataEntity.WrkDirectory = lstrWorkingDirectory
                    lrsCSVRecordset = objMoneyCostAutoDataClass.GetCsvRecords(objcDataEntity).CsvOutput

                    If Not IsNothing(lrsCSVRecordset) Then
                        lstrProcessDateFieldName = lrsCSVRecordset.Columns(0).ColumnName.ToString()
                    Else
                        lstrErrorDetails = "Error while getting first column name from " & lstrMC_CODE & " file."
                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment():" & lstrErrorDetails)
                        lstrCommonErrorDetails = lstrErrorDetails
                        Throw New Exception(lstrErrorDetails)
                    End If
                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Fetch first row from CSV file and Check the name of First Column.")


                    ' ------------------------------------------------------------------------
                    ' Delete any record for the Process date if already exist in the Recordset
                    ' ------------------------------------------------------------------------
                    objcDataEntity.CommonSQL = "SELECT TOP 1 * FROM " & lstrMC_CODE & " WHERE " & lstrProcessDateFieldName & " = #" & CDate(ProcessDate) & "#"
                    objcDataEntity.WrkDirectory = lstrWorkingDirectory
                    lrsCSVRecordset = objMoneyCostAutoDataClass.GetCsvRecords(objcDataEntity).CsvOutput

                    If lrsCSVRecordset.Rows.Count > 0 Then
                        objcDataEntity.astrFileName = lstrWorkingDirectory & lstrMC_CODE
                        objcDataEntity.ProcessDate = lstrProcessDate
                        lbDataDeletionFlag = Convert.ToBoolean(DeleteExistingData(objcDataEntity).OutputString)

                        If lbDataDeletionFlag = False Then
                            lstrErrorDetails = "Error while deleting records from " & lstrMC_CODE & " for Process Date : " & ProcessDate
                            STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment():" & lstrErrorDetails)
                            lstrCommonErrorDetails = lstrErrorDetails
                            Throw New Exception(lstrErrorDetails)
                        End If
                    End If
                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Delete any record for the Process date if already exist in the Recordset END")


                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Load the CSV file from Working DIR into an ADO object and fetch last process date. STRAT:-")
                    'We require to fetch data from CSV as per latest date. So filter out latest EffectiveDate used in CSV.
                    objcDataEntity.CommonSQL = "SELECT TOP 1 " & lstrProcessDateFieldName & " FROM " & lstrMC_CODE & " Order by " & lstrProcessDateFieldName & " Desc"
                    objcDataEntity.WrkDirectory = lstrWorkingDirectory
                    lrsCSVRecordset = objMoneyCostAutoDataClass.GetCsvRecords(objcDataEntity).CsvOutput

                    If lrsCSVRecordset.Rows.Count > 0 Then
                        strLastProcessDateInCSV = Trim(lrsCSVRecordset.Rows(0)(0).ToString())
                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Load CSV file from working DIR into an ADO object and last process date FOUND. END:-")
                    Else
                        strLastProcessDateInCSV = ""
                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Load CSV file from working DIR into an ADO object and last process date NOT FOUND. END:-")
                    End If



                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Get Cost Types from MoneyCost Database. Start:-")
                    ' Get Data for Cost_Types
                    Dim dsCostTypes As DataSet = lobjBSMCIMoneyCostService.GetCostTypes().OutputDataSet

                    If Not IsNothing(dsCostTypes) Then
                        strCostTypes = "'"
                        For Each drCostType As DataRow In dsCostTypes.Tables("COST_TYPES").Rows()
                            strCostTypes = strCostTypes & drCostType("Cost_Type").ToString() & "','"
                        Next

                        dtCSVColNames = dsCostTypes.Tables(1)
                    Else
                        lstrErrorDetails = "Cost Types not defined in Database."
                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment():" & lstrErrorDetails)
                        lstrCommonErrorDetails = lstrErrorDetails
                        Throw New Exception(lstrErrorDetails)
                    End If

                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Get Cost Types from MoneyCost Database. End.-")


                    If Right(strCostTypes, 2) = ",'" Then strCostTypes = Mid(strCostTypes, 1, Len(strCostTypes) - 2)


                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Fetch Data from DWH. Start:-")
                    'Data from DWH
                    objcDataEntity.ProcessDate = ProcessDate
                    objcDataEntity.CostTypes = strCostTypes
                    dsDWHData = lobjcDataClass.GetIndexDataForTreasuryAssessment(objcDataEntity)
                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Fetch Data from DWH. End.")


                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Fetch Data from CSV. Start:-")
                    If (Trim(strLastProcessDateInCSV) <> "") Then
                        ' Fetch data from CSV to dataset based on [EFFECTIVE DATE] and then Filer data starting from current quarter.
                        objcDataEntity.CommonSQL = "SELECT * FROM " & lstrMC_CODE & " Where " & lstrProcessDateFieldName & "= #" & CDate(strLastProcessDateInCSV) & "#"
                        objcDataEntity.WrkDirectory = lstrWorkingDirectory
                        lrsCSVRecordset = lobjcDataClass.GetCsvRecords(objcDataEntity).CsvOutput
                    Else
                        ' [EFFECTIVE DATE] is blank so fetch blank dataset as we need column names for further processing.
                        objcDataEntity.CommonSQL = "SELECT * FROM " & lstrMC_CODE & " Where 1=2"
                    End If

                    objcDataEntity.WrkDirectory = lstrWorkingDirectory
                    lrsCSVRecordset = lobjcDataClass.GetCsvRecords(objcDataEntity).CsvOutput
                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Fetch Data from CSV. End.")


                    'Select Case Now.Month()
                    Select Case Convert.ToDateTime(ProcessDate).Month()
                        Case 1, 2, 3
                            strQuater = "1"
                        Case 4, 5, 6
                            strQuater = "4"
                        Case 7, 8, 9
                            strQuater = "7"
                        Case 10, 11, 12
                            strQuater = "10"
                    End Select

                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Filter data from DWH. Start:-")
                    UpdateDWHData(dsDWHData, strQuater, ProcessDate)
                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Filter data from DWH. End.")

                    ' compare data with DWH. If data matches then send ERROR mail. IF data do not matches then insert data in csv.
                    If Not IsNothing(lrsCSVRecordset) Then
                        ' If data for Effective Date found then continue ELSE Insert Data into CSV From dsDWHData
                        If (lrsCSVRecordset.Rows.Count > 0) Then

                            lrsCSVRecordset.DefaultView.RowFilter = "Year >= " & Convert.ToDateTime(ProcessDate).Year()

                            If (lrsCSVRecordset.DefaultView.ToTable.Rows.Count > 0) Then
                                Dim dtFilteredData As New DataTable
                                dtFilteredData = lrsCSVRecordset.DefaultView(0).DataView.ToTable()

                                For Each dr As DataRow In dtFilteredData.Select("Year <= " & Convert.ToDateTime(ProcessDate).Year() & " and Quarter< " & strQuater)
                                    dtFilteredData.Rows.RemoveAt(0)
                                Next

                                ' Compare CSV data with [Resultant Dataset]
                                ' If in comparison some difference is found then insert data with new Calender date ELSE do nothing.

                                If (dtFilteredData.Rows.Count > 0) Then
                                    If dtFilteredData.Columns.Contains("EffectiveDate") Then
                                        'If column "EffectiveDate" exists then Remove column "EffectiveDate" for data comparison
                                        dtFilteredData.Columns.RemoveAt(0)
                                    End If

                                    If (dsDWHData.Tables(0).Rows.Count <> dtFilteredData.Rows.Count) Then
                                        'Row count of data from Dataware house and CSV are not matching.
                                        bDatasetComparison = False
                                    Else
                                        Dim intColCounter As Int32
                                        Dim intMainTable As Int32

                                        For intMainTable = 0 To dsDWHData.Tables(0).Rows.Count - 1
                                            drDwhData = dsDWHData.Tables(0).Rows(intMainTable)
                                            drFilteredData = dtFilteredData.Rows(intMainTable)

                                            If bDatasetComparison Then
                                                For intColCounter = 0 To dsDWHData.Tables(0).Columns.Count - 1
                                                    If String.Compare(drDwhData(intColCounter).ToString(), drFilteredData(intColCounter).ToString(), True) <> 0 Then
                                                        bDatasetComparison = False
                                                        Exit For
                                                    End If
                                                Next
                                            Else
                                                Exit For
                                            End If
                                        Next
                                    End If



                                    'Check if data from csv and data ware house is matching 
                                    If (bDatasetComparison) Then
                                        'if Matching then send out a mail stating "Data is matching so data available to write"
                                        '' Data inserted Successfully, send mail to inform business contacts.
                                        '' No data to update CSV file -- As data from CSV and DWH is matching, send mail to inform business contacts.
                                        STLogger.Error("No data available to update Treasury Assessment file on " & ProcessDate & " as Data from DWH and CSV is matching." + Environment.NewLine + "Process not completed today for Treasury Assessment file.")
                                        NotificationMailFlag = True
                                        strMailBody = "No data available to update Treasury Assessment file on " & ProcessDate & " as Data from DWH and CSV is matching."
                                        SendMailNotification(lstrClarifyQName, lstrBusinessContact, strMailBody, False, True, lstrMC_CODE, lstrProcessDate, True)
                                        strLogMessage = "No data available to update Treasury Assessment file on " & ProcessDate & " as Data from DWH and CSV is matching."

                                    Else
                                        ' Insert new data at end of csv file using data from DWH and then send a success mail "Data inserted successfully in csv".                                        
                                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Inserting data in CSV. Start:-")
                                        InsertDataFromDWH(dsDWHData, ProcessDate, dtCSVColNames, strLogMessage, intStatus)
                                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Inserting data in CSV. End.")
                                    End If


                                Else
                                    ' No data found in dtFilteredData.Rows.Count then there is not data to compare.
                                    ' Send out a mail stating "There is no data to process."
                                    '' No data to update CSV file, send mail to inform business contacts.
                                    STLogger.Error("No data available to update Treasury Assessment file." + Environment.NewLine + "Process not completed today for Treasury Assessment file on process date " & ProcessDate & "")
                                    NotificationMailFlag = True
                                    strMailBody = "No data available to update Treasury Assessment file."
                                    SendMailNotification(lstrClarifyQName, lstrBusinessContact, strMailBody, False, True, lstrMC_CODE, lstrProcessDate, True)
                                    strLogMessage = "No Treasury Data to Update."
                                End If
                            Else
                                'No related data found in CSV.
                                ' so check if there is data in Datawarehouse. If yes then insert Datawarehouse data in CSV.
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Inserting data in CSV. Start:-")
                                InsertDataFromDWH(dsDWHData, ProcessDate, dtCSVColNames, strLogMessage, intStatus)
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Inserting data in CSV. End.")

                            End If
                        Else
                            ' If data for Effective Date found then continue ELSE Insert Data into CSV From dsDWHData
                            ' i.e. If (lrsCSVRecordset.Rows.Count = 0) Then                           
                            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Inserting data in CSV. Start:-")
                            InsertDataFromDWH(dsDWHData, ProcessDate, dtCSVColNames, strLogMessage, intStatus)
                            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Inserting data in CSV. End.")
                        End If
                    Else
                        'No data found in CSV, check if data exist in DWH then Insert data.
                        ' Insert new data at end of csv file using data from DWH and then send a success mail "Data inserted successfully in csv".                        
                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Inserting data in CSV. Start:-")
                        InsertDataFromDWH(dsDWHData, ProcessDate, dtCSVColNames, strLogMessage, intStatus)
                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Inserting data in CSV. End.")
                    End If

                Else
                    '' FTP is not working. Drop a Error mail to inform business contacts.
                    STLogger.Error("Unable to map or connect ftp location: " + lstrFTP_Location + "," + Environment.NewLine + "Process not completed today for Treasury Assessment file on process date " & ProcessDate & " ")
                    NotificationMailFlag = True
                    strMailBody = "Unable to map or connect ftp location: " + lstrFTP_Location + "," + Environment.NewLine + "Process not completed today for Treasury Assessment file on process date " & ProcessDate & " "
                    SendMailNotification(lstrClarifyQName, lstrBusinessContact, strMailBody, False, True, lstrMC_CODE, lstrProcessDate, True)
                    strLogMessage = "Unable to map or connect ftp location: " + lstrFTP_Location + "," + Environment.NewLine + "Process not completed today for Treasury Assessment file on process date " & ProcessDate & " "
                End If

            Else
                '' lbServiceRunFlag is FALSE. Process Date is not matching with date specified in MC_File Table.
                '' Send a mail that PROCESS is not SCHDEULED for This date.
                STLogger.Error("No entry for Process run in Database found. " + Environment.NewLine + "Process not completed today for Treasury Assessment on process date " & ProcessDate & "")
                NotificationMailFlag = True
                strMailBody = "No entry for Process run in Database found. " + Environment.NewLine + "Process not completed today for Treasury Assessment on process date " & ProcessDate & ""
                SendMailNotification(lstrClarifyQName, lstrBusinessContact, strMailBody, False, True, lstrMC_CODE, lstrProcessDate, True)
                strLogMessage = "No entry for Process run in Database found. " + Environment.NewLine + "Process not completed today for Treasury Assessment on process date " & ProcessDate & ""
            End If


            strProcessEndDate = Now

            '' Update MC_File and MC_File_Log table with Last_Schdule_Process_Date and Process Status
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Update MC_Log and MC_file tables. Start:-")
            UpdateTables(ProcessDate, strProcessStartDate, strProcessEndDate, intStatus, strLogMessage)
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment(): Update MC_Log and MC_file tables. End.")


        Catch ex As Exception

            'Send an Error mail            
            strMailBody = IIf(lstrCommonErrorDetails = "", llErrNbr & " : " & lstrErrDesc & " Source: " & lstrErrSrc, lstrCommonErrorDetails)
            SendMailNotification(lstrClarifyQName, lstrBusinessContact, strMailBody, False, True, lstrMC_CODE, lstrProcessDate, True)

        Finally
            'If File.Exists(lstrWorkingDirectory & lstrMC_CODE) Then
            '    Kill(lstrWorkingDirectory & lstrMC_CODE)
            'End If

            If File.Exists(lstrWorkingDirectory & "schema.ini") Then
                Kill(lstrWorkingDirectory & "schema.ini")
            End If

            If Not IsNothing(lobjBSMCIMoneyCostService) Then
                lobjBSMCIMoneyCostService = Nothing
            End If

            If Not IsNothing(lobjcDataClass) Then
                lobjcDataClass = Nothing
            End If

            If Not IsNothing(lobjcMoneyCostUISvc) Then
                lobjcMoneyCostUISvc = Nothing
            End If

            If Not IsNothing(lrsCSVRecordset) Then
                lrsCSVRecordset = Nothing
            End If

            If Not IsNothing(objFileEntity) Then
                objFileEntity = Nothing
            End If

            If Not IsNothing(objMoneyCostAutoDataClass) Then
                objMoneyCostAutoDataClass = Nothing
            End If

            If Not IsNothing(lobjMCFileNode) Then
                lobjMCFileNode = Nothing
            End If


            If Not IsNothing(lobjIndexRateNodeList) Then
                lobjIndexRateNodeList = Nothing
            End If

            If Not IsNothing(lobjIndexRateNode) Then
                lobjIndexRateNode = Nothing
            End If

            If Not IsNothing(lobjIndexRateDOM) Then
                lobjIndexRateDOM = Nothing
            End If

            If Not IsNothing(lobjIndexDataNodeList) Then
                lobjIndexDataNodeList = Nothing
            End If

            If Not IsNothing(lobjIndexDataDOMXml) Then
                lobjIndexDataDOMXml = Nothing
            End If

            If Not IsNothing(lobjIndexDataNode) Then
                lobjIndexDataNode = Nothing
            End If

            If Not IsNothing(lobjCSVInsertNodeList) Then
                lobjCSVInsertNodeList = Nothing
            End If

            If Not IsNothing(lobjCSVInsertNode) Then
                lobjCSVInsertNode = Nothing
            End If

            If Not IsNothing(lobjCSVInsertDOMXml) Then
                lobjCSVInsertDOMXml = Nothing
            End If

            If Not IsNothing(lobjAllMCFileDOM) Then
                lobjAllMCFileDOM = Nothing
            End If

            If Not IsNothing(lobjMCFileNodeList) Then
                lobjMCFileNodeList = Nothing
            End If

            If Not IsNothing(objFileEntity) Then
                objFileEntity = Nothing
            End If

            If Not IsNothing(objxmlErrEntity) Then
                objxmlErrEntity = Nothing
            End If
        End Try
    End Sub

    Public Sub UpdateDWHData(ByRef dsDWHData As DataSet, ByVal strQuarter As String, ByVal ProcessDate As String)
        Try
            If Not IsNothing(dsDWHData) Then
                If (dsDWHData.Tables(0).Rows.Count > 0) Then

                    dsDWHData.Tables(0).DefaultView.RowFilter = "Year >= " & Convert.ToDateTime(ProcessDate).Year()

                    If (dsDWHData.Tables(0).DefaultView.ToTable.Rows.Count > 0) Then
                        Dim dtDWHdata As New DataTable, bRemoveData As Boolean = False
                        dtDWHdata = dsDWHData.Tables(0).DefaultView(0).DataView.ToTable()
                        For Each dr As DataRow In dtDWHdata.Select("Year <= " & Convert.ToDateTime(ProcessDate).Year() & " and Quarter< " & strQuarter)
                            dtDWHdata.Rows.RemoveAt(0)
                            bRemoveData = True
                        Next

                        If (bRemoveData) Then
                            dsDWHData = Nothing
                            dsDWHData = New DataSet
                            dsDWHData.Tables.Add(dtDWHdata)
                        End If
                    End If
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Sub InsertDataFromDWH(ByVal dsDWHData As DataSet, ByVal ProcessDate As String, ByVal dtCSVColNames As DataTable, ByRef strLogMessage As String, ByRef intStatus As Int32)
        Dim strMailBody As String = ""

        Try
            If Not IsNothing(dsDWHData) Then
                If (dsDWHData.Tables(0).Rows.Count > 0) Then
                    InsertData(dsDWHData, ProcessDate, lrsCSVRecordset, dtCSVColNames)

                    'Add these lines to encrypt the files #Sumit
                    If EncrptFileFlag = "True" Then
                        If EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE).ToString.ToUpper = "DECRYPTED" Then
                            EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE)
                        End If
                    End If
                    'Copy Working File back to the FTP Location.
                    STLogger.Debug("Copy Working File back to the FTP Location.")
                    objFileEntity.Source = lstrWorkingDirectory & lstrMC_CODE
                    objFileEntity.Destination = lstrFTP_Location & lstrFTP_Directory & lstrMC_CODE
                    lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString
                    STLogger.Debug("Copied Working File back to the FTP Location. " & objFileEntity.Source & "-->" & objFileEntity.Destination)

                    'if any error occurred, raise clarify case
                    If lstrCopyFileResponse <> "" Then
                        lstrErrorDetails = lstrErrorDetails & "Error while updating " & lstrMC_CODE & ".csv file as on the FTP location : " & lstrFTP_Location & " : " & lstrCopyFileResponse
                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ProcessTreasuryAssessment():" & lstrErrorDetails)
                        lstrCommonErrorDetails = lstrErrorDetails
                        Throw New Exception(lstrErrorDetails)
                    End If

                    '' Data inserted Successfully, send mail to inform business contacts.
                    NotificationMailFlag = True
                    strMailBody = "Data successfully copied in Treasury Assessment file for " & ProcessDate & "."
                    SendMailNotification(lstrClarifyQName, lstrBusinessContact, strMailBody, False, True, lstrMC_CODE, lstrProcessDate, False)
                    strLogMessage = "Data successfully copied in Treasury Assessment file for " & ProcessDate & "."
                    intStatus = 1

                Else
                    '' No data to update CSV file, send mail to inform business contacts.
                    STLogger.Error("No data available in data ware house for process date " & ProcessDate & " to update Treasury Assessment file." + Environment.NewLine + "Process not completed today for Treasury Assessment file.")
                    NotificationMailFlag = True
                    strMailBody = "No data available in data ware house for process date " & ProcessDate & " to update Treasury Assessment file." + Environment.NewLine + "Process not completed today for Treasury Assessment file."
                    SendMailNotification(lstrClarifyQName, lstrBusinessContact, strMailBody, False, True, lstrMC_CODE, lstrProcessDate, True)
                    strLogMessage = "No data available in data ware house for process date " & ProcessDate & " to update Treasury Assessment file." + Environment.NewLine + "Process not completed today for Treasury Assessment file."
                    intStatus = 0
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub SendMailNotification(ByVal strQueueName As String, ByVal strBusinessContact As String, ByVal strBody As String, _
                 ByVal bCutTicket As Boolean, ByVal bSendNotification As Boolean, ByVal strMCCode As String, ByVal strProcessDate As String, ByVal bErrorMail As Boolean)

        Try
            If objFileEntity Is Nothing Then
                objFileEntity = New FileInfoEntity
            End If

            objFileEntity.QueueName = strQueueName
            objFileEntity.BusinessContact = strBusinessContact
            objFileEntity.Body = strBody
            objFileEntity.CutTicket = bCutTicket
            objFileEntity.SendNotification = bSendNotification
            objFileEntity.MCCode = strMCCode
            objFileEntity.ProcessDates = strProcessDate

            If (bErrorMail) Then
                lstrSendErrNotiResult = SendErrNotification(objFileEntity).OutputString
            Else
                lstrSendErrNotiResult = SendNotification(objFileEntity).OutputString
            End If

        Catch ex As Exception
        End Try
    End Sub

    Public Sub InsertData(ByVal dsDWHData As DataSet, ByVal ProcessDate As String, ByVal dtCSVData As DataTable, ByVal dtCSVColNamesFromDB As DataTable)
        Dim drDwhData As DataRow
        Dim strQueryBlankValue As String = ""
        Dim objMoneyCostAutoDataClass As New MoneyCostAutoDataClass

        Try
            For Each drDwhData In dsDWHData.Tables(0).Rows()

                lstrCSVInsertSQL = "Insert into " & lstrMC_CODE & " Values('" & ProcessDate & "','" & drDwhData("Year").ToString() & "','" & drDwhData("Quarter").ToString() & "',"


                For i As Int32 = 2 To dtCSVData.Columns.Count - 2
                    dtCSVColNamesFromDB.DefaultView.RowFilter = "CURRENCY='" & dtCSVData.Columns(i + 1).ColumnName.ToString() & "' and DISPLAY=1"

                    If (dtCSVColNamesFromDB.DefaultView.ToTable.Rows.Count > 0) Then
                        If (i = 2) Then
                            lstrCSVInsertSQL = lstrCSVInsertSQL & "'" & drDwhData("UTH").ToString() & "',"
                        Else
                            lstrCSVInsertSQL = lstrCSVInsertSQL & "'" & drDwhData("USD").ToString() & "',"
                        End If
                    Else
                        lstrCSVInsertSQL = lstrCSVInsertSQL & "'0.0000',"
                    End If
                Next i

                lstrCSVInsertSQL = Mid(lstrCSVInsertSQL, 1, Len(lstrCSVInsertSQL) - 1)
                lstrCSVInsertSQL = lstrCSVInsertSQL & ");"

                If (lstrCSVInsertSQL <> "") Then
                    'Save row back to the working file as TreasuryAssessment.csv.
                    objcDataEntity.CommonSQL = lstrCSVInsertSQL
                    objcDataEntity.WrkDirectory = lstrWorkingDirectory
                    objMoneyCostAutoDataClass.UpdateCsvRecords(objcDataEntity)
                End If
            Next

        Catch ex As Exception

        Finally
            If Not IsNothing(objMoneyCostAutoDataClass) Then
                objMoneyCostAutoDataClass = Nothing
            End If
        End Try
    End Sub

    Public Sub UpdateTables(ByVal strProcessDate As String, ByVal strProcessStartDate As String, ByVal strProcessEndDate As String, ByVal intStatus As Int32, ByVal strLogMessage As String)
        Dim lobjMoneyCostDataClass As New cDataClass

        Try
            '' Update MC_File and MC_File_Log table with Last_Schdule_Process_Date and Process Status
            '' Update MC_FILE
            lstrCSVInsertSQL = "Update MC_File Set Last_Schedule_Process_Date = '" & strProcessDate.ToString() & "' Where SQ_MC_ID = " & liSQ_MC_ID
            objcDataEntity.CommonSQL = lstrCSVInsertSQL
            lobjMoneyCostDataClass.UpdateMCFile(objcDataEntity)

            '' Update MC_Logs
            lstrCSVInsertSQL = "Insert into MC_Logs Values(" & liSQ_MC_ID & ",'" & strProcessStartDate.ToString() & "','" & strProcessEndDate.ToString() & "' ," & intStatus & ",'" & strLogMessage & "')"
            objcDataEntity.CommonSQL = lstrCSVInsertSQL
            lobjMoneyCostDataClass.UpdateMCFile(objcDataEntity)

        Catch ex As Exception

        Finally
            If Not IsNothing(lobjMoneyCostDataClass) Then
                lobjMoneyCostDataClass = Nothing
            End If
        End Try
    End Sub


    <AutoComplete()> _
    Public Function DeleteExistingData(ByVal lobjEntity As cDataEntity) As cDataEntity

        Dim lTSFileStreamHandle As StreamReader = Nothing
        Dim lTSFileWriteHandle As StreamWriter = Nothing
        Dim lvarData() As Object
        Dim lstrSearch As String
        Dim lstrText As String
        Dim llLine As Long
        Dim liCount As Long
        SetLog4Net()
        Try
            lstrSearch = lobjEntity.ProcessDate.ToString("MM/dd/yyyy")
            lTSFileStreamHandle = File.OpenText(lobjEntity.astrFileName)
            llLine = 0

            lstrText = lTSFileStreamHandle.ReadLine
            ReDim Preserve lvarData(llLine)
            lvarData(llLine) = lstrText
            llLine = llLine + 1

            If lstrSearch <> "" Then
                While lTSFileStreamHandle.EndOfStream = False
                    lstrText = lTSFileStreamHandle.ReadLine

                    If InStr(1, lstrText, lstrSearch, vbBinaryCompare) = 0 Then
                        ReDim Preserve lvarData(llLine)

                        lvarData(llLine) = lstrText
                        llLine = llLine + 1
                    End If
                End While
            End If

            lTSFileStreamHandle.Close()

            If lvarData.Length = 0 Then
                lobjEntity.OutputString = "True"
                Return lobjEntity
                Exit Function
            End If

            lTSFileWriteHandle = New StreamWriter(lobjEntity.astrFileName)

            For liCount = 0 To UBound(lvarData)
                lTSFileWriteHandle.WriteLine(lvarData(liCount))
            Next
            lTSFileWriteHandle.Close()
            lobjEntity.OutputString = "True"
            Return lobjEntity
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            lobjEntity.OutputString = "False"
            Return lobjEntity
            Throw
        Finally
            lTSFileStreamHandle.Close()
            lTSFileWriteHandle.Close()
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
            If Not IsNothing(lTSFileStreamHandle) Then
                lTSFileStreamHandle = Nothing
            End If
            If Not IsNothing(lTSFileWriteHandle) Then
                lTSFileWriteHandle = Nothing
            End If
        End Try

    End Function
    
#End Region

#Region "Executive Service Flow Block"
    <AutoComplete()> _
   Public Sub ExecuteServiceFlow() Implements IMoneyCostAutoSvc.ExecuteServiceFlow
        Dim lobjBSMCIMoneyCostService As New BSMoneyCostBL.cMoneyCostUISvc
        Dim objMoneyCostAutoDataClass As New MoneyCostAutoDataClass
        Dim lstrRetMsg As String
        Dim ProcessMoneyCostFiles As String = ""
        Dim ProcessMoneyCostFiles_CurrencyCode As String = ""
        Dim SendErrMailForCurrencyFlag As Boolean
        SetLog4Net()
        Try
            SetConfigSettings()
            NotificationMailFlag = False
            lstrMethodName = "ExecuteServiceFlow"
            STLogger.Debug(cMODULE_NAME & lstrMethodName & " In " & lstrMethodName & "() method ")
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():  Checking necessary registry keys")

            'Setting Configuration Values To Variables
            SetConfigValue()

            If lstrCommonErrorDetails = "" Then
                If Right(lstrWorkingDirectory, 1) <> "\" Then lstrWorkingDirectory = lstrWorkingDirectory & "\"
                If Right(lstrBackup_Location, 1) <> "\" Then lstrBackup_Location = lstrBackup_Location & "\"
                If Right(lstrNetwork_Location, 1) <> "\" Then lstrNetwork_Location = lstrNetwork_Location & "\"
                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get List of MC ALL Files Start")
                ' -------------------------------------------
                ' Get List of MC ALL Files
                ' -------------------------------------------                
                lstrGetMCFilesRespXml = lobjBSMCIMoneyCostService.GetAllMCFilesForMCRun().OutputString
                'validate, if response xml is well-formed
                Try
                    lobjAllMCFileDOM.LoadXml(lstrGetMCFilesRespXml)
                Catch ex As Exception
                    lstrCommonErrorDetails = "Error No. : " & Err.Number & vbCrLf & "Error Description : " & _
                                        "Error loading response XML from BSMOneyCost.IMoneyCostService.GetAllMCFilesForMCRun(). " & _
                                        Err.Description
                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get List of MC ALL Files Error" & lstrCommonErrorDetails)
                    Throw
                End Try
                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get List of MC ALL Files Complete")
                lobjMCFileNodeList = lobjAllMCFileDOM.SelectNodes("/MC_FILE_RESPONSE/MC_FILESet/MC_FILE")

                If lobjMCFileNodeList.Count <= 0 Then
                    lstrCommonErrorDetails = "No Money Cost file found to process."
                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): No Money Cost file found to process.")
                    Throw New Exception(lstrErrorDetails)
                End If

                lstrUpdateMCLogReqXml = "<UPDATE_MC_FILE_LOGS_REQUEST><MC_FILE_LOGSet>"
                lstrUpdateMCFileReqXml = lstrUpdateMCFileReqXml & "<UPDATE_MC_FILE_REQUEST><MC_FILESet>"
                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Iterate through the Recordset with all MC Files list START:-")
                ' ----------------------------------------------------
                ' Iterate through the Recordset with all MC Files list
                ' ----------------------------------------------------

                For liCounter1 = 0 To lobjMCFileNodeList.Count - 1
                    'If Any Exception in MC Files [Getting  Next MC File ]
                    Try
                        lstrMCFileStartExec = Now()
                        SendErrMailForCurrencyFlag = False
                        ' ----------------------------------------------------
                        ' Iterate through the Recordset with all MC Files list
                        ' ----------------------------------------------------                      
                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Iterate through the Recordset with all MC Files list")
                        ' --------------------------------------------------------------
                        ' For each MC File record in the recordset process the following
                        ' --------------------------------------------------------------                        
                        Call SetXmlValue(lstrFTP_DirectoryForNewDateFormat, lstrFTP_Directory)
                        If lstrSTART_TIME <> "" And lstrEND_TIME <> "" Then
                            'if MC file process time is out of specified time, skip to next MC file
                            If DateDiff("s", Format(Now.TimeOfDay, "hh:mm:ss"), lstrSTART_TIME) < 0 Or DateDiff("s", Format(Now.TimeOfDay, "hh:mm:ss"), lstrEND_TIME) > 0 Then

                                STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): if MC file process time is out of specified time, skip to next MC file")
                                lstrErrorDetails = "MC file process time is out of specified time, skip to next MC file"
                                Throw New Exception(lstrErrorDetails)
                            End If
                        End If
                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Process date  START")
                        ''''Set Process Flag lbServiceRunFlag True Or False
                        Call SetDateProcessFlag()
                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Process date  END ; Process Flag=" & lbServiceRunFlag.ToString)
                        ProcessMoneyCostFiles_CurrencyCode = ""
                        If lbServiceRunFlag = True Then
                            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Process date = FTP Location Mapped start")
                            'ProcessMoneyCostFiles_CurrencyCode = lstrMC_CODE       '#### Added By Atul On 18-09-2011
                            'FTP Location Mapped start
                            Dim objFtp As New FTPEntity
                            objFtp.FTPLocation = lstrFTP_Location
                            objFtp.FTPUser = lstrFTP_User
                            objFtp.FTPPassword = lstrFTP_Password
                            If Convert.ToBoolean(IsFTPLocationMapped(objFtp).OutputString) Then
                                NotificationMailFlag = False
                                ProcessMoneyCostFiles_CurrencyCode = lstrMC_CODE        '#### Added By Atul On 18-09-2011
                                If Right(lstrFTP_Location, 1) <> "\" Then lstrFTP_Location = lstrFTP_Location & "\"
                                If Right(lstrFTP_Directory, 1) <> "\" Then lstrFTP_Directory = lstrFTP_Directory & "\"
                                ' -----------------------------------------------------------------
                                ' Create a Backup Copy of existing MCXXX File as BKMMDDYY.MCUSD.csv
                                ' on Backup location, from FTP location
                                ' -----------------------------------------------------------------
                                objcDataEntity.ProcessDate = Now.Date
                                objFileEntity.Source = lstrFTP_Location & lstrFTP_Directory & lstrMC_CODE & ".csv"
                                objFileEntity.Destination = lstrBackup_Location & "BK" & GetDateFormat(objcDataEntity).OutputString & "." & lstrMC_CODE & ".csv"
                                lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString

                                '                    If lobjcFTP.Connected Then lobjcFTP.Disconnect
                                '                    lobjcFTP.Connect lstrFTP_Location, lstrFTP_User, lstrFTP_Password
                                '                    lobjcFTP.Directory = lstrFTP_Directory
                                '                    lstrCopyFileResponse = lobjcFTP.GetFile(lstrMC_CODE & ".csv", lstrBackup_Location & "BK" & GetDateFormat(Date) & "." & lstrMC_CODE & ".csv")
                                '                    lobjcFTP.Disconnect

                                'if any error occurred, raise clarify case
                                If lstrCopyFileResponse <> "" Then
                                    objcDataEntity.ProcessDate = Now.Date
                                    lstrErrorDetails = "Error while creating backup copy of existing " & lstrMC_CODE & ".csv file as " & _
                                                        lstrBackup_Location & "BK" & GetDateFormat(objcDataEntity).OutputString & "." & lstrMC_CODE & ".csv on " & _
                                                        lstrBackup_Location & " : " & lstrCopyFileResponse

                                    'If Not lrsCSVRecordset. Then lrsCSVRecordset.Close()
                                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                    Throw New Exception(lstrErrorDetails)
                                End If

                                ' Invoke Common Method.CopyFiles() to copy latest MCXXX file from Backup Location to working Direcotry.
                                ' Working directory value should be defined in Registry
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get Latest MCXXX file from Network Path Start")
                                ' ---------------------------------------
                                ' Get Latest MCXXX file from Network Path
                                ' ---------------------------------------
                                ' Invoke Common Method.CopyFiles() to copy latest MCXXX file from Backup Location to working Direcotry.
                                ' Working directory value should be defined in Registry 
                                objcDataEntity.ProcessDate = Format(Convert.ToDateTime(Now.Date), "MM/dd/yyyy")

                                objFileEntity.Source = lstrBackup_Location & "BK" & GetDateFormat(objcDataEntity).OutputString & "." & lstrMC_CODE & ".csv"
                                objFileEntity.Destination = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString

                                'if error occurred, raise clarify case
                                If lstrCopyFileResponse <> "" Then
                                    lstrErrorDetails = "Error while getting latest " & lstrMC_CODE & ".csv to working " & "directory " & _
                                                        lstrWorkingDirectory & " : " & lstrCopyFileResponse
                                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                    Throw New Exception(lstrErrorDetails)
                                End If

                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get Latest MCXXX file from Network Path Start END")
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Copy MCXXX.ini file as schema.ini file in working direcotry Start")
                                ' -----------------------------------------------------------
                                ' Copy MCXXX.ini file as schema.ini file in working direcotry
                                ' -----------------------------------------------------------
                                ' Invoke Common Method.CopyFiles()
                                objFileEntity.Source = lstrWorkingDirectory & lstrMC_CODE & ".ini"
                                objFileEntity.Destination = lstrWorkingDirectory & "schema.ini"
                                lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString

                                'if any error occurred, raise clarify case
                                If lstrCopyFileResponse <> "" Then
                                    'raise clarify case
                                    'schema file not present for MCXXX.csv file.
                                    'Send notification to business MCXXX file not updated for Process Date.
                                    'Details: Schema file not defined for MCXXX.csv

                                    lstrErrorDetails = "Schema file not present for " & lstrMC_CODE & ".csv" & vbCrLf & _
                                                        "Error while defining schema file : " & lstrCopyFileResponse
                                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                    Throw New Exception(lstrErrorDetails)
                                End If
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Copy MCXXX.ini file as schema.ini file in working direcotry END")
                                '*****************************************************************
                                'change to incorporate Encryption and Decryption of CSV files.
                                If EncrptFileFlag = "True" Then
                                    If EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv").ToString.ToUpper = "ENCRYPTED" Then
                                        lstrRetMsg = EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv")
                                    End If
                                End If

                                '*****************************************************************
                                'get first data row from MC file, to get date field name
                                objcDataEntity.CommonSQL = "SELECT TOP 1 * FROM " & lstrMC_CODE & ".csv"
                                objcDataEntity.WrkDirectory = lstrWorkingDirectory
                                lrsCSVRecordset = objMoneyCostAutoDataClass.GetCsvRecords(objcDataEntity).CsvOutput

                                If lrsCSVRecordset.Rows.Count > 0 Then lstrProcessDateFieldName = Trim(lrsCSVRecordset.Columns(0).ColumnName)

                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Delete any record for the Process date if already exist in the Recordset START:-")

                                ' ------------------------------------------------------------------------
                                ' Delete any record for the Process date if already exist in the Recordset
                                ' ------------------------------------------------------------------------
                                'If Not lrsCSVRecordset.IsClosed Then lrsCSVRecordset.Close()
                                objcDataEntity.CommonSQL = "SELECT TOP 1 * FROM " & lstrMC_CODE & ".csv WHERE " & lstrProcessDateFieldName & " = #" & CDate(lstrProcessDate) & "#"
                                objcDataEntity.WrkDirectory = lstrWorkingDirectory
                                lrsCSVRecordset = objMoneyCostAutoDataClass.GetCsvRecords(objcDataEntity).CsvOutput

                                If lrsCSVRecordset.Rows.Count > 0 Then
                                    'If Not lrsCSVRecordset.IsClosed Then lrsCSVRecordset.Close()
                                    objcDataEntity.astrFileName = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                    objcDataEntity.ProcessDate = lstrProcessDate
                                    lbDataDeletionFlag = Convert.ToBoolean(UpdateFile(objcDataEntity).OutputString)

                                    If lbDataDeletionFlag = False Then
                                        lstrErrorDetails = "Error while deleting records from " & lstrMC_CODE & ".csv for Process Date : " & lstrProcessDate
                                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                        Throw New Exception(lstrErrorDetails)
                                    End If
                                End If
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Delete any record for the Process date if already exist in the Recordset END")
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Load the CSV file from Working DIR into an ADO object. STRAT:-")
                                ' ------------------------------------------------------
                                ' Load the CSV file from Working DIR into an ADO object.
                                ' ------------------------------------------------------
                                'If Not lrsCSVRecordset.IsClosed Then lrsCSVRecordset.Close()

                                lstrSqlQry = "SELECT * FROM " & lstrMC_CODE & ".csv Order By " & lstrProcessDateFieldName
                                objcDataEntity.CommonSQL = lstrSqlQry
                                objcDataEntity.WrkDirectory = lstrWorkingDirectory
                                lrsCSVRecordset = objMoneyCostAutoDataClass.GetCsvRecords(objcDataEntity).CsvOutput

                                ' ----------------------------------------------------
                                ' Check if Process date = 01/01 and Rs.Recordcount > 0
                                ' ----------------------------------------------------
                                If DatePart("d", lstrProcessDate) = 1 And DatePart("m", lstrProcessDate) = 1 And lrsCSVRecordset.Rows.Count > 0 Then
                                    '*****************************************************************
                                    'change to incorporate Encryption and Decryption of CSV files.
                                    If EncrptFileFlag = "True" Then
                                        If EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv").ToString.ToUpper = "DECRYPTED" Then
                                            lstrRetMsg = EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv")
                                        End If
                                    End If
                                    '*****************************************************************
                                    'Yes
                                    ' -------------------------------------------------------
                                    ' Save MCXXX.csv file to FTP location as MCXXX & YYYY.csv
                                    ' -------------------------------------------------------
                                    'If Not lrsCSVRecordset.IsClosed Then lrsCSVRecordset.Close()
                                    objFileEntity.Source = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                    objFileEntity.Destination = lstrFTP_Location & lstrFTP_Directory & lstrMC_CODE & CStr(CInt(DatePart("yyyy", lstrProcessDate)) - 1) & ".csv"
                                    lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString

                                    '                            If lobjcFTP.Connected Then lobjcFTP.Disconnect
                                    '                            lobjcFTP.Connect lstrFTP_Location, lstrFTP_User, lstrFTP_Password
                                    '                            lobjcFTP.Directory = lstrFTP_Directory
                                    '                            lstrCopyFileResponse = lobjcFTP.PutFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrMC_CODE & CStr(CInt(DatePart("yyyy", lstrProcessDate)) - 1) & ".csv")
                                    '                            lobjcFTP.Disconnect

                                    'if any error occurred, raise clarify case
                                    If lstrCopyFileResponse <> "" Then
                                        lstrErrorDetails = "Error while saving " & lstrMC_CODE & ".csv file as " & lstrMC_CODE & _
                                                            CStr(CInt(DatePart("yyyy", lstrProcessDate)) - 1) & ".csv on FTP location : " & lstrCopyFileResponse
                                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                        Throw New Exception(lstrErrorDetails)
                                    End If
                                    '*****************************************************************
                                    'change to incorporate Encryption and Decryption of CSV files.
                                    If EncrptFileFlag = "True" Then
                                        If EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv").ToString.ToUpper = "ENCRYPTED" Then
                                            lstrRetMsg = EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv")
                                        End If
                                    End If
                                    '*****************************************************************
                                    ' -----------------------------------
                                    ' Delete all the data from ADO object
                                    ' -----------------------------------
                                    objcDataEntity.astrFileName = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                    objcDataEntity.ProcessDate = Nothing
                                    lbDataDeletionFlag = Convert.ToBoolean(UpdateFile(objcDataEntity).OutputString)

                                    If lbDataDeletionFlag = False Then
                                        lstrErrorDetails = "Error creating blank " & lstrMC_CODE & ".csv for current year"
                                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                        Throw New Exception(lstrErrorDetails)
                                    End If
                                    'End
                                End If

                                'If Not lrsCSVRecordset.IsClosed Then lrsCSVRecordset.Close()
                                objcDataEntity.CommonSQL = "SELECT TOP 1 * FROM " & lstrMC_CODE & ".csv WHERE " & lstrProcessDateFieldName & " = #" & CDate(lstrProcessDate) & "#"
                                objcDataEntity.WrkDirectory = lstrWorkingDirectory
                                lrsCSVRecordset = objMoneyCostAutoDataClass.GetCsvRecords(objcDataEntity).CsvOutput
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get Index Rate information for the MC File being processed Start:-")

                                ' ----------------------------------------------------------
                                ' Get Index Rate information for the MC File being processed
                                ' ----------------------------------------------------------
                                ' Invoke BSMoneyCost.GetIndexRates(SQ_MC_ID, Process Date)
                                lstrIndexRateReqXml = "<INDEX_RATE_REQUEST>" & _
                                                            "<SQ_MC_ID>" & liSQ_MC_ID & "</SQ_MC_ID>" & _
                                                            "<PROCESS_DATE>" & lstrProcessDate & "</PROCESS_DATE>" & _
                                                        "</INDEX_RATE_REQUEST>"
                                objcDataEntity.OutputString = lstrIndexRateReqXml
                                lstrIndexRateRespXml = lobjBSMCIMoneyCostService.GetIndexRates(objcDataEntity).OutputString

                                'validate, if Index Rate Response XML is well-formed
                                Try
                                    lobjIndexRateDOM.LoadXml(lstrIndexRateRespXml)
                                Catch ex As Exception
                                    'if error, raise clarify case
                                    lstrErrorDetails = "Error while loading Response XML from lobjBSMCIMoneyCostService.GetIndexRates " & _
                                                        "method." & vbCrLf & "Error no. : " & Err.Number & _
                                                        vbCrLf & "Error description : " & Err.Description
                                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get Index Rate information for the MC File being processed Start:-")
                                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                    Throw
                                End Try
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Get Index Rate information for the MC File being processed End:-")
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Check whether no. of columns in CSV and Money Cost DB match Start:-")
                                '-------------------------------------------------------------
                                ' Check whether no. of columns in CSV and Money Cost DB match.
                                ' if not, raise error       
                                '-------------------------------------------------------------
                                lobjIndexRateNodeList = lobjIndexRateDOM.SelectNodes("/INDEX_RATE_RESPONSE/INDEX_RATESet/INDEX_RATE")

                                If Val(lobjIndexRateNodeList.Count) <> lrsCSVRecordset.Columns.Count - 1 Then
                                    'get list of all Index Codes and Index Terms, available in Money Cost DB
                                    For liCounter2 = 0 To lobjIndexRateNodeList.Count - 1
                                        lobjIndexRateNode = lobjIndexRateNodeList.Item(liCounter2)

                                        lstrMCDBAllIndexCode = lstrMCDBAllIndexCode & "," & Trim(lobjIndexRateNode.SelectNodes("INDEX_CODE").Item(0).InnerText)
                                        lstrMCDBAllIndexTerm = lstrMCDBAllIndexTerm & "," & Trim(lobjIndexRateNode.SelectNodes("INDEX_TERM").Item(0).InnerText)
                                    Next

                                    'remove first comma
                                    If Left(lstrMCDBAllIndexCode, 1) = "," Then lstrMCDBAllIndexCode = Mid(lstrMCDBAllIndexCode, 2, Len(lstrMCDBAllIndexCode))
                                    If Left(lstrMCDBAllIndexTerm, 1) = "," Then lstrMCDBAllIndexTerm = Mid(lstrMCDBAllIndexTerm, 2, Len(lstrMCDBAllIndexTerm))

                                    lstrCSVColumnHeaders = ""

                                    'get comma separated list of Money Cost file header
                                    For liCounter2 = 0 To lrsCSVRecordset.Columns.Count - 1
                                        lstrCSVColumnHeaders = lstrCSVColumnHeaders & Trim(lrsCSVRecordset.Columns(liCounter2).ColumnName.ToString) & ","
                                    Next

                                    If Right(lstrCSVColumnHeaders, 1) = "," Then lstrCSVColumnHeaders = Left(lstrCSVColumnHeaders, Len(lstrCSVColumnHeaders) - 1)

                                    'raise clarify case
                                    lstrErrorDetails = "Columns defined in schema file not matching to what are defined in Money Cost DB" & _
                                                        vbCrLf & "CSV column headers : " & lstrCSVColumnHeaders & vbCrLf & "Index Codes " & _
                                                        "in Money Cost DB: " & lstrMCDBAllIndexCode & vbCrLf & "Index Terms in Money Cost " & _
                                                        "DB : " & lstrMCDBAllIndexTerm
                                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                    Throw New Exception(lstrErrorDetails)
                                End If
                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Check whether no. of columns in CSV and Money Cost DB match END:-")

                                'get comma separated list of all unique Index Codes and Index Terms from MoneyCost DB
                                lobjIndexRateNodeList = lobjIndexRateDOM.SelectNodes("/INDEX_RATE_RESPONSE/INDEX_RATESet/INDEX_RATE[IND_QUERYDB = 'True']")

                                For liCounter2 = 0 To lobjIndexRateNodeList.Count - 1
                                    lobjIndexRateNode = lobjIndexRateNodeList.Item(liCounter2)

                                    If InStr("," + lstrINDEX_CODEList + ",", ",'" + UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_CODE").Item(0).InnerText)) + "',") <= 0 Then
                                        lstrINDEX_CODEList = lstrINDEX_CODEList + ",'" + UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_CODE").Item(0).InnerText)) + "'"
                                    End If

                                    If InStr("," + lstrINDEX_TERMList + ",", ",'" + UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_TERM").Item(0).InnerText)) + "',") <= 0 Then
                                        lstrINDEX_TERMList = lstrINDEX_TERMList + ",'" + UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_TERM").Item(0).InnerText)) + "'"
                                    End If
                                Next

                                'remove first comma
                                If Left(lstrINDEX_CODEList, 1) = "," Then lstrINDEX_CODEList = Mid(lstrINDEX_CODEList, 2, Len(lstrINDEX_CODEList))
                                If Left(lstrINDEX_TERMList, 1) = "," Then lstrINDEX_TERMList = Mid(lstrINDEX_TERMList, 2, Len(lstrINDEX_TERMList))

                                Do While True
                                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Generate CurveTypes list to pull Data from DW only with DB query flag = 1 START:-")

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
                                    lstrIndexDataRespXml = UCase(GetIndexData(lstrIndexDataReqXml))

                                    'validate, if response xml is well-formed
                                    Try
                                        lobjIndexDataDOMXml.LoadXml(lstrIndexDataRespXml)
                                    Catch ex As Exception
                                        'if error, raise clarify case
                                        lstrErrorDetails = "Error, while loading Response XML from lobjcMoneyCostAutoService.GetIndexData " & _
                                                            "method. " & vbCrLf & "Error No. : " & Err.Number & _
                                                            vbCrLf & "Error Description : " & Err.Description
                                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                        Throw
                                    End Try
                                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): loading Response XML from lobjcMoneyCostAutoService.GetIndexData END")
                                    lobjIndexDataNodeList = lobjIndexDataDOMXml.SelectNodes("/INDEX_DATA_RESPONSE/INDEX_DATASET/INDEX_DATA")
                                    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Check if data was fetched from DW START:-")
                                    'Check if data was fetched from DW
                                    If lobjIndexDataNodeList.Count > 0 Then

                                        ' ------------------------------------------------------------------------
                                        ' Delete any record for the Process date if already exist in the Recordset
                                        ' ------------------------------------------------------------------------                                        
                                        objcDataEntity.CommonSQL = "SELECT TOP 1 * FROM " & lstrMC_CODE & ".csv WHERE " & lstrProcessDateFieldName & " = #" & CDate(lstrProcessDate) & "#"
                                        objcDataEntity.WrkDirectory = lstrWorkingDirectory
                                        lrsCSVDeleteRecordset = objMoneyCostAutoDataClass.GetCsvRecords(objcDataEntity).CsvOutput
                                        If lrsCSVDeleteRecordset.Rows.Count > 0 Then
                                            'If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close()
                                            objcDataEntity.astrFileName = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                            objcDataEntity.ProcessDate = lstrProcessDate.ToString
                                            lbDataDeletionFlag = Convert.ToBoolean(UpdateFile(objcDataEntity).OutputString)

                                            If lbDataDeletionFlag = False Then
                                                lstrErrorDetails = "Error while deleting records from " & lstrMC_CODE & ".csv for Process Date : " & lstrProcessDate
                                                STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                                Throw New Exception(lstrErrorDetails)
                                            End If
                                        End If
                                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Load the CSV file from Working DIR into an ADO object. START:-")
                                        ' ------------------------------------------------------
                                        ' Load the CSV file from Working DIR into an ADO object.
                                        ' ------------------------------------------------------
                                        objcDataEntity.CommonSQL = "SELECT * FROM " & lstrMC_CODE & ".csv Order By " & lstrProcessDateFieldName
                                        objcDataEntity.WrkDirectory = lstrWorkingDirectory
                                        lrsCSVRecordset = objMoneyCostAutoDataClass.GetCsvRecords(objcDataEntity).CsvOutput


                                        'Reach the Last record in the existing MC recordsset. Value from this record would be picked
                                        'to fill missing data in new row being inserted.
                                        'lrsCSVRecordset.MoveLast()

                                        'For each Index Rate from MC_FILE with DB_Query = 1

                                        For liCounter2 = 0 To lobjIndexRateNodeList.Count - 1
                                            lobjIndexRateNode = lobjIndexRateNodeList.Item(liCounter2)

                                            ' -----------------------------------
                                            ' If Index Rate found in DW resultset  1
                                            ' -----------------------------------

                                            '''''lobjIndexDataNodeList = lobjIndexDataDOMXml.SelectNodes("/INDEX_DATA_RESPONSE/INDEX_DATASET/INDEX_DATA[YIELD_CURVE_TYPE = '" & UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_CODE").Item(0).InnerText)) & "' and TERM_PERIOD = '" & UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_TERM").Item(0).InnerText)) & "']")

                                            '''''If lobjIndexDataNodeList.Count > 0 Then
                                            '''''    'True
                                            '''''    lobjIndexDataNode = lobjIndexDataNodeList.Item(0)

                                            '''''    ' Get rate fetched from DW.
                                            '''''    lstrINTEREST_RATE_DWH = lobjIndexDataNode.SelectNodes("INTEREST_RATE").Item(0).InnerText

                                            'Change by Sumit START
                                            lstrINTEREST_RATE_DWH = "0"
                                            Dim bFlagIndexFound As Boolean = True
                                            Dim aSeperator As String() = {"','"}
                                            Dim aListIndexCode As ArrayList = New ArrayList(UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_CODE").Item(0).InnerText)).Split(aSeperator, StringSplitOptions.RemoveEmptyEntries))
                                            Dim IndexCodeIndicatorForRound As Char = UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_CODE_IND").Item(0).InnerText))


                                            For Each Item As String In aListIndexCode
                                                lobjIndexDataNodeList = lobjIndexDataDOMXml.SelectNodes("/INDEX_DATA_RESPONSE/INDEX_DATASET/INDEX_DATA[YIELD_CURVE_TYPE = '" & UCase(Trim(Item)) & "' and TERM_PERIOD = '" & UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_TERM").Item(0).InnerText)) & "']")

                                                If lobjIndexDataNodeList.Count > 0 Then
                                                    lobjIndexDataNode = lobjIndexDataNodeList.Item(0)
                                                    'If (String.Compare(lstrMC_CODE, "MCUTH", True) = 0) Then
                                                    '    lstrINTEREST_RATE_DWH = Convert.ToDouble(lstrINTEREST_RATE_DWH) + Math.Round(Convert.ToDouble(lobjIndexDataNode.SelectNodes("INTEREST_RATE").Item(0).InnerText), 4)
                                                    'Else
                                                    '    lstrINTEREST_RATE_DWH = Convert.ToDouble(lstrINTEREST_RATE_DWH) + Convert.ToDouble(lobjIndexDataNode.SelectNodes("INTEREST_RATE").Item(0).InnerText)
                                                    'End If

                                                    lstrINTEREST_RATE_DWH = Convert.ToDouble(lstrINTEREST_RATE_DWH) + Convert.ToDouble(lobjIndexDataNode.SelectNodes("INTEREST_RATE").Item(0).InnerText)

                                                Else
                                                    bFlagIndexFound = False
                                                    Exit For
                                                End If
                                            Next
                                            If String.Compare(IndexCodeIndicatorForRound, "Y", True) = 0 Then
                                                lstrINTEREST_RATE_DWH = Math.Round(Convert.ToDouble(lstrINTEREST_RATE_DWH), 4)
                                            End If

                                            'lobjIndexDataNodeList = lobjIndexDataDOMXml.SelectNodes("/INDEX_DATA_RESPONSE/INDEX_DATASET/INDEX_DATA[YIELD_CURVE_TYPE = '" & UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_CODE").Item(0).InnerText)) & "' and TERM_PERIOD = '" & UCase(Trim(lobjIndexRateNode.SelectNodes("INDEX_TERM").Item(0).InnerText)) & "']")

                                            'If lobjIndexDataNodeList.Count > 0 Then
                                            If bFlagIndexFound Then
                                                'True
                                                'lobjIndexDataNode = lobjIndexDataNodeList.Item(0)

                                                ' Get rate fetched from DW.
                                                'lstrINTEREST_RATE_DWH = lobjIndexDataNode.SelectNodes("INTEREST_RATE").Item(0).InnerText

                                                'Change by Sumit END

                                                'check if the interest rate returned from DWH is not blank... 1
                                                '---------------------------------------
                                                If IsNumeric(lstrINTEREST_RATE_DWH) Then
                                                    If UCase(Trim(lobjIndexRateNode.SelectNodes("IND_PERCENTILE").Item(0).InnerText)) = "TRUE" Then
                                                        ' Apply Percentile for the Index Rates requiring the same.
                                                        lstrINTEREST_RATE_DWH = Val(lstrINTEREST_RATE_DWH) / 100
                                                    End If

                                                    ' Add Adder value defined for each Index Rate to the processed value above
                                                    lstrINTEREST_RATE_DWH = Val(lstrINTEREST_RATE_DWH) + Val(lobjIndexRateNode.SelectNodes("AMT_ADDER").Item(0).InnerText)
                                                Else
                                                    'if the interest rate returned from the DWH is blank or some non-numeric value , then check for previous day's index data required or not
                                                    'if IND_PREV_INDEXRATES_REQ = 0 then not required to copy previous day's index data
                                                    If Trim(lobjIndexRateNode.SelectNodes("IND_PREV_INDEXRATES_REQ").Item(0).InnerText) = "0" Then
                                                        lstrMissingYieldCurve_NotCopied = lstrMissingYieldCurve_NotCopied & "," & Trim(lobjIndexRateNode.SelectNodes("DESCRIPTION").Item(0).InnerText)
                                                    Else
                                                        'if the interest rate returned from the DWH is blank or some non-numeric value , then copy the previous day's index data.
                                                        '---------------------------------------
                                                        ' Get the value for Index Rate from Previous days record in MC File Recordset

                                                        If lrsCSVRecordset.Rows.Count > 0 Then

                                                            Count = Val(lobjIndexRateNode.SelectNodes("MC_FILE_COL_POSITION").Item(0).InnerText)

                                                            If Not (lrsCSVRecordset.Rows(lrsCSVRecordset.Rows.Count - 1).Item(Count).Value) Is Nothing Then
                                                                Count = Val(lobjIndexRateNode.SelectNodes("MC_FILE_COL_POSITION").Item(0).InnerText)
                                                                lstrINTEREST_RATE_DWH = lrsCSVRecordset.Rows(lrsCSVRecordset.Rows.Count - 1).Item(Count).ToString
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
                                                        objcDataEntity.OutputString = lstrIndexRateReqXml
                                                        lstrIndexRateRespXml_PrevDay = lobjBSMCIMoneyCostService.GetIndexRates(objcDataEntity).OutputString
                                                        'validate, if Index Rate Response XML is well-formed
                                                        Try
                                                            lobjIndexRateDOM_PrevDay.LoadXml(lstrIndexRateRespXml_PrevDay)
                                                        Catch ex As Exception
                                                            'if error, raise clarify case
                                                            lstrErrorDetails = "Error while loading Response XML from lobjBSMCIMoneyCostService.GetIndexRates " & _
                                                                                "method for Previous Day." & vbCrLf & "Error no. : " & Err.Number & _
                                                                                vbCrLf & "Error description : " & Err.Description
                                                            STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                                            Throw
                                                        End Try


                                                        'Not Checking whether no. of columns in CSV and Money Cost DB match as it has already been checked in the first call.

                                                        'get comma separated list of all unique Index Codes and Index Terms from MoneyCost DB
                                                        lobjIndexRateNodeList_PrevDay = lobjIndexRateDOM_PrevDay.SelectNodes("/INDEX_RATE_RESPONSE/INDEX_RATESet/INDEX_RATE[IND_QUERYDB = 'True']")

                                                        For liCounter3 = 0 To lobjIndexRateNodeList_PrevDay.Count - 1
                                                            lobjIndexRateNode_PrevDay = lobjIndexRateNodeList_PrevDay.Item(liCounter3)
                                                            If liCounter2 = liCounter3 Then
                                                                lstrINTEREST_RATE_DWH = Val(lstrINTEREST_RATE_DWH) - Val(lobjIndexRateNode_PrevDay.SelectNodes("AMT_ADDER").Item(0).InnerText) + Val(lobjIndexRateNode.SelectNodes("AMT_ADDER").Item(0).InnerText)
                                                            End If
                                                        Next

                                                        '------------------------------------------------------------------
                                                        ' Maintain a list of all such Index Rates and the values copied.
                                                        lstrMissingYieldCurve = lstrMissingYieldCurve & "," & Trim(lobjIndexRateNode.SelectNodes("DESCRIPTION").Item(0).InnerText)
                                                        lstrMissingInterestRate = lstrMissingInterestRate & "," & lstrINTEREST_RATE_DWH
                                                    End If
                                                    '---------------------------------------
                                                End If
                                                'Else
                                            Else
                                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): if the interest rate returned from the DWH is blank or some non-numeric value , then check for previous day's index data required or not")
                                                'if the interest rate returned from the DWH is blank or some non-numeric value , then check for previous day's index data required or not
                                                'if IND_PREV_INDEXRATES_REQ = 0 then not required to copy previous day's index data
                                                If Trim(lobjIndexRateNode.SelectNodes("IND_PREV_INDEXRATES_REQ").Item(0).InnerText) = "0" Then
                                                    lstrMissingYieldCurve_NotCopied = lstrMissingYieldCurve_NotCopied & "," & Trim(lobjIndexRateNode.SelectNodes("DESCRIPTION").Item(0).InnerText)
                                                Else
                                                    'if the interest rate returned from the DWH is blank or some non-numeric value , then copy the previous day's index data.
                                                    '---------------------------------------
                                                    ' Get the value for Index Rate from Previous days record in MC File Recordset
                                                    If lrsCSVRecordset.Rows.Count > 0 Then
                                                        Count = Val(lobjIndexRateNode.SelectNodes("MC_FILE_COL_POSITION").Item(0).InnerText)
                                                        If Not (lrsCSVRecordset.Rows(lrsCSVRecordset.Rows.Count - 1).Item(Count)) Is Nothing Then
                                                            lstrINTEREST_RATE_DWH = lrsCSVRecordset.Rows(lrsCSVRecordset.Rows.Count - 1).Item(Count).ToString
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
                                                    objcDataEntity.OutputString = lstrIndexRateReqXml
                                                    lstrIndexRateRespXml_PrevDay = lobjBSMCIMoneyCostService.GetIndexRates(objcDataEntity).OutputString
                                                    'validate, if Index Rate Response XML is well-formed
                                                    Try
                                                        lobjIndexRateDOM_PrevDay.LoadXml(lstrIndexRateRespXml_PrevDay)
                                                    Catch ex As Exception
                                                        'if error, raise clarify case
                                                        lstrErrorDetails = "Error while loading Response XML from lobjBSMCIMoneyCostService.GetIndexRates " & _
                                                                            "method for Previous Day." & vbCrLf & "Error no. : " & Err.Number & _
                                                                            vbCrLf & "Error description : " & Err.Description
                                                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                                        Throw
                                                    End Try


                                                    'Not Checking whether no. of columns in CSV and Money Cost DB match as it has already been checked in the first call.

                                                    'get comma separated list of all unique Index Codes and Index Terms from MoneyCost DB
                                                    lobjIndexRateNodeList_PrevDay = lobjIndexRateDOM_PrevDay.SelectNodes("/INDEX_RATE_RESPONSE/INDEX_RATESet/INDEX_RATE[IND_QUERYDB = 'True']")

                                                    For liCounter3 = 0 To lobjIndexRateNodeList_PrevDay.Count - 1
                                                        lobjIndexRateNode_PrevDay = lobjIndexRateNodeList_PrevDay.Item(liCounter3)
                                                        If liCounter2 = liCounter3 Then
                                                            lstrINTEREST_RATE_DWH = Val(lstrINTEREST_RATE_DWH) - Val(lobjIndexRateNode_PrevDay.SelectNodes("AMT_ADDER").Item(0).InnerText) + Val(lobjIndexRateNode.SelectNodes("AMT_ADDER").Item(0).InnerText)
                                                        End If
                                                    Next
                                                    '------------------------------------------------------------------
                                                    ' Maintain a list of all such Index Rates and the values copied.
                                                    lstrMissingYieldCurve = lstrMissingYieldCurve & "," & Trim(lobjIndexRateNode.SelectNodes("DESCRIPTION").Item(0).InnerText)
                                                    lstrMissingInterestRate = lstrMissingInterestRate & "," & lstrINTEREST_RATE_DWH
                                                End If
                                                'Endif
                                            End If
                                            'If NOT lstrMissingYieldCurve_NotCopied = "" Then
                                            lstrCSV_INSERT_RECORDXml = lstrCSV_INSERT_RECORDXml & _
                                                                        "<CSV_INSERT_DATA>" & _
                                                                            "<COL_POSITION>" & Trim(lobjIndexRateNode.SelectNodes("MC_FILE_COL_POSITION").Item(0).InnerText) & "</COL_POSITION>" & _
                                                                            "<INTEREST_RATE>" & lstrINTEREST_RATE_DWH & "</INTEREST_RATE>" & _
                                                                        "</CSV_INSERT_DATA>"
                                            'End If
                                            'Next
                                        Next

                                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): remove first comma START:-")
                                        'remove first comma
                                        If Not lstrMissingYieldCurve_NotCopied = "" Then

                                            lstrMissingYieldCurve_NotCopied = lstrMissingYieldCurve_NotCopied & lstrMissingYieldCurve

                                            If Left(lstrMissingYieldCurve_NotCopied, 1) = "," Then lstrMissingYieldCurve_NotCopied = Mid(lstrMissingYieldCurve_NotCopied, 2, Len(lstrMissingYieldCurve_NotCopied))

                                            'If data was found missing for some Index Rates, Create a p3 case for the respective queue and include the list of Index Rates missing.
                                            lstrErrorDetails = "Errors were reported while updating " & lstrMC_CODE & " file for " & Format(Convert.ToDateTime(lstrProcessDate), "MM/dd/yyyy") & ". " & vbCrLf & _
                                                                "Interest rates for the following were missing. So " & lstrMC_CODE & " file for " & Format(Convert.ToDateTime(lstrProcessDate), "MM/dd/yyyy") & " is not updated. " & vbCrLf & vbCrLf & _
                                                                "S.No." & vbTab & "Index Name" & vbCrLf & vbCrLf

                                            larrMissingIndexCode_NotCopied = Split(lstrMissingYieldCurve_NotCopied, ",")
                                            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)

                                            For liCounter2 = 0 To UBound(larrMissingIndexCode_NotCopied, 1)
                                                lstrErrorDetails = lstrErrorDetails & CStr(liCounter2 + 1) & ")" & vbTab & _
                                                                    larrMissingIndexCode_NotCopied(liCounter2) & vbCrLf
                                            Next
                                            objFileEntity.QueueName = lstrClarifyQName
                                            objFileEntity.BusinessContact = lstrBusinessContact
                                            objFileEntity.Body = lstrErrorDetails
                                            objFileEntity.CutTicket = "False"
                                            objFileEntity.SendNotification = "True"
                                            objFileEntity.MCCode = lstrMC_CODE
                                            objFileEntity.ProcessDates = lstrProcessDate
                                            SendErrNotification(objFileEntity)
                                            SendErrMailForCurrencyFlag = True
                                            '*****************************************************************
                                            'change to incorporate Encryption and Decryption of CSV files.
                                            If EncrptFileFlag = "True" Then
                                                If EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv").ToString.ToUpper = "DECRYPTED" Then
                                                    lstrRetMsg = EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv")
                                                End If
                                            End If
                                            '*****************************************************************
                                            'lstrErrorDetails = ""
                                            '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
                                            'close the connection, if not closed
                                            'If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()
                                            'If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close()

                                            'Copy Working File back to the FTP Location.
                                            STLogger.Debug("Copy Working File back to the FTP Location.")
                                            objFileEntity.Source = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                            objFileEntity.Destination = lstrFTP_Location & lstrFTP_Directory & lstrMC_CODE & ".csv"
                                            lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString
                                            STLogger.Debug("Copied Working File back to the FTP Location. " & objFileEntity.Source & "-->" & objFileEntity.Destination)

                                            '                                    If lobjcFTP.Connected Then lobjcFTP.Disconnect
                                            '                                    lobjcFTP.Connect lstrFTP_Location, lstrFTP_User, lstrFTP_Password
                                            '                                    lobjcFTP.Directory = lstrFTP_Directory
                                            '                                    lstrCopyFileResponse = lobjcFTP.PutFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrMC_CODE & ".csv")
                                            '                                    lobjcFTP.Disconnect

                                            'if any error occurred, raise clarify case
                                            If lstrCopyFileResponse <> "" Then
                                                lstrErrorDetails = lstrErrorDetails & "Error while updating " & lstrMC_CODE & ".csv file as on the FTP location : " & lstrFTP_Location & " : " & lstrCopyFileResponse
                                                STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                                'If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()
                                                Throw New Exception(lstrErrorDetails)
                                            End If

                                            'Copy working File over network location, if Network Location is defined in Registry
                                            If lstrNetwork_Location <> "" Then
                                                STLogger.Debug("Copy working File over network location ")
                                                objFileEntity.Source = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                                objFileEntity.Destination = lstrNetwork_Location & lstrMC_CODE & ".csv"
                                                lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString
                                                STLogger.Debug("Copied working File over network location " & objFileEntity.Source & "-->" & objFileEntity.Destination)
                                                'if any error occurred, raise clarify case
                                                If lstrCopyFileResponse <> "" Then
                                                    lstrErrorDetails = lstrErrorDetails & "Error while saving " & lstrMC_CODE & ".csv to network location : " & _
                                                                        lstrNetwork_Location & " as " & lstrMC_CODE & ".csv" & vbCrLf & _
                                                                        lstrCopyFileResponse
                                                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                                    Throw New Exception(lstrErrorDetails)
                                                End If
                                            End If
                                            '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

                                            Throw New Exception
                                        End If


                                        'remove first comma
                                        If Left(lstrMissingYieldCurve, 1) = "," Then lstrMissingYieldCurve = Mid(lstrMissingYieldCurve, 2, Len(lstrMissingYieldCurve))
                                        If Left(lstrMissingInterestRate, 1) = "," Then lstrMissingInterestRate = Mid(lstrMissingInterestRate, 2, Len(lstrMissingInterestRate))
                                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): remove first comma END")

                                        'If data was found missing for some Index Rates, Create a p3 case for the respective queue and include the list of Index Rates missing.
                                        If lstrMissingYieldCurve <> "" And lstrMissingInterestRate <> "" Then

                                            lstrErrorDetails = "Errors were reported while updating " & lstrMC_CODE & " file for " & Format(Convert.ToDateTime(lstrProcessDate), "MM/dd/yyyy") & ". " & vbCrLf & _
                                                                "Interest rates for following were copied with prior days values. " & vbCrLf & vbCrLf & _
                                                                "S.No." & vbTab & "Index Name & Value" & vbCrLf & vbCrLf
                                            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                            larrMissingIndexCode = Split(lstrMissingYieldCurve, ",")
                                            larrMissingIntRate = Split(lstrMissingInterestRate, ",")
                                            For liCounter2 = 0 To UBound(larrMissingIndexCode, 1)
                                                lstrErrorDetails = lstrErrorDetails & CStr(liCounter2 + 1) & ")" & vbTab & _
                                                                    larrMissingIndexCode(liCounter2) & " = " & _
                                                                    larrMissingIntRate(liCounter2) & vbCrLf
                                            Next
                                            ' Added By Sanjay For getting Currency Code in Sucess mail for Prior rate Updation
                                            If ProcessMoneyCostFiles_CurrencyCode.Trim <> "" Then
                                                ProcessMoneyCostFiles = ProcessMoneyCostFiles & vbCrLf & lstrMC_CODE
                                            End If
                                            ''''''''''''''''''''''''''''
                                        End If

                                        lobjIndexRateNodeList = lobjIndexRateDOM.SelectNodes("/INDEX_RATE_RESPONSE/INDEX_RATESet/INDEX_RATE[IND_QUERYDB = 'False']")
                                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): For each Index Rate from MC_FILE with DB_Query = 0")
                                        'For each Index Rate from MC_FILE with DB_Query = 0
                                        For liCounter2 = 0 To lobjIndexRateNodeList.Count - 1
                                            lobjIndexRateNode = lobjIndexRateNodeList.Item(liCounter2)

                                            ' Add Adder value defined for each Index Rate to the processed value above
                                            lstrINTEREST_RATE_DWH = Trim(lobjIndexRateNode.SelectNodes("AMT_ADDER").Item(0).InnerText)

                                            lstrCSV_INSERT_RECORDXml = lstrCSV_INSERT_RECORDXml & _
                                                                        "<CSV_INSERT_DATA>" & _
                                                                            "<COL_POSITION>" & Trim(lobjIndexRateNode.SelectNodes("MC_FILE_COL_POSITION").Item(0).InnerText) & "</COL_POSITION>" & _
                                                                            "<INTEREST_RATE>" & lstrINTEREST_RATE_DWH & "</INTEREST_RATE>" & _
                                                                        "</CSV_INSERT_DATA>"
                                            'Next
                                        Next

                                        lstrCSV_INSERT_RECORDXml = lstrCSV_INSERT_RECORDXml & "</CSV_INSERT_DATASet></CSV_INSERT_DATA_REQUEST>"

                                        'Arrange all Index Rates based upon the Column Position and prepare an Insert Statement
                                        objxmlErrEntity.XmlDoc = lstrCSV_INSERT_RECORDXml
                                        objxmlErrEntity.XslFilePath = lstrWorkingDirectory & "CSVInsertXMLSorting.xsl"
                                        lstrCSV_INSERT_RECORDXml = fnSortXmlData(objxmlErrEntity)

                                        lstrCSV_INSERT_RECORDXml = Replace(lstrCSV_INSERT_RECORDXml, "<?xml version=""1.0""?>", "")

                                        'validate, if dynamically build XMl for CSV insert is well-formed
                                        Try
                                            lobjCSVInsertDOMXml.LoadXml(lstrCSV_INSERT_RECORDXml)
                                        Catch ex As Exception
                                            lstrErrorDetails = "Error while loading request XML for CSV data insert. " & vbCrLf & _
                                                                "Error No. : " & Err.Number & vbCrLf & _
                                                                "Error Description : " & Err.Description
                                            STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails)
                                            Throw
                                        End Try


                                        lstrCSVInsertSQL = "INSERT INTO " & lstrMC_CODE & ".csv VALUES('" & lstrProcessDate & "', "

                                        'Insert new row to MC File Recordset
                                        lobjCSVInsertNodeList = lobjCSVInsertDOMXml.SelectNodes("/CSV_INSERT_DATA_REQUEST/CSV_INSERT_DATASet/CSV_INSERT_DATA")

                                        For liCounter2 = 0 To lobjCSVInsertNodeList.Count - 1
                                            lobjCSVInsertNode = lobjCSVInsertNodeList.Item(liCounter2)
                                            'If (String.Compare(lstrMC_CODE, "MCUTH", True) = 0) Then
                                            '    lstrCSVInsertSQL = lstrCSVInsertSQL & "'" & Math.Round(Convert.ToDecimal(Trim(lobjCSVInsertNode.SelectNodes("INTEREST_RATE").Item(0).InnerText)), 4) & "', "
                                            'Else
                                            '    lstrCSVInsertSQL = lstrCSVInsertSQL & "'" & Trim(lobjCSVInsertNode.SelectNodes("INTEREST_RATE").Item(0).InnerText) & "', "
                                            'End If
                                            lstrCSVInsertSQL = lstrCSVInsertSQL & "'" & Trim(lobjCSVInsertNode.SelectNodes("INTEREST_RATE").Item(0).InnerText) & "', "
                                            'lstrCSVInsertSQL = lstrCSVInsertSQL & "'" & Math.Round(Convert.ToDecimal(Trim(lobjCSVInsertNode.SelectNodes("INTEREST_RATE").Item(0).InnerText)), 4) & "', "
                                        Next
                                        lstrCSVInsertSQL = Mid(lstrCSVInsertSQL, 1, Len(lstrCSVInsertSQL) - 2)
                                        lstrCSVInsertSQL = lstrCSVInsertSQL & ");"
                                        'Save ADO recordset back to the working file as MCXXX.csv.
                                        objcDataEntity.CommonSQL = lstrCSVInsertSQL
                                        objcDataEntity.WrkDirectory = lstrWorkingDirectory
                                        objMoneyCostAutoDataClass.UpdateCsvRecords(objcDataEntity)

                                        '*****************************************************************
                                        'change to incorporate Encryption and Decryption of CSV files.
                                        If EncrptFileFlag = "True" Then
                                            If EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv").ToString.ToUpper = "DECRYPTED" Then
                                                lstrRetMsg = EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv")
                                            End If
                                        End If
                                        '*****************************************************************
                                        'close the connection, if closed
                                        'If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()
                                        'If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close()

                                        'Copy Working File back to the FTP Location.
                                        STLogger.Debug("Copy Working File back to the FTP Location.")
                                        objFileEntity.Source = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                        objFileEntity.Destination = lstrFTP_Location & lstrFTP_Directory & lstrMC_CODE & ".csv"
                                        lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString
                                        STLogger.Debug("Copied Working File back to the FTP Location " & objFileEntity.Source & " -->" & objFileEntity.Destination)

                                        '                                If lobjcFTP.Connected Then lobjcFTP.Disconnect
                                        '                                lobjcFTP.Connect lstrFTP_Location, lstrFTP_User, lstrFTP_Password
                                        '                                lobjcFTP.Directory = lstrFTP_Directory
                                        '                                lstrCopyFileResponse = lobjcFTP.PutFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrMC_CODE & ".csv")
                                        '                                lobjcFTP.Disconnect

                                        'if any error occurred, raise clarify case
                                        If lstrCopyFileResponse <> "" Then
                                            lstrErrorDetails = "Error while updating " & lstrMC_CODE & ".csv file as on the FTP location : " & lstrFTP_Location & " : " & lstrCopyFileResponse
                                            STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrErrorDetails)
                                            'If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()
                                            Throw New Exception(lstrErrorDetails)
                                        End If

                                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Copy working File over network location, if Network Location is defined in Registry")
                                        'Copy working File over network location, if Network Location is defined in Registry
                                        If lstrNetwork_Location <> "" Then
                                            objFileEntity.Source = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                            objFileEntity.Destination = lstrNetwork_Location & lstrMC_CODE & ".csv"
                                            lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString
                                            STLogger.Debug("Copied working File over network location " & objFileEntity.Source & "-->" & objFileEntity.Destination)
                                            'if any error occurred, raise clarify case
                                            If lstrCopyFileResponse <> "" Then
                                                lstrErrorDetails = "Error while saving " & lstrMC_CODE & ".csv to network location : " & _
                                                                    lstrNetwork_Location & " as " & lstrMC_CODE & ".csv" & vbCrLf & _
                                                                    lstrCopyFileResponse
                                                STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails)
                                                Throw New Exception(lstrErrorDetails)
                                            End If
                                        End If

                                        STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Take the generated files , change the date format and move it to the specified location as defined in registry.")

                                        'Take the generated files , change the date format and move it to the specified location as defined in registry.
                                        If (lstrFTP_LocationForNewDateFormat <> "") Then

                                            If Right(lstrFTP_LocationForNewDateFormat, 1) <> "\" Then lstrFTP_LocationForNewDateFormat = lstrFTP_LocationForNewDateFormat & "\"
                                            If Right(lstrFTP_DirectoryForNewDateFormat, 1) <> "\" Then lstrFTP_DirectoryForNewDateFormat = lstrFTP_DirectoryForNewDateFormat & "\"

                                            'check if Date Formatting and copying is required for this particular Currency
                                            If lstrDateFormatRequired = True Then

                                                'Copy the new ini file with the extn of MCCode_AUD.ini which stores the date in dd/mm/yyyy format.
                                                objFileEntity.Source = lstrWorkingDirectory & lstrMC_CODE & "_AUD.ini"
                                                objFileEntity.Destination = lstrWorkingDirectory & "schema.ini"
                                                lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString

                                                'if any error occurred, raise clarify case
                                                If lstrCopyFileResponse <> "" Then
                                                    'raise clarify case
                                                    'schema file not present for MCXXX.csv file.
                                                    'Send notification to business MCXXX file not updated for Process Date.
                                                    'Details: Schema file not defined for MCXXX.csv

                                                    lstrErrorDetails = "Schema file with new Date Format not present for " & lstrMC_CODE & ".csv" & vbCrLf & _
                                                                        "Error while defining schema file : " & lstrCopyFileResponse
                                                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails)
                                                    Throw New Exception(lstrErrorDetails)
                                                End If


                                                'copy from new ftp location to working directory
                                                objFileEntity.Source = lstrFTP_LocationForNewDateFormat & lstrFTP_DirectoryForNewDateFormat & lstrMC_CODE & ".csv"
                                                objFileEntity.Destination = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                                lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString

                                                '                                        If lobjcFTP.Connected Then lobjcFTP.Disconnect
                                                '                                        lobjcFTP.Connect lstrFTP_LocationForNewDateFormat, lstrFTP_UserForNewDateFormat, lstrFTP_PasswordForNewDateFormat
                                                '                                        lobjcFTP.Directory = lstrFTP_DirectoryForNewDateFormat
                                                '                                        lstrCopyFileResponse = lobjcFTP.GetFile(lstrMC_CODE & ".csv", lstrWorkingDirectory & lstrMC_CODE & ".csv")
                                                '                                        lobjcFTP.Disconnect

                                                'if any error occurred, raise clarify case
                                                '*****************************************************************
                                                'change to incorporate Encryption and Decryption of CSV files.
                                                If EncrptFileFlag = "True" Then
                                                    If EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv").ToString.ToUpper = "ENCRYPTED" Then
                                                        lstrRetMsg = EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv")
                                                    End If
                                                End If
                                                '*****************************************************************
                                                If lstrCopyFileResponse <> "" Then
                                                    lstrErrorDetails = "Error while creating working copy of existing " & lstrMC_CODE & ".csv file as " & _
                                                                        lstrMC_CODE & ".csv from New FTP on " & lstrWorkingDirectory & " : " & lstrCopyFileResponse

                                                    ' If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()                                                   
                                                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails)
                                                    Throw New Exception(lstrErrorDetails)
                                                End If
                                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Delete any record for the Process date if already exist in the Recordset START:-")

                                                ' ------------------------------------------------------------------------
                                                ' Delete any record for the Process date if already exist in the Recordset
                                                ' ------------------------------------------------------------------------
                                                'If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()
                                                objcDataEntity.CommonSQL = "SELECT TOP 1 * FROM " & lstrMC_CODE & ".csv WHERE " & "#" & Format(Convert.ToDateTime(lstrProcessDate), "dd/MM/yyyy") & "#"
                                                objcDataEntity.WrkDirectory = lstrWorkingDirectory
                                                lrsCSVRecordset = objMoneyCostAutoDataClass.GetCsvRecords(objcDataEntity).CsvOutput

                                                lstrDay_NewDateFormat = Day(lstrProcessDate)
                                                If Len(lstrDay_NewDateFormat) = 1 Then lstrDay_NewDateFormat = "0" & lstrDay_NewDateFormat
                                                lstrMonth_NewDateFormat = Month(lstrProcessDate)
                                                If Len(lstrMonth_NewDateFormat) = 1 Then lstrMonth_NewDateFormat = "0" & lstrMonth_NewDateFormat

                                                lstrProcessDate_NewDateFormat = lstrMonth_NewDateFormat & "/" & lstrDay_NewDateFormat & "/" & Year(lstrProcessDate)

                                                If lrsCSVRecordset.Rows.Count > 0 Then
                                                    'If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()
                                                    objcDataEntity.astrFileName = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                                    objcDataEntity.ProcessDate = Convert.ToDateTime(lstrProcessDate_NewDateFormat)
                                                    lbDataDeletionFlag = Convert.ToBoolean(UpdateFile(objcDataEntity).OutputString)

                                                    If lbDataDeletionFlag = False Then
                                                        lstrErrorDetails = "Error while deleting records from " & lstrMC_CODE & ".csv for Process Date : " & lstrProcessDate_NewDateFormat

                                                        'If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()                                                        
                                                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails)
                                                        Throw New Exception(lstrErrorDetails)
                                                    End If
                                                End If

                                                'Insert  with the new formatted date....
                                                lstrCSVInsertSQL = "INSERT INTO " & lstrMC_CODE & ".csv VALUES('" & Format(Convert.ToDateTime(lstrProcessDate), lstrDateFormat.Replace("m", "M")) & "', "

                                                'Insert new row to MC File Recordset with the new formatted date.
                                                lobjCSVInsertNodeList = lobjCSVInsertDOMXml.SelectNodes("/CSV_INSERT_DATA_REQUEST/CSV_INSERT_DATASet/CSV_INSERT_DATA")

                                                For liCounter2 = 0 To lobjCSVInsertNodeList.Count - 1
                                                    lobjCSVInsertNode = lobjCSVInsertNodeList.Item(liCounter2)

                                                    lstrCSVInsertSQL = lstrCSVInsertSQL & "'" & Trim(lobjCSVInsertNode.SelectNodes("INTEREST_RATE").Item(0).InnerText) & "', "
                                                    'lstrCSVInsertSQL = lstrCSVInsertSQL & "'" & Math.Round(Convert.ToDecimal(Trim(lobjCSVInsertNode.SelectNodes("INTEREST_RATE").Item(0).InnerText)), 4) & "', "
                                                Next

                                                lstrCSVInsertSQL = Mid(lstrCSVInsertSQL, 1, Len(lstrCSVInsertSQL) - 2)

                                                lstrCSVInsertSQL = lstrCSVInsertSQL & ");"

                                                objcDataEntity.CommonSQL = lstrCSVInsertSQL
                                                objcDataEntity.WrkDirectory = lstrWorkingDirectory
                                                objMoneyCostAutoDataClass.UpdateCsvRecords(objcDataEntity)
                                                '*****************************************************************
                                                'change to incorporate Encryption and Decryption of CSV files.
                                                If EncrptFileFlag = "True" Then
                                                    If EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv").ToString.ToUpper = "DECRYPTED" Then
                                                        lstrRetMsg = EncrptFile(lstrWorkingDirectory & "EncryptFiles.Exe", lstrWorkingDirectory & lstrMC_CODE & ".csv")
                                                    End If
                                                End If
                                                '*****************************************************************

                                                'close the connection, if closed
                                                'If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()
                                                'If lrsCSVDeleteRecordset.State = adStateOpen Then lrsCSVDeleteRecordset.Close()
                                                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): copy this version of the file in the new FTP location. START:-")
                                                'copy this version of the file in the new FTP location.
                                                objFileEntity.Source = lstrWorkingDirectory & lstrMC_CODE & ".csv"
                                                objFileEntity.Destination = lstrFTP_LocationForNewDateFormat & lstrFTP_DirectoryForNewDateFormat & lstrMC_CODE & ".csv"
                                                lstrCopyFileResponse = CopyFiles(objFileEntity).OutputString

                                                '                                        If lobjcFTP.Connected Then lobjcFTP.Disconnect
                                                '                                        lobjcFTP.Connect lstrFTP_LocationForNewDateFormat, lstrFTP_UserForNewDateFormat, lstrFTP_PasswordForNewDateFormat
                                                '                                        lobjcFTP.Directory = lstrFTP_DirectoryForNewDateFormat
                                                '                                        lstrCopyFileResponse = lobjcFTP.PutFile(lstrWorkingDirectory & lstrMC_CODE & ".csv", lstrMC_CODE & ".csv")
                                                '                                        lobjcFTP.Disconnect

                                                'if any error occurred, raise clarify case
                                                If lstrCopyFileResponse <> "" Then
                                                    lstrErrorDetails = "Error while updating " & lstrMC_CODE & ".csv file as on the New FTP location : " & lstrFTP_LocationForNewDateFormat & " : " & lstrCopyFileResponse
                                                    STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): " & lstrErrorDetails)
                                                    'If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()
                                                    Throw New Exception(lstrErrorDetails)
                                                End If

                                            End If

                                        End If


                                        'If data was found missing for some Index Rates, Create a p3 case for the respective queue
                                        'and include the list of Index Rates missing.
                                        If lstrMissingYieldCurve <> "" And lstrMissingInterestRate <> "" Then
                                            objFileEntity.QueueName = lstrClarifyQName
                                            objFileEntity.BusinessContact = lstrBusinessContact
                                            objFileEntity.Body = lstrErrorDetails
                                            objFileEntity.CutTicket = "False"
                                            objFileEntity.SendNotification = "True"
                                            objFileEntity.MCCode = lstrMC_CODE
                                            objFileEntity.ProcessDates = lstrProcessDate
                                            SendErrNotification(objFileEntity)
                                            SendErrMailForCurrencyFlag = True
                                            lstrErrorDetails = ""
                                            STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): If data was found missing for some Index Rates, Create a p3 case for the respective queue and include the list of Index Rates missing.")
                                            Throw New Exception
                                        End If

                                        lstrErrorDetails = ""

                                        Exit Do

                                    Else
                                        If liMARKET_CLOSED_DWH_CHECK_COUNTER = 0 Then
                                            'If lrsCSVRecordset.State = adStateOpen Then lrsCSVRecordset.Close()
                                            ' Assumption that Markets were closed for Process date. Stop the process for current file and skip to next file.
                                            ProcessMoneyCostFiles_CurrencyCode = ""
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
                            Else        '#### Added By Atul On 18-09-2011
                                STLogger.Error("Unable to map or connect ftp location: " + lstrFTP_Location + "," + Environment.NewLine + "Process not completed today for moneycost file " + lstrMC_CODE)
                                NotificationMailFlag = True
                                objFileEntity.QueueName = lstrClarifyQName
                                objFileEntity.BusinessContact = lstrBusinessContact
                                objFileEntity.Body = "Unable to map or connect ftp location " + lstrFTP_Location + "," + Environment.NewLine + "Process not completed today for moneycost file " + lstrMC_CODE
                                objFileEntity.CutTicket = "False"
                                objFileEntity.SendNotification = "True"
                                objFileEntity.MCCode = lstrMC_CODE
                                objFileEntity.ProcessDates = lstrProcessDate
                                ProcessMoneyCostFiles = ProcessMoneyCostFiles.Replace(lstrMC_CODE, "") '& vbCrLf & lstrMC_CODE
                                lstrSendErrNotiResult = SendErrNotification(objFileEntity).OutputString
                            End If
                            'FTP Location Mapped End
                        End If
                        If ProcessMoneyCostFiles_CurrencyCode.Trim <> "" Then
                            ProcessMoneyCostFiles = ProcessMoneyCostFiles & vbCrLf & lstrMC_CODE
                        End If

                    Catch ex As Exception
                        STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow(): Error while updating " & lstrMC_CODE & ".csv file ErrorMsg:-" & ex.Message)
                        lstrErrorDetails = ex.Message.ToString()
                    Finally
                        lstrMCFileEndExec = Now()

                        If File.Exists(lstrWorkingDirectory & lstrMC_CODE & ".csv") Then
                            Kill(lstrWorkingDirectory & lstrMC_CODE & ".csv")
                        End If

                        If File.Exists(lstrWorkingDirectory & lstrMC_CODE & ".csv") Then
                            Kill(lstrWorkingDirectory & "schema.ini")
                        End If

                        'if some error found during processing of Money Cost file
                        If (lstrErrorDetails <> "" Or Err.Description <> "") And SendErrMailForCurrencyFlag = False Then
                            'raise clarify ticket

                            objFileEntity.QueueName = lstrClarifyQName
                            objFileEntity.BusinessContact = lstrBusinessContact
                            objFileEntity.Body = IIf(lstrErrorDetails = "", "Error No. : " & Err.Number & vbCrLf & _
                                                    "Error Description : " & Err.Description, lstrErrorDetails) & " For " & lstrMC_CODE & ".csv file on Proessing Date : " & Format(Convert.ToDateTime(lstrProcessDate), "MM/dd/yyyy") & "."
                            objFileEntity.CutTicket = "False"
                            objFileEntity.SendNotification = "True"
                            objFileEntity.MCCode = lstrMC_CODE
                            objFileEntity.ProcessDates = lstrProcessDate
                            ProcessMoneyCostFiles = ProcessMoneyCostFiles.Replace(lstrMC_CODE, "") '& vbCrLf & lstrMC_CODE
                            lstrSendErrNotiResult = SendErrNotification(objFileEntity).OutputString

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
                                                                ".csv file on Proessing Date : " & Format(Convert.ToDateTime(lstrProcessDate), "MM/dd/yyyy") & "." & _
                                                            "</DETAILS>" & _
                                                        "</MC_FILE_LOG>"

                                'send notification email to business contacts, if no data found in DWH,
                                'provided execution day is not Saturday or Sunday
                                Select Case Weekday(Now)
                                    Case 2, 3, 4, 5, 6  'all weekday Monday thru Friday
                                        objFileEntity.QueueName = lstrClarifyQName
                                        objFileEntity.BusinessContact = lstrBusinessContact
                                        objFileEntity.Body = "No data found in the Data warehouse for " & lstrMC_CODE & ".csv file on Proessing Date : " & Format(Convert.ToDateTime(lstrProcessDate), "MM/dd/yyyy") & "."
                                        objFileEntity.CutTicket = "False"
                                        objFileEntity.SendNotification = "True"
                                        objFileEntity.MCCode = lstrMC_CODE
                                        objFileEntity.ProcessDates = lstrProcessDate
                                        SendErrNotification(objFileEntity)
                                End Select
                            End If
                        End If

                        'append MC_FILE update request XML, with last schedules process date
                        lstrUpdateMCFileReqXml = lstrUpdateMCFileReqXml & _
                                                    "<MC_FILE>" & _
                                                        "<SQ_MC_ID>" & liSQ_MC_ID & "</SQ_MC_ID>" & _
                                                        "<LAST_SCHEDULE_PROCESS_DATE>" & lstrScheduleProcessDate & "</LAST_SCHEDULE_PROCESS_DATE>" & _
                                                    "</MC_FILE>"
                    End Try
                    'Next MC File
                Next


                lstrUpdateMCLogReqXml = lstrUpdateMCLogReqXml & "</MC_FILE_LOGSet></UPDATE_MC_FILE_LOGS_REQUEST>"
                lstrUpdateMCFileReqXml = lstrUpdateMCFileReqXml & "</MC_FILESet></UPDATE_MC_FILE_REQUEST>"

                'call BSMoneyCost component for updation of MC_LOGS table
                objcDataEntity.OutputString = lstrUpdateMCLogReqXml
                lstrUpdateMCLogRespXml = lobjBSMCIMoneyCostService.UpdateMCLogs(objcDataEntity).OutputString

                'call BSMoneyCost component for updation of MC_FILE table, for last schedule process date
                objcDataEntity.OutputString = lstrUpdateMCFileReqXml
                lstrUpdateMCFileRespXml = lobjBSMCIMoneyCostService.UpdateMCFile(objcDataEntity).OutputString
            Else
                Throw New Exception
            End If

            If NotificationMailFlag = False Then        '#### Added By Atul On 18-09-2011
                objFileEntity.QueueName = lstrClarifyQName
                objFileEntity.BusinessContact = lstrBusinessContact
                objFileEntity.SendNotification = "True"
                objFileEntity.ProcessDates = lstrProcessDate
                objFileEntity.Body = ProcessMoneyCostFiles & vbCrLf
                lstrSendErrNotiResult = SendNotification(objFileEntity).OutputString
            End If

            ''******************************************************
            '' New Code inserted for sending Sucess Email in All Condition :- Sanjay
            ''******************************************************
            ''If NotificationMailFlag = False Then
            'objFileEntity.QueueName = lstrClarifyQName
            'objFileEntity.BusinessContact = lstrBusinessContact
            'objFileEntity.SendNotification = "True"
            'objFileEntity.ProcessDates = lstrProcessDate
            'objFileEntity.Body = ProcessMoneyCostFiles & vbCrLf
            'lstrSendErrNotiResult = SendNotification(objFileEntity).OutputString
            ''End If
            ''******************************************************
            ''******************************************************

            'write details in log file
            STLogger.Debug(cMODULE_NAME & lstrMethodName)
            STLogger.Debug(cMODULE_NAME & lstrMethodName & "Exit " & lstrMethodName & "() Method ")
            'If Any Exception in Main Block
        Catch ex As Exception
            If InStr(lstrUpdateMCLogReqXml, "</MC_FILE_LOGSet></UPDATE_MC_FILE_LOGS_REQUEST>") <= 0 Then
                lstrUpdateMCLogReqXml = lstrUpdateMCLogReqXml & "</MC_FILE_LOGSet></UPDATE_MC_FILE_LOGS_REQUEST>"
            End If

            'update MC_LOGS table in MoneyCost DB, using BSMoneyCost component
            'lstrUpdateMCLogRespXml = lobjBSMCIMoneyCostMgr.UpdateMCLogs(lstrUpdateMCLogReqXml)

            llErrNbr = Err.Number
            lstrErrSrc = cCOMPONENT_NAME & "." & cMODULE_NAME & ":" & lstrMethodName & "/" & Err.Source
            lstrErrDesc = Err.Description

            'raise clarify ticket
            If objFileEntity Is Nothing Then
                objFileEntity = New FileInfoEntity
            End If
            objFileEntity.QueueName = lstrClarifyQName
            objFileEntity.BusinessContact = lstrBusinessContact
            objFileEntity.Body = IIf(lstrCommonErrorDetails = "", llErrNbr & " : " & lstrErrDesc & " Source: " & lstrErrSrc, lstrCommonErrorDetails)
            objFileEntity.CutTicket = "False"
            objFileEntity.SendNotification = "False"
            objFileEntity.MCCode = lstrMC_CODE
            objFileEntity.ProcessDates = lstrProcessDate
            lstrSendErrNotiResult = SendErrNotification(objFileEntity).OutputString

            'write error to log file            

            If objxmlErrEntity Is Nothing Then
                objxmlErrEntity = New XmlErrEntity
            End If
            objxmlErrEntity.ErrNbr = Err.Number
            objxmlErrEntity.ErrSource = Err.Source
            objxmlErrEntity.ErrDesc = Err.Description
            STLogger.Debug(cCOMPONENT_NAME & "." & cMODULE_NAME & ".ExecuteServiceFlow " & BuildErrXMLauto(objxmlErrEntity).OutputString)
            objxmlErrEntity = Nothing
            objFileEntity = Nothing
            Throw
        Finally
            'clear object variables and recordset variables from memory
            'If Not IsNothing(lobjMCFileNode) Then
            '    lobjMCFileNode = Nothing
            'End If
            'If Not IsNothing(lobjIndexRateNodeList) Then
            '    lobjIndexRateNodeList = Nothing
            'End If
            'If Not IsNothing(lobjIndexRateNode) Then
            '    lobjIndexRateNode = Nothing
            'End If
            'If Not IsNothing(lobjIndexRateDOM) Then
            '    lobjIndexRateDOM = Nothing
            'End If
            'If Not IsNothing(lobjIndexDataNodeList) Then
            '    lobjIndexDataNodeList = Nothing
            'End If
            'If Not IsNothing(lobjIndexDataDOMXml) Then
            '    lobjIndexDataDOMXml = Nothing
            'End If
            'If Not IsNothing(lobjIndexDataNode) Then
            '    lobjIndexDataNode = Nothing
            'End If
            'If Not IsNothing(lobjCSVInsertNodeList) Then
            '    lobjCSVInsertNodeList = Nothing
            'End If
            'If Not IsNothing(lobjCSVInsertNode) Then
            '    lobjCSVInsertNode = Nothing
            'End If
            'If Not IsNothing(lobjCSVInsertDOMXml) Then
            '    lobjCSVInsertDOMXml = Nothing
            'End If
            If Not IsNothing(lobjBSMCIMoneyCostService) Then
                lobjBSMCIMoneyCostService.Dispose()
                lobjBSMCIMoneyCostService = Nothing
            End If
            If Not IsNothing(objMoneyCostAutoDataClass) Then
                objMoneyCostAutoDataClass.Dispose()
                objMoneyCostAutoDataClass = Nothing
            End If

            'If Not IsNothing(lobjAllMCFileDOM) Then
            '    lobjAllMCFileDOM = Nothing
            'End If

            'If Not IsNothing(lobjMCFileNodeList) Then
            '    lobjMCFileNodeList = Nothing
            'End If
            'If Not IsNothing(objFileEntity) Then
            '    objFileEntity = Nothing
            'End If
            If Not IsNothing(objcDataEntity) Then
                objcDataEntity = Nothing
            End If
            'If Not IsNothing(objxmlErrEntity) Then
            '    objxmlErrEntity = Nothing
            'End If

            If Not IsNothing(lrsCSVRecordset) Then
                lrsCSVRecordset = Nothing
            End If
            If Not IsNothing(lrsCSVDeleteRecordset) Then
                lrsCSVDeleteRecordset = Nothing
            End If
        End Try
    End Sub
    <AutoComplete()> _
    Private Sub SetConfigValue()
        SetLog4Net()
        Try
            ' --------------------------------------------
            ' Check if necessary registry keys are defined
            ' --------------------------------------------
            lstrWorkingDirectory = GetConfigurationKey(cWORKING_DIRECTORY_PATH)
            lstrClarifyQName = GetConfigurationKey(cClarifyQNameKey)
            lstrBackup_Location = GetConfigurationKey(cBACKUP_LOCATION)
            lstrNetwork_Location = GetConfigurationKey(cNETWORK_LOCATION)
            lstrFTP_Location = GetConfigurationKey(cFTP_LOCATION)
            ' Commenting New Path Will be active By Database lstrFTP_Directory = GetConfigurationKey(cFTP_DIRECTORY)
            lstrFTP_User = GetConfigurationKey(cFTP_USER)
            lstrFTP_Password = GetConfigurationKey(cFTP_PASSWORD)

            lstrFTP_LocationForNewDateFormat = GetConfigurationKey(cFTP_LOCATION_NEWDATEFORMAT)
            'Commenting New path will be Active By Database lstrFTP_DirectoryForNewDateFormat = GetConfigurationKey(cFTP_DIRECTORY_NEWDATEFORMAT)
            lstrFTP_UserForNewDateFormat = GetConfigurationKey(cFTP_USER_NEWDATEFORMAT)
            lstrFTP_PasswordForNewDateFormat = GetConfigurationKey(cFTP_PASSWORD_NEWDATEFORMAT)

            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():  Check if necessary registry entries are blank")
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
                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_ExecuteServiceFlow():" & lstrCommonErrorDetails)
            End If

        Catch ex As Exception
            STLogger.Error(ex.Message)
            Throw
        End Try

    End Sub
    <AutoComplete()> _
    Private Sub SetXmlValue(ByRef lstrFTP_DirectoryForNewDateFormat As String, ByRef lstrFTP_Directory As String)

        lobjMCFileNode = lobjMCFileNodeList.Item(liCounter1)
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
        lstrFTP_DirectoryForNewDateFormat = ""
        lstrFTP_Directory = ""
        Try
            'fetch MC file details into local variables, from XML
            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "SQ_MC_ID"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                liSQ_MC_ID = Val(lobjMCFileNode.SelectNodes("SQ_MC_ID").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "MC_CODE"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrMC_CODE = Trim(lobjMCFileNode.SelectNodes("MC_CODE").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "DESCRIPTION"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrDESCRIPTION = Trim(lobjMCFileNode.SelectNodes("DESCRIPTION").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "CURRENCY_CODE"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrCURRENCY_CODE = Trim(lobjMCFileNode.SelectNodes("CURRENCY_CODE").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "START_TIME"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrSTART_TIME = Trim(lobjMCFileNode.SelectNodes("START_TIME").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "END_TIME"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrEND_TIME = Trim(lobjMCFileNode.SelectNodes("END_TIME").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "DAYS_TO_SKIP"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                liDAYS_TO_SKIP = Val(lobjMCFileNode.SelectNodes("DAYS_TO_SKIP").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "CLARIFY_QUEUE"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrClarifyQName = Trim(lobjMCFileNode.SelectNodes("CLARIFY_QUEUE").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "BUSINESS_CONTACT"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrBusinessContact = Trim(lobjMCFileNode.SelectNodes("BUSINESS_CONTACT").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "FREQUENCY"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrFREQUENCY = Trim(lobjMCFileNode.SelectNodes("FREQUENCY").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "FREQUENCY_COUNT"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                liFREQUENCY_COUNT = Trim(lobjMCFileNode.SelectNodes("FREQUENCY_COUNT").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "LAST_SCHEDULE_PROCESS_DATE"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrLAST_SCHEDULE_PROCESS_DATE = Trim(lobjMCFileNode.SelectNodes("LAST_SCHEDULE_PROCESS_DATE").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "MARKET_CLOSED_DWH_CHECK_COUNTER"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                liMARKET_CLOSED_DWH_CHECK_COUNTER = Trim(lobjMCFileNode.SelectNodes("MARKET_CLOSED_DWH_CHECK_COUNTER").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "LAST_UPDATED_IND"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrLAST_UPDATED_IND = Trim(lobjMCFileNode.SelectNodes("LAST_UPDATED_IND").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "DATE_FORMAT"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrDateFormat = Trim(lobjMCFileNode.SelectNodes("DATE_FORMAT").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "DATE_FORMAT_REQUIRED"

            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrDateFormatRequired = Trim(lobjMCFileNode.SelectNodes("DATE_FORMAT_REQUIRED").Item(0).InnerText)
            End If

            'Adding FTP Dirctory By Database column [9th Mar 2010] By Sanjay Srivsatva
            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "FTP_DIRECTORY_NewDateFormat"
            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrFTP_DirectoryForNewDateFormat = Trim(lobjMCFileNode.SelectNodes("FTP_DIRECTORY_NewDateFormat").Item(0).InnerText)
            End If

            objxmlErrEntity.IXMLDOMNode = lobjMCFileNode
            objxmlErrEntity.ElementXPath = "FTP_DIRECTORY"
            If Convert.ToBoolean(IsXMLElementPresent(objxmlErrEntity).OutputString) Then
                lstrFTP_Directory = Trim(lobjMCFileNode.SelectNodes("FTP_DIRECTORY").Item(0).InnerText)
            End If
        Catch ex As Exception
            Throw
        End Try

    End Sub
    <AutoComplete()> _
    Private Sub SetDateProcessFlag()
        Try
            ' -------------------------------------------------
            ' Process date = System Date - MC_FILE.Days_TO_SKIP
            ' -------------------------------------------------
            lstrProcessDate = DateAdd("d", -liDAYS_TO_SKIP, Now.Date)

            'get the previous processdate as well to get the previous day's adder rates.
            lstrProcessDate_PrevDay = DateAdd("d", -(liDAYS_TO_SKIP + 1), Now.Date)

            'The week end check of the previous day is based on the current process date.
            'dont change the order of the cases.
            Select Case Weekday(lstrProcessDate)
                Case 7, 1 'Saturday , Sunday
                    lstrProcessDate_PrevDay = DateAdd("d", -(liDAYS_TO_SKIP + 1 + 2), Now.Date)
            End Select

            Select Case Weekday(lstrProcessDate)
                Case 7, 1 'Saturday , Sunday
                    lstrProcessDate = DateAdd("d", -(liDAYS_TO_SKIP + 2), Now.Date)
            End Select


            lstrScheduleProcessDate = DateAdd(lstrFREQUENCY, liFREQUENCY_COUNT, lstrLAST_SCHEDULE_PROCESS_DATE)
            Select Case Weekday(lstrScheduleProcessDate)
                Case 7, 1 'Saturday , Sunday
                    lstrScheduleProcessDate = DateAdd(lstrFREQUENCY, liFREQUENCY_COUNT + 2, lstrLAST_SCHEDULE_PROCESS_DATE)
            End Select

            If DateDiff("d", lstrProcessDate, lstrScheduleProcessDate) = 0 And Weekday(Now) <> 1 And Weekday(Now) <> 7 Then
                lbServiceRunFlag = True
            End If

        Catch ex As Exception
            Throw
        End Try

    End Sub

#End Region

End Class
