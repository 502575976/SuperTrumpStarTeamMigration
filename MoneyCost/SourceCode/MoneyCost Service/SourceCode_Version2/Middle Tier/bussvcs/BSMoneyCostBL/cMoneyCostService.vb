Imports System.Reflection
Imports System.Runtime.InteropServices
Imports BSMoneyCostDL
Imports BSMoneyCostEntity
Imports System.EnterpriseServices
Public Interface IMoneyCostService
    Function Ping() As String
    Function Test() As String
    Function GetMCFiles(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function GetMCFileDetails(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function GetAllMCFiles() As cDataEntity
    Function GetIndexRates(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function GetMCSecurity(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function GetAllMCFilesForMCRUN() As cDataEntity
    '5555555555555
    Function GetTreasuryAssessmentData(ByVal lobjEntity As cDataEntity) As DataSet
    Function GetTreasuryDetails() As cDataEntity
    Function GetCostTypes() As cDataEntity
End Interface
<JustInTimeActivation(), _
 EventTrackingEnabled(), _
 ClassInterface(ClassInterfaceType.None), _
 Transaction(TransactionOption.NotSupported, Isolation:=TransactionIsolationLevel.Serializable, Timeout:=120), _
 ComponentAccessControl(True)> _
Public Class cMoneyCostService
    Inherits ServicedComponent
    Implements IMoneyCostService
    Private Const cMODULE_NAME As String = "cMoneyCostService"
    Dim STLogger As log4net.ILog
    '================================================================

    '================================================================
    'METHOD  :  Ping
    'PURPOSE :  Allows component to be pinged to verify it can be
    '           instantiated
    'PARMS   :  none
    'RETURN  :  String with date and time
    '================================================================
    <AutoComplete()> _
    Public Function Ping() As String Implements IMoneyCostService.Ping
        Try
            Return "Ping request to " & cCOMPONENT_NAME & "." & cMODULE_NAME & " returned at " & Format(Now, "mm/dd/yyyy Hh:Nn:Ss AM/PM") & " server time."
        Catch ex As Exception
            Throw
        End Try
    End Function

    '================================================================
    'METHOD  : Test
    'PURPOSE : Returns a string that indicates that the component
    '          can connect to the database and the registry.
    'PARMS   : NONE
    'RETURN  : String
    '================================================================
    <AutoComplete()> _
    Public Function Test() As String Implements IMoneyCostService.Test
        Dim lobjDataClass As New cDataClass
        Dim lrsTest As New DataTable
        Try
            'Execute the Test SQL statement which returns a count of the records
            lrsTest = lobjDataClass.Test
            'Return the total records
            Return "Retrieved " & lrsTest.Rows(0)(0).Value & " records."
        Catch ex As Exception
            Return vbNullString
            Throw
        Finally
            If Not IsNothing(lrsTest) Then
                lrsTest = Nothing
            End If
            If Not IsNothing(lobjDataClass) Then
                'lobjDataClass.Dispose()
                lobjDataClass = Nothing
            End If
        End Try
    End Function
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
    <AutoComplete()> _
    Public Function GetMCFiles(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostService.GetMCFiles

        Dim lobjDataClass As New cDataClass                       'object variable to access cDataClass method(s)
        Dim lobjRequestXmlDOM As New Xml.XmlDocument                    'object variable for DOM, to load Request XML
        Dim lstrUSER_SSOID As String = ""                               'to store logged-in User Gessouid, from Request XML        
        Dim lstrResult As String                               'to store final Response XML
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Validate Request XML: " & lobjEntity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            'Validate if the Request XML is well-formed
            Try
                lobjRequestXmlDOM.LoadXml(lobjEntity.OutputString)
            Catch ex As Exception
                STLogger.Error("Error loading Request XML to GetMCFiles(). " & ex.Message, ex)
                Throw
            End Try

            'If logged-in User SSO ID found in Input XML, fetch the value in local variable

            If IsXMLElementPresent(lobjRequestXmlDOM.DocumentElement, "USER_SSOID") Then
                If Trim(lobjRequestXmlDOM.GetElementsByTagName("USER_SSOID").Item(0).InnerText) <> "" Then
                    lstrUSER_SSOID = lobjRequestXmlDOM.GetElementsByTagName("USER_SSOID").Item(0).InnerText
                Else
                    STLogger.Error(cINVALID_PARMS & "User's SSO ID not specified.")
                End If
            Else
                STLogger.Error(cINVALID_PARMS & "User's SSO ID not specified.")
            End If

            'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
            'method of Recordset Utilities component to form the Output XML, in local variable

            'lstrResult = lobjRSUtils.RSToXML("MC_FILE", lobjDataClass.Execute(ecGetMCFiles, ecRSExecuteRecords, lstrUSER_SSOID))            
            lobjEntity.UserSSOID = lstrUSER_SSOID
            lstrResult = DsToXML(lobjDataClass.GetMCFiles(lobjEntity).OutputDataSet)

            'Return the XML as output
            lobjEntity.OutputString = "<MC_FILE_RESPONSE>" & lstrResult & "</MC_FILE_RESPONSE>"
            Return lobjEntity
            STLogger.Debug(GetMCFiles.OutputString & ": " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear all local object variables from memory
            If Not IsNothing(lobjDataClass) Then
                'lobjDataClass.Dispose()
                lobjDataClass = Nothing
            End If
            If Not IsNothing(lobjRequestXmlDOM) Then
                lobjRequestXmlDOM = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

    End Function

    '================================================================
    'METHOD  : GetMCSeucurity
    'PURPOSE : Returns details of selected Money Cost file as per given
    '          filter criteria. Filter Criteria will be Money Cost File
    '          Id. Data will be in ascending order of column position
    'PARMS   :
    '          astrGetMCFileDetailsXML [String] = Filter criteria
    '          Sample XML structure:
    '           <MC_FILE_DETAIL_REQUEST>
    '               <SQ_MC_ID>1</SQ_MC_ID>
    '           </MC_FILE_DETAIL_REQUEST>
    '================================================================
    <AutoComplete()> _
    Public Function GetMCSecurity(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostService.GetMCSecurity

        Dim lobjDataClass As New cDataClass                       'object variable to access cDataClass method(s)
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
            GetMCSecurity = lobjDataClass.GetMCSecurity(lobjEntity)

            STLogger.Debug(GetMCSecurity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear all local object variables from memory
            If Not IsNothing(lobjDataClass) Then
                'lobjDataClass.Dispose()
                lobjDataClass = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

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
    <AutoComplete()> _
    Public Function GetMCFileDetails(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostService.GetMCFileDetails

        Dim lobjDataClass As New cDataClass                       'object variable to access cDataClass method(s)
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
            GetMCFileDetails = lobjDataClass.GetMCFileDetails(lobjEntity)

            STLogger.Debug(GetMCFileDetails.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear all local object variables from memory
            If Not IsNothing(lobjDataClass) Then
                'lobjDataClass.Dispose()
                lobjDataClass = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

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
    <AutoComplete()> _
    Public Function GetAllMCFiles() As cDataEntity Implements IMoneyCostService.GetAllMCFiles
        Dim lobjEntity As cDataEntity = Nothing                              ' to store passing argument to data layer       
        Dim lobjDataClass As New cDataClass                       'object variable to access cDataClass method(s)
        Dim lstrResult As String                               'to store final Response XML
        Dim cdataEntity As New cDataEntity

        SetLog4Net()

        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
            'method of Recordset Utilities component to form the Output XML, in local variable
            lobjEntity = New cDataEntity
            lstrResult = DsToXML(lobjDataClass.GetAllMCFiles(lobjEntity).OutputDataSet)

            'Return the XML as output
            cdataEntity.OutputString = "<MC_FILE_RESPONSE>" & lstrResult & "</MC_FILE_RESPONSE>"
            Return cdataEntity

            STLogger.Debug(GetAllMCFiles.OutputString & " " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear all local object variables from memory
            If Not IsNothing(lobjDataClass) Then
                'lobjDataClass.Dispose()
                lobjDataClass = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try

    End Function
    '555555555555555555
    ' Treasury Assessment method will be used to fetch Adder from Moneycost DB
    <AutoComplete()> _
       Public Function GetTreasuryAssessmentData(ByVal lobjEntity As cDataEntity) As DataSet Implements IMoneyCostService.GetTreasuryAssessmentData
        Dim lobjDataClass As New cDataClass
        Return lobjDataClass.GetTreasuryAssessmentData(lobjEntity)
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



    <AutoComplete()> _
    Public Function GetIndexRates(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostService.GetIndexRates
        Dim lobjDataClass As New cDataClass                       'object variable to access cDataClass method(s)
        Dim lobjRequestXmlDOM As New Xml.XmlDocument                    'object variable for DOM, to load Request XML
        Dim liSQ_MC_ID As Integer                              'to store Money Cost file ID, from Request XML
        Dim lstrPROCESS_DATE As String = ""                               'to store MC File Process Date, from Request XML
        Dim lstrResult As String                               'to store Response XML
        Dim cdataEntity As New cDataEntity
        SetLog4Net()

        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("validate Request xml: " & lobjEntity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'validate, if Request xml is well-formed
            Try
                lobjRequestXmlDOM.LoadXml(lobjEntity.OutputString)
            Catch ex As Exception
                STLogger.Error("Error loading Input Request XML to GetIndexRates(). " & ex.Message, ex)
                Throw
            End Try

            'check, whether MC File Id found, in Request XML
            If IsXMLElementPresent(lobjRequestXmlDOM.DocumentElement, "SQ_MC_ID") Then
                If lobjRequestXmlDOM.GetElementsByTagName("SQ_MC_ID").Count > 0 Then
                    liSQ_MC_ID = Val(lobjRequestXmlDOM.GetElementsByTagName("SQ_MC_ID").Item(0).InnerText)
                Else
                    STLogger.Error(cINVALID_PARMS & "Money Cost File ID not specified.")
                End If
            Else
                STLogger.Error(cINVALID_PARMS & "Money Cost File ID not specified.")
            End If

            'check, whether MC File Process Date found, in Request XML
            If IsXMLElementPresent(lobjRequestXmlDOM.DocumentElement, "PROCESS_DATE") Then
                If lobjRequestXmlDOM.GetElementsByTagName("PROCESS_DATE").Count > 0 Then
                    lstrPROCESS_DATE = Trim(lobjRequestXmlDOM.GetElementsByTagName("PROCESS_DATE").Item(0).InnerText)
                Else
                    'Err.Raise(cINVALID_PARMS, , "Money Cost File Process Date not specified.")
                    STLogger.Error(cINVALID_PARMS & "Money Cost File Process Date not specified.")
                End If
            Else
                STLogger.Error(cINVALID_PARMS & "Money Cost File Process Date not specified.")
            End If

            'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
            'method of Recordset Utilities component to form the Output XML, in local variable           
            cdataEntity.MoneyCostID = liSQ_MC_ID
            cdataEntity.ProcessDate = lstrPROCESS_DATE
            'lstrResult = lobjRSUtils.RSToXML("INDEX_RATE", lobjDataClass.Execute(ecGetIndexRates, ecRSExecuteRecords, liSQ_MC_ID, lstrPROCESS_DATE))
            lstrResult = DsToXML(lobjDataClass.GetIndexRates(cdataEntity).OutputDataSet)

            'Return the XML as output
            cdataEntity.OutputString = "<INDEX_RATE_RESPONSE>" & lstrResult & "</INDEX_RATE_RESPONSE>"
            Return cdataEntity
            STLogger.Debug(GetIndexRates.OutputString & " " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear all local object variables from memory
            If Not IsNothing(lobjDataClass) Then
                'lobjDataClass.Dispose()
                lobjDataClass = Nothing
            End If
            If Not IsNothing(lobjRequestXmlDOM) Then
                lobjRequestXmlDOM = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try
    End Function
    <AutoComplete()> _
    Public Function GetAllMCFilesForMCRUN() As cDataEntity Implements IMoneyCostService.GetAllMCFilesForMCRUN
        Dim lobjEntity As cDataEntity = Nothing                              ' to store passing argument to data layer       
        Dim lobjDataClass As New cDataClass                       'object variable to access cDataClass method(s)
        Dim lstrResult As String                               'to store final Response XML
        Dim cdataEntity As New cDataEntity

        SetLog4Net()

        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
            'method of Recordset Utilities component to form the Output XML, in local variable
            lobjEntity = New cDataEntity
            lstrResult = DsToXML(lobjDataClass.GetAllMCFilesForMCRUN(lobjEntity).OutputDataSet)

            'Return the XML as output
            cdataEntity.OutputString = "<MC_FILE_RESPONSE>" & lstrResult & "</MC_FILE_RESPONSE>"
            Return cdataEntity

            STLogger.Debug(GetAllMCFiles.OutputString & " " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear all local object variables from memory
            If Not IsNothing(lobjDataClass) Then
                'lobjDataClass.Dispose()
                lobjDataClass = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try

    End Function

    ' Added for Treasury Assessment
    <AutoComplete()> _
    Public Function GetTreasuryDetails() As cDataEntity Implements IMoneyCostService.GetTreasuryDetails
        Dim lobjEntity As cDataEntity = Nothing                              ' to store passing argument to data layer       
        Dim lobjDataClass As New cDataClass                       'object variable to access cDataClass method(s)
        Dim lstrResult As String                               'to store final Response XML
        Dim cdataEntity As New cDataEntity

        SetLog4Net()

        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
            'method of Recordset Utilities component to form the Output XML, in local variable
            lobjEntity = New cDataEntity
            lstrResult = DsToXML(lobjDataClass.GetTreasuryDetails(lobjEntity).OutputDataSet)

            'Return the XML as output
            cdataEntity.OutputString = "<MC_FILE_RESPONSE>" & lstrResult & "</MC_FILE_RESPONSE>"
            Return cdataEntity

            STLogger.Debug(GetAllMCFiles.OutputString & " " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear all local object variables from memory
            If Not IsNothing(lobjDataClass) Then
                'lobjDataClass.Dispose()
                lobjDataClass = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try

    End Function

    ' Added for Treasury Assessment
    '<AutoComplete()> _
    'Public Function GetCostTypes() As cDataEntity Implements IMoneyCostService.GetCostTypes
    '     Dim lobjEntity As cDataEntity = Nothing
    '     Dim lobjDataClass As New cDataClass                       'object variable to access cDataClass method(s)
    '     SetLog4Net()
    '     Try
    '         STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

    '         'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
    '         GetCostTypes = lobjDataClass.GetCostTypes(lobjEntity)

    '         STLogger.Debug(GetCostTypes.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
    '         STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
    '     Catch ex As Exception
    '         STLogger.Error(ex.Message, ex)
    '         Throw
    '     Finally
    '         'clear all local object variables from memory
    '         If Not IsNothing(lobjDataClass) Then
    '             'lobjDataClass.Dispose()
    '             lobjDataClass = Nothing
    '         End If
    '     End Try

    ' End Function


    <AutoComplete()> _
    Public Function GetCostTypes() As cDataEntity Implements IMoneyCostService.GetCostTypes
        Dim lobjEntity As cDataEntity = Nothing                              ' to store passing argument to data layer       
        Dim lobjDataClass As New cDataClass                       'object variable to access cDataClass method(s)
        Dim dsResult As DataSet
        Dim cdataEntity As New cDataEntity

        SetLog4Net()

        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'call Execute method of cDataClass to fetch required dataset and send recordset to RSToXML
            'method of Recordset Utilities component to form the Output XML, in local variable
            lobjEntity = New cDataEntity
            dsResult = lobjDataClass.GetCostTypes(lobjEntity).OutputDataSet
            cdataEntity.OutputDataSet = dsResult
            'Return the XML as output
            Return cdataEntity

            STLogger.Debug(GetAllMCFiles.OutputString & " " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear all local object variables from memory
            If Not IsNothing(lobjDataClass) Then
                'lobjDataClass.Dispose()
                lobjDataClass = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try

    End Function
End Class
