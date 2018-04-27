Imports System.Reflection
Imports System.Runtime.InteropServices
Imports BSMoneyCostDL
Imports BSMoneyCostEntity
Imports System.EnterpriseServices
Public Interface IMoneyCostUISvc
    Function Ping() As String
    Function Test() As String
    Function UpdateMCDetails(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function UpdateMCLogs(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function UpdateMCFile(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function GetMCFiles(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function GetMCFileDetails(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function GetAllMCFiles() As cDataEntity
    Function GetIndexRates(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function GetMCSecurity(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function GetAllMCFilesForMCRUN() As cDataEntity

    '555555555555555555555555
    Function GetTreasuryAssessmentData(ByVal lobjEntity As cDataEntity) As DataSet
    Function GetTreasuryDetails() As cDataEntity
    Function GetCostTypes() As cDataEntity

End Interface
<JustInTimeActivation(), _
 EventTrackingEnabled(), _
 ClassInterface(ClassInterfaceType.None), _
 Transaction(TransactionOption.NotSupported, Isolation:=TransactionIsolationLevel.Serializable, Timeout:=120), _
 ComponentAccessControl(True)> _
Public Class cMoneyCostUISvc
    Inherits ServicedComponent
    Implements IMoneyCostUISvc
    Private Const cMODULE_NAME As String = "cMoneyCostUIsvc"
    Dim STLogger As log4net.ILog
    '================================================================
    'METHOD  :  Ping
    'PURPOSE :  Allows component to be pinged to verify it can be
    '           instantiated
    'PARMS   :  none
    'RETURN  :  String with date and time
    '================================================================
    <AutoComplete()> _
    Public Function Ping() As String Implements IMoneyCostUISvc.Ping
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
    Public Function Test() As String Implements IMoneyCostUISvc.Test

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

    <AutoComplete()> _
    Public Function UpdateMCDetails(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostUISvc.UpdateMCDetails
        Dim lobjcMoneyCostMgr As New cMoneyCostMgr    'object of cMoneyCostMgr class module, to access its methods     
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug(lobjEntity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            Return lobjcMoneyCostMgr.UpdateMCDetails(lobjEntity)

            STLogger.Debug(UpdateMCDetails.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjcMoneyCostMgr) Then
                'lobjcMoneyCostMgr.Dispose()
                lobjcMoneyCostMgr = Nothing

            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

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
    <AutoComplete()> _
    Public Function UpdateMCLogs(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostUISvc.UpdateMCLogs
        Dim lobjcMoneyCostMgr As New cMoneyCostMgr    'object of cMoneyCostMgr class module, to access its methods      
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug(lobjEntity.OutputString & " " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            'call UpdateMCLogs() method of cMoneyCostMgr class module to update MC_Logs table,
            'on execution of BSSTMoneyCost service, and return final output XML          
            Return lobjcMoneyCostMgr.UpdateMCLogs(lobjEntity)
            STLogger.Debug(UpdateMCLogs.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjcMoneyCostMgr) Then
                'lobjcMoneyCostMgr.Dispose()
                lobjcMoneyCostMgr = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

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
    <AutoComplete()> _
    Public Function UpdateMCFile(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostUISvc.UpdateMCFile
        Dim lobjcMoneyCostMgr As New cMoneyCostMgr    'object of cMoneyCostMgr class module, to access its methods       
        SetLog4Net()
        Try

            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug(lobjEntity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            'call UpdateMCFile() method of cMoneyCostMgr class module to update MC_File table,           
            Return lobjcMoneyCostMgr.UpdateMCFile(lobjEntity)
            STLogger.Debug(UpdateMCFile.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjcMoneyCostMgr) Then
                'lobjcMoneyCostMgr.Dispose()
                lobjcMoneyCostMgr = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

    End Function

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
    <AutoComplete()> _
    Public Function GetMCFiles(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostUISvc.GetMCFiles

        Dim lobjcMoneyCostService As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods
        'Dim lobjBSLDAPIService2 As New cMoneyCostLDAPSvc
        Dim lobjBSLDAPIService2 As New BSMoneyCostBL.cMoneyCostLDAP   'object of BSLDAP component for getting logged-in user's details
        Dim lobjRequestXmlDOM As New Xml.XmlDocument        'DOM object variable for loading Request XML
        Dim lobjUserDetailXmlDOM As New Xml.XmlDocument        'DOM object variable for loading Response XML from LDAP
        Dim lstrUSER_GESSOUID As String = ""                   'to store logged-in user's GESSOUID from Request XML
        Dim lstrUSER_SSOID As String = ""                   'to store logged-in user's SSO ID, retrieved from LDAP
        Dim lstrBSLDAPRequestXML As String                   'to store Request XML for LDAP component
        Dim lstrBSLDAPResponseXML As String                   'to store Response XML from LDAP component
        Dim lstrMCFileRequestXML As String                   'to store dynamically build Request XML for cMoneyCostService.GetMCFile() method        
        SetLog4Net()
        Try

            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("validate Request XML: " & lobjEntity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            'validate, if the Input Request XML is well-formed
            Try
                lobjRequestXmlDOM.LoadXml(lobjEntity.OutputString)
            Catch ex As Exception
                STLogger.Error("Error loading Input XML to GetMCFiles(). " & ex.Message, ex)
                Throw
            End Try
            'If logged-in User SSO ID found in Input XML, fetch the value in local variable
            If IsXMLElementPresent(lobjRequestXmlDOM.DocumentElement, "USER_GESSOUID") Then
                If Trim(lobjRequestXmlDOM.GetElementsByTagName("USER_GESSOUID").Item(0).InnerText) <> "" Then
                    lstrUSER_GESSOUID = lobjRequestXmlDOM.GetElementsByTagName("USER_GESSOUID").Item(0).InnerText
                End If
            End If
            'build Request XML for BSLDAP component for fetching logged-in User's details
            lstrBSLDAPRequestXML = "<LDAPSearch FETCH='uid,givenname,sn' ou='geworker'>" & _
                                        "<LDAP_ATTRIB NAME='gessouid' OPERATOR='EQ'>" & UCase(lstrUSER_GESSOUID) & "</LDAP_ATTRIB>" & _
                                    "</LDAPSearch>"

            'call BSLDAP component to fetch logged-in user's details
            'lstrBSLDAPResponseXML = lobjBSLDAPIService2.GetUserDetailsByAttributes(lstrBSLDAPRequestXML)
            lstrBSLDAPResponseXML = lobjBSLDAPIService2.GetUserDetailsByAttributes("(gessouid=" & UCase(lstrUSER_GESSOUID) & ")")

            'Dim lobj As New cMoneyCostLDAP

            'lstrBSLDAPResponseXML = lobj.GetUserDetailsByAttributes("(gessouid=" & UCase(lstrUSER_GESSOUID) & ")")


            'validate, if the Input Request XML is well-formed
            Try
                lobjUserDetailXmlDOM.LoadXml(lstrBSLDAPResponseXML)
            Catch ex As Exception
                STLogger.Error("Error loading Input XML to GetMCFiles(). " & ex.Message, ex)
                Throw
            End Try
            If IsXMLElementPresent(lobjUserDetailXmlDOM.DocumentElement, "//uid") And lobjUserDetailXmlDOM.GetElementsByTagName("uid").Count > 0 Then
                lstrUSER_SSOID = Trim(lobjUserDetailXmlDOM.GetElementsByTagName("uid").Item(0).InnerText)
            Else
                STLogger.Error(cINVALID_PARMS & "User's SSO ID could not found from LDAP.")
            End If

            lstrMCFileRequestXML = "<MC_FILES_REQUEST><USER_SSOID>" & lstrUSER_SSOID & "</USER_SSOID></MC_FILES_REQUEST>"

            'call GetMCFiles() method of cMoneyCostService class module to get logged-in user's Money Cost File List

            lobjEntity.OutputString = lstrMCFileRequestXML
            lobjEntity = lobjcMoneyCostService.GetMCFiles(lobjEntity)

            'return final output XML
            lobjEntity.OutputString = "<USER_MC_FILE_RESPONSE>" & lstrBSLDAPResponseXML & lobjEntity.OutputString & "</USER_MC_FILE_RESPONSE>"
            Return lobjEntity

            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjBSLDAPIService2) Then
                lobjBSLDAPIService2 = Nothing
            End If
            If Not IsNothing(lobjcMoneyCostService) Then
                'lobjcMoneyCostService.Dispose()
                lobjcMoneyCostService = Nothing
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
    'METHOD  : GetMCSecurity
    'PURPOSE : Wrapper method to fetch the details of the selected
    '          Money Cost Security, selected by the logged-in user
    'PARMS   :
    '          astrGetMCFileDetailsXML [String] = Input paramters
    '          to retrieve the information for the selected MC File
    '          details.
    '          Sample XML structure:
    '================================================================
    <AutoComplete()> _
    Public Function GetMCSecurity(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostUISvc.GetMCSecurity
        Dim lobjcMoneyCostService As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods    
        SetLog4Net()
        Try

            STLogger.Debug(lobjEntity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            'call GetMCFileDetails() method of cMoneyCostService class module to get selected Money Cost File details,
            'and return final outpur Response XML
            Return lobjcMoneyCostService.GetMCSecurity(lobjEntity)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjcMoneyCostService) Then
                'lobjcMoneyCostService.Dispose()
                lobjcMoneyCostService = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

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
    '================================================================
    <AutoComplete()> _
    Public Function GetMCFileDetails(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostUISvc.GetMCFileDetails
        Dim lobjcMoneyCostService As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods    
        SetLog4Net()
        Try

            STLogger.Debug(lobjEntity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            'call GetMCFileDetails() method of cMoneyCostService class module to get selected Money Cost File details,
            'and return final outpur Response XML
            Return lobjcMoneyCostService.GetMCFileDetails(lobjEntity)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjcMoneyCostService) Then
                'lobjcMoneyCostService.Dispose()
                lobjcMoneyCostService = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

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
    <AutoComplete()> _
    Public Function GetAllMCFiles() As cDataEntity Implements IMoneyCostUISvc.GetAllMCFiles
        Dim lobjEntity As cDataEntity = Nothing
        Dim lobjcMoneyCostService As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'call GetAllMCFiles() method of cMoneyCostService class module
            'to get all Money Cost File List, to return response XML
            lobjEntity = lobjcMoneyCostService.GetAllMCFiles()
            Return lobjEntity
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjcMoneyCostService) Then
                'lobjcMoneyCostService.Dispose()
                lobjcMoneyCostService = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

    End Function

    '5555555555555555555555
    ' Treasury Assessment method will be used to fetch Adder from Moneycost DB
    <AutoComplete()> _
   Public Function GetTreasuryAssessmentData(ByVal lobjEntity As cDataEntity) As DataSet Implements IMoneyCostUISvc.GetTreasuryAssessmentData
        Dim lobjcMoneyCostService As New cMoneyCostService
        Return lobjcMoneyCostService.GetTreasuryAssessmentData(lobjEntity)

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
    <AutoComplete()> _
    Public Function GetIndexRates(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostUISvc.GetIndexRates
        Dim lobjcMoneyCostService As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            'call GetIndexRates() method of cMoneyCostService class module
            'to get Index Rates for particular Money Cost File, to return response XML
            Return lobjcMoneyCostService.GetIndexRates(lobjEntity)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjcMoneyCostService) Then
                'lobjcMoneyCostService.Dispose()
                lobjcMoneyCostService = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If

        End Try

    End Function

    Public Sub New()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    <AutoComplete()> _
   Public Function GetAllMCFilesForMCRun() As cDataEntity Implements IMoneyCostUISvc.GetAllMCFilesForMCRUN
        Dim lobjEntity As cDataEntity = Nothing
        Dim lobjcMoneyCostService As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'call GetAllMCFiles() method of cMoneyCostService class module
            'to get all Money Cost File List, to return response XML
            lobjEntity = lobjcMoneyCostService.GetAllMCFilesForMCRUN()
            Return lobjEntity
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjcMoneyCostService) Then
                'lobjcMoneyCostService.Dispose()
                lobjcMoneyCostService = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

    End Function

    ' Added for Treasury Assessment
    <AutoComplete()> _
   Public Function GetTreasuryDetails() As cDataEntity Implements IMoneyCostUISvc.GetTreasuryDetails
        Dim lobjEntity As cDataEntity = Nothing
        Dim lobjcMoneyCostService As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'call GetTreasuryDetails() method of cMoneyCostService class module
            'to get treasury details, to return response XML
            lobjEntity = lobjcMoneyCostService.GetTreasuryDetails()
            Return lobjEntity
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjcMoneyCostService) Then
                'lobjcMoneyCostService.Dispose()
                lobjcMoneyCostService = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

    End Function


    <AutoComplete()> _
    Public Function GetCostTypes() As cDataEntity Implements IMoneyCostUISvc.GetCostTypes
        Dim lobjEntity As cDataEntity = Nothing
        Dim lobjcMoneyCostService As New cMoneyCostService    'object of cMoneyCostService class module, to access its methods
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'call GetCostTypes() method of cMoneyCostService class module
            'to get GetCostTypes
            lobjEntity = lobjcMoneyCostService.GetCostTypes()
            Return lobjEntity
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not IsNothing(lobjcMoneyCostService) Then
                'lobjcMoneyCostService.Dispose()
                lobjcMoneyCostService = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try
    End Function
End Class