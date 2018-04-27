Imports System.EnterpriseServices
Imports System.Runtime.InteropServices
Imports BSMoneyCostEntity
Imports BSMoneyCostDL
Imports System.Reflection
Public Interface IMoneyCostMgr
    Function UpdateMCDetails(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function UpdateMCLogs(ByVal lobjEntity As cDataEntity) As cDataEntity
    Function UpdateMCFile(ByVal lobjEntity As cDataEntity) As cDataEntity
End Interface
<JustInTimeActivation(), _
 EventTrackingEnabled(), _
 ClassInterface(ClassInterfaceType.None), _
 Transaction(TransactionOption.Required, Isolation:=TransactionIsolationLevel.Serializable, Timeout:=120), _
 ComponentAccessControl(True)> _
Public Class cMoneyCostMgr
    Inherits ServicedComponent
    Implements IMoneyCostMgr
    Dim STLogger As log4net.ILog
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
    '================================================================
    'METHOD  : UpdateMCDetails
    'PURPOSE : Wrapper method to update the details of selected Money
    '          Cost file.    
    '================================================================
    <AutoComplete()> _
    Public Function UpdateMCDetails(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostMgr.UpdateMCDetails
        SetLog4Net()
        Dim lobjcDataClass As New cDataClass       'object of cDataClass to access its methods
        Dim lstrAction As String = ""
        Try
            If lobjEntity.ActionID = 1 Then
                lstrAction = "Index Rates "
            ElseIf lobjEntity.ActionID = 2 Then
                lstrAction = "MC File "
            ElseIf lobjEntity.ActionID = 3 Then
                lstrAction = "MC Security "
            End If
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Update " & lstrAction & " Query : " & lobjEntity.CommonSQL & " " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            Call lobjcDataClass.UpdateMCDetails(lobjEntity)

            STLogger.Debug("" & lstrAction & "Updated " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            UpdateMCDetails = lobjEntity
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear object variables from memory
            If Not IsNothing(lobjcDataClass) Then
                'lobjcDataClass.Dispose()
                lobjcDataClass = Nothing
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
    Public Function UpdateMCLogs(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostMgr.UpdateMCLogs

        Dim lobjcDataClass As New cDataClass       'object of cDataClass to access its methods
        Dim lobjRequestXmlDOM As New Xml.XmlDocument    'object of XML DOM to load input Request XML
        Dim lobjMCFileLogsNdLst As Xml.XmlNodeList = Nothing      'to store node list of MC_FILE_LOG node set
        Dim lobjMCFileLogsElem As Xml.XmlElement = Nothing      'to store single node of MC_FILE_LOG node set
        Dim lstrBatchUpdateQry As String               'to store dynamically batch update query
        Dim liCounter As Integer              'to store loop counter
        Dim liSQ_MC_ID As Integer              'to store Money Cost file Id
        Dim lstrDATE_START As String               'to store start date/time of money cost file processing
        Dim lstrDATE_END As String               'to store end date/time of money cost file processed
        Dim lbySTATUS As Byte                 'to store status of money cost file processing (0: Success; 1: Failure)
        Dim lstrDETAILS As String               'to store details of money cost file processing result
        SetLog4Net()
        Try

            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Validate Input Request: " & lobjEntity.OutputString & " " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            Try
                lobjRequestXmlDOM.LoadXml(lobjEntity.OutputString)
            Catch ex As Exception
                STLogger.Error("Error loading Input Request XML to UpdateMCDetails(). " & ex.Message, ex)
            End Try


            'get all node list of MC_FILE_DETAIL node set
            lobjMCFileLogsNdLst = lobjRequestXmlDOM.SelectNodes("/UPDATE_MC_FILE_LOGS_REQUEST/MC_FILE_LOGSet/MC_FILE_LOG")

            lstrBatchUpdateQry = ""

            'loop for each individual MC_FILE_DETAIL node set
            For liCounter = 0 To lobjMCFileLogsNdLst.Count - 1

                'get single node set of MC_FILE_DETAIL
                lobjMCFileLogsElem = lobjMCFileLogsNdLst.Item(liCounter)

                're-initialize local variables
                liSQ_MC_ID = 0
                lstrDATE_START = ""
                lstrDATE_END = ""
                lbySTATUS = 1
                lstrDETAILS = ""

                'check, if SQ_MC_ID found in Input XML, store in local variable
                If IsXMLElementPresent(lobjMCFileLogsElem, "SQ_MC_ID") Then
                    If Trim(lobjMCFileLogsElem.SelectNodes("SQ_MC_ID").Item(0).InnerText) <> "" Then
                        liSQ_MC_ID = Val(lobjMCFileLogsElem.SelectNodes("SQ_MC_ID").Item(0).InnerText)
                    Else
                        STLogger.Error(cINVALID_PARMS & " SQ_MC_ID is not specified.")
                    End If
                Else
                    STLogger.Error(cINVALID_PARMS & " SQ_MC_ID is not specified.")
                End If

                'check, if DATE_START found in Input XML, store in local variable
                If IsXMLElementPresent(lobjMCFileLogsElem, "DATE_START") Then
                    If Trim(lobjMCFileLogsElem.SelectNodes("DATE_START").Item(0).InnerText) <> "" Then
                        lstrDATE_START = Trim(lobjMCFileLogsElem.SelectNodes("DATE_START").Item(0).InnerText)
                    Else
                        STLogger.Error(cINVALID_PARMS & " Date Start is not specified.")
                    End If
                Else
                    STLogger.Error(cINVALID_PARMS & " Date Start is not specified.")
                End If

                'check, if DATE_START found in Input XML, store in local variable
                If IsXMLElementPresent(lobjMCFileLogsElem, "DATE_END") Then
                    If Trim(lobjMCFileLogsElem.SelectNodes("DATE_END").Item(0).InnerText) <> "" Then
                        lstrDATE_END = Trim(lobjMCFileLogsElem.SelectNodes("DATE_END").Item(0).InnerText)
                    Else
                        STLogger.Error(cINVALID_PARMS & " Date End is not specified.")
                    End If
                Else
                    STLogger.Error(cINVALID_PARMS & " Date End is not specified.")
                End If

                'check, if STATUS found in Input XML, store in local variable
                If IsXMLElementPresent(lobjMCFileLogsElem, "STATUS") Then
                    If Trim(lobjMCFileLogsElem.SelectNodes("STATUS").Item(0).InnerText) <> "" Then
                        lbySTATUS = Trim(lobjMCFileLogsElem.SelectNodes("STATUS").Item(0).InnerText)
                    Else
                        STLogger.Error(cINVALID_PARMS & " Status is not specified.")
                    End If
                Else
                    STLogger.Error(cINVALID_PARMS & " Status is not specified.")
                End If

                'check, if STATUS found in Input XML, store in local variable
                If IsXMLElementPresent(lobjMCFileLogsElem, "DETAILS") Then
                    If Trim(lobjMCFileLogsElem.SelectNodes("DETAILS").Item(0).InnerText) <> "" Then
                        lstrDETAILS = Trim(lobjMCFileLogsElem.SelectNodes("DETAILS").Item(0).InnerText)

                        'replace one Single Quote to two Single Quotes
                        lstrDETAILS = Replace(lstrDETAILS, "'", "''")
                    Else
                        STLogger.Error(cINVALID_PARMS & " Details is not specified.")
                    End If
                Else
                    STLogger.Error(cINVALID_PARMS & " Details is not specified.")
                End If

                'dynamically build insert statement for MC_LOGS table
                lstrBatchUpdateQry = lstrBatchUpdateQry & _
                    "INSERT INTO MC_LOGS (SQ_MC_ID, DATE_START, DATE_END, STATUS, DETAILS) VALUES(" & liSQ_MC_ID & ", '" & _
                    lstrDATE_START & "', '" & lstrDATE_END & "', " & lbySTATUS & ", '" & lstrDETAILS & "'); " & vbCrLf
            Next

            'if some data found, which needs to be updated/ inserted            
            If lstrBatchUpdateQry <> "" Then
                'execute batch update SQL                
                lobjEntity.CommonSQL = lstrBatchUpdateQry
                'execute batch update SQL              
                Call lobjcDataClass.UpdateMCLogs(lobjEntity)
                'return Output XML with status as SUCCESS
                lobjEntity.OutputString = "<UPDATE_MC_FILE_DETAIL_RESPONSE><STATUS>SUCCESS</STATUS></UPDATE_MC_FILE_DETAIL_RESPONSE>"
                Return lobjEntity
            Else
                Return lobjEntity 'Statement Added By Sanjay Has To Confirm
            End If

            'write details in log file            
            STLogger.Debug(lobjEntity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug(" Exit:" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear object variables from memory
            If Not IsNothing(lobjcDataClass) Then
                'lobjcDataClass.Dispose()
                lobjcDataClass = Nothing
            End If
            If Not IsNothing(lobjMCFileLogsElem) Then
                lobjMCFileLogsElem = Nothing
            End If
            If Not IsNothing(lobjMCFileLogsNdLst) Then
                lobjMCFileLogsNdLst = Nothing
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
    'METHOD  : UpdateMCFile
    'PURPOSE : Wrapper method to update the MC File, on execution of
    '          BSSTMoneyCostAuto service
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
    Public Function UpdateMCFile(ByVal lobjEntity As cDataEntity) As cDataEntity Implements IMoneyCostMgr.UpdateMCFile

        Dim lobjcDataClass As New cDataClass       'object of cDataClass to access its methods
        Dim lobjRequestXmlDOM As New Xml.XmlDocument    'object of XML DOM to load input Request XML
        Dim lobjMCFileNdLst As Xml.XmlNodeList = Nothing      'to store node list of MC_FILE node set
        Dim lobjMCFileElem As Xml.XmlElement = Nothing       'to store single node of MC_FILE node set
        Dim lstrBatchUpdateQry As String               'to store dynamically batch update query
        Dim liCounter As Integer              'to store loop counter
        Dim liSQ_MC_ID As Integer              'to store Money Cost file Id
        Dim lstrLastScheduleDate As String               'to store last schedule process date for a particular MC File
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Validate Request XML: " & lobjEntity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'Validate if the Input Request XML is well-formed
            Try
                lobjRequestXmlDOM.LoadXml(lobjEntity.OutputString)
            Catch ex As Exception
                STLogger.Error("Error loading Input Request XML to UpdateMCDetails().  " & ex.Message, ex)
            End Try

            'get all node list of MC_FILE_DETAIL node set
            lobjMCFileNdLst = lobjRequestXmlDOM.SelectNodes("/UPDATE_MC_FILE_REQUEST/MC_FILESet/MC_FILE")

            lstrBatchUpdateQry = ""

            'loop for each individual MC_FILE_DETAIL node set
            For liCounter = 0 To lobjMCFileNdLst.Count - 1

                'get single node set of MC_FILE
                lobjMCFileElem = lobjMCFileNdLst.Item(liCounter)

                're-initialize local variables
                liSQ_MC_ID = 0
                lstrLastScheduleDate = ""

                'check, if SQ_MC_ID found in Input XML, store in local variable
                If IsXMLElementPresent(lobjMCFileElem, "SQ_MC_ID") Then
                    If Trim(lobjMCFileElem.SelectNodes("SQ_MC_ID").Item(0).InnerText) <> "" Then
                        liSQ_MC_ID = Val(lobjMCFileElem.SelectNodes("SQ_MC_ID").Item(0).InnerText)
                    Else
                        STLogger.Error(cINVALID_PARMS & "SQ_MC_ID is not specified.")
                    End If
                Else
                    STLogger.Error(cINVALID_PARMS & "SQ_MC_ID is not specified.")
                End If

                'check, if DATE_START found in Input XML, store in local variable
                If IsXMLElementPresent(lobjMCFileElem, "LAST_SCHEDULE_PROCESS_DATE") Then
                    If Trim(lobjMCFileElem.SelectNodes("LAST_SCHEDULE_PROCESS_DATE").Item(0).InnerText) <> "" Then
                        lstrLastScheduleDate = Trim(lobjMCFileElem.SelectNodes("LAST_SCHEDULE_PROCESS_DATE").Item(0).InnerText)
                    Else
                        STLogger.Error(cINVALID_PARMS & "Last Schedule Process Date is not specified.")
                    End If
                Else
                    STLogger.Error(cINVALID_PARMS & "Last Schedule Process Date is not specified.")
                End If

                'dynamically build insert statement for MC_LOGS table
                lstrBatchUpdateQry = lstrBatchUpdateQry & "UPDATE MC_FILE SET LAST_SCHEDULE_PROCESS_DATE = '" & _
                                                    lstrLastScheduleDate & "' WHERE SQ_MC_ID = " & liSQ_MC_ID & ";" & vbCrLf
            Next

            'if some data found, which needs to be updated/ inserted
            If lstrBatchUpdateQry <> "" Then
                'execute batch update SQL
                ''''''''Call lobjcDataClass.Execute(ecExecuteSQL, ecRSExecuteNoRecords, lstrBatchUpdateQry)
                lobjEntity = New cDataEntity
                lobjEntity.CommonSQL = lstrBatchUpdateQry

                Call lobjcDataClass.UpdateMCFile(lobjEntity)
                'return Output XML with status as SUCCESS
                lobjEntity.OutputString = "<UPDATE_MC_FILE_RESPONSE><STATUS>SUCCESS</STATUS></UPDATE_MC_FILE_RESPONSE>"
                Return lobjEntity
            Else
                lobjEntity.OutputString = vbNullString
                Return lobjEntity 'Added By Sanjay To Avoid Function Return Warning
            End If

            STLogger.Debug(lobjEntity.OutputString & " :" & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            'clear object variables from memory
            If Not IsNothing(lobjcDataClass) Then
                'lobjcDataClass.Dispose()
                lobjcDataClass = Nothing
            End If
            If Not IsNothing(lobjMCFileElem) Then
                lobjMCFileElem = Nothing
            End If
            If Not IsNothing(lobjMCFileNdLst) Then
                lobjMCFileNdLst = Nothing
            End If
            If Not IsNothing(lobjRequestXmlDOM) Then
                lobjRequestXmlDOM = Nothing
            End If
            If Not IsNothing(lobjEntity) Then
                lobjEntity = Nothing
            End If
        End Try

    End Function

End Class
