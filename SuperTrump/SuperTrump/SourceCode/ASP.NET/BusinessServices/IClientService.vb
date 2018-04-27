Imports System.EnterpriseServices
Imports System.ComponentModel
Imports System.Configuration
Imports System.Runtime.InteropServices


Public Class IClientService
    Dim STLogger As log4net.ILog

#Region "Class Inilization"
    Private Sub New()
        MyBase.New()
        If log4net.LogManager.GetRepository.Configured = False Then
            log4net.Config.XmlConfigurator.ConfigureAndWatch(New System.IO.FileInfo(GetConfigurationKey("DebugLogFilePath_LogForNet")))
        End If
        STLogger = log4net.LogManager.GetLogger("SUPER_TRUMP")
    End Sub
#End Region

#Region "IClientService"

    '================================================================
    'MODULE  : IClientService
    'PURPOSE : This interface provides all customized methods for the
    '          Client applications. These methods internally call the
    '          methods in the ISuperTrumpService interface.
    '================================================================

    '================================================================
    'METHOD  : ProcessMQMessage
    'PURPOSE : To process messages sent asynchronously by Client
    '          Applications through MQ Series.
    'PARMS   :
    '          astrMQMsgInfoXML [String] = XML string containing
    '          the Message Id, data, reply queue manager name, reply
    '          queue name and Correlation Id.
    '
    '          Sample Input parameter structure:
    '            <?xml version="1.0" encoding="UTF-8"?>
    '            <MQ_MESSAGE_INFO>
    '                <MESSAGE_ID>214D512053494542454C50524F4420204464C33C73201F00</MESSAGE_ID>
    '                <MESSAGE_DATA><![CDATA[<MY_MSG>...</MY_MSG>]]></MESSAGE_DATA>
    '                <REPLY_QUEUE_MANAGER>MY_Q_MGR</REPLY_QUEUE_MANAGER>
    '                <REPLY_QUEUE>MY_Q</REPLY_QUEUE>
    '                <CORRELATION_ID>414D512053494542454C50524F4420204464C33C73201F00</CORRELATION_ID>
    '            </MQ_MESSAGE_INFO>
    'RETURN  : String = XML string containing Instructions to MQ.
    '================================================================ 
    Public Function ProcessMQMessage(ByVal astrMQMsgInfoXML As String) As String
        Dim lstrResponseXML As String
        Dim lstrSTResponseQMgr As String
        Dim lstrSTResponseQ As String
        Dim lobjMQMsgInfoXMLDOM As New Xml.XmlDocument
        Dim lobjXMLSchemaSpace As New Xml.Schema.XmlSchemaSet
        Dim lstrFileLoc As String
        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_ProcessMQMessage(): In ProcessMQMessage() method")
        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_ProcessMQMessage(): Input Argument 1:" & astrMQMsgInfoXML)
        Try
            'Get the MQMessageInfo.xsd Schema
            lstrFileLoc = GetConfigurationKey("SchemaFilePath")
            Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcMQMsgInfoXMLSchemaName)

            'Assign Schema to the XML DOM object
            lobjMQMsgInfoXMLDOM.Schemas = lobjXMLSchemaSpace
            Try
                lobjMQMsgInfoXMLDOM.LoadXml(astrMQMsgInfoXML)
            Catch ex As Exception
                'Raise Error                            
                STLogger.Error(Err.Number & "BSCEFSuperTrump_IClientService:ProcessMQMessage()/" & Err.Source & Err.Description)
                Return ""
            End Try

            'Process the Pricing Request
            lstrResponseXML = ProcessPricingRequest(lobjMQMsgInfoXMLDOM.DocumentElement.SelectSingleNode("MESSAGE_DATA").InnerText)

            'Build the Response XML
            lstrSTResponseQMgr = GetConfigurationKey("ResponseQueueManager")
            lstrSTResponseQ = GetConfigurationKey("ResponseQueue")
            lstrResponseXML = "<CLIENT_APP_RESPONSE_INSTRUCTIONS>" & "<RESPONSE_MSG><![CDATA[" & lstrResponseXML & "]]></RESPONSE_MSG>" & "<RECEIVER_APP_QUEUE_MGR>" & lobjMQMsgInfoXMLDOM.DocumentElement.SelectSingleNode("REPLY_QUEUE_MANAGER").InnerText & "</RECEIVER_APP_QUEUE_MGR>" & "<RECEIVER_APP_QUEUE>" & lobjMQMsgInfoXMLDOM.DocumentElement.SelectSingleNode("REPLY_QUEUE").InnerText & "</RECEIVER_APP_QUEUE>" & "<CLIENT_RESPONSE_QUEUE_MGR>" & lstrSTResponseQMgr & "</CLIENT_RESPONSE_QUEUE_MGR>" & "<CLIENT_RESPONSE_QUEUE>" & lstrSTResponseQ & "</CLIENT_RESPONSE_QUEUE>" & "<DELETE_REQUEST_MSG>true</DELETE_REQUEST_MSG>" & "<SEND_RESPONSE_MSG_ID>false</SEND_RESPONSE_MSG_ID>" & "</CLIENT_APP_RESPONSE_INSTRUCTIONS>"
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_ProcessMQMessage(): Return value: " & lstrResponseXML)
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_ProcessMQMessage(): Exit ProcessMQMessage() method")

            'Send the Response XML back to MQ
            Return lstrResponseXML
        Catch ex As Exception
            ProcessMQMessage = ""
            STLogger.Error(Err.Number & "BSCEFSuperTrump_IClientService:ProcessMQMessage()/" & Err.Source & Err.Description)
        Finally
            If Not (lobjMQMsgInfoXMLDOM Is Nothing) Then
                lobjMQMsgInfoXMLDOM = Nothing
            End If
            If Not (lobjXMLSchemaSpace Is Nothing) Then
                lobjXMLSchemaSpace = Nothing
            End If
        End Try
    End Function

    '================================================================
    'METHOD  : ProcessPricingRequest
    'PURPOSE : To process the Pricing Request XML
    'PARMS   :
    '          astrPricingRequestXML [String] = Pricing Request XML
    'RETURN  : String = Pricing Response XML
    '================================================================
    Public Function ProcessPricingRequest(ByVal astrPricingRequestXML As String) As String
        'Added by Gaurav
        gstrExceptionFlag = ""
        'MS XML DOM Objects Declarations
        Dim lobjXMLSchemaSpace As New Xml.Schema.XmlSchemaSet
        Dim lobjPricingRequestXMLDOM As New Xml.XmlDocument
        Dim lobjPricingResponseXMLDOM As New Xml.XmlDocument

        'Other Declarations
        Dim lstrPricingRequestXML As String
        Dim llErrNbr As Integer
        Dim lstrErrDesc As String
        Dim lstrPricingResponseXML As String
        Dim lstrFileLoc As String
        Try
            lstrPricingResponseXML = Nothing ' To Avoid Null Reference Exception Added By Sanjay
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_ProcessPricingRequest(): In ProcessPricingRequest() method")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_ProcessPricingRequest(): Input Argument 1:" & astrPricingRequestXML)

            'Get the PricingRequestInfoXML.xsd Schema
            lstrFileLoc = GetConfigurationKey("SchemaFilePath")
            Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPricingReqInfoXMLSchemaName)

            'Assign Schema to the XML DOM object
            lobjPricingRequestXMLDOM.Schemas = lobjXMLSchemaSpace
            lobjXMLSchemaSpace = Nothing
            lstrPricingRequestXML = Replace(astrPricingRequestXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")

            'Load the Input XML into the XML DOM object           
            Try
                lobjPricingRequestXMLDOM.LoadXml(lstrPricingRequestXML)
            Catch ex As Exception
                'Raise Error                            
                llErrNbr = Err.Number
                lstrErrDesc = Err.Description

                'Load response XML
                Call lobjPricingResponseXMLDOM.LoadXml("<PRICING_RESPONSE_INFO><ERROR></ERROR></PRICING_RESPONSE_INFO>")

                'Add the <ERROR> node to the response XML          
                AddXMLElement(lobjPricingResponseXMLDOM, lobjPricingResponseXMLDOM.DocumentElement, "ERROR", "")
                AddXMLElement(lobjPricingResponseXMLDOM, lobjPricingResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
                AddXMLElement(lobjPricingResponseXMLDOM, lobjPricingResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", "Error on line number " & Err.Number & " of the XML. " & lstrErrDesc)
                Return lobjPricingResponseXMLDOM.OuterXml
            End Try

            'Determine which private method is to be called to process the pricing request
            Select Case lobjPricingRequestXMLDOM.DocumentElement.ChildNodes(0).Name
                'Get the prm file and amortizaton schedule for the given input
                Case "PRM_FILE_AND_AMORT_SCHED_INFO"
                    lstrPricingResponseXML = GetPRMFileAndAmortSched(lobjPricingRequestXMLDOM.DocumentElement.ChildNodes(0).OuterXml)

                    'Get the prm file and amortizaton schedule for the given input
                    '(using the Solve for payments method)
                Case "PRM_FILE_AND_AMORT_SCHED_INFO2"
                    lstrPricingResponseXML = GetPRMFileAndAmortSched2(lobjPricingRequestXMLDOM.DocumentElement.ChildNodes(0).OuterXml)

                    'Get the prm parameters and amortizaton schedule for the given input
                Case "PRM_PARAMS_AND_AMORT_SCHED_INFO"
                    lstrPricingResponseXML = GetPRMParamsAndAmortSched(lobjPricingRequestXMLDOM.DocumentElement.ChildNodes(0).OuterXml)

                    'Get the prm file, prm parameters and amortizaton schedule for the given input
                    '(using the Solve for payments method)
                Case "PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO"
                    lstrPricingResponseXML = GetPRMFileAmortSchedAndPRMParams(lobjPricingRequestXMLDOM.DocumentElement.ChildNodes(0).OuterXml)

                    'Get the prm file, prm parameters and amortizaton schedule for the given input
                    '(using the Solve for rate method)
                Case "PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2"
                    lstrPricingResponseXML = GetPRMFileAmortSchedAndPRMParams2(lobjPricingRequestXMLDOM.DocumentElement.ChildNodes(0).OuterXml)

                    'Get the prm parameters and amortization schedule but not the prm file.
                    '(using the Solve for payments method)
                Case "PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO"
                    lstrPricingResponseXML = GetPRMParamsAmortSchedAndNoPRMFile(lobjPricingRequestXMLDOM.DocumentElement.ChildNodes(0).OuterXml)
            End Select

            'Build the response XML                      
            Call lobjPricingResponseXMLDOM.LoadXml("<PRICING_RESPONSE_INFO>" & lstrPricingResponseXML & "</PRICING_RESPONSE_INFO>")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_ProcessPricingRequest(): Return value: " & lobjPricingResponseXMLDOM.OuterXml)
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_ProcessPricingRequest(): Exit ProcessPricingRequest() method")
            Return lobjPricingResponseXMLDOM.OuterXml
        Catch ex As Exception
            llErrNbr = Err.Number
            lstrErrDesc = Err.Description
            lobjPricingResponseXMLDOM.RemoveAll() ' Changed By Sanjay lobjPricingResponseXMLDOM=nothing
            Call lobjPricingResponseXMLDOM.LoadXml("<PRICING_RESPONSE_INFO><ERROR></ERROR></PRICING_RESPONSE_INFO>")
            AddXMLElement(lobjPricingResponseXMLDOM, lobjPricingResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjPricingResponseXMLDOM, lobjPricingResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)

            'Return the response XML      
            Return lobjPricingResponseXMLDOM.OuterXml
        Finally
            If Not (lobjPricingRequestXMLDOM Is Nothing) Then
                lobjPricingRequestXMLDOM = Nothing
            End If
            If Not (lobjXMLSchemaSpace Is Nothing) Then
                lobjXMLSchemaSpace = Nothing
            End If
            If Not (lobjPricingResponseXMLDOM Is Nothing) Then
                lobjPricingResponseXMLDOM = Nothing
            End If
        End Try
    End Function

    '================================================================
    'METHOD  : GetPRMFileAndAmortSched
    'PURPOSE : To get the binary PRM files and amortization schedule
    '          data.
    'PARMS   :
    '          astrPRMFileAndAmortSchedInfoXML [String] = XML string
    '          contaning information used to obtain the binary PRM
    '          files and amortization schedule data.
    'RETURN  : String = XML string containing the binary PRM files
    '          and amortization schedule data.
    '================================================================
    <STAThreadAttribute()> _
    Private Function GetPRMFileAndAmortSched(ByVal astrPRMFileAndAmortSchedInfoXML As String) As String

        'MS XML DOM Objects Declarations
        Dim lobjPRMFileAndAmortSchedInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileAndAmortSchedResponseXMLDOM As New Xml.XmlDocument
        Dim lobjPRMInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileLstXMLDOM As New Xml.XmlDocument
        Dim lobjErrorNode As Xml.XmlNode
        Dim lobjAmortSchedLstXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileAndAmortSchedNode As Xml.XmlNode
        Dim lobjCloneNode As Xml.XmlNode

        'Other Declarations
        Dim llErrNbr As Integer
        Dim lstrErrDesc As String
        Dim lstrRootTagIDAttribVal As String
        Dim liCnt As Short
        Dim liPRMFileIndex As Short
        Dim lstrPRMFileLstXML As String
        Dim lstrAmortSchedLstXML As String
        Dim lstrPRMFileName As String
        Dim lstrTotalPRMFiles As String
        Dim lobjSTSvc As BusinessServices.ISuperTrumpService

        lstrRootTagIDAttribVal = ""
        Try
            'Load the Input XML
            Call lobjPRMFileAndAmortSchedInfoXMLDOM.LoadXml(astrPRMFileAndAmortSchedInfoXML)

            'Get the ID attribute value of the root tag        
            lstrRootTagIDAttribVal = lobjPRMFileAndAmortSchedInfoXMLDOM.DocumentElement.GetAttributeNode("ID").Value

            'Load the response XML
            Call lobjPRMFileAndAmortSchedResponseXMLDOM.LoadXml("<PRM_FILE_AND_AMORT_SCHED_INFO ID=""" & lstrRootTagIDAttribVal & """></PRM_FILE_AND_AMORT_SCHED_INFO>")

            'Load the PRM Info XML which will be send as input to the GeneratePRMFiles() method
            Call lobjPRMInfoXMLDOM.LoadXml("<PRM_INFO></PRM_INFO>")

            'For each set of PRM parameters in the Input XML
            For liCnt = 0 To lobjPRMFileAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes.Count - 1

                'Add the PRM parameters to the PRM Info XML            
                AddXMLElement(lobjPRMInfoXMLDOM, lobjPRMInfoXMLDOM.DocumentElement, "PRM_FILE", "")

                lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_META_DATA").CloneNode(True)
                lobjPRMInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_PARAMS").CloneNode(True)
                lobjPRMInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Add the PRM parameters Parent tag to the response XML
                lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(False)
                lobjPRMFileAndAmortSchedResponseXMLDOM.DocumentElement.AppendChild(lobjPRMFileAndAmortSchedResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing
            Next
            'Object Destroyed For Optimization
            lobjPRMFileAndAmortSchedInfoXMLDOM = Nothing

            'Call the GeneratePRMFiles() method to get the binary PRM file List XML
            lobjSTSvc = BusinessServices.ISuperTrumpService.Instance
            lstrPRMFileLstXML = lobjSTSvc.GeneratePRMFiles(lobjPRMInfoXMLDOM.OuterXml)
            'lobjSTSvc.Dispose()

            lobjPRMInfoXMLDOM = Nothing

            Call lobjPRMFileLstXMLDOM.LoadXml(lstrPRMFileLstXML)

            'For each binary PRM file in the PRM file List XML
            liCnt = 0
            liPRMFileIndex = 0
            lstrTotalPRMFiles = CStr(lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes.Count - 1)
            While liCnt <= CDbl(lstrTotalPRMFiles)

                'Add the binary PRM file to the response XML
                lobjCloneNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).CloneNode(True)
                lobjPRMFileAndAmortSchedResponseXMLDOM.DocumentElement.ChildNodes(liCnt).AppendChild(lobjPRMFileAndAmortSchedResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Check if there was an error generating the PRM file
                lobjErrorNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("ERROR")
                If Not (lobjErrorNode Is Nothing) Then

                    'If there was an error remove the binary PRM file from the PRM file List XML
                    lobjPRMFileLstXMLDOM.DocumentElement.RemoveChild(lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex))
                    liCnt = liCnt + 1
                Else
                    liCnt = liCnt + 1
                    liPRMFileIndex = liPRMFileIndex + 1
                End If

                lobjErrorNode = Nothing
            End While

            'Check if the PRM file list XML is not empty
            If lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes.Count > 0 Then

                'Get the Amortization Schedule list for the PRM file list XML
                lobjSTSvc = ISuperTrumpService.Instance
                lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMFileLstXMLDOM.OuterXml)
                'lobjSTSvc.Dispose()
                lobjPRMFileLstXMLDOM = Nothing

                'Load the return Amortization Schedule list
                'Object Destroyed For Optimization
                lobjAmortSchedLstXMLDOM = New Xml.XmlDocument
                Call lobjAmortSchedLstXMLDOM.LoadXml(lstrAmortSchedLstXML)

                'For each Amortization Schedule in the list
                For liCnt = 0 To lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes.Count - 1

                    'Get the PRM file name in the amortization schedule data
                    lstrPRMFileName = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE_NAME").InnerText

                    'Locate the PRM file in the Response XML
                    lobjPRMFileAndAmortSchedNode = lobjPRMFileAndAmortSchedResponseXMLDOM.DocumentElement.SelectSingleNode("/PRM_FILE_AND_AMORT_SCHED_INFO/PRM_FILE_AND_AMORT_SCHED[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]")
                    If Not (lobjPRMFileAndAmortSchedNode Is Nothing) Then

                        'Add the amortization schedule to the Response XML
                        lobjCloneNode = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(True)
                        lobjPRMFileAndAmortSchedNode.AppendChild(lobjPRMFileAndAmortSchedResponseXMLDOM.ImportNode(lobjCloneNode, True))
                        lobjCloneNode = Nothing
                    End If

                    lobjPRMFileAndAmortSchedNode = Nothing
                Next

                lobjAmortSchedLstXMLDOM = Nothing
            End If
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMFileAndAmortSched(): Return value: " & lobjPRMFileAndAmortSchedResponseXMLDOM.OuterXml)
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMFileAndAmortSched(): Exit GetPRMFileAndAmortSched() method")

            'Return the response XML 
            Return lobjPRMFileAndAmortSchedResponseXMLDOM.OuterXml
        Catch ex As Exception
            lobjPRMFileAndAmortSchedResponseXMLDOM.RemoveAll()
            llErrNbr = Err.Number
            lstrErrDesc = Err.Description
            Call lobjPRMFileAndAmortSchedResponseXMLDOM.LoadXml("<PRM_FILE_AND_AMORT_SCHED_INFO ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_FILE_AND_AMORT_SCHED_INFO>")
            AddXMLElement(lobjPRMFileAndAmortSchedResponseXMLDOM, lobjPRMFileAndAmortSchedResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))

            AddXMLElement(lobjPRMFileAndAmortSchedResponseXMLDOM, lobjPRMFileAndAmortSchedResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)
            Return lobjPRMFileAndAmortSchedResponseXMLDOM.OuterXml
        Finally
            If Not (lobjPRMFileAndAmortSchedInfoXMLDOM Is Nothing) Then
                lobjPRMFileAndAmortSchedInfoXMLDOM = Nothing
            End If
            If Not (lobjPRMFileAndAmortSchedResponseXMLDOM Is Nothing) Then
                lobjPRMFileAndAmortSchedResponseXMLDOM = Nothing
            End If
        End Try

    End Function

    '================================================================
    'METHOD  : GetPRMParamsAndAmortSched
    'PURPOSE : To get the PRM parameters and the Amortization
    '          Schedule data.
    'PARMS   :
    '          astrPRMParamsAndAmortSchedInfoXML [String] = XML string
    '          contaning information used to obtain the PRM parameters
    '          and amortization schedule data.
    'RETURN  : String = XML string containing the PRM parameters
    '          and amortization schedule data.
    '================================================================

    Private Function GetPRMParamsAndAmortSched(ByVal astrPRMParamsAndAmortSchedInfoXML As String) As String
        'MS XML DOM Objects Declarations
        Dim lobjPRMParamsAndAmortSchedInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMParamsAndAmortSchedResponseXMLDOM As New Xml.XmlDocument
        Dim lobjPRMParamsInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMParamsLstXMLDOM As New Xml.XmlDocument
        Dim lobjAmortSchedLstXMLDOM As New Xml.XmlDocument
        Dim lobjPRMParamsAndAmortSchedNode As Xml.XmlNode
        Dim lobjCloneNode As Xml.XmlNode
        Dim lobjPRMFileNodeLst As Xml.XmlNodeList

        'Other Declarations
        Dim llErrNbr As Integer
        Dim lstrErrDesc As String
        Dim lstrRootTagIDAttribVal As String
        Dim liCnt As Short
        Dim lstrPRMParamsLstXML As String
        Dim lstrAmortSchedLstXML As String
        Dim lstrPRMFileName As String
        Dim lobjSTSvc As ISuperTrumpService

        lstrRootTagIDAttribVal = ""
        Try
            'Load the Input XML
            Call lobjPRMParamsAndAmortSchedInfoXMLDOM.LoadXml(astrPRMParamsAndAmortSchedInfoXML)

            'Get the ID attribute value of the root tag        
            lstrRootTagIDAttribVal = lobjPRMParamsAndAmortSchedInfoXMLDOM.DocumentElement.GetAttributeNode("ID").Value

            'Load the response XML
            Call lobjPRMParamsAndAmortSchedResponseXMLDOM.LoadXml("<PRM_PARAMS_AND_AMORT_SCHED_INFO ID=""" & lstrRootTagIDAttribVal & """></PRM_PARAMS_AND_AMORT_SCHED_INFO>")

            'Load the PRM Params Info XML which will be send as input to the GetPRMParams() method
            Call lobjPRMParamsInfoXMLDOM.LoadXml("<PRM_PARAMS_INFO></PRM_PARAMS_INFO>")

            'For each PRM Param Info in the input XML
            For liCnt = 0 To lobjPRMParamsAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes.Count - 1

                'Add the PRM Param Info to the PRM Params Info XML            
                AddXMLElement(lobjPRMParamsInfoXMLDOM, lobjPRMParamsInfoXMLDOM.DocumentElement, "PRM_PARAMS", "")

                lobjCloneNode = lobjPRMParamsAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_PARAMS_SPECS").CloneNode(True)
                lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                lobjCloneNode = lobjPRMParamsAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE").CloneNode(True)
                lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Add the PRM param Info Parent tag to the response XML
                lobjCloneNode = lobjPRMParamsAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(False)
                lobjPRMParamsAndAmortSchedResponseXMLDOM.DocumentElement.AppendChild(lobjPRMParamsAndAmortSchedResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing
            Next

            'Get the PRM parameters specified in the PRM Params Info XML
            lobjSTSvc = ISuperTrumpService.Instance
            lstrPRMParamsLstXML = lobjSTSvc.GetPRMParams(lobjPRMParamsInfoXMLDOM.OuterXml)
            'lobjSTSvc.Dispose()

            lobjPRMParamsInfoXMLDOM = Nothing

            'Object Destroyed For Optimization
            lobjPRMParamsLstXMLDOM = New Xml.XmlDocument

            Call lobjPRMParamsLstXMLDOM.LoadXml(lstrPRMParamsLstXML)

            'For each PRM parameter set in the return XML
            For liCnt = 0 To lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes.Count - 1

                'Add the PRM parameter set to the response XML
                lobjCloneNode = lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(True)
                lobjPRMParamsAndAmortSchedResponseXMLDOM.DocumentElement.ChildNodes(liCnt).AppendChild(lobjPRMParamsAndAmortSchedResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing
            Next

            lobjPRMParamsLstXMLDOM.RemoveAll()


            'Build the PRM File list XML
            lobjPRMFileNodeLst = lobjPRMParamsAndAmortSchedInfoXMLDOM.GetElementsByTagName("PRM_FILE")

            'Object Destroyed For Optimization
            lobjPRMParamsAndAmortSchedInfoXMLDOM = Nothing

            If Not (lobjPRMFileNodeLst Is Nothing) Then

                Call lobjPRMParamsLstXMLDOM.LoadXml("<PRM_FILE_LIST></PRM_FILE_LIST>")

                For liCnt = 0 To lobjPRMFileNodeLst.Count - 1

                    lobjCloneNode = lobjPRMFileNodeLst.Item(liCnt).CloneNode(True)
                    lobjPRMParamsLstXMLDOM.DocumentElement.AppendChild(lobjPRMParamsLstXMLDOM.ImportNode(lobjCloneNode, True))
                    lobjCloneNode = Nothing
                Next

                'Get the Amortization Schedule list for the PRM file list XML
                lobjSTSvc = ISuperTrumpService.Instance
                lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMParamsLstXMLDOM.OuterXml)               
                lobjPRMParamsLstXMLDOM = Nothing

                'Load the return Amortization Schedule list
                'Object Destroyed For Optimization
                lobjAmortSchedLstXMLDOM = New Xml.XmlDocument

                Call lobjAmortSchedLstXMLDOM.LoadXml(lstrAmortSchedLstXML)

                'For each Amortization Schedule in the list
                For liCnt = 0 To lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes.Count - 1

                    'Get the PRM file name in the amortization schedule data
                    lstrPRMFileName = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE_NAME").InnerText

                    'Locate the PRM file in the Response XML
                    lobjPRMParamsAndAmortSchedNode = lobjPRMParamsAndAmortSchedResponseXMLDOM.DocumentElement.SelectSingleNode("/PRM_PARAMS_AND_AMORT_SCHED_INFO/PRM_PARAMS_AND_AMORT_SCHED[PRM_PARAMS/PRM_FILE_NAME=""" & lstrPRMFileName & """]")
                    If Not (lobjPRMParamsAndAmortSchedNode Is Nothing) Then

                        'Add the amortization schedule to the Response XML
                        lobjCloneNode = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(True)
                        lobjPRMParamsAndAmortSchedNode.AppendChild(lobjPRMParamsAndAmortSchedResponseXMLDOM.ImportNode(lobjCloneNode, True))
                        lobjCloneNode = Nothing
                    End If

                    lobjPRMParamsAndAmortSchedNode = Nothing
                Next

                lobjAmortSchedLstXMLDOM = Nothing
            End If
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMParamsAndAmortSched(): Return value: " & lobjPRMParamsAndAmortSchedResponseXMLDOM.OuterXml)
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMParamsAndAmortSched(): Exit GetPRMParamsAndAmortSched() method")
            'Return the response XML
            Return lobjPRMParamsAndAmortSchedResponseXMLDOM.OuterXml
        Catch ex As Exception
            lobjPRMParamsAndAmortSchedResponseXMLDOM.RemoveAll()

            llErrNbr = Err.Number
            lstrErrDesc = Err.Description
            Call lobjPRMParamsAndAmortSchedResponseXMLDOM.LoadXml("<PRM_PARAMS_AND_AMORT_SCHED_INFO ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_PARAMS_AND_AMORT_SCHED_INFO>")
            AddXMLElement(lobjPRMParamsAndAmortSchedResponseXMLDOM, lobjPRMParamsAndAmortSchedResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjPRMParamsAndAmortSchedResponseXMLDOM, lobjPRMParamsAndAmortSchedResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)
            Return lobjPRMParamsAndAmortSchedResponseXMLDOM.OuterXml
        Finally
            If Not (lobjPRMParamsAndAmortSchedInfoXMLDOM Is Nothing) Then
                lobjPRMParamsAndAmortSchedInfoXMLDOM = Nothing
            End If
            If Not (lobjPRMParamsAndAmortSchedResponseXMLDOM Is Nothing) Then
                lobjPRMParamsAndAmortSchedResponseXMLDOM = Nothing
            End If
        End Try
    End Function

    '================================================================
    'METHOD  : Ping
    'PURPOSE : Returns a string that indicates that the component
    '          is registered properly.
    'PARMS   : NONE
    'RETURN  : String
    '================================================================


    '================================================================
    'METHOD  : GetPRMFileAndAmortSched2
    'PURPOSE : To get the binary PRM files and amortization schedule
    '          data for the specified payment structure.
    '          (Using Solve by payments method)
    'PARMS   :
    '          astrPRMFileAndAmortSchedInfoXML [String] = XML string
    '          contaning information used to obtain the binary PRM
    '          files and amortization schedule data for the specified
    '          payment structure.
    'RETURN  : String = XML string containing the binary PRM files
    '          and amortization schedule data for the specified
    '          payment structure.
    '================================================================

    Private Function GetPRMFileAndAmortSched2(ByVal astrPRMFileAndAmortSchedInfoXML As String) As String

        'MS XML DOM Objects Declarations
        Dim lobjPRMFileAndAmortSchedInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileAndAmortSchedResponseXMLDOM As New Xml.XmlDocument
        Dim lobjPRMInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileLstXMLDOM As New Xml.XmlDocument
        Dim lobjErrorNode As Xml.XmlNode
        Dim lobjAmortSchedLstXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileAndAmortSchedNode As Xml.XmlNode
        Dim lobjCloneNode As Xml.XmlNode

        'Other Declarations
        Dim llErrNbr As Integer
        Dim lstrErrDesc As String
        Dim lstrRootTagIDAttribVal As String
        Dim liCnt As Short
        Dim liPRMFileIndex As Short
        Dim lstrPRMFileLstXML As String
        Dim lstrAmortSchedLstXML As String
        Dim lstrPRMFileName As String
        Dim lstrTotalPRMFiles As String
        Dim lobjSTSvc As ISuperTrumpService

        lstrRootTagIDAttribVal = ""
        Try
            'Load the Input XML
            Call lobjPRMFileAndAmortSchedInfoXMLDOM.LoadXml(astrPRMFileAndAmortSchedInfoXML)

            'Get the ID attribute value of the root tag        
            lstrRootTagIDAttribVal = lobjPRMFileAndAmortSchedInfoXMLDOM.DocumentElement.GetAttributeNode("ID").Value

            'Load the response XML
            Call lobjPRMFileAndAmortSchedResponseXMLDOM.LoadXml("<PRM_FILE_AND_AMORT_SCHED_INFO2 ID=""" & lstrRootTagIDAttribVal & """></PRM_FILE_AND_AMORT_SCHED_INFO2>")

            'Load the PRM Info XML which will be send as input to the GeneratePRMFilesForPmtStruct() method
            Call lobjPRMInfoXMLDOM.LoadXml("<PRM_INFO></PRM_INFO>")

            'For each set of PRM parameters in the Input XML
            For liCnt = 0 To lobjPRMFileAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes.Count - 1

                'Add the PRM parameters to the PRM Info XML          
                AddXMLElement(lobjPRMInfoXMLDOM, lobjPRMInfoXMLDOM.DocumentElement, "PRM_FILE", "")

                lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_META_DATA").CloneNode(True)
                lobjPRMInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_PARAMS").CloneNode(True)
                lobjPRMInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Add the PRM parameters Parent tag to the response XML
                lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(False)
                lobjPRMFileAndAmortSchedResponseXMLDOM.DocumentElement.AppendChild(lobjPRMFileAndAmortSchedResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing
            Next

            'Object Destroyed For Optimization
            lobjPRMFileAndAmortSchedInfoXMLDOM = Nothing

            'Call the GeneratePRMFilesForPmtStruct() method to get the binary PRM file List XML
            lobjSTSvc = ISuperTrumpService.Instance
            lstrPRMFileLstXML = lobjSTSvc.GeneratePRMFilesForPmtStruct(lobjPRMInfoXMLDOM.OuterXml)
            'lobjSTSvc.Dispose()
            lobjPRMInfoXMLDOM = Nothing

            'Object Destroyed For Optimization
            lobjPRMFileLstXMLDOM = New Xml.XmlDocument

            Call lobjPRMFileLstXMLDOM.LoadXml(lstrPRMFileLstXML)

            'For each binary PRM file in the PRM file List XML
            liCnt = 0
            liPRMFileIndex = 0
            lstrTotalPRMFiles = CStr(lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes.Count - 1)
            While liCnt <= CDbl(lstrTotalPRMFiles)

                'Add the binary PRM file to the response XML
                lobjCloneNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).CloneNode(True)
                lobjPRMFileAndAmortSchedResponseXMLDOM.DocumentElement.ChildNodes(liCnt).AppendChild(lobjPRMFileAndAmortSchedResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Check if there was an error generating the PRM file
                lobjErrorNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("ERROR")
                If Not (lobjErrorNode Is Nothing) Then

                    'If there was an error remove the binary PRM file from the PRM file List XML
                    lobjPRMFileLstXMLDOM.DocumentElement.RemoveChild(lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex))
                    liCnt = liCnt + 1
                Else
                    liCnt = liCnt + 1
                    liPRMFileIndex = liPRMFileIndex + 1
                End If
                lobjErrorNode = Nothing
            End While

            'Check if the PRM file list XML is not empty
            If lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes.Count > 0 Then

                'Get the Amortization Schedule list for the PRM file list XML
                lobjSTSvc = ISuperTrumpService.Instance
                lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMFileLstXMLDOM.OuterXml)             

                lobjPRMFileLstXMLDOM = Nothing

                lobjAmortSchedLstXMLDOM = New Xml.XmlDocument
                'Load the return Amortization Schedule list
                Call lobjAmortSchedLstXMLDOM.LoadXml(lstrAmortSchedLstXML)

                'For each Amortization Schedule in the list
                For liCnt = 0 To lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes.Count - 1

                    'Get the PRM file name in the amortization schedule data
                    lstrPRMFileName = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE_NAME").InnerText

                    'Locate the PRM file in the Response XML
                    lobjPRMFileAndAmortSchedNode = lobjPRMFileAndAmortSchedResponseXMLDOM.DocumentElement.SelectSingleNode("/PRM_FILE_AND_AMORT_SCHED_INFO2/PRM_FILE_AND_AMORT_SCHED2[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]")
                    If Not (lobjPRMFileAndAmortSchedNode Is Nothing) Then

                        'Add the amortization schedule to the Response XML
                        lobjCloneNode = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(True)
                        lobjPRMFileAndAmortSchedNode.AppendChild(lobjPRMFileAndAmortSchedResponseXMLDOM.ImportNode(lobjCloneNode, True))
                        lobjCloneNode = Nothing
                    End If

                    lobjPRMFileAndAmortSchedNode = Nothing
                Next

                lobjAmortSchedLstXMLDOM = Nothing
            End If
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMFileAndAmortSched2(): Return value: " & lobjPRMFileAndAmortSchedResponseXMLDOM.OuterXml)
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMFileAndAmortSched2(): Exit GetPRMFileAndAmortSched2() method")
            Return lobjPRMFileAndAmortSchedResponseXMLDOM.OuterXml
        Catch ex As Exception
            lobjPRMFileAndAmortSchedResponseXMLDOM.RemoveAll()
            llErrNbr = Err.Number
            lstrErrDesc = Err.Description
            Call lobjPRMFileAndAmortSchedResponseXMLDOM.LoadXml("<PRM_FILE_AND_AMORT_SCHED_INFO2 ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_FILE_AND_AMORT_SCHED_INFO2>")
            AddXMLElement(lobjPRMFileAndAmortSchedResponseXMLDOM, lobjPRMFileAndAmortSchedResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjPRMFileAndAmortSchedResponseXMLDOM, lobjPRMFileAndAmortSchedResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)
            Return lobjPRMFileAndAmortSchedResponseXMLDOM.OuterXml
        Finally
            If Not (lobjPRMFileAndAmortSchedInfoXMLDOM Is Nothing) Then
                lobjPRMFileAndAmortSchedInfoXMLDOM = Nothing
            End If
            If Not (lobjPRMFileAndAmortSchedResponseXMLDOM Is Nothing) Then
                lobjPRMFileAndAmortSchedResponseXMLDOM = Nothing
            End If
        End Try

    End Function


    '================================================================
    'METHOD  : GetPRMFileAmortSchedAndPRMParams
    'PURPOSE : To get the binary PRM files and amortization schedule
    '          data for the specified payment structure and the PRM
    '          parameters. (Using Solve by payments method)
    'PARMS   :
    '          astrPRMFileAmortSchedPRMParamsInfoXML [String] = XML
    '          string contaning information used to obtain the binary
    '          PRM files, amortization schedule and PRM parameters
    '          data for the specified payment structure.
    'RETURN  : String = XML string containing the binary PRM files,
    '          amortization schedule and PRM parameters data for the
    '          specified payment structure.
    '================================================================
    Private Function GetPRMFileAmortSchedAndPRMParams(ByVal astrPRMFileAmortSchedPRMParamsInfoXML As String) As String

        ' Dim lobjSTSvc As ISuperTrumpService = New ISuperTrumpService

        'MS XML DOM Objects Declarations
        Dim lobjRequestInfoXMLDOM As New Xml.XmlDocument
        Dim lobjResponseXMLDOM As New Xml.XmlDocument
        Dim lobjPRMInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileLstXMLDOM As Xml.XmlDocument = Nothing
        Dim lobjErrorNode As Xml.XmlNode
        Dim lobjAmortSchedLstXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileAndAmortSchedNode As Xml.XmlNode
        Dim lobjCloneNode As Xml.XmlNode
        Dim lobjPRMParamsInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMParamsLstXMLDOM As New Xml.XmlDocument
        'Other Declarations
        Dim llErrNbr As Integer
        Dim lstrErrDesc As String
        Dim lstrRootTagIDAttribVal As String
        Dim liCnt As Short
        Dim liPRMFileIndex As Short
        Dim lstrPRMFileLstXML As String
        Dim lstrAmortSchedLstXML As String
        Dim lstrPRMFileName As String
        Dim lstrTotalPRMFiles As String
        Dim lstrPRMParamsLstXML As String
        Dim lobjSTSvc As ISuperTrumpService
        lstrRootTagIDAttribVal = ""
        Try
            'Load the Input XML
            Call lobjRequestInfoXMLDOM.LoadXml(astrPRMFileAmortSchedPRMParamsInfoXML)

            'Get the ID attribute value of the root tag        
            lstrRootTagIDAttribVal = lobjRequestInfoXMLDOM.DocumentElement.GetAttributeNode("ID").Value

            'Load the response XML
            Call lobjResponseXMLDOM.LoadXml("<PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO ID=""" & lstrRootTagIDAttribVal & """></PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO>")

            'Load the PRM Info XML which will be send as input to the GeneratePRMFilesForPmtStruct() method
            Call lobjPRMInfoXMLDOM.LoadXml("<PRM_INFO></PRM_INFO>")

            'Load the PRM Params Info XML which will be send as input to the GetPRMParams() method
            Call lobjPRMParamsInfoXMLDOM.LoadXml("<PRM_PARAMS_INFO></PRM_PARAMS_INFO>")

            'For each set of PRM parameters in the Input XML
            For liCnt = 0 To lobjRequestInfoXMLDOM.DocumentElement.ChildNodes.Count - 1

                'Add the PRM parameters to the PRM Info XML            
                AddXMLElement(lobjPRMInfoXMLDOM, lobjPRMInfoXMLDOM.DocumentElement, "PRM_FILE", "")

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_META_DATA").CloneNode(True)
                lobjPRMInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_PARAMS").CloneNode(True)
                lobjPRMInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Add the PRM parameters Parent tag to the response XML
                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(False)
                lobjResponseXMLDOM.DocumentElement.AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Add the PRM Param Info to the PRM Params Info XML           
                AddXMLElement(lobjPRMParamsInfoXMLDOM, lobjPRMParamsInfoXMLDOM.DocumentElement, "PRM_PARAMS", "")

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_PARAMS_SPECS").CloneNode(True)
                lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                AddXMLElement(lobjPRMParamsInfoXMLDOM, lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild, "PRM_FILE", "")

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_META_DATA/FILE_NAME").CloneNode(True)
                lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild.SelectSingleNode("PRM_FILE").AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing
            Next

            'Object Destroyed For Optimization
            lobjRequestInfoXMLDOM = Nothing
            'Call the GeneratePRMFilesForPmtStruct() method to get the binary PRM file List XML
            '(Using Solve by payments method)
            lobjSTSvc = ISuperTrumpService.Instance
            lstrPRMFileLstXML = lobjSTSvc.GeneratePRMFilesForPmtStruct(lobjPRMInfoXMLDOM.OuterXml)
            'lobjSTSvc.Dispose()
            lobjPRMInfoXMLDOM = Nothing

            'Object Destroyed For Optimization
            lobjPRMFileLstXMLDOM = New Xml.XmlDocument
            Call lobjPRMFileLstXMLDOM.LoadXml(lstrPRMFileLstXML)

            'For each binary PRM file in the PRM file List XML
            liCnt = 0
            liPRMFileIndex = 0
            lstrTotalPRMFiles = CStr(lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes.Count - 1)
            While liCnt <= CDbl(lstrTotalPRMFiles)

                'Add the binary PRM file to the response XML
                lobjCloneNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).CloneNode(True)
                lobjResponseXMLDOM.DocumentElement.ChildNodes(liCnt).AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Check if there was an error generating the PRM file
                lobjErrorNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("ERROR")
                If Not (lobjErrorNode Is Nothing) Then

                    'If there was an error remove the binary PRM file from the PRM file List XML
                    lobjPRMFileLstXMLDOM.DocumentElement.RemoveChild(lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex))
                    liCnt = liCnt + 1
                Else
                    'Get the PRM file name in the PRM file List XML
                    lstrPRMFileName = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("FILE_NAME").OuterXml

                    'Add the binary PRM file to the PRM Params Info XML
                    lobjCloneNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("FILE_DATA").CloneNode(True)
                    lobjPRMParamsInfoXMLDOM.DocumentElement.SelectSingleNode("//PRM_PARAMS/PRM_FILE[FILE_NAME=""" & lstrPRMFileName & """]").AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                    lobjCloneNode = Nothing

                    liCnt = liCnt + 1
                    liPRMFileIndex = liPRMFileIndex + 1
                End If
                lobjErrorNode = Nothing
            End While

            'Check if the PRM file list XML is not empty
            If lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes.Count > 0 Then

                'Get the Amortization Schedule list for the PRM file list XML
                lobjSTSvc = ISuperTrumpService.Instance
                lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMFileLstXMLDOM.OuterXml)
                'lobjSTSvc.Dispose()
                'Object Destroyed For Optimization
                lobjAmortSchedLstXMLDOM = New Xml.XmlDocument

                'Load the return Amortization Schedule list
                Call lobjAmortSchedLstXMLDOM.LoadXml(lstrAmortSchedLstXML)

                'For each Amortization Schedule in the list
                For liCnt = 0 To lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes.Count - 1

                    'Get the PRM file name in the amortization schedule data
                    lstrPRMFileName = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE_NAME").InnerText

                    'Locate the PRM file in the Response XML
                    lobjPRMFileAndAmortSchedNode = lobjResponseXMLDOM.DocumentElement.SelectSingleNode("/PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO/PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]")
                    If Not (lobjPRMFileAndAmortSchedNode Is Nothing) Then

                        'Add the amortization schedule to the Response XML
                        lobjCloneNode = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(True)
                        lobjPRMFileAndAmortSchedNode.AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                        lobjCloneNode = Nothing
                    End If
                    lobjPRMFileAndAmortSchedNode = Nothing
                Next
                lobjAmortSchedLstXMLDOM = Nothing

                'Get the PRM parameters specified in the PRM Params Info XML
                lobjSTSvc = ISuperTrumpService.Instance
                lstrPRMParamsLstXML = lobjSTSvc.GetPRMParams(lobjPRMParamsInfoXMLDOM.OuterXml)

                lobjPRMParamsInfoXMLDOM = Nothing

                'Object Destroyed For Optimization
                lobjPRMParamsLstXMLDOM = New Xml.XmlDocument

                Call lobjPRMParamsLstXMLDOM.LoadXml(lstrPRMParamsLstXML)

                'For each PRM parameter set in the return XML
                For liCnt = 0 To lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes.Count - 1

                    'Get the PRM file name in the amortization schedule data
                    lstrPRMFileName = lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE_NAME").InnerText

                    'Add the PRM parameter set to the response XML
                    lobjCloneNode = lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(True)
                    lobjResponseXMLDOM.DocumentElement.SelectSingleNode("//PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]").AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                    lobjCloneNode = Nothing
                Next

                lobjPRMParamsLstXMLDOM = Nothing

            End If
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMFileAmortSchedAndPRMParams(): Return value: " & lobjResponseXMLDOM.OuterXml)
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMFileAmortSchedAndPRMParams(): Exit GetPRMFileAmortSchedAndPRMParams() method")
            'Return the response XML
            Return lobjResponseXMLDOM.OuterXml
        Catch ex As Exception
            lobjResponseXMLDOM.RemoveAll()
            llErrNbr = Err.Number
            lstrErrDesc = Err.Description
            Call lobjResponseXMLDOM.LoadXml("<PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO>")
            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)
            Return lobjResponseXMLDOM.OuterXml
        Finally
            If Not (lobjPRMFileLstXMLDOM Is Nothing) Then
                lobjPRMFileLstXMLDOM = Nothing
            End If
            If Not (lobjRequestInfoXMLDOM Is Nothing) Then
                lobjRequestInfoXMLDOM = Nothing
            End If
            If Not (lobjResponseXMLDOM Is Nothing) Then
                lobjResponseXMLDOM = Nothing
            End If
        End Try
    End Function

    '================================================================
    'METHOD  : GetPRMFileAmortSchedAndPRMParams2
    'PURPOSE : To get the binary PRM files and amortization schedule
    '          data for the specified payment structure and the PRM
    '          parameters. (Using Solve by Rate method)
    'PARMS   :
    '          astrPRMFileAmortSchedPRMParamsInfoXML [String] = XML
    '          string contaning information used to obtain the binary
    '          PRM files, amortization schedule and PRM parameters
    '          data for the specified payment structure.
    'RETURN  : String = XML string containing the binary PRM files,
    '          amortization schedule and PRM parameters data for the
    '          specified payment structure. 
    '================================================================

    Private Function GetPRMFileAmortSchedAndPRMParams2(ByVal astrPRMFileAmortSchedPRMParamsInfoXML As String) As String

        'Dim lobjSTSvc As ISuperTrumpService = New ISuperTrumpService

        'MS XML DOM Objects Declarations
        Dim lobjRequestInfoXMLDOM As New Xml.XmlDocument
        Dim lobjResponseXMLDOM As New Xml.XmlDocument
        Dim lobjPRMInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileLstXMLDOM As Xml.XmlDocument = Nothing
        Dim lobjErrorNode As Xml.XmlNode
        Dim lobjAmortSchedLstXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileAndAmortSchedNode As Xml.XmlNode
        Dim lobjCloneNode As Xml.XmlNode
        Dim lobjPRMParamsInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMParamsLstXMLDOM As New Xml.XmlDocument

        'Other Declarations
        Dim llErrNbr As Integer
        Dim lstrErrDesc As String
        Dim lstrRootTagIDAttribVal As String
        Dim liCnt As Short
        Dim liPRMFileIndex As Short
        Dim lstrPRMFileLstXML As String
        Dim lstrAmortSchedLstXML As String
        Dim lstrPRMFileName As String
        Dim lstrTotalPRMFiles As String
        Dim lstrPRMParamsLstXML As String
        Dim lobjSTSvc As ISuperTrumpService

        lstrRootTagIDAttribVal = ""
        Try
            'Load the Input XML
            Call lobjRequestInfoXMLDOM.LoadXml(astrPRMFileAmortSchedPRMParamsInfoXML)

            'Get the ID attribute value of the root tag        
            lstrRootTagIDAttribVal = lobjRequestInfoXMLDOM.DocumentElement.GetAttributeNode("ID").Value

            'Load the response XML
            Call lobjResponseXMLDOM.LoadXml("<PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2 ID=""" & lstrRootTagIDAttribVal & """></PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2>")

            'Load the PRM Info XML which will be send as input to the GeneratePRMFilesForPmtStruct2() method
            Call lobjPRMInfoXMLDOM.LoadXml("<PRM_INFO></PRM_INFO>")

            'Load the PRM Params Info XML which will be send as input to the GetPRMParams() method
            Call lobjPRMParamsInfoXMLDOM.LoadXml("<PRM_PARAMS_INFO></PRM_PARAMS_INFO>")

            'For each set of PRM parameters in the Input XML
            For liCnt = 0 To lobjRequestInfoXMLDOM.DocumentElement.ChildNodes.Count - 1

                'Add the PRM parameters to the PRM Info XML           
                AddXMLElement(lobjPRMInfoXMLDOM, lobjPRMInfoXMLDOM.DocumentElement, "PRM_FILE", "")

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_META_DATA").CloneNode(True)
                lobjPRMInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_PARAMS").CloneNode(True)
                lobjPRMInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Add the PRM parameters Parent tag to the response XML
                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(False)
                lobjResponseXMLDOM.DocumentElement.AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Add the PRM Param Info to the PRM Params Info XML            
                AddXMLElement(lobjPRMParamsInfoXMLDOM, lobjPRMParamsInfoXMLDOM.DocumentElement, "PRM_PARAMS", "")

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_PARAMS_SPECS").CloneNode(True)
                lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                AddXMLElement(lobjPRMParamsInfoXMLDOM, lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild, "PRM_FILE", "")

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_META_DATA/FILE_NAME").CloneNode(True)
                lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild.SelectSingleNode("PRM_FILE").AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing
            Next

            'Object Destroyed For Optimization
            lobjRequestInfoXMLDOM = Nothing

            'Call the GeneratePRMFilesForPmtStruct2() method to get the binary PRM file List XML
            '(Using Solve by Rate method)
            lobjSTSvc = ISuperTrumpService.Instance
            lstrPRMFileLstXML = lobjSTSvc.GeneratePRMFilesForPmtStruct2(lobjPRMInfoXMLDOM.OuterXml)
            'lobjSTSvc.Dispose()

            lobjPRMInfoXMLDOM = Nothing

            'Object Destroyed For Optimization
            lobjPRMFileLstXMLDOM = New Xml.XmlDocument

            Call lobjPRMFileLstXMLDOM.LoadXml(lstrPRMFileLstXML)

            'For each binary PRM file in the PRM file List XML
            liCnt = 0
            liPRMFileIndex = 0
            lstrTotalPRMFiles = CStr(lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes.Count - 1)
            While liCnt <= CDbl(lstrTotalPRMFiles)

                'Add the binary PRM file to the response XML
                lobjCloneNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).CloneNode(True)
                lobjResponseXMLDOM.DocumentElement.ChildNodes(liCnt).AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Check if there was an error generating the PRM file
                lobjErrorNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("ERROR")
                If Not (lobjErrorNode Is Nothing) Then

                    'If there was an error remove the binary PRM file from the PRM file List XML
                    lobjPRMFileLstXMLDOM.DocumentElement.RemoveChild(lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex))
                    liCnt = liCnt + 1
                Else
                    'Get the PRM file name in the PRM file List XML
                    lstrPRMFileName = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("FILE_NAME").InnerXml

                    'Add the binary PRM file to the PRM Params Info XML
                    lobjCloneNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("FILE_DATA").CloneNode(True)
                    lobjPRMParamsInfoXMLDOM.DocumentElement.SelectSingleNode("//PRM_PARAMS/PRM_FILE[FILE_NAME=""" & lstrPRMFileName & """]").AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                    lobjCloneNode = Nothing

                    liCnt = liCnt + 1
                    liPRMFileIndex = liPRMFileIndex + 1
                End If
                lobjErrorNode = Nothing
            End While

            'Check if the PRM file list XML is not empty
            If lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes.Count > 0 Then

                'Get the Amortization Schedule list for the PRM file list XML
                lobjSTSvc = ISuperTrumpService.Instance
                lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMFileLstXMLDOM.OuterXml)
                'lobjSTSvc.Dispose()

                'Object Destroyed For Optimization
                lobjAmortSchedLstXMLDOM = New Xml.XmlDocument

                'Load the return Amortization Schedule list
                Call lobjAmortSchedLstXMLDOM.LoadXml(lstrAmortSchedLstXML)

                'For each Amortization Schedule in the list
                For liCnt = 0 To lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes.Count - 1

                    'Get the PRM file name in the amortization schedule data
                    lstrPRMFileName = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE_NAME").InnerText

                    'Locate the PRM file in the Response XML
                    lobjPRMFileAndAmortSchedNode = lobjResponseXMLDOM.DocumentElement.SelectSingleNode("/PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2/PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS2[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]")
                    If Not (lobjPRMFileAndAmortSchedNode Is Nothing) Then

                        'Add the amortization schedule to the Response XML
                        lobjCloneNode = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(True)
                        lobjPRMFileAndAmortSchedNode.AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                        lobjCloneNode = Nothing
                    End If

                    lobjPRMFileAndAmortSchedNode = Nothing
                Next

                lobjAmortSchedLstXMLDOM = Nothing

                'Get the PRM parameters specified in the PRM Params Info XML
                lobjSTSvc = ISuperTrumpService.Instance
                lstrPRMParamsLstXML = lobjSTSvc.GetPRMParams(lobjPRMParamsInfoXMLDOM.OuterXml)
                'lobjSTSvc.Dispose()
                lobjPRMParamsInfoXMLDOM = Nothing

                'Object Destroyed For Optimization
                lobjPRMParamsLstXMLDOM = New Xml.XmlDocument

                Call lobjPRMParamsLstXMLDOM.LoadXml(lstrPRMParamsLstXML)

                'For each PRM parameter set in the return XML
                For liCnt = 0 To lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes.Count - 1

                    'Get the PRM file name in the amortization schedule data
                    lstrPRMFileName = lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE_NAME").InnerText

                    'Add the PRM parameter set to the response XML
                    lobjCloneNode = lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(True)
                    lobjResponseXMLDOM.DocumentElement.SelectSingleNode("//PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS2[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]").AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                    lobjCloneNode = Nothing
                Next


                lobjPRMParamsLstXMLDOM = Nothing

            End If
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMFileAmortSchedAndPRMParams2(): Return value: " & lobjResponseXMLDOM.OuterXml)
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMFileAmortSchedAndPRMParams2(): Exit GetPRMFileAmortSchedAndPRMParams2() method")
            'Return the response XML
            Return lobjResponseXMLDOM.OuterXml
        Catch ex As Exception
            lobjResponseXMLDOM.RemoveAll()
            llErrNbr = Err.Number
            lstrErrDesc = Err.Description
            Call lobjResponseXMLDOM.LoadXml("<PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2 ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2>")
            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)
            Return lobjResponseXMLDOM.OuterXml
        Finally
            If Not (lobjPRMFileLstXMLDOM Is Nothing) Then
                lobjPRMFileLstXMLDOM = Nothing
            End If
            If Not (lobjRequestInfoXMLDOM Is Nothing) Then
                lobjRequestInfoXMLDOM = Nothing
            End If
            If Not (lobjResponseXMLDOM Is Nothing) Then
                lobjResponseXMLDOM = Nothing
            End If
        End Try

    End Function

    '================================================================
    'METHOD  : GetPRMParamsAmortSchedAndNoPRMFile
    'PURPOSE : To get the amortization schedule
    '          data for the specified payment structure and the PRM
    '          parameters. (Using Solve by payments method)
    '          The generated PRM file will be ignored.
    'PARMS   :
    '          astrPRMFileAmortSchedPRMParamsInfoXML [String] = XML
    '          string contaning information used to obtain the
    '          amortization schedule and PRM parameters
    '          data for the specified payment structure.
    'RETURN  : String = XML string containing the amortization
    '          schedule and PRM parameters data for the
    '          specified payment structure.
    '================================================================

    Private Function GetPRMParamsAmortSchedAndNoPRMFile(ByVal astrPRMFileAmortSchedPRMParamsInfoXML As String) As String

        'Dim lobjSTSvc As ISuperTrumpService = New ISuperTrumpService

        'MS XML DOM Objects Declarations
        Dim lobjRequestInfoXMLDOM As New Xml.XmlDocument
        Dim lobjResponseXMLDOM As New Xml.XmlDocument
        Dim lobjPRMInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileLstXMLDOM As Xml.XmlDocument = Nothing
        Dim lobjErrorNode As Xml.XmlNode
        Dim lobjAmortSchedLstXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileAndAmortSchedNode As Xml.XmlNode
        Dim lobjCloneNode As Xml.XmlNode
        Dim lobjPRMParamsInfoXMLDOM As New Xml.XmlDocument
        Dim lobjPRMParamsLstXMLDOM As New Xml.XmlDocument
        Dim lobjBinaryPRMFile As Xml.XmlNode

        'Other Declarations
        Dim llErrNbr As Integer
        Dim lstrErrDesc As String
        Dim lstrRootTagIDAttribVal As String
        Dim liCnt As Short
        Dim liPRMFileIndex As Short
        Dim lstrPRMFileLstXML As String
        Dim lstrAmortSchedLstXML As String
        Dim lstrPRMFileName As String
        Dim lstrTotalPRMFiles As String
        Dim lstrPRMParamsLstXML As String
        Dim lobjSTSvc As ISuperTrumpService
        lstrRootTagIDAttribVal = ""
        Try
            'Load the Input XML
            Call lobjRequestInfoXMLDOM.LoadXml(astrPRMFileAmortSchedPRMParamsInfoXML)

            'Get the ID attribute value of the root tag       
            lstrRootTagIDAttribVal = lobjRequestInfoXMLDOM.DocumentElement.GetAttributeNode("ID").Value

            'Load the response XML
            Call lobjResponseXMLDOM.LoadXml("<PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO ID=""" & lstrRootTagIDAttribVal & """></PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO>")

            'Load the PRM Info XML which will be send as input to the GeneratePRMFilesForPmtStruct() method
            Call lobjPRMInfoXMLDOM.LoadXml("<PRM_INFO></PRM_INFO>")

            'Load the PRM Params Info XML which will be send as input to the GetPRMParams() method
            Call lobjPRMParamsInfoXMLDOM.LoadXml("<PRM_PARAMS_INFO></PRM_PARAMS_INFO>")

            'For each set of PRM parameters in the Input XML
            For liCnt = 0 To lobjRequestInfoXMLDOM.DocumentElement.ChildNodes.Count - 1

                'Add the PRM parameters to the PRM Info XML            
                AddXMLElement(lobjPRMInfoXMLDOM, lobjPRMInfoXMLDOM.DocumentElement, "PRM_FILE", "")

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_META_DATA").CloneNode(True)
                lobjPRMInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_PARAMS").CloneNode(True)
                lobjPRMInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Add the PRM parameters Parent tag to the response XML
                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(False)
                lobjResponseXMLDOM.DocumentElement.AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Add the PRM Param Info to the PRM Params Info XML            
                AddXMLElement(lobjPRMParamsInfoXMLDOM, lobjPRMParamsInfoXMLDOM.DocumentElement, "PRM_PARAMS", "")

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_PARAMS_SPECS").CloneNode(True)
                lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild.AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                AddXMLElement(lobjPRMParamsInfoXMLDOM, lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild, "PRM_FILE", "")

                lobjCloneNode = lobjRequestInfoXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_META_DATA/FILE_NAME").CloneNode(True)
                lobjPRMParamsInfoXMLDOM.DocumentElement.LastChild.SelectSingleNode("PRM_FILE").AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing
            Next

            'Object Destroyed For Optimization
            lobjRequestInfoXMLDOM = Nothing

            'Call the GeneratePRMFilesForPmtStruct() method to get the binary PRM file List XML
            '(Using Solve by payments method)
            lobjSTSvc = ISuperTrumpService.Instance
            lstrPRMFileLstXML = lobjSTSvc.GeneratePRMFilesForPmtStruct(lobjPRMInfoXMLDOM.OuterXml)
            'lobjSTSvc.Dispose()
            lobjPRMInfoXMLDOM = Nothing

            'Object Destroyed For Optimization
            lobjPRMFileLstXMLDOM = New Xml.XmlDocument

            Call lobjPRMFileLstXMLDOM.LoadXml(lstrPRMFileLstXML)

            'For each binary PRM file in the PRM file List XML
            liCnt = 0
            liPRMFileIndex = 0
            lstrTotalPRMFiles = CStr(lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes.Count - 1)
            While liCnt <= CDbl(lstrTotalPRMFiles)

                'Add the binary PRM file to the response XML
                lobjCloneNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).CloneNode(True)
                lobjResponseXMLDOM.DocumentElement.ChildNodes(liCnt).AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                lobjCloneNode = Nothing

                'Remove the binary PRM file.
                lobjBinaryPRMFile = lobjResponseXMLDOM.DocumentElement.ChildNodes(liCnt).ChildNodes(0).SelectSingleNode("FILE_DATA")
                If Not (lobjBinaryPRMFile Is Nothing) Then
                    lobjResponseXMLDOM.DocumentElement.ChildNodes(liCnt).ChildNodes(0).RemoveChild(lobjBinaryPRMFile)
                End If

                'Check if there was an error generating the PRM file
                lobjErrorNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("ERROR")
                If Not (lobjErrorNode Is Nothing) Then

                    'If there was an error remove the binary PRM file from the PRM file List XML
                    lobjPRMFileLstXMLDOM.DocumentElement.RemoveChild(lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex))
                    liCnt = liCnt + 1
                Else
                    'Get the PRM file name in the PRM file List XML
                    lstrPRMFileName = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("FILE_NAME").InnerText

                    'Add the binary PRM file to the PRM Params Info XML
                    lobjCloneNode = lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes(liPRMFileIndex).SelectSingleNode("FILE_DATA").CloneNode(True)

                    lobjPRMParamsInfoXMLDOM.DocumentElement.SelectSingleNode("//PRM_PARAMS/PRM_FILE[FILE_NAME=""" & lstrPRMFileName & """]").AppendChild(lobjPRMParamsInfoXMLDOM.ImportNode(lobjCloneNode, True))
                    lobjCloneNode = Nothing

                    liCnt = liCnt + 1
                    liPRMFileIndex = liPRMFileIndex + 1
                End If
                lobjErrorNode = Nothing
            End While

            'Check if the PRM file list XML is not empty
            If lobjPRMFileLstXMLDOM.DocumentElement.ChildNodes.Count > 0 Then

                'Get the Amortization Schedule list for the PRM file list XML
                lobjSTSvc = ISuperTrumpService.Instance
                lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMFileLstXMLDOM.OuterXml)
                'lobjSTSvc.Dispose()

                'Object Destroyed For Optimization
                lobjAmortSchedLstXMLDOM = New Xml.XmlDocument
                'Load the return Amortization Schedule list
                Call lobjAmortSchedLstXMLDOM.LoadXml(lstrAmortSchedLstXML)

                'For each Amortization Schedule in the list
                For liCnt = 0 To lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes.Count - 1

                    'Get the PRM file name in the amortization schedule data
                    lstrPRMFileName = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE_NAME").InnerText

                    'Locate the PRM file in the Response XML
                    lobjPRMFileAndAmortSchedNode = lobjResponseXMLDOM.DocumentElement.SelectSingleNode("/PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO/PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]")
                    If Not (lobjPRMFileAndAmortSchedNode Is Nothing) Then

                        'Add the amortization schedule to the Response XML
                        lobjCloneNode = lobjAmortSchedLstXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(True)
                        lobjPRMFileAndAmortSchedNode.AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                        lobjCloneNode = Nothing
                    End If

                    lobjPRMFileAndAmortSchedNode = Nothing
                Next

                lobjAmortSchedLstXMLDOM = Nothing

                'Get the PRM parameters specified in the PRM Params Info XML
                lobjSTSvc = ISuperTrumpService.Instance
                lstrPRMParamsLstXML = lobjSTSvc.GetPRMParams(lobjPRMParamsInfoXMLDOM.OuterXml)
                'lobjSTSvc.Dispose()
                lobjPRMParamsInfoXMLDOM = Nothing

                'Object Destroyed For Optimization
                lobjPRMParamsLstXMLDOM = New Xml.XmlDocument

                Call lobjPRMParamsLstXMLDOM.LoadXml(lstrPRMParamsLstXML)

                'For each PRM parameter set in the return XML
                For liCnt = 0 To lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes.Count - 1

                    'Get the PRM file name in the amortization schedule data
                    lstrPRMFileName = lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE_NAME").InnerText

                    'Add the PRM parameter set to the response XML
                    lobjCloneNode = lobjPRMParamsLstXMLDOM.DocumentElement.ChildNodes(liCnt).CloneNode(True)
                    lobjResponseXMLDOM.DocumentElement.SelectSingleNode("//PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]").AppendChild(lobjResponseXMLDOM.ImportNode(lobjCloneNode, True))
                    lobjCloneNode = Nothing
                Next

                lobjPRMParamsLstXMLDOM = Nothing

            End If
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMParamsAmortSchedAndNoPRMFile(): Return value: " & lobjResponseXMLDOM.OuterXml)
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.IClientService_GetPRMParamsAmortSchedAndNoPRMFile(): Exit GetPRMParamsAmortSchedAndNoPRMFile() method")
            Return lobjResponseXMLDOM.OuterXml
        Catch ex As Exception
            lobjResponseXMLDOM.RemoveAll()
            llErrNbr = Err.Number
            lstrErrDesc = Err.Description
            Call lobjResponseXMLDOM.LoadXml("<PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO>")
            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)
            Return lobjResponseXMLDOM.OuterXml
        Finally
            If Not (lobjPRMFileLstXMLDOM Is Nothing) Then
                lobjPRMFileLstXMLDOM = Nothing
            End If
            If Not (lobjRequestInfoXMLDOM Is Nothing) Then
                lobjRequestInfoXMLDOM = Nothing
            End If
            If Not (lobjResponseXMLDOM Is Nothing) Then
                lobjResponseXMLDOM = Nothing
            End If
        End Try
    End Function
    Public Function GetThradApartment() As String
        Try
            Return System.Threading.Thread.CurrentThread.GetApartmentState.ToString()
        Catch ex As Exception
            Return "Error occured while getting Thread Apartment - " & Err.Description
        Finally
        End Try
    End Function
#End Region

#Region "Shared Function for Sigleton"
    Private Shared _instanceIClient As IClientService
    Public Shared Function Instance() As IClientService
        If _instanceIClient Is Nothing Then
            _instanceIClient = New IClientService
        End If
        Return _instanceIClient
    End Function
#End Region

End Class
