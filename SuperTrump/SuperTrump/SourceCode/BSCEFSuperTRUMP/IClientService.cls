VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "IClientService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'================================================================
'MODULE  : IClientService
'PURPOSE : This interface provides all customized methods for the
'          Client applications. These methods internally call the
'          methods in the ISuperTrumpService interface.
'================================================================
Option Explicit

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
On Error GoTo ErrHandler

Dim lstrResponseXML         As String
Dim lstrSTResponseQMgr      As String
Dim lstrSTResponseQ         As String
Dim lobjMQMsgInfoXMLDOM    As New DOMDocument40
Dim lobjXMLSchemaSpace      As New XMLSchemaCache40
Dim lstrFileLoc             As String

    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_ProcessMQMessage(): In ProcessMQMessage() method"
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_ProcessMQMessage(): Input Argument 1:" & astrMQMsgInfoXML
    
    'Get the MQMessageInfo.xsd Schema
    lstrFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcSchemaFilePathKey)
    Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcMQMsgInfoXMLSchemaName)
    
    'Assign Schema to the XML DOM object
    Set lobjMQMsgInfoXMLDOM.schemas = lobjXMLSchemaSpace
    lobjMQMsgInfoXMLDOM.async = False
    
    If lobjMQMsgInfoXMLDOM.loadXML(astrMQMsgInfoXML) Then
        
        'Process the Pricing Request
        lstrResponseXML = ProcessPricingRequest(lobjMQMsgInfoXMLDOM.documentElement.selectSingleNode("MESSAGE_DATA").Text)
        
        'Build the Response XML
        lstrSTResponseQMgr = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcResponseQMgrKey)
        lstrSTResponseQ = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcResponseQKey)
        
        lstrResponseXML = "<CLIENT_APP_RESPONSE_INSTRUCTIONS>" & _
                            "<RESPONSE_MSG><![CDATA[" & lstrResponseXML & "]]></RESPONSE_MSG>" & _
                            "<RECEIVER_APP_QUEUE_MGR>" & lobjMQMsgInfoXMLDOM.documentElement.selectSingleNode("REPLY_QUEUE_MANAGER").Text & "</RECEIVER_APP_QUEUE_MGR>" & _
                            "<RECEIVER_APP_QUEUE>" & lobjMQMsgInfoXMLDOM.documentElement.selectSingleNode("REPLY_QUEUE").Text & "</RECEIVER_APP_QUEUE>" & _
                            "<CLIENT_RESPONSE_QUEUE_MGR>" & lstrSTResponseQMgr & "</CLIENT_RESPONSE_QUEUE_MGR>" & _
                            "<CLIENT_RESPONSE_QUEUE>" & lstrSTResponseQ & "</CLIENT_RESPONSE_QUEUE>" & _
                            "<DELETE_REQUEST_MSG>true</DELETE_REQUEST_MSG>" & _
                            "<SEND_RESPONSE_MSG_ID>false</SEND_RESPONSE_MSG_ID>" & _
                        "</CLIENT_APP_RESPONSE_INSTRUCTIONS>"
        
        'Send the Response XML back to MQ
        ProcessMQMessage = lstrResponseXML
    Else
        Err.Raise lobjMQMsgInfoXMLDOM.parseError.errorCode, "BSCEFSuperTrump_IClientService:ProcessMQMessage()/" & Err.Source, lobjMQMsgInfoXMLDOM.parseError.reason
    End If
    
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_ProcessMQMessage(): Return value: " & ProcessMQMessage
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_ProcessMQMessage(): Exit ProcessMQMessage() method"
    
    Exit Function
    
ErrHandler:
    ProcessMQMessage = ""
    Err.Raise Err.Number, "BSCEFSuperTrump_IClientService:ProcessMQMessage()/" & Err.Source, Err.Description
End Function

'================================================================
'METHOD  : ProcessPricingRequest
'PURPOSE : To process the Pricing Request XML
'PARMS   :
'          astrPricingRequestXML [String] = Pricing Request XML
'RETURN  : String = Pricing Response XML
'================================================================
Public Function ProcessPricingRequest(ByVal astrPricingRequestXML As String) As String
                                
On Error GoTo ErrHandler
'Added by Gaurav
gstrExceptionFlag = ""
'MS XML DOM Objects Declarations
Dim lobjXMLSchemaSpace          As New XMLSchemaCache40
Dim lobjPricingRequestXMLDOM    As New DOMDocument40
Dim lobjPricingResponseXMLDOM   As New DOMDocument40

'Other Declarations
Dim lstrPricingRequestXML       As String
Dim llErrNbr                    As Long
Dim lstrErrDesc                 As String
Dim lstrPricingResponseXML      As String
Dim lstrFileLoc                 As String

    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_ProcessPricingRequest(): In ProcessPricingRequest() method"
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_ProcessPricingRequest(): Input Argument 1:" & astrPricingRequestXML
    
    'Get the PricingRequestInfoXML.xsd Schema
    lstrFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcSchemaFilePathKey)
    Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPricingReqInfoXMLSchemaName)
    
    'Assign Schema to the XML DOM object
    Set lobjPricingRequestXMLDOM.schemas = lobjXMLSchemaSpace
    lobjPricingRequestXMLDOM.async = False
    
    'Remove "xmlns:dt" & "dt:dt" attributes from the <FILE_DATA> elements in the Input XML
    'as we cannot include these attributes in the schema as they contain special
    'characters(:). We will add them after schema validation.
    lstrPricingRequestXML = Replace(astrPricingRequestXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
    
    'Load the Input XML into the XML DOM object
    Call lobjPricingRequestXMLDOM.loadXML(lstrPricingRequestXML)
    
    'Check , if any parser error has occured (Validate Input XML)
    If lobjPricingRequestXMLDOM.parseError.errorCode = 0 Then
        
        'Determine which private method is to be called to process the pricing request
        Select Case lobjPricingRequestXMLDOM.documentElement.childNodes(0).nodeName
        
            'Get the prm file and amortizaton schedule for the given input
            Case "PRM_FILE_AND_AMORT_SCHED_INFO"
                lstrPricingResponseXML = GetPRMFileAndAmortSched(lobjPricingRequestXMLDOM.documentElement.childNodes(0).xml)
            
            'Get the prm file and amortizaton schedule for the given input
            '(using the Solve for payments method)
            Case "PRM_FILE_AND_AMORT_SCHED_INFO2"
                lstrPricingResponseXML = GetPRMFileAndAmortSched2(lobjPricingRequestXMLDOM.documentElement.childNodes(0).xml)
            
            'Get the prm parameters and amortizaton schedule for the given input
            Case "PRM_PARAMS_AND_AMORT_SCHED_INFO"
                lstrPricingResponseXML = GetPRMParamsAndAmortSched(lobjPricingRequestXMLDOM.documentElement.childNodes(0).xml)
            
            'Get the prm file, prm parameters and amortizaton schedule for the given input
            '(using the Solve for payments method)
            Case "PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO"
                lstrPricingResponseXML = GetPRMFileAmortSchedAndPRMParams(lobjPricingRequestXMLDOM.documentElement.childNodes(0).xml)
            
            'Get the prm file, prm parameters and amortizaton schedule for the given input
            '(using the Solve for rate method)
            Case "PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2"
                lstrPricingResponseXML = GetPRMFileAmortSchedAndPRMParams2(lobjPricingRequestXMLDOM.documentElement.childNodes(0).xml)
            
            'Get the prm parameters and amortization schedule but not the prm file.
            '(using the Solve for payments method)
            Case "PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO"
                lstrPricingResponseXML = GetPRMParamsAmortSchedAndNoPRMFile(lobjPricingRequestXMLDOM.documentElement.childNodes(0).xml)
        End Select
        
        'Build the response XML
        Call lobjPricingResponseXMLDOM.loadXML("<PRICING_RESPONSE_INFO>" & _
                                                    lstrPricingResponseXML & _
                                                "</PRICING_RESPONSE_INFO>")
    Else
        llErrNbr = lobjPricingRequestXMLDOM.parseError.errorCode
        lstrErrDesc = lobjPricingRequestXMLDOM.parseError.reason
            
        'Load response XML
        Call lobjPricingResponseXMLDOM.loadXML("<PRICING_RESPONSE_INFO><ERROR></ERROR></PRICING_RESPONSE_INFO>")
    
        'Add the <ERROR> node to the response XML
        AddXMLElement lobjPricingResponseXMLDOM, _
                    lobjPricingResponseXMLDOM.documentElement, _
                    "ERROR", _
                    ""
                    
        AddXMLElement lobjPricingResponseXMLDOM, _
                    lobjPricingResponseXMLDOM.documentElement.childNodes(0), _
                    "ERROR_NBR", _
                    llErrNbr
                
        AddXMLElement lobjPricingResponseXMLDOM, _
                    lobjPricingResponseXMLDOM.documentElement.childNodes(0), _
                    "ERROR_DESC", _
                    "Error on line number " & lobjPricingRequestXMLDOM.parseError.Line & " of the XML. " & lstrErrDesc
    End If
    
CleanUp:

    'Return the response XML
    ProcessPricingRequest = lobjPricingResponseXMLDOM.xml
        
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_ProcessPricingRequest(): Return value: " & ProcessPricingRequest
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_ProcessPricingRequest(): Exit ProcessPricingRequest() method"
            
    Set lobjPricingRequestXMLDOM = Nothing
    Set lobjXMLSchemaSpace = Nothing
    Set lobjPricingResponseXMLDOM = Nothing
    
    Exit Function
    
ErrHandler:
    Set lobjPricingResponseXMLDOM = Nothing
    
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description
    
    Call lobjPricingResponseXMLDOM.loadXML("<PRICING_RESPONSE_INFO><ERROR></ERROR></PRICING_RESPONSE_INFO>")
        
    AddXMLElement lobjPricingResponseXMLDOM, _
                lobjPricingResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_NBR", _
                llErrNbr
                
    AddXMLElement lobjPricingResponseXMLDOM, _
                lobjPricingResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_DESC", _
                lstrErrDesc
                
    Resume CleanUp
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
Private Function GetPRMFileAndAmortSched(ByVal astrPRMFileAndAmortSchedInfoXML As String) As String
On Error GoTo ErrHandler
    
Dim lobjSTSvc                               As New BSCEFSuperTrump.ISuperTrumpService
    
'MS XML DOM Objects Declarations
Dim lobjPRMFileAndAmortSchedInfoXMLDOM      As New DOMDocument40
Dim lobjPRMFileAndAmortSchedResponseXMLDOM  As New DOMDocument40
Dim lobjPRMInfoXMLDOM                       As New DOMDocument40
Dim lobjPRMFileLstXMLDOM                    As New DOMDocument40
Dim lobjErrorNode                           As IXMLDOMNode
Dim lobjAmortSchedLstXMLDOM                 As New DOMDocument40
Dim lobjPRMFileAndAmortSchedNode            As IXMLDOMNode
Dim lobjCloneNode                           As IXMLDOMNode

'Other Declarations
Dim llErrNbr                                As Long
Dim lstrErrDesc                             As String
Dim lstrRootTagIDAttribVal                  As String
Dim liCnt                                   As Integer
Dim liPRMFileIndex                          As Integer
Dim lstrPRMFileLstXML                       As String
Dim lstrAmortSchedLstXML                    As String
Dim lstrPRMFileName                         As String
Dim lstrTotalPRMFiles                       As String
    
    'Load the Input XML
    Call lobjPRMFileAndAmortSchedInfoXMLDOM.loadXML(astrPRMFileAndAmortSchedInfoXML)
    
    'Get the ID attribute value of the root tag
    lstrRootTagIDAttribVal = lobjPRMFileAndAmortSchedInfoXMLDOM.documentElement.getAttributeNode("ID").Value
    
    'Load the response XML
    Call lobjPRMFileAndAmortSchedResponseXMLDOM.loadXML("<PRM_FILE_AND_AMORT_SCHED_INFO ID=""" & lstrRootTagIDAttribVal & """></PRM_FILE_AND_AMORT_SCHED_INFO>")
    
    'Load the PRM Info XML which will be send as input to the GeneratePRMFiles() method
    Call lobjPRMInfoXMLDOM.loadXML("<PRM_INFO></PRM_INFO>")
    
    'For each set of PRM parameters in the Input XML
    For liCnt = 0 To lobjPRMFileAndAmortSchedInfoXMLDOM.documentElement.childNodes.length - 1
        
        'Add the PRM parameters to the PRM Info XML
        AddXMLElement lobjPRMInfoXMLDOM, _
                        lobjPRMInfoXMLDOM.documentElement, _
                        "PRM_FILE", _
                        ""
        
        Set lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_META_DATA").cloneNode(True)
        lobjPRMInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        Set lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_PARAMS").cloneNode(True)
        lobjPRMInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Add the PRM parameters Parent tag to the response XML
        Set lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.documentElement.childNodes(liCnt).cloneNode(False)
        lobjPRMFileAndAmortSchedResponseXMLDOM.documentElement.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
    Next
        
    'Call the GeneratePRMFiles() method to get the binary PRM file List XML
    lstrPRMFileLstXML = lobjSTSvc.GeneratePRMFiles(lobjPRMInfoXMLDOM.xml)
    
    Set lobjPRMInfoXMLDOM = Nothing
    
    Call lobjPRMFileLstXMLDOM.loadXML(lstrPRMFileLstXML)
    
    'For each binary PRM file in the PRM file List XML
    liCnt = 0
    liPRMFileIndex = 0
    lstrTotalPRMFiles = lobjPRMFileLstXMLDOM.documentElement.childNodes.length - 1
    While liCnt <= lstrTotalPRMFiles
        
        'Add the binary PRM file to the response XML
        Set lobjCloneNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).cloneNode(True)
        lobjPRMFileAndAmortSchedResponseXMLDOM.documentElement.childNodes(liCnt).appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Check if there was an error generating the PRM file
        Set lobjErrorNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("ERROR")
        If Not (lobjErrorNode Is Nothing) Then
            
            'If there was an error remove the binary PRM file from the PRM file List XML
            lobjPRMFileLstXMLDOM.documentElement.removeChild lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex)
            liCnt = liCnt + 1
        Else
            liCnt = liCnt + 1
            liPRMFileIndex = liPRMFileIndex + 1
        End If
        Set lobjErrorNode = Nothing
    Wend
    
    'Check if the PRM file list XML is not empty
    If lobjPRMFileLstXMLDOM.documentElement.childNodes.length > 0 Then
    
        'Get the Amortization Schedule list for the PRM file list XML
        lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMFileLstXMLDOM.xml)
        
        Set lobjPRMFileLstXMLDOM = Nothing
        
        'Load the return Amortization Schedule list
        Call lobjAmortSchedLstXMLDOM.loadXML(lstrAmortSchedLstXML)
        
        'For each Amortization Schedule in the list
        For liCnt = 0 To lobjAmortSchedLstXMLDOM.documentElement.childNodes.length - 1
            
            'Get the PRM file name in the amortization schedule data
            lstrPRMFileName = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE_NAME").Text
            
            'Locate the PRM file in the Response XML
            Set lobjPRMFileAndAmortSchedNode = lobjPRMFileAndAmortSchedResponseXMLDOM.documentElement.selectSingleNode("/PRM_FILE_AND_AMORT_SCHED_INFO/PRM_FILE_AND_AMORT_SCHED[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]")
            If Not (lobjPRMFileAndAmortSchedNode Is Nothing) Then
                
                'Add the amortization schedule to the Response XML
                Set lobjCloneNode = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).cloneNode(True)
                lobjPRMFileAndAmortSchedNode.appendChild lobjCloneNode
                Set lobjCloneNode = Nothing
            End If
            
            Set lobjPRMFileAndAmortSchedNode = Nothing
        Next
        
        Set lobjAmortSchedLstXMLDOM = Nothing
    End If
CleanUp:

    'Return the response XML
    GetPRMFileAndAmortSched = lobjPRMFileAndAmortSchedResponseXMLDOM.xml
        
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMFileAndAmortSched(): Return value: " & GetPRMFileAndAmortSched
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMFileAndAmortSched(): Exit GetPRMFileAndAmortSched() method"
            
    Set lobjPRMFileAndAmortSchedInfoXMLDOM = Nothing
    Set lobjPRMFileAndAmortSchedResponseXMLDOM = Nothing
    
    Exit Function
    
ErrHandler:
    Set lobjPRMFileAndAmortSchedResponseXMLDOM = Nothing
    
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description
    
    Call lobjPRMFileAndAmortSchedResponseXMLDOM.loadXML("<PRM_FILE_AND_AMORT_SCHED_INFO ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_FILE_AND_AMORT_SCHED_INFO>")
        
    AddXMLElement lobjPRMFileAndAmortSchedResponseXMLDOM, _
                lobjPRMFileAndAmortSchedResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_NBR", _
                llErrNbr
                
    AddXMLElement lobjPRMFileAndAmortSchedResponseXMLDOM, _
                lobjPRMFileAndAmortSchedResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_DESC", _
                lstrErrDesc
                
    Resume CleanUp
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
On Error GoTo ErrHandler
    
Dim lobjSTSvc                                   As New BSCEFSuperTrump.ISuperTrumpService
    
'MS XML DOM Objects Declarations
Dim lobjPRMParamsAndAmortSchedInfoXMLDOM        As New DOMDocument40
Dim lobjPRMParamsAndAmortSchedResponseXMLDOM    As New DOMDocument40
Dim lobjPRMParamsInfoXMLDOM                     As New DOMDocument40
Dim lobjPRMParamsLstXMLDOM                      As New DOMDocument40
Dim lobjErrorNode                               As IXMLDOMNode
Dim lobjAmortSchedLstXMLDOM                     As New DOMDocument40
Dim lobjPRMParamsAndAmortSchedNode              As IXMLDOMNode
Dim lobjCloneNode                               As IXMLDOMNode
Dim lobjPRMFileLstXMLDOM                        As New DOMDocument40
Dim lobjPRMFileNodeLst                          As IXMLDOMNodeList

'Other Declarations
Dim llErrNbr                                    As Long
Dim lstrErrDesc                                 As String
Dim lstrRootTagIDAttribVal                      As String
Dim liCnt                                       As Integer
Dim liPRMFileIndex                              As Integer
Dim lstrPRMParamsLstXML                         As String
Dim lstrAmortSchedLstXML                        As String
Dim lstrPRMFileName                             As String
Dim lstrTotalPRMFiles                           As String
    
    'Load the Input XML
    Call lobjPRMParamsAndAmortSchedInfoXMLDOM.loadXML(astrPRMParamsAndAmortSchedInfoXML)
    
    'Get the ID attribute value of the root tag
    lstrRootTagIDAttribVal = lobjPRMParamsAndAmortSchedInfoXMLDOM.documentElement.getAttributeNode("ID").Value
    
    'Load the response XML
    Call lobjPRMParamsAndAmortSchedResponseXMLDOM.loadXML("<PRM_PARAMS_AND_AMORT_SCHED_INFO ID=""" & lstrRootTagIDAttribVal & """></PRM_PARAMS_AND_AMORT_SCHED_INFO>")
    
    'Load the PRM Params Info XML which will be send as input to the GetPRMParams() method
    Call lobjPRMParamsInfoXMLDOM.loadXML("<PRM_PARAMS_INFO></PRM_PARAMS_INFO>")
    
    'For each PRM Param Info in the input XML
    For liCnt = 0 To lobjPRMParamsAndAmortSchedInfoXMLDOM.documentElement.childNodes.length - 1
        
        'Add the PRM Param Info to the PRM Params Info XML
        AddXMLElement lobjPRMParamsInfoXMLDOM, _
                        lobjPRMParamsInfoXMLDOM.documentElement, _
                        "PRM_PARAMS", _
                        ""
        
        Set lobjCloneNode = lobjPRMParamsAndAmortSchedInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_PARAMS_SPECS").cloneNode(True)
        lobjPRMParamsInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        Set lobjCloneNode = lobjPRMParamsAndAmortSchedInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE").cloneNode(True)
        lobjPRMParamsInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Add the PRM param Info Parent tag to the response XML
        Set lobjCloneNode = lobjPRMParamsAndAmortSchedInfoXMLDOM.documentElement.childNodes(liCnt).cloneNode(False)
        lobjPRMParamsAndAmortSchedResponseXMLDOM.documentElement.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
    Next
        
    'Get the PRM parameters specified in the PRM Params Info XML
    lstrPRMParamsLstXML = lobjSTSvc.GetPRMParams(lobjPRMParamsInfoXMLDOM.xml)
    
    Set lobjPRMParamsInfoXMLDOM = Nothing
    
    Call lobjPRMParamsLstXMLDOM.loadXML(lstrPRMParamsLstXML)
    
    'For each PRM parameter set in the return XML
    For liCnt = 0 To lobjPRMParamsLstXMLDOM.documentElement.childNodes.length - 1
        
        'Add the PRM parameter set to the response XML
        Set lobjCloneNode = lobjPRMParamsLstXMLDOM.documentElement.childNodes(liCnt).cloneNode(True)
        lobjPRMParamsAndAmortSchedResponseXMLDOM.documentElement.childNodes(liCnt).appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
    Next
    
    Set lobjPRMParamsLstXMLDOM = Nothing
    
    'Build the PRM File list XML
    Set lobjPRMFileNodeLst = lobjPRMParamsAndAmortSchedInfoXMLDOM.getElementsByTagName("PRM_FILE")
    
    If Not (lobjPRMFileNodeLst Is Nothing) Then
        
        Call lobjPRMParamsLstXMLDOM.loadXML("<PRM_FILE_LIST></PRM_FILE_LIST>")
        
        For liCnt = 0 To lobjPRMFileNodeLst.length - 1
            
            Set lobjCloneNode = lobjPRMFileNodeLst.Item(liCnt).cloneNode(True)
            lobjPRMParamsLstXMLDOM.documentElement.appendChild lobjCloneNode
            Set lobjCloneNode = Nothing
        Next
        
        'Get the Amortization Schedule list for the PRM file list XML
        lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMParamsLstXMLDOM.xml)
        Set lobjPRMParamsLstXMLDOM = Nothing
        
        'Load the return Amortization Schedule list
        Call lobjAmortSchedLstXMLDOM.loadXML(lstrAmortSchedLstXML)
        
        'For each Amortization Schedule in the list
        For liCnt = 0 To lobjAmortSchedLstXMLDOM.documentElement.childNodes.length - 1
            
            'Get the PRM file name in the amortization schedule data
            lstrPRMFileName = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE_NAME").Text
            
            'Locate the PRM file in the Response XML
            Set lobjPRMParamsAndAmortSchedNode = lobjPRMParamsAndAmortSchedResponseXMLDOM.documentElement.selectSingleNode("/PRM_PARAMS_AND_AMORT_SCHED_INFO/PRM_PARAMS_AND_AMORT_SCHED[PRM_PARAMS/PRM_FILE_NAME=""" & lstrPRMFileName & """]")
            If Not (lobjPRMParamsAndAmortSchedNode Is Nothing) Then
                
                'Add the amortization schedule to the Response XML
                Set lobjCloneNode = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).cloneNode(True)
                lobjPRMParamsAndAmortSchedNode.appendChild lobjCloneNode
                Set lobjCloneNode = Nothing
            End If
            
            Set lobjPRMParamsAndAmortSchedNode = Nothing
        Next
        
        Set lobjAmortSchedLstXMLDOM = Nothing
    End If
    
CleanUp:

    'Return the response XML
    GetPRMParamsAndAmortSched = lobjPRMParamsAndAmortSchedResponseXMLDOM.xml
        
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMParamsAndAmortSched(): Return value: " & GetPRMParamsAndAmortSched
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMParamsAndAmortSched(): Exit GetPRMParamsAndAmortSched() method"
            
    Set lobjPRMParamsAndAmortSchedInfoXMLDOM = Nothing
    Set lobjPRMParamsAndAmortSchedResponseXMLDOM = Nothing
    
    Exit Function
    
ErrHandler:
    Set lobjPRMParamsAndAmortSchedResponseXMLDOM = Nothing
    
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description
    
    Call lobjPRMParamsAndAmortSchedResponseXMLDOM.loadXML("<PRM_PARAMS_AND_AMORT_SCHED_INFO ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_PARAMS_AND_AMORT_SCHED_INFO>")
        
    AddXMLElement lobjPRMParamsAndAmortSchedResponseXMLDOM, _
                lobjPRMParamsAndAmortSchedResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_NBR", _
                llErrNbr
                
    AddXMLElement lobjPRMParamsAndAmortSchedResponseXMLDOM, _
                lobjPRMParamsAndAmortSchedResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_DESC", _
                lstrErrDesc
                
    Resume CleanUp
End Function

'================================================================
'METHOD  : Ping
'PURPOSE : Returns a string that indicates that the component
'          is registered properly.
'PARMS   : NONE
'RETURN  : String
'================================================================

Public Function Ping() As String
    
    Ping = "Ping request to BSCEFSuperTrump.IClientService returned at " & Format(Now, "mm/dd/yyyy Hh:Nn:Ss AM/PM") & " server time."
End Function

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
On Error GoTo ErrHandler
    
Dim lobjSTSvc                               As New BSCEFSuperTrump.ISuperTrumpService
    
'MS XML DOM Objects Declarations
Dim lobjPRMFileAndAmortSchedInfoXMLDOM      As New DOMDocument40
Dim lobjPRMFileAndAmortSchedResponseXMLDOM  As New DOMDocument40
Dim lobjPRMInfoXMLDOM                       As New DOMDocument40
Dim lobjPRMFileLstXMLDOM                    As New DOMDocument40
Dim lobjErrorNode                           As IXMLDOMNode
Dim lobjAmortSchedLstXMLDOM                 As New DOMDocument40
Dim lobjPRMFileAndAmortSchedNode            As IXMLDOMNode
Dim lobjCloneNode                           As IXMLDOMNode

'Other Declarations
Dim llErrNbr                                As Long
Dim lstrErrDesc                             As String
Dim lstrRootTagIDAttribVal                  As String
Dim liCnt                                   As Integer
Dim liPRMFileIndex                          As Integer
Dim lstrPRMFileLstXML                       As String
Dim lstrAmortSchedLstXML                    As String
Dim lstrPRMFileName                         As String
Dim lstrTotalPRMFiles                       As String
    
    'Load the Input XML
    Call lobjPRMFileAndAmortSchedInfoXMLDOM.loadXML(astrPRMFileAndAmortSchedInfoXML)
    
    'Get the ID attribute value of the root tag
    lstrRootTagIDAttribVal = lobjPRMFileAndAmortSchedInfoXMLDOM.documentElement.getAttributeNode("ID").Value
    
    'Load the response XML
    Call lobjPRMFileAndAmortSchedResponseXMLDOM.loadXML("<PRM_FILE_AND_AMORT_SCHED_INFO2 ID=""" & lstrRootTagIDAttribVal & """></PRM_FILE_AND_AMORT_SCHED_INFO2>")
    
    'Load the PRM Info XML which will be send as input to the GeneratePRMFilesForPmtStruct() method
    Call lobjPRMInfoXMLDOM.loadXML("<PRM_INFO></PRM_INFO>")
    
    'For each set of PRM parameters in the Input XML
    For liCnt = 0 To lobjPRMFileAndAmortSchedInfoXMLDOM.documentElement.childNodes.length - 1
        
        'Add the PRM parameters to the PRM Info XML
        AddXMLElement lobjPRMInfoXMLDOM, _
                        lobjPRMInfoXMLDOM.documentElement, _
                        "PRM_FILE", _
                        ""
        
        Set lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_META_DATA").cloneNode(True)
        lobjPRMInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        Set lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_PARAMS").cloneNode(True)
        lobjPRMInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Add the PRM parameters Parent tag to the response XML
        Set lobjCloneNode = lobjPRMFileAndAmortSchedInfoXMLDOM.documentElement.childNodes(liCnt).cloneNode(False)
        lobjPRMFileAndAmortSchedResponseXMLDOM.documentElement.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
    Next
        
    'Call the GeneratePRMFilesForPmtStruct() method to get the binary PRM file List XML
    lstrPRMFileLstXML = lobjSTSvc.GeneratePRMFilesForPmtStruct(lobjPRMInfoXMLDOM.xml)
    
    Set lobjPRMInfoXMLDOM = Nothing
    
    Call lobjPRMFileLstXMLDOM.loadXML(lstrPRMFileLstXML)
    
    'For each binary PRM file in the PRM file List XML
    liCnt = 0
    liPRMFileIndex = 0
    lstrTotalPRMFiles = lobjPRMFileLstXMLDOM.documentElement.childNodes.length - 1
    While liCnt <= lstrTotalPRMFiles
        
        'Add the binary PRM file to the response XML
        Set lobjCloneNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).cloneNode(True)
        lobjPRMFileAndAmortSchedResponseXMLDOM.documentElement.childNodes(liCnt).appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Check if there was an error generating the PRM file
        Set lobjErrorNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("ERROR")
        If Not (lobjErrorNode Is Nothing) Then
            
            'If there was an error remove the binary PRM file from the PRM file List XML
            lobjPRMFileLstXMLDOM.documentElement.removeChild lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex)
            liCnt = liCnt + 1
        Else
            liCnt = liCnt + 1
            liPRMFileIndex = liPRMFileIndex + 1
        End If
        Set lobjErrorNode = Nothing
    Wend
    
    'Check if the PRM file list XML is not empty
    If lobjPRMFileLstXMLDOM.documentElement.childNodes.length > 0 Then
    
        'Get the Amortization Schedule list for the PRM file list XML
        lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMFileLstXMLDOM.xml)
        
        Set lobjPRMFileLstXMLDOM = Nothing
        
        'Load the return Amortization Schedule list
        Call lobjAmortSchedLstXMLDOM.loadXML(lstrAmortSchedLstXML)
        
        'For each Amortization Schedule in the list
        For liCnt = 0 To lobjAmortSchedLstXMLDOM.documentElement.childNodes.length - 1
            
            'Get the PRM file name in the amortization schedule data
            lstrPRMFileName = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE_NAME").Text
            
            'Locate the PRM file in the Response XML
            Set lobjPRMFileAndAmortSchedNode = lobjPRMFileAndAmortSchedResponseXMLDOM.documentElement.selectSingleNode("/PRM_FILE_AND_AMORT_SCHED_INFO2/PRM_FILE_AND_AMORT_SCHED2[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]")
            If Not (lobjPRMFileAndAmortSchedNode Is Nothing) Then
                
                'Add the amortization schedule to the Response XML
                Set lobjCloneNode = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).cloneNode(True)
                lobjPRMFileAndAmortSchedNode.appendChild lobjCloneNode
                Set lobjCloneNode = Nothing
            End If
            
            Set lobjPRMFileAndAmortSchedNode = Nothing
        Next
        
        Set lobjAmortSchedLstXMLDOM = Nothing
    End If
CleanUp:

    'Return the response XML
    GetPRMFileAndAmortSched2 = lobjPRMFileAndAmortSchedResponseXMLDOM.xml
        
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMFileAndAmortSched2(): Return value: " & GetPRMFileAndAmortSched2
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMFileAndAmortSched2(): Exit GetPRMFileAndAmortSched2() method"
            
    Set lobjPRMFileAndAmortSchedInfoXMLDOM = Nothing
    Set lobjPRMFileAndAmortSchedResponseXMLDOM = Nothing
    
    Exit Function
    
ErrHandler:
    Set lobjPRMFileAndAmortSchedResponseXMLDOM = Nothing
    
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description
    
    Call lobjPRMFileAndAmortSchedResponseXMLDOM.loadXML("<PRM_FILE_AND_AMORT_SCHED_INFO2 ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_FILE_AND_AMORT_SCHED_INFO2>")
        
    AddXMLElement lobjPRMFileAndAmortSchedResponseXMLDOM, _
                lobjPRMFileAndAmortSchedResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_NBR", _
                llErrNbr
                
    AddXMLElement lobjPRMFileAndAmortSchedResponseXMLDOM, _
                lobjPRMFileAndAmortSchedResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_DESC", _
                lstrErrDesc
                
    Resume CleanUp
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
On Error GoTo ErrHandler
    
Dim lobjSTSvc                               As New BSCEFSuperTrump.ISuperTrumpService
    
'MS XML DOM Objects Declarations
Dim lobjRequestInfoXMLDOM                   As New DOMDocument40
Dim lobjResponseXMLDOM                      As New DOMDocument40
Dim lobjPRMInfoXMLDOM                       As New DOMDocument40
Dim lobjPRMFileLstXMLDOM                    As New DOMDocument40
Dim lobjErrorNode                           As IXMLDOMNode
Dim lobjAmortSchedLstXMLDOM                 As New DOMDocument40
Dim lobjPRMFileAndAmortSchedNode            As IXMLDOMNode
Dim lobjCloneNode                           As IXMLDOMNode
Dim lobjPRMParamsInfoXMLDOM                 As New DOMDocument40
Dim lobjPRMParamsLstXMLDOM                  As New DOMDocument40

'Other Declarations
Dim llErrNbr                                As Long
Dim lstrErrDesc                             As String
Dim lstrRootTagIDAttribVal                  As String
Dim liCnt                                   As Integer
Dim liPRMFileIndex                          As Integer
Dim lstrPRMFileLstXML                       As String
Dim lstrAmortSchedLstXML                    As String
Dim lstrPRMFileName                         As String
Dim lstrTotalPRMFiles                       As String
Dim lstrPRMParamsLstXML                     As String

    'Load the Input XML
    Call lobjRequestInfoXMLDOM.loadXML(astrPRMFileAmortSchedPRMParamsInfoXML)
    
    'Get the ID attribute value of the root tag
    lstrRootTagIDAttribVal = lobjRequestInfoXMLDOM.documentElement.getAttributeNode("ID").Value
    
    'Load the response XML
    Call lobjResponseXMLDOM.loadXML("<PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO ID=""" & lstrRootTagIDAttribVal & """></PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO>")
    
    'Load the PRM Info XML which will be send as input to the GeneratePRMFilesForPmtStruct() method
    Call lobjPRMInfoXMLDOM.loadXML("<PRM_INFO></PRM_INFO>")
    
    'Load the PRM Params Info XML which will be send as input to the GetPRMParams() method
    Call lobjPRMParamsInfoXMLDOM.loadXML("<PRM_PARAMS_INFO></PRM_PARAMS_INFO>")
    
    'For each set of PRM parameters in the Input XML
    For liCnt = 0 To lobjRequestInfoXMLDOM.documentElement.childNodes.length - 1
        
        'Add the PRM parameters to the PRM Info XML
        AddXMLElement lobjPRMInfoXMLDOM, _
                        lobjPRMInfoXMLDOM.documentElement, _
                        "PRM_FILE", _
                        ""
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_META_DATA").cloneNode(True)
        lobjPRMInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_PARAMS").cloneNode(True)
        lobjPRMInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Add the PRM parameters Parent tag to the response XML
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).cloneNode(False)
        lobjResponseXMLDOM.documentElement.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Add the PRM Param Info to the PRM Params Info XML
        AddXMLElement lobjPRMParamsInfoXMLDOM, _
                        lobjPRMParamsInfoXMLDOM.documentElement, _
                        "PRM_PARAMS", _
                        ""
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_PARAMS_SPECS").cloneNode(True)
        lobjPRMParamsInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        AddXMLElement lobjPRMParamsInfoXMLDOM, _
                        lobjPRMParamsInfoXMLDOM.documentElement.lastChild, _
                        "PRM_FILE", _
                        ""
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_META_DATA/FILE_NAME").cloneNode(True)
        lobjPRMParamsInfoXMLDOM.documentElement.lastChild.selectSingleNode("PRM_FILE").appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
    Next
        
    'Call the GeneratePRMFilesForPmtStruct() method to get the binary PRM file List XML
    '(Using Solve by payments method)
    lstrPRMFileLstXML = lobjSTSvc.GeneratePRMFilesForPmtStruct(lobjPRMInfoXMLDOM.xml)
    
    Set lobjPRMInfoXMLDOM = Nothing
    
    Call lobjPRMFileLstXMLDOM.loadXML(lstrPRMFileLstXML)
    
    'For each binary PRM file in the PRM file List XML
    liCnt = 0
    liPRMFileIndex = 0
    lstrTotalPRMFiles = lobjPRMFileLstXMLDOM.documentElement.childNodes.length - 1
    While liCnt <= lstrTotalPRMFiles
        
        'Add the binary PRM file to the response XML
        Set lobjCloneNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).cloneNode(True)
        lobjResponseXMLDOM.documentElement.childNodes(liCnt).appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Check if there was an error generating the PRM file
        Set lobjErrorNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("ERROR")
        If Not (lobjErrorNode Is Nothing) Then
            
            'If there was an error remove the binary PRM file from the PRM file List XML
            lobjPRMFileLstXMLDOM.documentElement.removeChild lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex)
            liCnt = liCnt + 1
        Else
            'Get the PRM file name in the PRM file List XML
            lstrPRMFileName = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("FILE_NAME").Text
            
            'Add the binary PRM file to the PRM Params Info XML
            Set lobjCloneNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("FILE_DATA").cloneNode(True)
            lobjPRMParamsInfoXMLDOM.documentElement.selectSingleNode("//PRM_PARAMS/PRM_FILE[FILE_NAME=""" & lstrPRMFileName & """]").appendChild lobjCloneNode
            Set lobjCloneNode = Nothing
            
            liCnt = liCnt + 1
            liPRMFileIndex = liPRMFileIndex + 1
        End If
        Set lobjErrorNode = Nothing
    Wend
    
    'Check if the PRM file list XML is not empty
    If lobjPRMFileLstXMLDOM.documentElement.childNodes.length > 0 Then
    
        'Get the Amortization Schedule list for the PRM file list XML
        lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMFileLstXMLDOM.xml)
        
        'Load the return Amortization Schedule list
        Call lobjAmortSchedLstXMLDOM.loadXML(lstrAmortSchedLstXML)
        
        'For each Amortization Schedule in the list
        For liCnt = 0 To lobjAmortSchedLstXMLDOM.documentElement.childNodes.length - 1
            
            'Get the PRM file name in the amortization schedule data
            lstrPRMFileName = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE_NAME").Text
            
            'Locate the PRM file in the Response XML
            Set lobjPRMFileAndAmortSchedNode = lobjResponseXMLDOM.documentElement.selectSingleNode("/PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO/PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]")
            If Not (lobjPRMFileAndAmortSchedNode Is Nothing) Then
                
                'Add the amortization schedule to the Response XML
                Set lobjCloneNode = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).cloneNode(True)
                lobjPRMFileAndAmortSchedNode.appendChild lobjCloneNode
                Set lobjCloneNode = Nothing
            End If
            
            Set lobjPRMFileAndAmortSchedNode = Nothing
        Next
        
        Set lobjAmortSchedLstXMLDOM = Nothing
        
        'Get the PRM parameters specified in the PRM Params Info XML
        lstrPRMParamsLstXML = lobjSTSvc.GetPRMParams(lobjPRMParamsInfoXMLDOM.xml)
    
        Set lobjPRMParamsInfoXMLDOM = Nothing
    
        Call lobjPRMParamsLstXMLDOM.loadXML(lstrPRMParamsLstXML)
    
        'For each PRM parameter set in the return XML
        For liCnt = 0 To lobjPRMParamsLstXMLDOM.documentElement.childNodes.length - 1
            
            'Get the PRM file name in the amortization schedule data
            lstrPRMFileName = lobjPRMParamsLstXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE_NAME").Text
            
            'Add the PRM parameter set to the response XML
            Set lobjCloneNode = lobjPRMParamsLstXMLDOM.documentElement.childNodes(liCnt).cloneNode(True)
            lobjResponseXMLDOM.documentElement.selectSingleNode("//PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]").appendChild lobjCloneNode
            Set lobjCloneNode = Nothing
        Next
        
        Set lobjPRMParamsLstXMLDOM = Nothing
    
    End If
CleanUp:

    'Return the response XML
    GetPRMFileAmortSchedAndPRMParams = lobjResponseXMLDOM.xml
        
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMFileAmortSchedAndPRMParams(): Return value: " & GetPRMFileAmortSchedAndPRMParams
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMFileAmortSchedAndPRMParams(): Exit GetPRMFileAmortSchedAndPRMParams() method"
    
    Set lobjPRMFileLstXMLDOM = Nothing
    Set lobjRequestInfoXMLDOM = Nothing
    Set lobjResponseXMLDOM = Nothing
    
    Exit Function
    
ErrHandler:
    Set lobjResponseXMLDOM = Nothing
    
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description
    
    Call lobjResponseXMLDOM.loadXML("<PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO>")
        
    AddXMLElement lobjResponseXMLDOM, _
                lobjResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_NBR", _
                llErrNbr
                
    AddXMLElement lobjResponseXMLDOM, _
                lobjResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_DESC", _
                lstrErrDesc
                
    Resume CleanUp
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
On Error GoTo ErrHandler
    
Dim lobjSTSvc                               As New BSCEFSuperTrump.ISuperTrumpService
    
'MS XML DOM Objects Declarations
Dim lobjRequestInfoXMLDOM                   As New DOMDocument40
Dim lobjResponseXMLDOM                      As New DOMDocument40
Dim lobjPRMInfoXMLDOM                       As New DOMDocument40
Dim lobjPRMFileLstXMLDOM                    As New DOMDocument40
Dim lobjErrorNode                           As IXMLDOMNode
Dim lobjAmortSchedLstXMLDOM                 As New DOMDocument40
Dim lobjPRMFileAndAmortSchedNode            As IXMLDOMNode
Dim lobjCloneNode                           As IXMLDOMNode
Dim lobjPRMParamsInfoXMLDOM                 As New DOMDocument40
Dim lobjPRMParamsLstXMLDOM                  As New DOMDocument40

'Other Declarations
Dim llErrNbr                                As Long
Dim lstrErrDesc                             As String
Dim lstrRootTagIDAttribVal                  As String
Dim liCnt                                   As Integer
Dim liPRMFileIndex                          As Integer
Dim lstrPRMFileLstXML                       As String
Dim lstrAmortSchedLstXML                    As String
Dim lstrPRMFileName                         As String
Dim lstrTotalPRMFiles                       As String
Dim lstrPRMParamsLstXML                     As String

    'Load the Input XML
    Call lobjRequestInfoXMLDOM.loadXML(astrPRMFileAmortSchedPRMParamsInfoXML)
    
    'Get the ID attribute value of the root tag
    lstrRootTagIDAttribVal = lobjRequestInfoXMLDOM.documentElement.getAttributeNode("ID").Value
    
    'Load the response XML
    Call lobjResponseXMLDOM.loadXML("<PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2 ID=""" & lstrRootTagIDAttribVal & """></PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2>")
    
    'Load the PRM Info XML which will be send as input to the GeneratePRMFilesForPmtStruct2() method
    Call lobjPRMInfoXMLDOM.loadXML("<PRM_INFO></PRM_INFO>")
    
    'Load the PRM Params Info XML which will be send as input to the GetPRMParams() method
    Call lobjPRMParamsInfoXMLDOM.loadXML("<PRM_PARAMS_INFO></PRM_PARAMS_INFO>")
    
    'For each set of PRM parameters in the Input XML
    For liCnt = 0 To lobjRequestInfoXMLDOM.documentElement.childNodes.length - 1
        
        'Add the PRM parameters to the PRM Info XML
        AddXMLElement lobjPRMInfoXMLDOM, _
                        lobjPRMInfoXMLDOM.documentElement, _
                        "PRM_FILE", _
                        ""
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_META_DATA").cloneNode(True)
        lobjPRMInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_PARAMS").cloneNode(True)
        lobjPRMInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Add the PRM parameters Parent tag to the response XML
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).cloneNode(False)
        lobjResponseXMLDOM.documentElement.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Add the PRM Param Info to the PRM Params Info XML
        AddXMLElement lobjPRMParamsInfoXMLDOM, _
                        lobjPRMParamsInfoXMLDOM.documentElement, _
                        "PRM_PARAMS", _
                        ""
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_PARAMS_SPECS").cloneNode(True)
        lobjPRMParamsInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        AddXMLElement lobjPRMParamsInfoXMLDOM, _
                        lobjPRMParamsInfoXMLDOM.documentElement.lastChild, _
                        "PRM_FILE", _
                        ""
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_META_DATA/FILE_NAME").cloneNode(True)
        lobjPRMParamsInfoXMLDOM.documentElement.lastChild.selectSingleNode("PRM_FILE").appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
    Next
        
    'Call the GeneratePRMFilesForPmtStruct2() method to get the binary PRM file List XML
    '(Using Solve by Rate method)
    lstrPRMFileLstXML = lobjSTSvc.GeneratePRMFilesForPmtStruct2(lobjPRMInfoXMLDOM.xml)
    
    Set lobjPRMInfoXMLDOM = Nothing
    
    Call lobjPRMFileLstXMLDOM.loadXML(lstrPRMFileLstXML)
    
    'For each binary PRM file in the PRM file List XML
    liCnt = 0
    liPRMFileIndex = 0
    lstrTotalPRMFiles = lobjPRMFileLstXMLDOM.documentElement.childNodes.length - 1
    While liCnt <= lstrTotalPRMFiles
        
        'Add the binary PRM file to the response XML
        Set lobjCloneNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).cloneNode(True)
        lobjResponseXMLDOM.documentElement.childNodes(liCnt).appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Check if there was an error generating the PRM file
        Set lobjErrorNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("ERROR")
        If Not (lobjErrorNode Is Nothing) Then
            
            'If there was an error remove the binary PRM file from the PRM file List XML
            lobjPRMFileLstXMLDOM.documentElement.removeChild lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex)
            liCnt = liCnt + 1
        Else
            'Get the PRM file name in the PRM file List XML
            lstrPRMFileName = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("FILE_NAME").Text
            
            'Add the binary PRM file to the PRM Params Info XML
            Set lobjCloneNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("FILE_DATA").cloneNode(True)
            lobjPRMParamsInfoXMLDOM.documentElement.selectSingleNode("//PRM_PARAMS/PRM_FILE[FILE_NAME=""" & lstrPRMFileName & """]").appendChild lobjCloneNode
            Set lobjCloneNode = Nothing
            
            liCnt = liCnt + 1
            liPRMFileIndex = liPRMFileIndex + 1
        End If
        Set lobjErrorNode = Nothing
    Wend
    
    'Check if the PRM file list XML is not empty
    If lobjPRMFileLstXMLDOM.documentElement.childNodes.length > 0 Then
    
        'Get the Amortization Schedule list for the PRM file list XML
        lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMFileLstXMLDOM.xml)
        
        'Load the return Amortization Schedule list
        Call lobjAmortSchedLstXMLDOM.loadXML(lstrAmortSchedLstXML)
        
        'For each Amortization Schedule in the list
        For liCnt = 0 To lobjAmortSchedLstXMLDOM.documentElement.childNodes.length - 1
            
            'Get the PRM file name in the amortization schedule data
            lstrPRMFileName = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE_NAME").Text
            
            'Locate the PRM file in the Response XML
            Set lobjPRMFileAndAmortSchedNode = lobjResponseXMLDOM.documentElement.selectSingleNode("/PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2/PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS2[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]")
            If Not (lobjPRMFileAndAmortSchedNode Is Nothing) Then
                
                'Add the amortization schedule to the Response XML
                Set lobjCloneNode = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).cloneNode(True)
                lobjPRMFileAndAmortSchedNode.appendChild lobjCloneNode
                Set lobjCloneNode = Nothing
            End If
            
            Set lobjPRMFileAndAmortSchedNode = Nothing
        Next
        
        Set lobjAmortSchedLstXMLDOM = Nothing
        
        'Get the PRM parameters specified in the PRM Params Info XML
        lstrPRMParamsLstXML = lobjSTSvc.GetPRMParams(lobjPRMParamsInfoXMLDOM.xml)
    
        Set lobjPRMParamsInfoXMLDOM = Nothing
    
        Call lobjPRMParamsLstXMLDOM.loadXML(lstrPRMParamsLstXML)
    
        'For each PRM parameter set in the return XML
        For liCnt = 0 To lobjPRMParamsLstXMLDOM.documentElement.childNodes.length - 1
            
            'Get the PRM file name in the amortization schedule data
            lstrPRMFileName = lobjPRMParamsLstXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE_NAME").Text
            
            'Add the PRM parameter set to the response XML
            Set lobjCloneNode = lobjPRMParamsLstXMLDOM.documentElement.childNodes(liCnt).cloneNode(True)
            lobjResponseXMLDOM.documentElement.selectSingleNode("//PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS2[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]").appendChild lobjCloneNode
            Set lobjCloneNode = Nothing
        Next
        
        Set lobjPRMParamsLstXMLDOM = Nothing
    
    End If
CleanUp:

    'Return the response XML
    GetPRMFileAmortSchedAndPRMParams2 = lobjResponseXMLDOM.xml
        
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMFileAmortSchedAndPRMParams2(): Return value: " & GetPRMFileAmortSchedAndPRMParams2
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMFileAmortSchedAndPRMParams2(): Exit GetPRMFileAmortSchedAndPRMParams2() method"
    
    Set lobjPRMFileLstXMLDOM = Nothing
    Set lobjRequestInfoXMLDOM = Nothing
    Set lobjResponseXMLDOM = Nothing
    
    Exit Function
    
ErrHandler:
    Set lobjResponseXMLDOM = Nothing
    
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description
    
    Call lobjResponseXMLDOM.loadXML("<PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2 ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_FILE_AMORT_SCHED_AND_PRM_PARAMS_INFO2>")
        
    AddXMLElement lobjResponseXMLDOM, _
                lobjResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_NBR", _
                llErrNbr
                
    AddXMLElement lobjResponseXMLDOM, _
                lobjResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_DESC", _
                lstrErrDesc
                
    Resume CleanUp
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
On Error GoTo ErrHandler
    
Dim lobjSTSvc                               As New BSCEFSuperTrump.ISuperTrumpService
    
'MS XML DOM Objects Declarations
Dim lobjRequestInfoXMLDOM                   As New DOMDocument40
Dim lobjResponseXMLDOM                      As New DOMDocument40
Dim lobjPRMInfoXMLDOM                       As New DOMDocument40
Dim lobjPRMFileLstXMLDOM                    As New DOMDocument40
Dim lobjErrorNode                           As IXMLDOMNode
Dim lobjAmortSchedLstXMLDOM                 As New DOMDocument40
Dim lobjPRMFileAndAmortSchedNode            As IXMLDOMNode
Dim lobjCloneNode                           As IXMLDOMNode
Dim lobjPRMParamsInfoXMLDOM                 As New DOMDocument40
Dim lobjPRMParamsLstXMLDOM                  As New DOMDocument40
Dim lobjBinaryPRMFile                       As IXMLDOMNode

'Other Declarations
Dim llErrNbr                                As Long
Dim lstrErrDesc                             As String
Dim lstrRootTagIDAttribVal                  As String
Dim liCnt                                   As Integer
Dim liPRMFileIndex                          As Integer
Dim lstrPRMFileLstXML                       As String
Dim lstrAmortSchedLstXML                    As String
Dim lstrPRMFileName                         As String
Dim lstrTotalPRMFiles                       As String
Dim lstrPRMParamsLstXML                     As String

    'Load the Input XML
    Call lobjRequestInfoXMLDOM.loadXML(astrPRMFileAmortSchedPRMParamsInfoXML)
    
    'Get the ID attribute value of the root tag
    lstrRootTagIDAttribVal = lobjRequestInfoXMLDOM.documentElement.getAttributeNode("ID").Value
    
    'Load the response XML
    Call lobjResponseXMLDOM.loadXML("<PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO ID=""" & lstrRootTagIDAttribVal & """></PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO>")
    
    'Load the PRM Info XML which will be send as input to the GeneratePRMFilesForPmtStruct() method
    Call lobjPRMInfoXMLDOM.loadXML("<PRM_INFO></PRM_INFO>")
    
    'Load the PRM Params Info XML which will be send as input to the GetPRMParams() method
    Call lobjPRMParamsInfoXMLDOM.loadXML("<PRM_PARAMS_INFO></PRM_PARAMS_INFO>")
    
    'For each set of PRM parameters in the Input XML
    For liCnt = 0 To lobjRequestInfoXMLDOM.documentElement.childNodes.length - 1
        
        'Add the PRM parameters to the PRM Info XML
        AddXMLElement lobjPRMInfoXMLDOM, _
                        lobjPRMInfoXMLDOM.documentElement, _
                        "PRM_FILE", _
                        ""
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_META_DATA").cloneNode(True)
        lobjPRMInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_PARAMS").cloneNode(True)
        lobjPRMInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Add the PRM parameters Parent tag to the response XML
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).cloneNode(False)
        lobjResponseXMLDOM.documentElement.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Add the PRM Param Info to the PRM Params Info XML
        AddXMLElement lobjPRMParamsInfoXMLDOM, _
                        lobjPRMParamsInfoXMLDOM.documentElement, _
                        "PRM_PARAMS", _
                        ""
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_PARAMS_SPECS").cloneNode(True)
        lobjPRMParamsInfoXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        AddXMLElement lobjPRMParamsInfoXMLDOM, _
                        lobjPRMParamsInfoXMLDOM.documentElement.lastChild, _
                        "PRM_FILE", _
                        ""
        
        Set lobjCloneNode = lobjRequestInfoXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_META_DATA/FILE_NAME").cloneNode(True)
        lobjPRMParamsInfoXMLDOM.documentElement.lastChild.selectSingleNode("PRM_FILE").appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
    Next
        
    'Call the GeneratePRMFilesForPmtStruct() method to get the binary PRM file List XML
    '(Using Solve by payments method)
    lstrPRMFileLstXML = lobjSTSvc.GeneratePRMFilesForPmtStruct(lobjPRMInfoXMLDOM.xml)
    
    Set lobjPRMInfoXMLDOM = Nothing
    
    Call lobjPRMFileLstXMLDOM.loadXML(lstrPRMFileLstXML)
    
    'For each binary PRM file in the PRM file List XML
    liCnt = 0
    liPRMFileIndex = 0
    lstrTotalPRMFiles = lobjPRMFileLstXMLDOM.documentElement.childNodes.length - 1
    While liCnt <= lstrTotalPRMFiles
        
        'Add the binary PRM file to the response XML
        Set lobjCloneNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).cloneNode(True)
        lobjResponseXMLDOM.documentElement.childNodes(liCnt).appendChild lobjCloneNode
        Set lobjCloneNode = Nothing
        
        'Remove the binary PRM file.
        Set lobjBinaryPRMFile = lobjResponseXMLDOM.documentElement.childNodes(liCnt).childNodes(0).selectSingleNode("FILE_DATA")
        If Not (lobjBinaryPRMFile Is Nothing) Then
            lobjResponseXMLDOM.documentElement.childNodes(liCnt).childNodes(0).removeChild lobjBinaryPRMFile
        End If
        
        'Check if there was an error generating the PRM file
        Set lobjErrorNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("ERROR")
        If Not (lobjErrorNode Is Nothing) Then
            
            'If there was an error remove the binary PRM file from the PRM file List XML
            lobjPRMFileLstXMLDOM.documentElement.removeChild lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex)
            liCnt = liCnt + 1
        Else
            'Get the PRM file name in the PRM file List XML
            lstrPRMFileName = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("FILE_NAME").Text
            
            'Add the binary PRM file to the PRM Params Info XML
            Set lobjCloneNode = lobjPRMFileLstXMLDOM.documentElement.childNodes(liPRMFileIndex).selectSingleNode("FILE_DATA").cloneNode(True)
            lobjPRMParamsInfoXMLDOM.documentElement.selectSingleNode("//PRM_PARAMS/PRM_FILE[FILE_NAME=""" & lstrPRMFileName & """]").appendChild lobjCloneNode
            Set lobjCloneNode = Nothing
            
            liCnt = liCnt + 1
            liPRMFileIndex = liPRMFileIndex + 1
        End If
        Set lobjErrorNode = Nothing
    Wend
    
    'Check if the PRM file list XML is not empty
    If lobjPRMFileLstXMLDOM.documentElement.childNodes.length > 0 Then
    
        'Get the Amortization Schedule list for the PRM file list XML
        lstrAmortSchedLstXML = lobjSTSvc.GetAmortizationSchedule(lobjPRMFileLstXMLDOM.xml)
        
        'Load the return Amortization Schedule list
        Call lobjAmortSchedLstXMLDOM.loadXML(lstrAmortSchedLstXML)
        
        'For each Amortization Schedule in the list
        For liCnt = 0 To lobjAmortSchedLstXMLDOM.documentElement.childNodes.length - 1
            
            'Get the PRM file name in the amortization schedule data
            lstrPRMFileName = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE_NAME").Text
            
            'Locate the PRM file in the Response XML
            Set lobjPRMFileAndAmortSchedNode = lobjResponseXMLDOM.documentElement.selectSingleNode("/PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO/PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]")
            If Not (lobjPRMFileAndAmortSchedNode Is Nothing) Then
                
                'Add the amortization schedule to the Response XML
                Set lobjCloneNode = lobjAmortSchedLstXMLDOM.documentElement.childNodes(liCnt).cloneNode(True)
                lobjPRMFileAndAmortSchedNode.appendChild lobjCloneNode
                Set lobjCloneNode = Nothing
            End If
            
            Set lobjPRMFileAndAmortSchedNode = Nothing
        Next
        
        Set lobjAmortSchedLstXMLDOM = Nothing
        
        'Get the PRM parameters specified in the PRM Params Info XML
        lstrPRMParamsLstXML = lobjSTSvc.GetPRMParams(lobjPRMParamsInfoXMLDOM.xml)
    
        Set lobjPRMParamsInfoXMLDOM = Nothing
    
        Call lobjPRMParamsLstXMLDOM.loadXML(lstrPRMParamsLstXML)
    
        'For each PRM parameter set in the return XML
        For liCnt = 0 To lobjPRMParamsLstXMLDOM.documentElement.childNodes.length - 1
            
            'Get the PRM file name in the amortization schedule data
            lstrPRMFileName = lobjPRMParamsLstXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE_NAME").Text
            
            'Add the PRM parameter set to the response XML
            Set lobjCloneNode = lobjPRMParamsLstXMLDOM.documentElement.childNodes(liCnt).cloneNode(True)
            lobjResponseXMLDOM.documentElement.selectSingleNode("//PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE[PRM_FILE/FILE_NAME=""" & lstrPRMFileName & """]").appendChild lobjCloneNode
            Set lobjCloneNode = Nothing
        Next
        
        Set lobjPRMParamsLstXMLDOM = Nothing
    
    End If
CleanUp:

    'Return the response XML
    GetPRMParamsAmortSchedAndNoPRMFile = lobjResponseXMLDOM.xml
        
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMParamsAmortSchedAndNoPRMFile(): Return value: " & GetPRMParamsAmortSchedAndNoPRMFile
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.IClientService_GetPRMParamsAmortSchedAndNoPRMFile(): Exit GetPRMParamsAmortSchedAndNoPRMFile() method"
    
    Set lobjPRMFileLstXMLDOM = Nothing
    Set lobjRequestInfoXMLDOM = Nothing
    Set lobjResponseXMLDOM = Nothing
    
    Exit Function
    
ErrHandler:
    Set lobjResponseXMLDOM = Nothing
    
    llErrNbr = Err.Number
    lstrErrDesc = Err.Description
    
    Call lobjResponseXMLDOM.loadXML("<PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO ID=""" & lstrRootTagIDAttribVal & """><ERROR></ERROR></PRM_PARAMS_AMORT_SCHED_AND_NO_PRM_FILE_INFO>")
        
    AddXMLElement lobjResponseXMLDOM, _
                lobjResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_NBR", _
                llErrNbr
                
    AddXMLElement lobjResponseXMLDOM, _
                lobjResponseXMLDOM.documentElement.childNodes(0), _
                "ERROR_DESC", _
                lstrErrDesc
                
    Resume CleanUp
End Function


Private Sub Class_Initialize()
        
On Error Resume Next

    'Get Debugging parameters from registry
    giDebugLevel = Val(ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\Debug\DebugLevel"))
    If giDebugLevel > 0 Then
        gstrDebugFile = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\Debug\DebugFile")
        glMaxDebugFileSize = Val(ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\Debug\MaxDebugFileSize"))
    End If
End Sub
