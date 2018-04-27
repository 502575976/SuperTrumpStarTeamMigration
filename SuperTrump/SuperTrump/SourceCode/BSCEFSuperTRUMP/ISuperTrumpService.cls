VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ISuperTrumpService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'================================================================
'MODULE  : ISuperTrumpService
'PURPOSE : This interface provides all functionality relating to
'          retrieval of pricing and PRM file data as specified
'          by the client applications.
'================================================================
Option Explicit

Private Enum eSolveMethod
    ecSolveForPayments = 0
    ecSolveForRate = 1
End Enum

Private Const cADHOC_QUERY_RESULT_XML As String = "<PRM_INFO><PRM_FILE><AD_HOC_QUERY></AD_HOC_QUERY></PRM_FILE></PRM_INFO>"

'================================================================
'METHOD  : ConvertPRMToXML
'PURPOSE : To convert the binary PRM file(s) to their XML
'          equivalent.
'PARMS   :
'          astrPRMFileListXML [String] = XML string containing
'          the binary PRM files. This XML will conform to the
'          PRMFileListXML.xsd schema.
'
'          Sample Input Parameter structure:
'           <PRM_FILE_LIST>
'               <PRM_FILE>
'                   <FILE_NAME>…</FILE_NAME>
'                   <FILE_DATA>…</FILE_ DATA>
'               </PRM_FILE>
'               <PRM_FILE>
'                   <FILE_NAME>…</FILE_NAME>
'                   <FILE_DATA>…</FILE_ DATA>
'               </PRM_FILE>
'               …
'           </PRM_FILE_LIST>
'RETURN  : String = An XML string containing XML equivalent for
'          each PRM File. It will also contain an error message
'          for each erroneous PRM File.
'
'          Sample Return XML structure:
'           <PRM_FILE_LIST>
'               <PRM_FILE>
'                   <FILE_NAME>…</FILE_NAME>
'                   <PRM_XML>…</PRM_XML>
'               </PRM_FILE>
'               <PRM_FILE>
'                   <FILE_NAME>…</FILE_NAME>
'                   <ERROR>
'                       <ERROR_NBR>…</ERROR_NBR>
'                       <ERROR_DESC>…</ERROR_DESC>
'                   </ERROR>
'               </PRM_FILE>
'               …
'           </PRM_FILE_LIST>
'
'           OR in case of application error
'
'           <PRM_FILE_LIST>
'               <ERROR>
'                   <ERROR_NBR>…</ERROR_NBR>
'                   <ERROR_DESC>…</ERROR_DESC>
'               </ERROR>
'           </PRM_FILE_LIST>
'================================================================
Function ConvertPRMToXML(ByVal astrPRMFileListXML As String) As String
On Error GoTo ErrHandler

'Super Trump Server Objects Declarations
Dim lobjSTApplication       As New STApplication

'MS XML DOM Objects Declarations
Dim lobjPRMXMLDOM           As New DOMDocument40
Dim lobjPRMExceptionXMLDOM  As New DOMDocument40
Dim lobjFileNameList        As IXMLDOMNodeList
Dim lobjFileDataList        As IXMLDOMNodeList
Dim lobjExecpList           As IXMLDOMNodeList
Dim lobjXMLSchemaSpace      As New XMLSchemaCache40
Dim lobjPRMFileNameDOM      As New DOMDocument40
Dim lobjElem                As IXMLDOMElement

'Other Declarations
Dim lintCtrLoop             As Integer
Dim lstrPRM2XML             As String
Dim lstrFileName            As String
Dim lstrPRMFileName         As String
Dim lstrFilePath            As String
Dim lstrSupertrumpQuery     As String
Dim lstrPRMBIN2XML          As String
Dim lstrFileLoc             As String
Dim lstrPRMFileListXML      As String

Dim llErrNbr                As Long
Dim lstrErrSrc              As String
Dim lstrErrDesc             As String

Dim lobjErrComment          As IXMLDOMNode
Dim lstrErrComment          As String

    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): In ConvertPRMToXML() method"
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Input Argument 1:" & astrPRMFileListXML
    
    'Get the PRMFileListXML.xsd Schema
    lstrFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcSchemaFilePathKey)
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Schema file read from registry"
    
    Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPRMFileLstXMLSchemaName)
    
    'Assign Schema to the XML DOM object
    Set lobjPRMXMLDOM.schemas = lobjXMLSchemaSpace
    lobjPRMXMLDOM.async = False
    
    'Remove "xmlns:dt" & "dt:dt" attributes from the <FILE_DATA> elements in the Input XML
    'as we cannot include these attributes in the schema as they contain special
    'characters(:). We will add them after schema validation.
    lstrPRMFileListXML = Replace(astrPRMFileListXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
        
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Validating Input XML"
    
    'Load the Input XML into the XML DOM object
    Call lobjPRMXMLDOM.loadXML(lstrPRMFileListXML)
    
    'Check , if any parser error has occured (Validate Input XML)
    If lobjPRMXMLDOM.parseError.errorCode = 0 Then
        
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Input XML Valid"
        
        'Get the FileName and FileData list from DOM
        Set lobjFileNameList = lobjPRMXMLDOM.getElementsByTagName("FILE_NAME")
        Set lobjFileDataList = lobjPRMXMLDOM.getElementsByTagName("FILE_DATA")
        
        lstrFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMFilePathKey)
        
        'Traverse Each PRM File.
        For lintCtrLoop = 0 To (lobjFileNameList.length - 1)
        
            lstrFileName = lobjFileNameList.Item(lintCtrLoop).Text
            If UCase(Right(lstrFileName, 4)) <> ".PRM" Then
                lstrPRMFileName = lstrFileName & ".PRM"
            Else
                lstrPRMFileName = lstrFileName
            End If
            
            'Save the PRM Binary data to a physical file.
            lobjFileDataList.Item(lintCtrLoop).dataType = "bin.base64"
            lstrFilePath = SavePRMBinaryFile(lstrPRMFileName, lstrFileLoc, lobjFileDataList.Item(lintCtrLoop).nodeTypedValue)
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): " & lstrPRMFileName & " saved to " & lstrFileLoc
            
            'Create Supertrump XML Query.
            lstrSupertrumpQuery = GetSuperTrumpQuery("TRANS_ID_" & Mid(lstrPRMFileName, 1, InStrRev(UCase(lstrPRMFileName), ".PRM") - 1), lstrFilePath)
              
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Calling XMLInOut() method with input - " & lstrSupertrumpQuery
            
            'Get XML representation for the PRM file.
            lstrPRMBIN2XML = lobjSTApplication.XmlInOut(lstrSupertrumpQuery)
                        
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): XML representation for the PRM file returned."
            
            'Load the XML representation into DOM
            Call lobjPRMExceptionXMLDOM.loadXML(lstrPRMBIN2XML)
            
            'Check for any Exception.
            If (lobjPRMExceptionXMLDOM.getElementsByTagName("Exception").length) > 0 Then
                
                'Create the <ERROR> node for the exception
                Set lobjExecpList = lobjPRMExceptionXMLDOM.getElementsByTagName("Exception")
                Set lobjErrComment = lobjExecpList.Item(0).selectSingleNode("Comment")
                
                lstrErrComment = ""
                If Not (lobjErrComment Is Nothing) Then lstrErrComment = lobjErrComment.Text
                
                
                lstrPRMBIN2XML = "<ERROR>" & _
                                    "<ERROR_NBR>" & lobjExecpList.Item(0).selectSingleNode("Number").Text & "</ERROR_NBR>" & _
                                    "<ERROR_DESC><![CDATA[" & lobjExecpList.Item(0).selectSingleNode("Description").Text & " " & lstrErrComment & "]]></ERROR_DESC>" & _
                                    "<PRM_XML>" & lobjPRMExceptionXMLDOM.documentElement.xml & "</PRM_XML>" & _
                                 "</ERROR>"
                
                If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Exception from STServer : " & lobjPRMExceptionXMLDOM.selectSingleNode("//Exceptions").xml
                
            'Else if no exception
            Else
                'Create <PRM_XML> node
                lstrPRMBIN2XML = "<PRM_XML>" & lobjPRMExceptionXMLDOM.documentElement.xml & "</PRM_XML>"
            End If
                
            'Create <PRM_FILE> node
            If lobjPRMFileNameDOM.loadXML("<PRM_FILE>" & lstrPRMBIN2XML & "</PRM_FILE>") Then
                Set lobjElem = lobjPRMFileNameDOM.createElement("FILE_NAME")
                lobjElem.Text = lstrFileName
                lobjPRMFileNameDOM.documentElement.insertBefore lobjElem, lobjPRMFileNameDOM.documentElement.childNodes(0)
                Set lobjElem = Nothing
                
                lstrPRM2XML = lstrPRM2XML & lobjPRMFileNameDOM.xml
            End If
            Set lobjPRMFileNameDOM = Nothing
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Data added to the output XML."
        Next
        
        'Create <PRM_FILE_LIST> node
        lstrPRM2XML = "<PRM_FILE_LIST>" & lstrPRM2XML & "</PRM_FILE_LIST>"
        
    'Else if Input XML is invalid
    Else
    
        'Create the return XML with <ERROR> node
        lstrPRM2XML = "<PRM_FILE_LIST>" & _
                            "<ERROR>" & _
                                "<ERROR_NBR>" & lobjPRMXMLDOM.parseError.errorCode & "</ERROR_NBR>" & _
                                "<ERROR_DESC><![CDATA[Error on line number " & lobjPRMXMLDOM.parseError.Line & " of the XML." & lobjPRMXMLDOM.parseError.reason & "]]></ERROR_DESC>" & _
                            "</ERROR>" & _
                       "</PRM_FILE_LIST>"
                                
    End If
    
    'Return the Final XML
    ConvertPRMToXML = lstrPRM2XML
    
    'Delete the saved PRM files
    Call DeletePRMBinaryFile(lobjPRMXMLDOM.getElementsByTagName("FILE_NAME"), ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMFilePathKey))
        
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): PRM files deleted from temp location."
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Return value : " & ConvertPRMToXML
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Exit ConvertPRMToXML() method"
    
CleanUp:

    Set lobjSTApplication = Nothing
    Set lobjPRMXMLDOM = Nothing
    Set lobjPRMExceptionXMLDOM = Nothing
    Set lobjFileNameList = Nothing
    Set lobjFileDataList = Nothing
    Set lobjExecpList = Nothing
    Set lobjXMLSchemaSpace = Nothing
    
    Exit Function
    
'In error handler
ErrHandler:
    
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description
    
    'Return the Final XML with <ERROR> node specifying the application error
    ConvertPRMToXML = "<PRM_FILE_LIST>" & _
                             "<ERROR>" & _
                                 "<ERROR_NBR>" & llErrNbr & "</ERROR_NBR>" & _
                                 "<ERROR_DESC><![CDATA[" & lstrErrDesc & "]]></ERROR_DESC>" & _
                             "</ERROR>" & _
                        "</PRM_FILE_LIST>"
    
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): General Error : " & ConvertPRMToXML
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Exit ConvertPRMToXML() method"
    
    Resume CleanUp
End Function

'================================================================
'METHOD  : GetPricingReports
'PURPOSE : To get the specified report(s) for the inputted binary
'          PRM file(s).
'PARMS   :
'          astrPricingRepInfoXML [String] = XML string containing
'          the binary PRM files and information specifying what
'          report(s) needs to be generated for each PRM file.
'          This XML will conform to the PricingRepInfoXML.xsd
'          schema.
'
'          Sample Input Parameter structure:
'           <PRICING_REPORT_INFO>
'               <PRICING_REPORT>
'                   <PRM_FILE>
'                       <FILE_NAME>…</FILE_NAME>
'                       <FILE_DATA>…</FILE_ DATA>
'                   </PRM_FILE>
'                   <REPORT_LIST>
'                       <REPORT_TYPE>…</REPORT_TYPE>
'                       <REPORT_TYPE>…</REPORT_TYPE>
'                       …
'                   </REPORT_LIST>
'               </PRICING_REPORT>
'               …
'           </PRICING_REPORT_INFO>
'RETURN  : String = An XML string containing the pricing reports
'          for each PRM File. It will also contain an error
'          message for each erroneous PRM File and each pricing
'          reports, which couldn't be generated.
'
'          Sample Return XML structure:
'           <PRICING_REPORT_LIST>
'               <PRICING_REPORT>
'                   <PRM_FILE_NAME>…</ PRM_FILE_NAME>
'                   <REPORT_LIST>
'                       <REPORT>
'                           <REPORT_TYPE>…</REPORT_TYPE>
'                           <TEXT_REPORT>…</TEXT_REPORT>
'                       </REPORT>
'                       <REPORT>
'                           <REPORT_TYPE>…</REPORT_TYPE>
'                           <TEXT_REPORT>…</TEXT_REPORT>
'                       </REPORT>
'                       …
'                   </REPORT_LIST>
'               </PRICING_REPORT>
'               <PRICING_REPORT>
'                   <PRM_FILE_NAME>…</ PRM_FILE_NAME>
'                   <ERROR>
'                       <ERROR_NBR>…</ERROR_NBR>
'                       <ERROR_DESC>…</ERROR_DESC>
'                   </ERROR>
'               </PRICING_REPORT>
'               <PRICING_REPORT>
'                   <PRM_FILE_NAME>…</ PRM_FILE_NAME>
'                   <REPORT_LIST>
'                       <REPORT>
'                           <REPORT_TYPE>…</REPORT_TYPE>
'                           <TEXT_REPORT>…</TEXT_REPORT>
'                       </REPORT>
'                       <REPORT>
'                           <REPORT_TYPE>…</REPORT_TYPE>
'                           <ERROR>
'                               <ERROR_NBR>…</ERROR_NBR>
'                               <ERROR_DESC>…</ERROR_DESC>
'                           </ERROR>
'                       </REPORT>
'                       …
'                   </REPORT_LIST>
'               </PRICING_REPORT>
'               …
'           </PRICING_REPORT_LIST>
'
'           OR in case of application error
'
'           <PRICING_REPORT_LIST>
'               <ERROR>
'                   <ERROR_NBR>…</ERROR_NBR>
'                   <ERROR_DESC>…</ERROR_DESC>
'               </ERROR>
'           </PRICING_REPORT_LIST>
'================================================================
Public Function GetPricingReports(ByVal astrPricingRepInfoXML As String) As String
On Error GoTo ErrHandler

'Declare Super Trump Variables
Dim lobjReportSTTrans               As New STTransaction
Dim lobjReportSTResults             As STResults
Dim lobjSTApp                       As New STApplication

'Declare XML Dom variables
Dim lobjPRICING_REPORT_INFO_DOM     As New DOMDocument40
Dim lobjPRMBIN2XML_DOM              As New DOMDocument40
Dim lobjPRICING_REPORTS             As IXMLDOMNodeList
Dim lobjPRICING_REPORT              As IXMLDOMNode
Dim lobjFILE_NAME                   As IXMLDOMNode
Dim lobjFILE_DATA                   As IXMLDOMNode
Dim lobjREPORT_TYPES                As IXMLDOMNodeList
Dim lobjREPORT_TYPE                 As IXMLDOMNode
Dim lobjXMLSchemaSpace              As New XMLSchemaCache40
Dim lobjDOM                         As New DOMDocument40
Dim lobjElem                        As IXMLDOMElement

'Other Declarations
Dim lstrReport                      As String
Dim lbReportProcessFlag             As Boolean
Dim lbReportProcessErrFlag          As Boolean
Dim lstrPRM2ReportXML               As String
Dim lstrPRICING_REPORT              As String
Dim lstrSupertrumpQuery             As String
Dim lstrFilePath                    As String
Dim lstrPRMBIN2XML                  As String
Dim lstrPRM_Err_Mesage              As String
Dim lstrErrCode                     As String
Dim lstrFileLoc                     As String
Dim lstrReportTemplateLoc           As String
Dim lstrFileName                    As String
Dim lstrPRMFileName                 As String
Dim lstrPricingRepInfoXML           As String

Dim llErrNbr                As Long
Dim lstrErrSrc              As String
Dim lstrErrDesc             As String


    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): In GetPricingReports() method"
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): Input Argument 1:" & astrPricingRepInfoXML

    'Get the PricingRepInfoXML.xsd Schema
    lstrFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcSchemaFilePathKey)
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): Schema file read from registry"
    
    Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPricingRepInfoXMLSchemaName)
    
    'Assign Schema to the XML DOM object
    Set lobjPRICING_REPORT_INFO_DOM.schemas = lobjXMLSchemaSpace
    lobjPRICING_REPORT_INFO_DOM.async = False
    
    'Remove "xmlns:dt" & "dt:dt" attributes from the <FILE_DATA> elements in the Input XML
    'as we cannot include these attributes in the schema as they contain special
    'characters(:). We will add them after schema validation.
    lstrPricingRepInfoXML = Replace(astrPricingRepInfoXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): Validating Input XML"
    
    'Load the Input XML into the XML DOM object
    Call lobjPRICING_REPORT_INFO_DOM.loadXML(lstrPricingRepInfoXML)
    
    'Check , if any parser error has occured (Validate Input XML)
    If lobjPRICING_REPORT_INFO_DOM.parseError.errorCode = 0 Then
            
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): Input XML Valid"
        
        'Get the Pricing Report List
        Set lobjPRICING_REPORTS = lobjPRICING_REPORT_INFO_DOM.getElementsByTagName("PRICING_REPORT")
        
        lstrFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMFilePathKey)
        
        lstrReportTemplateLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcReportTemplatePathKey)
        
        'Traverse Each Pricing Report
        For Each lobjPRICING_REPORT In lobjPRICING_REPORTS
        
            'Get the File Name & File data for Each Pricing Report
            Set lobjFILE_NAME = lobjPRICING_REPORT.selectSingleNode("PRM_FILE/FILE_NAME")
            Set lobjFILE_DATA = lobjPRICING_REPORT.selectSingleNode("PRM_FILE/FILE_DATA")
            
            'Save the PRM Binary data to a physical file.
            lstrFileName = lobjFILE_NAME.Text
            If UCase(Right(lstrFileName, 4)) <> ".PRM" Then
                lstrPRMFileName = lstrFileName & ".PRM"
            Else
                lstrPRMFileName = lstrFileName
            End If
            
            lobjFILE_DATA.dataType = "bin.base64"
            lstrFilePath = SavePRMBinaryFile(lstrPRMFileName, lstrFileLoc, lobjFILE_DATA.nodeTypedValue)
                
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): " & lstrPRMFileName & " saved to " & lstrFileLoc
            
            'Get the Report List for the PRM File
            Set lobjREPORT_TYPES = lobjPRICING_REPORT.selectNodes("REPORT_LIST/REPORT_TYPE")
            
            'Open the .PRM file
            Call lobjReportSTTrans.OpenFile(lstrFilePath)
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): PRM file opened."
            
            lobjReportSTTrans.Calculate
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): PRM file recalculated."
            
            Set lobjReportSTResults = lobjReportSTTrans.Results
                 
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): STResult object created."
            
            'Traverse Each Report List.
            For Each lobjREPORT_TYPE In lobjREPORT_TYPES
             
                 lbReportProcessFlag = True
                 
                 If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): report template name - " & lobjREPORT_TYPE.Text
                 
                 'Assign the Report File Name
                 lobjReportSTResults.ReportFileName = lstrReportTemplateLoc & "\" & lobjREPORT_TYPE.Text
                 
                 'Retrieve the Report
                 lstrReport = lobjReportSTResults.PrintBuffer
                 
                 If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): report retrieved."
                 
                 'Create the <REPORT> Node.
                 If lobjDOM.loadXML("<REPORT></REPORT>") Then
                    Set lobjElem = lobjDOM.createElement("REPORT_TYPE")
                    lobjElem.Text = lobjREPORT_TYPE.Text
                    lobjDOM.documentElement.appendChild lobjElem
                    Set lobjElem = Nothing
                    
                    Set lobjElem = lobjDOM.createElement("TEXT_REPORT")
                    lobjElem.Text = lstrReport
                    lobjDOM.documentElement.appendChild lobjElem
                    Set lobjElem = Nothing
                    
                    lstrPRM2ReportXML = lstrPRM2ReportXML & lobjDOM.xml
                 End If
                 Set lobjDOM = Nothing
                 
NEXT_REPORT_TYPE:
                
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): Report Data added to the output XML."
                
            Next
                
            Set lobjReportSTResults = Nothing
            Set lobjReportSTTrans = Nothing
                
            'Create the <PRICING_REPORT> node
            'lstrPRICING_REPORT = lstrPRICING_REPORT & "<PRICING_REPORT>" & _
                                                        "<PRM_FILE_NAME><![CDATA[" & lobjFILE_NAME.Text & "]]></PRM_FILE_NAME>" & _
                                                        "<REPORT_LIST><![CDATA[" & lstrPRM2ReportXML & "]]></REPORT_LIST>" & _
                                                    "</PRICING_REPORT>"
            
            If lobjDOM.loadXML("<PRICING_REPORT><REPORT_LIST>" & lstrPRM2ReportXML & "</REPORT_LIST></PRICING_REPORT>") Then
               Set lobjElem = lobjDOM.createElement("PRM_FILE_NAME")
               lobjElem.Text = lobjFILE_NAME.Text
               lobjDOM.documentElement.insertBefore lobjElem, lobjDOM.documentElement.childNodes(0)
               Set lobjElem = Nothing
               
               lstrPRICING_REPORT = lstrPRICING_REPORT & lobjDOM.xml
            End If
            Set lobjDOM = Nothing
                 
            lstrPRM2ReportXML = ""
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): PRM added to the output XML."
            
        Next
    
        'Create the <PRICING_REPORT_LIST> node
        lstrPRICING_REPORT = "<PRICING_REPORT_LIST>" & lstrPRICING_REPORT & "</PRICING_REPORT_LIST>"
        
    'Else if the Input XML is invalid
    Else
        
        'Create the return XML with <ERROR> node
        lstrPRICING_REPORT = "<PRICING_REPORT_LIST>" & _
                                "<ERROR>" & _
                                    "<ERROR_NBR>" & lobjPRICING_REPORT_INFO_DOM.parseError.errorCode & "</ERROR_NBR>" & _
                                    "<ERROR_DESC><![CDATA[Error on line number " & lobjPRICING_REPORT_INFO_DOM.parseError.Line & " of the XML." & lobjPRICING_REPORT_INFO_DOM.parseError.reason & "]]></ERROR_DESC>" & _
                                "</ERROR>" & _
                            "</PRICING_REPORT_LIST>"
    
    End If

    'Return the final XML
    GetPricingReports = lstrPRICING_REPORT
    
    'Delete the saved PRM files
    Call DeletePRMBinaryFile(lobjPRICING_REPORT_INFO_DOM.getElementsByTagName("FILE_NAME"), ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMFilePathKey))
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): PRM files deleted from temp location."
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): Return value: " & GetPricingReports
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): Exit GetPricingReports() method"

CleanUp:
    Set lobjReportSTTrans = Nothing
    Set lobjReportSTResults = Nothing
    Set lobjSTApp = Nothing
    Set lobjPRICING_REPORT_INFO_DOM = Nothing
    Set lobjPRMBIN2XML_DOM = Nothing
    Set lobjPRICING_REPORTS = Nothing
    Set lobjPRICING_REPORT = Nothing
    Set lobjFILE_NAME = Nothing
    Set lobjFILE_DATA = Nothing
    Set lobjREPORT_TYPES = Nothing
    Set lobjREPORT_TYPE = Nothing
    Set lobjXMLSchemaSpace = Nothing
    
    Exit Function
    
'In Error handler
ErrHandler:
        
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description
        
    'Check if error while generating report data
    If lbReportProcessFlag Then
        
        lbReportProcessFlag = False
        
        'Create an <ERROR> node for that <REPORT> node and resume processing of the next report
        lstrPRM2ReportXML = lstrPRM2ReportXML & "<REPORT>" & _
                                                    "<REPORT_TYPE>" & lobjREPORT_TYPE.Text & "</REPORT_TYPE>" & _
                                                    "<ERROR>" & _
                                                        "<ERROR_NBR>" & llErrNbr & "</ERROR_NBR>" & _
                                                        "<ERROR_DESC><![CDATA[" & lstrErrDesc & "]]></ERROR_DESC>" & _
                                                    "</ERROR>" & _
                                                "</REPORT>"
                                                        
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): Report Generation Error - " & lstrErrDesc
        
        lbReportProcessFlag = True
        
        Resume NEXT_REPORT_TYPE
            
    'Else if general application error
    Else
        
        'Return the Final XML with <ERROR> node specifying the application error
        GetPricingReports = "<PRICING_REPORT_LIST>" & _
                                    "<ERROR>" & _
                                        "<ERROR_NBR>" & llErrNbr & "</ERROR_NBR>" & _
                                        "<ERROR_DESC><![CDATA[" & lstrErrDesc & "]]></ERROR_DESC>" & _
                                    "</ERROR>" & _
                                 "</PRICING_REPORT_LIST>"
            
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): General Error : " & GetPricingReports
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPricingReports(): Exit GetPricingReports() method"
    
        Resume CleanUp
    End If
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

'================================================================
'METHOD  : Ping
'PURPOSE : Returns a string that indicates that the component
'          is registered properly.
'PARMS   : NONE
'RETURN  : String
'================================================================

Public Function Ping() As String
    Ping = "Ping request to BSCEFSuperTrump.ISuperTrumpService returned at " & Format(Now, "mm/dd/yyyy Hh:Nn:Ss AM/PM") & " server time."
End Function

'================================================================
'METHOD  : Test
'PURPOSE : Returns a string that this component is able to invoke
'          STServer (Ivory's SuperTrump Server component).
'PARMS   : NONE
'RETURN  : String
'================================================================

Public Function Test() As String
On Error GoTo ErrHandler
Dim lobjSTApp   As New STApplication
    
    Test = "Test to invoke Ivory's SuperTrump Server component - " & lobjSTApp.Version & " (" & lobjSTApp.BuildInfo & ") successful."
    Exit Function
    
ErrHandler:
    Test = "Error occured while invoking STServer - " & Err.Description
End Function
'================================================================
'METHOD  : GeneratePRMFiles
'PURPOSE : To generate binary PRM file for each set of PRM
'          parameters and Meta data.
'PARMS   :
'          astrPRMInfoXML [String] = XML string containing the
'          PRM Parameters and Meta data required to generate the
'          binary PRM file(s).
'
'          Sample Input Parameter structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_INFO>
'                <PRM_FILE>
'                    <PRM_META_DATA>
'                        <FILE_NAME>MyPRMFile.prm</FILE_NAME>
'                        <TEMPLATE_NAME>USA 5 MACRS.TEM</TEMPLATE_NAME>
'                        <MODE>Lessor</MODE>
'                    </PRM_META_DATA>
'                    <PRM_PARAMS>
'                        <TRANSACTIONAMOUNT>25000000</TRANSACTIONAMOUNT>
'                        <TRANSACTIONSTARTDATE>2002-08-20</TRANSACTIONSTARTDATE>
'                        <RESIDUALAMOUNT>100000</RESIDUALAMOUNT>
'                        <NUMBEROFPAYMENTS>60</NUMBEROFPAYMENTS>
'                        <PERIODICITY>Monthly</PERIODICITY>
'                        <PAYMENTTIMING>Advance</PAYMENTTIMING>
'                        <STRUCTURE>Level</STRUCTURE>
'                        <TARGETDATA>
'                            <TYPEOFSTATISTIC>Yield</TYPEOFSTATISTIC>
'                            <STATISTICINDEX>1</STATISTICINDEX>
'                            <NEPA>Pre-tax nominal</NEPA>
'                            <TARGETVALUE>0.075</TARGETVALUE>
'                            <ADJUST>Rent</ADJUST>
'                            <ADJUSTMENTMETHOD>Proportional</ADJUSTMENTMETHOD>
'                        </TARGETDATA>
'                    </PRM_PARAMS>
'                </PRM_FILE>
'                <PRM_FILE>
'                    <PRM_META_DATA>
'                        <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
'                        ...
'                    </PRM_META_DATA>
'                    …
'                </PRM_FILE>
'                …
'            </PRM_INFO>
'RETURN  : String= XML string containing, the binary PRM File or
'          <ERROR> tag, for each set of PRM Input Parameters.
'          It may also return an <ERROR> tag for any general
'          failure condition.
'
'            Sample Return XML structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_FILE_LIST>
'
'                <!-- Sucessful generation of PRM file -->
'                <PRM_FILE>
'                    <FILE_NAME>MyPRMFile.prm</FILE_NAME>
'                    <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
'                </PRM_FILE>
'
'                <!-- Error generating PRM file -->
'                <PRM_FILE>
'                    <FILE_NAME>ErrorPRMFile.prm </FILE_NAME>
'                    <ERROR>
'                        <ERROR_NBR>-1072896682</ERROR_NBR>
'                        <ERROR_DESC>Error!!!...</ERROR_DESC>
'                    </ERROR>
'                </PRM_FILE>
'                …
'            </PRM_FILE_LIST>
'
'            OR In case of general failure:
'
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_FILE_LIST>
'                <ERROR>
'                    <ERROR_NBR>-1072896682</ERROR_NBR>
'                    <ERROR_DESC>Error!!!...</ERROR_DESC>
'                </ERROR>
'            </PRM_FILE_LIST>
'================================================================
Public Function GeneratePRMFiles(ByVal astrPRMInfoXML As String) As String
On Error GoTo ErrHandler
    
'Declare Super Trump Variables
Dim lobjSTApplication       As New STApplication
Dim lobjSTTransaction       As New STTransaction
Dim lobjSTQuick             As STQuick

'Declare XML Dom variables
Dim lobjPRMInfoXMLDOM       As New DOMDocument40
Dim lobjXMLSchemaSpace      As New XMLSchemaCache40
Dim lobjReturnPRMLstXMLDOM  As New DOMDocument40
Dim lobjSTQueryXMLDOM       As New DOMDocument40
Dim lobjSTResponseXMLDOM    As New DOMDocument40
Dim lobjExeceptionlst       As IXMLDOMNodeList

'Other Declarations
Dim lstrFileLoc             As String
Dim lstrPRMInfoXML          As String
Dim liPRMParamsCnt          As Integer
Dim lstrPRMFilePath         As String
Dim lstrPRMTemplatePath     As String
Dim lstrPRMMode             As String
Dim lstrReturnXML           As String
Dim lstrPRMFileName         As String
Dim lvPRMFileData           As Variant
Dim lbGenPRM                As Boolean

Dim llErrNbr                As Long
Dim lstrErrSrc              As String
Dim lstrErrDesc             As String

Dim lobjErrComment          As IXMLDOMNode
Dim lstrErrComment          As String

    lbGenPRM = False
    
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): In GeneratePRMFiles() method"
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Input Argument 1:" & astrPRMInfoXML
    
    'Load Return XML
    Call lobjReturnPRMLstXMLDOM.loadXML("<PRM_FILE_LIST></PRM_FILE_LIST>")
    
    'Get the PRMInfoXML.xsd Schema
    lstrFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcSchemaFilePathKey)
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Schema file read from registry"
    
    Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPRMInfoXMLSchemaName)
    
    'Assign Schema to the XML DOM object
    Set lobjPRMInfoXMLDOM.schemas = lobjXMLSchemaSpace
    lobjPRMInfoXMLDOM.async = False
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Validating Input XML"
    
    'Load the Input XML into the XML DOM object
    Call lobjPRMInfoXMLDOM.loadXML(astrPRMInfoXML)
    
    'Check if Input XML is valid
    If lobjPRMInfoXMLDOM.parseError.errorCode = 0 Then
        
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Input XML Valid"
        
        'Get File Paths
        lstrPRMFilePath = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMFilePathKey)
        lstrPRMTemplatePath = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMTemplatePathKey)
        
        'For Each set of PRM Parameters in the Input XML
         For liPRMParamsCnt = 0 To lobjPRMInfoXMLDOM.documentElement.childNodes.length - 1
            
            lbGenPRM = True
            
            'Build the Input XML for the XMLInOut() method
            Call lobjSTQueryXMLDOM.loadXML("<SuperTRUMP>" & _
                                            "<Transaction id='TRANS_ID_GEN_PRM' query='true'/>" & _
                                        "</SuperTRUMP>")
            
            'Mode
            lstrPRMMode = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_META_DATA/MODE")
            
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "MODE", _
                        lstrPRMMode
            
            'Initialize
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "INITIALIZE", _
                        ""
            
            'Read Template
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "READFILE", _
                        ""
            
            AddXMLElementAttribute lobjSTQueryXMLDOM, _
                                    lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                                    "filename", _
                                    lstrPRMTemplatePath & "\" & GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                                                        "PRM_META_DATA/TEMPLATE_NAME")
            
            'Transaction Amount
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "TRANSACTIONAMOUNT", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/TRANSACTIONAMOUNT")
                                                
            'Transaction Date
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "TRANSACTIONSTARTDATE", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/TRANSACTIONSTARTDATE")
            
            'Residual Amout for Lease
            If UCase(lstrPRMMode) = "LESSOR" Then
                
                AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "ASSETS", _
                        ""
                
                AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                        "ASSET", _
                        ""
                                        
                AddXMLElementAttribute lobjSTQueryXMLDOM, _
                                    lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0), _
                                    "index", _
                                    0
                                    
                 AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0), _
                        "ResidualKeptAsAPercent", _
                        "false"
                                                     
                AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0), _
                            "RESIDUAL", _
                            GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                    "PRM_PARAMS/RESIDUALAMOUNT")
            End If
            
            'Term
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "NUMBEROFPAYMENTS", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                            "PRM_PARAMS/NUMBEROFPAYMENTS")
                        
            'Periodicity
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "PERIODICITY", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/PERIODICITY")
                                    
            'Payment timing
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "PAYMENTTIMING", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/PAYMENTTIMING")
                        
            'Structure
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "STRUCTURE", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/STRUCTURE")
                        
            'Yield Data
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "TARGETDATA", _
                        ""
                                    
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                        "TYPEOFSTATISTIC", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/TARGETDATA/TYPEOFSTATISTIC")
                        
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                        "STATISTICINDEX", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/TARGETDATA/STATISTICINDEX")
                        
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                        "NEPA", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/TARGETDATA/NEPA")
                        
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                        "TARGETVALUE", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/TARGETDATA/TARGETVALUE")
                        
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                        "ADJUST", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/TARGETDATA/ADJUST")
                        
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                        "ADJUSTMENTMETHOD", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/TARGETDATA/ADJUSTMENTMETHOD")
                        
            'Write PRM file
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "WRITEFILE", _
                        ""
                                    
            lstrPRMFileName = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                    "PRM_META_DATA/FILE_NAME")
            
            AddXMLElementAttribute lobjSTQueryXMLDOM, _
                                    lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                                    "filename", _
                                    lstrPRMFilePath & "\" & lstrPRMFileName
                                                
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Calling XMLInOut() method with input - " & lobjSTQueryXMLDOM.xml
            
            'Generate the PRM file.
            lstrReturnXML = lobjSTApplication.XmlInOut(lobjSTQueryXMLDOM.xml)
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): " & lstrPRMFileName & " file generated."
            
            AddXMLElement lobjReturnPRMLstXMLDOM, _
                                lobjReturnPRMLstXMLDOM.documentElement, _
                                "PRM_FILE", _
                                ""
            
            AddXMLElement lobjReturnPRMLstXMLDOM, _
                                lobjReturnPRMLstXMLDOM.documentElement.lastChild, _
                                "FILE_NAME", _
                                lstrPRMFileName
                                
            'Load the super Trump response XML
            Call lobjSTResponseXMLDOM.loadXML(lstrReturnXML)
            
            'Check for any Exception.
            Set lobjExeceptionlst = lobjSTResponseXMLDOM.getElementsByTagName("Exception")
            
            If (lobjExeceptionlst.length) > 0 Then
                                
                'Add the <ERROR> node to the Return XML for the PRM file
                AddXMLElement lobjReturnPRMLstXMLDOM, _
                                lobjReturnPRMLstXMLDOM.documentElement.lastChild, _
                                "ERROR", _
                                ""
                    
                AddXMLElement lobjReturnPRMLstXMLDOM, _
                                lobjReturnPRMLstXMLDOM.documentElement.lastChild.lastChild, _
                                "ERROR_NBR", _
                                lobjExeceptionlst.Item(0).selectSingleNode("Number").Text
                
                Set lobjErrComment = lobjExeceptionlst.Item(0).selectSingleNode("Comment")
                lstrErrComment = ""
                If Not (lobjErrComment Is Nothing) Then lstrErrComment = lobjErrComment.Text
                
                AddXMLElement lobjReturnPRMLstXMLDOM, _
                                lobjReturnPRMLstXMLDOM.documentElement.lastChild.lastChild, _
                                "ERROR_DESC", _
                                lobjExeceptionlst.Item(0).selectSingleNode("Description").Text & " " & lstrErrComment
                                
                If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Exception from STServer : " & lobjSTResponseXMLDOM.selectSingleNode("//Exceptions").xml
                
            'Else if no exception
            Else
                
                'Open the PRM file and ReTarget to recalculate the values
                lobjSTTransaction.OpenFile (lstrPRMFilePath & "\" & lstrPRMFileName)
                
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): PRM file read from " & lstrPRMFilePath
                
                Set lobjSTQuick = lobjSTTransaction.Quick
                lobjSTQuick.ReTarget
                
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): PRM file re-targeted"
                
                lobjSTTransaction.Calculate
                
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): PRM file values calculated"
                
                lobjSTTransaction.SaveFile lstrPRMFilePath & "\" & lstrPRMFileName
                
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): PRM file changes updated to disk"
                
                Set lobjSTQuick = Nothing
                Set lobjSTTransaction = Nothing
                
                'Read the generated PRM File from disk
                lvPRMFileData = GetBinaryFileData(lstrPRMFilePath & "\" & lstrPRMFileName)
                                                
                'Add it to the Return XML
                AddBinaryXMLElement lobjReturnPRMLstXMLDOM, _
                                        lobjReturnPRMLstXMLDOM.documentElement.lastChild, _
                                        "FILE_DATA", _
                                        lvPRMFileData
                                        
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Final PRM file read from disk and added to output XML"
            End If
NextRec:

            Set lobjSTResponseXMLDOM = Nothing
            Set lobjExeceptionlst = Nothing
            Set lobjSTQueryXMLDOM = Nothing
        Next
    
    'Else if input XML is invalid
    Else
        llErrNbr = lobjPRMInfoXMLDOM.parseError.errorCode
        lstrErrDesc = lobjPRMInfoXMLDOM.parseError.reason
        
        'Add the <ERROR> node to the Return XML
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                    lobjReturnPRMLstXMLDOM.documentElement, _
                    "ERROR", _
                    ""
                    
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                    lobjReturnPRMLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_NBR", _
                    llErrNbr
                
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                    lobjReturnPRMLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_DESC", _
                    "Error on line number " & lobjPRMInfoXMLDOM.parseError.Line & " of the XML. " & lstrErrDesc
    End If
    
    'Return the final PRM list
    GeneratePRMFiles = lobjReturnPRMLstXMLDOM.xml
                
    'Delete the saved PRM files from disk
    Call DeletePRMBinaryFile(lobjReturnPRMLstXMLDOM.getElementsByTagName("FILE_NAME"), lstrPRMFilePath)
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): PRM files deleted from temp location."
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Return value: " & GeneratePRMFiles
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Exit GeneratePRMFiles() method"

CleanUp:
    
    Set lobjSTApplication = Nothing
    Set lobjPRMInfoXMLDOM = Nothing
    Set lobjXMLSchemaSpace = Nothing
    Set lobjReturnPRMLstXMLDOM = Nothing
    
    Exit Function
    
ErrHandler:
    
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    If lbGenPRM Then
        
        lbGenPRM = False
        
        'Add the <ERROR> node to the Return XML for the PRM file
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                        lobjReturnPRMLstXMLDOM.documentElement.lastChild, _
                        "ERROR", _
                        ""
            
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                        lobjReturnPRMLstXMLDOM.documentElement.lastChild.lastChild, _
                        "ERROR_NBR", _
                        llErrNbr
        
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                        lobjReturnPRMLstXMLDOM.documentElement.lastChild.lastChild, _
                        "ERROR_DESC", _
                        lstrErrDesc
            
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): PRM file Generation error - " & lstrErrDesc
        
        lbGenPRM = True
        
        Resume NextRec
    Else
    
        Set lobjReturnPRMLstXMLDOM = Nothing
        
        'Build the Error XML
        Call lobjReturnPRMLstXMLDOM.loadXML("<PRM_FILE_LIST><ERROR></ERROR></PRM_FILE_LIST>")
            
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                    lobjReturnPRMLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_NBR", _
                    llErrNbr
                    
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                    lobjReturnPRMLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_DESC", _
                    lstrErrDesc
                        
        'Return error XML
        GeneratePRMFiles = lobjReturnPRMLstXMLDOM.xml
        
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): General error : " & GeneratePRMFiles
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Exit GeneratePRMFiles() method"
        
        Resume CleanUp
    End If
End Function

'================================================================
'METHOD  : GetPRMParams
'PURPOSE : To get specified PRM Parameters for the inputted
'          binary PRM file(s).
'          Note: This method is similar to the ConvertPRMToXML()
'          method, but it will return only a subset of data than
'          the one returned by the ConvertPRMToXML() method.
'PARMS   :
'          astrPRMParamsInfoXML [String]= XML string containing
'          the List of PRM parameters.
'
'            Sample Input Parameter structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_PARAMS_INFO>
'                <PRM_PARAMS>
'                    <PRM_PARAMS_SPECS>
'                        <TRANSACTIONAMOUNT query="true"/>
'                        <TRANSACTIONSTARTDATE query="true"/>
'                        <RESIDUALAMOUNT query="true"/>
'                        <STRUCTURE query="true"/>
'                        <PERIODICITY query="true"/>
'                        <PAYMENTTIMING query="true"/>
'                        <NUMBEROFPAYMENTS query="true"/>
'                        <TARGETDATA query="true"/>
'                    </PRM_PARAMS_SPECS>
'                    <PRM_FILE>
'                        <FILE_NAME>LeasePRMFile.prm</FILE_NAME>
'                        <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
'                    </PRM_FILE>
'                </PRM_PARAMS>
'                <PRM_PARAMS>
'                    <PRM_PARAMS_SPECS>
'                        <TRANSACTIONAMOUNT query="true"/>
'                        <TRANSACTIONSTARTDATE query="true"/>
'                        <STRUCTURE query="true"/>
'                        <PERIODICITY query="true"/>
'                        <PAYMENTTIMING query="true"/>
'                        <NUMBEROFPAYMENTS query="true"/>
'                        <TARGETDATA query="true"/>
'                    </PRM_PARAMS_SPECS>
'                    <PRM_FILE>
'                        <FILE_NAME>LoanPRMFile.prm</FILE_NAME>
'                        <FILE_DATA>M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAPgADAP7…</FILE_DATA>
'                    </PRM_FILE>
'                </PRM_PARAMS>
'                <PRM_PARAMS>
'                    <PRM_PARAMS_SPECS>
'                        <TRANSACTIONAMOUNT query="true"/>
'                        <TRANSACTIONSTARTDATE query="true"/>
'                        <STRUCTURE query="true"/>
'                        <PERIODICITY query="true"/>
'                        <PAYMENTTIMING query="true"/>
'                        <NUMBEROFPAYMENTS query="true"/>
'                        <TARGETDATA query="true"/>
'                    </PRM_PARAMS_SPECS>
'                    <PRM_FILE>
'                        <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
'                        <FILE_DATA>AAAAAAAAAAAAAAAAAAAAAPgADAP7…</FILE_DATA>
'                    </PRM_FILE>
'                </PRM_PARAMS>
'                …
'            </PRM_PARAMS_INFO>
'
'            Note:
'            1)  <FILE_NAME> tag must contain PRM file name with .prm extension.
'            2)  <FILE_DATA> tag must contain binary value of type base64Binary.
'
'RETURN  : String = XML string containing, the set of Input
'          Parameters or <ERROR> tag, for each binary PRM file.
'          It may also return an <ERROR> tag for any general
'          failure condition.
'
'            Sample Return XML structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_PARAMS_LIST>
'
'                <!-- Successfully converted Lease PRM file -->
'                <PRM_PARAMS>
'                    <PRM_FILE_NAME>LeasePRMFile.prm</PRM_FILE_NAME>
'                    <TransactionAmount>25000000</TransactionAmount>
'                    <TransactionStartDate>2002-08-20</TransactionStartDate>
'                    <ResidualAmount>100000</ResidualAmount>
'                    <Structure>High/Low</Structure>
'                    <Periodicity>Semiannual</Periodicity>
'                    <PaymentTiming>Advance</PaymentTiming>
'                    <NumberOfPayments>14</NumberOfPayments>
'                    <TargetData>
'                        <TypeOfStatistic>Yield</TypeOfStatistic>
'                        <StatisticIndex>1</StatisticIndex>
'                        <NEPA>Pre-tax nominal</NEPA>
'                        <TargetValue>0.075</TargetValue>
'                        <Adjust>Rent</Adjust>
'                        <AdjustmentMethod>Proportional</AdjustmentMethod>
'                    </TargetData>
'                </PRM_PARAMS>
'
'                <!-- Successfully converted Loan PRM file -->
'                <PRM_PARAMS>
'                    <PRM_FILE_NAME>LoanPRMFile.prm</PRM_FILE_NAME>
'                    <TransactionAmount>25000000</TransactionAmount>
'                    <TransactionStartDate>2002-08-20</TransactionStartDate>
'                    <ResidualAmount>100000</ResidualAmount>
'                    <Structure>High/Low</Structure>
'                    <Periodicity>Semiannual</Periodicity>
'                    <PaymentTiming>Advance</PaymentTiming>
'                    <NumberOfPayments>14</NumberOfPayments>
'                    <TargetData>
'                        <TypeOfStatistic>Yield</TypeOfStatistic>
'                        <StatisticIndex>1</StatisticIndex>
'                        <NEPA>Pre-tax nominal</NEPA>
'                        <TargetValue>0.075</TargetValue>
'                        <Adjust>Rent</Adjust>
'                        <AdjustmentMethod>Proportional</AdjustmentMethod>
'                    </TargetData>
'                </PRM_PARAMS>
'
'                <!-- Error reading PRM file -->
'                <PRM_PARAMS>
'                    <PRM_FILE_NAME>ErrorPRMFile.prm</PRM_FILE_NAME>
'                    <ERROR>
'                        <ERROR_NBR>-1072896682</ERROR_NBR>
'                        <ERROR_DESC>Error!!!...</ERROR_DESC>
'                    </ERROR>
'                </PRM_PARAMS>
'                …
'            </PRM_PARAMS_LIST>
'
'            OR In case of general failure:
'
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_PARAMS_LIST>
'                <ERROR>
'                    <ERROR_NBR>-1072896682</ERROR_NBR>
'                    <ERROR_DESC>Error!!!...</ERROR_DESC>
'                </ERROR>
'            </PRM_PARAMS_LIST>
'================================================================
Public Function GetPRMParams(ByVal astrPRMParamsInfoXML As String) As String
On Error GoTo ErrHandler

'Declare Super Trump Variables
Dim lobjSTApplication           As New STApplication
Dim lobjSTTransaction           As New STTransaction

'Declare XML Dom variables
Dim lobjReturnPRMParamLstXMLDOM As New DOMDocument40
Dim lobjXMLSchemaSpace          As New XMLSchemaCache40
Dim lobjPRMParamsInfoXMLDOM     As New DOMDocument40
Dim lobjSTQueryXMLDOM           As New DOMDocument40
Dim lobjSTResponseXMLDOM        As New DOMDocument40
Dim lobjExeceptionlst           As IXMLDOMNodeList
Dim lobjPRMFileBinDataElem      As IXMLDOMElement
Dim lobjPRMParamElem            As IXMLDOMElement
Dim lobjCloneNode               As IXMLDOMNode

'Other Declarations
Dim lstrPRMFilePath             As String
Dim liPRMParamsSpecCnt          As Integer
Dim lstrPRMParamQuery           As String
Dim lstrPRMFileName             As String
Dim lstrReturnXML               As String
Dim lstrSTQuery                 As String
Dim liPRMParamsCnt              As Integer
Dim lstrFileLoc                 As String
Dim lstrPRMParamsInfoXML        As String
Dim lbGetPRMParam               As Boolean
    
Dim liStart                     As Integer
Dim liEnd                       As Integer
Dim lstrResdElem                As String

Dim llErrNbr                As Long
Dim lstrErrSrc              As String
Dim lstrErrDesc             As String

Dim lobjErrComment          As IXMLDOMNode
Dim lstrErrComment          As String

'New variables declared to force the aniequity error in the output in 9549
'which is supressing it: 5th march 2007
Dim liPRMFile               As Integer
Dim liCntPRMs                 As Integer

    lbGetPRMParam = False
    
    liPRMFile = 1
    
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): In GetPRMParams() method"
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Input Argument 1:" & astrPRMParamsInfoXML

    'Load Return XML
    Call lobjReturnPRMParamLstXMLDOM.loadXML("<PRM_PARAMS_LIST></PRM_PARAMS_LIST>")
    
    'Get the PRMParamsInfoXML.xsd Schema
    lstrFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcSchemaFilePathKey)
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Schema file read from registry"
    
    Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPRMParamsInfoXMLSchemaName)
    
    'Assign Schema to the XML DOM object
    Set lobjPRMParamsInfoXMLDOM.schemas = lobjXMLSchemaSpace
    lobjPRMParamsInfoXMLDOM.async = False
    
    'Remove "xmlns:dt" & "dt:dt" attributes from the <FILE_DATA> elements in the Input XML
    'as we cannot include these attributes in the schema as they contain special
    'characters(:). We will add them after schema validation.
    lstrPRMParamsInfoXML = Replace(astrPRMParamsInfoXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Validating Input XML"
    
    'Load the Input XML into the XML DOM object
    Call lobjPRMParamsInfoXMLDOM.loadXML(lstrPRMParamsInfoXML)
    
    'Check if Input XML is valid
    If lobjPRMParamsInfoXMLDOM.parseError.errorCode = 0 Then
            
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Input XML Valid"
        
        'Get File Paths
        lstrPRMFilePath = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMFilePathKey)
        
        'For Each set of PRM Parameters Specs in the Input XML
         For liPRMParamsSpecCnt = 0 To lobjPRMParamsInfoXMLDOM.documentElement.childNodes.length - 1
                    
            lbGetPRMParam = True
            
            'Get Parameter Query for Super Trump
            lstrPRMParamQuery = lobjPRMParamsInfoXMLDOM.documentElement.childNodes(liPRMParamsSpecCnt).selectSingleNode("PRM_PARAMS_SPECS").xml
            lstrPRMParamQuery = Replace(Replace(lstrPRMParamQuery, "<PRM_PARAMS_SPECS>", ""), "</PRM_PARAMS_SPECS>", "")
            
            With lobjPRMParamsInfoXMLDOM.documentElement.childNodes(liPRMParamsSpecCnt)
                
                'Get PRM File name
                lstrPRMFileName = GetXMLElementValue(.selectSingleNode("PRM_FILE"), _
                                                        "FILE_NAME")
                
                'Save the PRM Binary data to a physical file.
                Set lobjPRMFileBinDataElem = .selectSingleNode("PRM_FILE").selectSingleNode("FILE_DATA")
                lobjPRMFileBinDataElem.dataType = "bin.base64"
                SavePRMBinaryFile lstrPRMFileName, lstrPRMFilePath, lobjPRMFileBinDataElem.nodeTypedValue
                Set lobjPRMFileBinDataElem = Nothing
                    
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): " & lstrPRMFileName & " saved to " & lstrPRMFilePath
                
                'Open the PRM file to determine the Mode
                If lobjSTTransaction.OpenFile(lstrPRMFilePath & "\" & lstrPRMFileName) Then
                    
                    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): PRM file opened to determine Mode value"
                    
                    'Remove Residual Amout for Loan from Parameter Query for Super Trump
                    If lobjSTTransaction.Mode = ST_Mode_Lender Then
                        
                        liStart = InStr(1, lstrPRMParamQuery, "<RESIDUALAMOUNT")
                        If liStart > 0 Then
                            liEnd = InStr(liStart, lstrPRMParamQuery, "</RESIDUALAMOUNT>")
                            
                            If liEnd <= 0 Then
                                liEnd = InStr(liStart, lstrPRMParamQuery, "/>") + 2
                            Else
                                liEnd = liEnd + Len("</RESIDUALAMOUNT>")
                            End If
                            
                            If liEnd > liStart Then
                                'Get the residual amount XML element
                                lstrResdElem = Mid(lstrPRMParamQuery, liStart, (liEnd - liStart))
                                
                                'Remove it
                                lstrPRMParamQuery = Replace(lstrPRMParamQuery, lstrResdElem, "")
                            End If
                        End If
                    End If
                End If
                Set lobjSTTransaction = Nothing
                
                'Build the Input XML for the XMLInOut() method
                Call lobjSTQueryXMLDOM.loadXML("<SuperTRUMP>" & _
                                                "<Transaction id='TRANS_ID_GET_PRM_PARAMS'/>" & _
                                            "</SuperTRUMP>")
                    
                    
                'Read PRM file
                AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                            "READFILE", _
                            ""
                                                                                                
                AddXMLElementAttribute lobjSTQueryXMLDOM, _
                                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                                        "filename", _
                                        lstrPRMFilePath & "\" & lstrPRMFileName
                    
                lstrSTQuery = lobjSTQueryXMLDOM.xml
                lstrSTQuery = Replace(lstrSTQuery, "</Transaction>", lstrPRMParamQuery & "</Transaction>")
                
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Calling XMLInOut() method with input - " & lstrSTQuery
                    
                'Get the PRM Parameters.
                lstrReturnXML = lobjSTApplication.XmlInOut(lstrSTQuery)
                    
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Parmeter values returned as XML string"
                
                AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                                    lobjReturnPRMParamLstXMLDOM.documentElement, _
                                    "PRM_PARAMS", _
                                    ""
                    
                AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                                    lobjReturnPRMParamLstXMLDOM.documentElement.lastChild, _
                                    "PRM_FILE_NAME", _
                                    lstrPRMFileName
          
          'Load the Super Trump response XML
                Call lobjSTResponseXMLDOM.loadXML(lstrReturnXML)
                    
                    If UCase(Trim(gstrExceptionFlag)) = "TRUE" Then
                    
                        For liCntPRMs = 0 To UBound(gliPRMFilearr)
                            If liPRMFile = gliPRMFilearr(liCntPRMs) Then
                        
                            'Add the <ERRORS> node to the Return XML
                                AddXMLElement lobjSTResponseXMLDOM, _
                                        lobjSTResponseXMLDOM.documentElement.lastChild, _
                                        "Exceptions", _
                                        gstrExcptionXMLDOMarr(liCntPRMs)
                            End If
                        Next
                    End If
                'This replace is reqd as AddXMLElement adds values to the tags as .text.
                'So the XMl tags in it get converted to text values , which need to be re-converted.
                Call lobjSTResponseXMLDOM.loadXML(Replace(Replace(lobjSTResponseXMLDOM.xml, "&lt;", "<"), "&gt;", ">"))
                'Check for any Exception.
                Set lobjExeceptionlst = lobjSTResponseXMLDOM.getElementsByTagName("Exception")
                
                If (lobjExeceptionlst.length) > 0 Then
                                    
                    'Add the <ERROR> node to the Return XML for the PRM file
                    AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                                    lobjReturnPRMParamLstXMLDOM.documentElement.lastChild, _
                                    "ERROR", _
                                    ""
                        
                    AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                                    lobjReturnPRMParamLstXMLDOM.documentElement.lastChild.lastChild, _
                                    "ERROR_NBR", _
                                    lobjExeceptionlst.Item(0).selectSingleNode("Number").Text
                    
                    Set lobjErrComment = lobjExeceptionlst.Item(0).selectSingleNode("Comment")
                    lstrErrComment = ""
                    If Not (lobjErrComment Is Nothing) Then lstrErrComment = lobjErrComment.Text

                    AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                                    lobjReturnPRMParamLstXMLDOM.documentElement.lastChild.lastChild, _
                                    "ERROR_DESC", _
                                    lobjExeceptionlst.Item(0).selectSingleNode("Description").Text & " " & lstrErrComment
                                                        
                    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Exception from STServer : " & lobjSTResponseXMLDOM.selectSingleNode("//Exceptions").xml
                    
                'Else if no exception
                Else
                    
                    'Add the response to the Return XML
                    For liPRMParamsCnt = 0 To lobjSTResponseXMLDOM.documentElement.childNodes(0).childNodes.length - 1
                        Set lobjPRMParamElem = lobjSTResponseXMLDOM.documentElement.childNodes(0).childNodes(liPRMParamsCnt)
                        If Not (lobjPRMParamElem Is Nothing) Then
                            
                            Set lobjCloneNode = lobjPRMParamElem.cloneNode(True)
                            lobjReturnPRMParamLstXMLDOM.documentElement.lastChild.appendChild lobjCloneNode
                            Set lobjCloneNode = Nothing
                        End If
                        Set lobjPRMParamElem = Nothing
                    Next
                    
                    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Data added to output XML."
                End If
                
NextRec:
                Set lobjSTResponseXMLDOM = Nothing
                Set lobjExeceptionlst = Nothing
                Set lobjSTQueryXMLDOM = Nothing
                    

            End With
            liPRMFile = liPRMFile + 1
        Next
        
    'Else if input XML is invalid
    Else
        llErrNbr = lobjPRMParamsInfoXMLDOM.parseError.errorCode
        lstrErrDesc = lobjPRMParamsInfoXMLDOM.parseError.reason
        
        'Add the <ERROR> node to the Return XML
        AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                    lobjReturnPRMParamLstXMLDOM.documentElement, _
                    "ERROR", _
                    ""
                    
        AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                    lobjReturnPRMParamLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_NBR", _
                    llErrNbr
                
        AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                    lobjReturnPRMParamLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_DESC", _
                    "Error on line number " & lobjPRMParamsInfoXMLDOM.parseError.Line & " of the XML. " & lstrErrDesc
    End If
    
    'Return the final PRM list
    GetPRMParams = lobjReturnPRMParamLstXMLDOM.xml
                
    'Delete the saved PRM files from disk
    Call DeletePRMBinaryFile(lobjPRMParamsInfoXMLDOM.getElementsByTagName("FILE_NAME"), lstrPRMFilePath)
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): PRM files deleted from temp location."
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Return value: " & GetPRMParams
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Exit GetPRMParams() method"
    
CleanUp:

    Set lobjSTApplication = Nothing
    Set lobjPRMParamsInfoXMLDOM = Nothing
    Set lobjXMLSchemaSpace = Nothing
    Set lobjReturnPRMParamLstXMLDOM = Nothing
    gstrExceptionFlag = ""
    ReDim gliPRMFilearr(0)
    ReDim gstrExcptionXMLDOMarr(0)
    
    Exit Function
    
ErrHandler:
    
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description
    
    If lbGetPRMParam Then
        
        lbGetPRMParam = False
        
        'Add the <ERROR> node to the Return XML for the PRM file
        AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                        lobjReturnPRMParamLstXMLDOM.documentElement.lastChild, _
                        "ERROR", _
                        ""
            
        AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                        lobjReturnPRMParamLstXMLDOM.documentElement.lastChild.lastChild, _
                        "ERROR_NBR", _
                        llErrNbr
        
        AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                        lobjReturnPRMParamLstXMLDOM.documentElement.lastChild.lastChild, _
                        "ERROR_DESC", _
                        lstrErrDesc
        
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): PRM file parameter retrieval error - " & lstrErrDesc
        
        lbGetPRMParam = True
        
        Resume NextRec
    Else
        
        Set lobjReturnPRMParamLstXMLDOM = Nothing
        
        'Build Error XML
        Call lobjReturnPRMParamLstXMLDOM.loadXML("<PRM_PARAMS_LIST><ERROR></ERROR></PRM_PARAMS_LIST>")
            
        AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                    lobjReturnPRMParamLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_NBR", _
                    llErrNbr
                    
        AddXMLElement lobjReturnPRMParamLstXMLDOM, _
                    lobjReturnPRMParamLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_DESC", _
                    lstrErrDesc
                
        'Return error XML
        GetPRMParams = lobjReturnPRMParamLstXMLDOM.xml
        
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): General error : " & GetPRMParams
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetPRMParams(): Exit GetPRMParams() method"
        
        Resume CleanUp
    End If
End Function

'================================================================
'METHOD  : GetAmortizationSchedule
'PURPOSE : To get amortization schedule for the inputted binary
'          PRM file(s).
'PARMS   :
'          astrPRMFileListXML [String] = XML string containing
'          the List of binary PRM files.
'
'            Sample Input Parameter structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_FILE_LIST>
'                <PRM_FILE>
'                    <FILE_NAME>LeasePRMFile.prm</FILE_NAME>
'                    <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
'                </PRM_FILE>
'                <PRM_FILE>
'                    <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
'                    <FILE_DATA>AAAAAAAAAAAAAAAAAAAAAPgADAP7…</FILE_DATA>
'                </PRM_FILE>
'                <PRM_FILE>
'                    <FILE_NAME>LoanPRMFile.prm</FILE_NAME>
'                    <FILE_DATA>M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAPgADAP7…</FILE_DATA>
'                </PRM_FILE>
'                …
'            </PRM_FILE_LIST>
'
'            Note:
'            1)  <FILE_NAME> tag must contain PRM file name with .prm extension.
'            2)  <FILE_DATA> tag must contain binary value of type base64Binary.
'RETURN  : String = XML string containing, the Rent Schedule data
'          or a <ERROR> tag, for each binary PRM file. It may also
'          return an <ERROR> tag for any general failure condition.
'
'            Sample Return XML structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <AMORTIZATION_SCHEDULE_LIST>
'
'                <!-- Sucessful generation of amortization schedule -->
'                <AMORTIZATION_SCHEDULE>
'                    <PRM_FILE_NAME>LeasePRMFile.prm</PRM_FILE_NAME>
'                    <PAYMENT_LIST>
'                        <PAYMENT>
'                            <PAYMENT_NUMBER>1</PAYMENT_NUMBER>
'                            <PAYMENT_START_DATE>8/22/2002</PAYMENT_START_DATE>
'                            <PAYMENT_AMOUNT>10000</PAYMENT_AMOUNT>
'                            <LEASE_FACTOR>0.0543</LEASE_FACTOR>
'                        </PAYMENT>
'                        ...
'                    </PAYMENT_LIST>
'                </AMORTIZATION_SCHEDULE>
'
'                <!-- Error reading PRM file  -->
'                <AMORTIZATION_SCHEDULE>
'                    <PRM_FILE_NAME>ErrorPRMFile.prm </PRM_FILE_NAME>
'                    <ERROR>
'                        <ERROR_NBR>-1072896682</ERROR_NBR>
'                        <ERROR_DESC>Error!!!...</ERROR_DESC>
'                    </ERROR>
'                </AMORTIZATION_SCHEDULE>
'                ...
'            </AMORTIZATION_SCHEDULE_LIST>
'
'            OR In case of general failure:
'
'            <?xml version="1.0" encoding="UTF-8"?>
'            <AMORTIZATION_SCHEDULE_LIST>
'                <ERROR>
'                    <ERROR_NBR>-1072896682</ERROR_NBR>
'                    <ERROR_DESC>Error!!!...</ERROR_DESC>
'                </ERROR>
'            </AMORTIZATION_SCHEDULE_LIST>
'================================================================
Public Function GetAmortizationSchedule(ByVal astrPRMFileListXML As String) As String
On Error GoTo ErrHandler

'Declare Super Trump Variables
Dim lobjSTTransaction               As New STTransaction
Dim lobjSTQuick                     As STQuick
Dim lobjSTStream                    As STStream
Dim lobjSTStreamPrincipal           As STStream
Dim lobjSTStreamInterest            As STStream

'Declare XML Dom variables
Dim lobjXMLSchemaSpace              As New XMLSchemaCache40
Dim lobjReturnAmortSchedLstXMLDOM   As New DOMDocument40
Dim lobjPRMFileListXMLDOM           As New DOMDocument40
Dim lobjPRMFileBinDataElem          As IXMLDOMElement

'Other Declarations
Dim lstrFileLoc                     As String
Dim lstrPRMFileListXML              As String
Dim lstrPRMFilePath                 As String
Dim liPRMFileLstCnt                 As Integer
Dim lstrPRMFileName                 As String
Dim liStreamCnt                     As Integer
Dim ldLeaseFactor                   As Double
Dim lbGetAmort                      As Boolean
    
Dim llErrNbr                        As Long
Dim lstrErrSrc                      As String
Dim lstrErrDesc                     As String

Dim lstrXMLInOutInput               As String
Dim lobjSTApplication               As New STApplication
Dim lstrReturnXMLInOut              As String
Dim lobjSTResponseXMLDOM            As New DOMDocument40


Dim lobjNodeList1                   As Object
Dim lobjNodeList2                   As Object
Dim lobjNodeList3                   As Object
Dim lobjRefrenceNode                As IXMLDOMElement
Dim lobjRefrenceNodePrincipal       As IXMLDOMElement
Dim lobjRefrenceNodeInterest        As IXMLDOMElement
Dim licount                         As Integer
Dim lstrPaymentStartDate            As String


    lbGetAmort = False
    
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): In GetAmortizationSchedule() method"
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Input Argument 1:" & astrPRMFileListXML

    'Load Return XML
    Call lobjReturnAmortSchedLstXMLDOM.loadXML("<AMORTIZATION_SCHEDULE_LIST></AMORTIZATION_SCHEDULE_LIST>")
    
    'Get the PRMFileLstXML.xsd Schema
    lstrFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcSchemaFilePathKey)
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Schema file read from registry"
    
    Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPRMFileLstXMLSchemaName)
    
    'Assign Schema to the XML DOM object
    Set lobjPRMFileListXMLDOM.schemas = lobjXMLSchemaSpace
    lobjPRMFileListXMLDOM.async = False
    
    'Remove "xmlns:dt" & "dt:dt" attributes from the <FILE_DATA> elements in the Input XML
    'as we cannot include these attributes in the schema as they contain special
    'characters(:). We will add them after schema validation.
    lstrPRMFileListXML = Replace(astrPRMFileListXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
        
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Validating Input XML"
    
    'Load the Input XML into the XML DOM object
    Call lobjPRMFileListXMLDOM.loadXML(lstrPRMFileListXML)
    
    'Check if Input XML is valid
    If lobjPRMFileListXMLDOM.parseError.errorCode = 0 Then
        
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Input XML Valid"
        
        'Get File Paths
        lstrPRMFilePath = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMFilePathKey)
        
        'For each binary PRM file in the Input XML
        For liPRMFileLstCnt = 0 To lobjPRMFileListXMLDOM.documentElement.childNodes.length - 1
                    
            lbGetAmort = True
            
            'Save the PRM Binary data to a physical file.
            lstrPRMFileName = GetXMLElementValue(lobjPRMFileListXMLDOM.documentElement.childNodes(liPRMFileLstCnt), _
                                                            "FILE_NAME")
            Set lobjPRMFileBinDataElem = lobjPRMFileListXMLDOM.documentElement.childNodes(liPRMFileLstCnt).selectSingleNode("FILE_DATA")
            lobjPRMFileBinDataElem.dataType = "bin.base64"
            SavePRMBinaryFile lstrPRMFileName, lstrPRMFilePath, lobjPRMFileBinDataElem.nodeTypedValue
            Set lobjPRMFileBinDataElem = Nothing
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): " & lstrPRMFileName & " saved to " & lstrPRMFilePath
            
            AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                lobjReturnAmortSchedLstXMLDOM.documentElement, _
                                "AMORTIZATION_SCHEDULE", _
                                ""
            
            AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild, _
                                "PRM_FILE_NAME", _
                                lstrPRMFileName
                    
            'Check if the PRM file can be read
            If lobjSTTransaction.OpenFile(lstrPRMFilePath & "\" & lstrPRMFileName) Then
                     
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): PRM file read from disk"
                
                lobjSTTransaction.Calculate
                
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): PRM file values calculated"
                
                'Get the Amortization Data
                'lobjSTTransaction.GetFreeStream
                If lobjSTTransaction.Mode = ST_Mode_Lender Then
                    'For Loans , Payment Amount , Principal Amount and Interest Amount is fetched.
                    'Build the Input XML for loans ..
                    lstrXMLInOutInput = "<SuperTRUMP>" & _
                                "<Transaction>" & _
                                "<Results>" & _
                                "<Stream name=""Lending Loans Debt Service"" query=""true"" Label=""Payment amount"" />" & _
                                "<Stream name=""Lending Loans Principal"" query=""true"" Label=""Principal amount"" />" & _
                                "<Stream name=""Lending Loans Interest"" query=""true"" Label=""Interest amount"" />" & _
                                "</Results>" & _
                                "</Transaction>" & _
                                "</SuperTRUMP>"
                                
                    '''Set lobjSTStream = lobjSTTransaction.GetEconStream(ST_EDA_LoanLendDS) 'For Loan - Monthly Payments
                    '''Set lobjSTStreamPrincipal = lobjSTTransaction.GetEconStream(ST_EDA_LoanLendPrin) 'For Loan - Principal
                    '''Set lobjSTStreamInterest = lobjSTTransaction.GetEconStream(ST_EDA_LoanLendInt) 'For Loan - Interest
                Else
                    'For Lease.....
                    lstrXMLInOutInput = "<SuperTRUMP>" & _
                                "<Transaction>" & _
                                "<Results>" & _
                                "<Stream name=""Rent"" query=""true"" Label=""Payment amount"" />" & _
                                "</Results>" & _
                                "</Transaction>" & _
                                "</SuperTRUMP>"
                    
                    '''Set lobjSTStream = lobjSTTransaction.GetEconStream(ST_EDA_Rent)  'For Lease
                End If
                
                'Call XMLInOut now to get the Ammortization schedule..
                lstrReturnXMLInOut = lobjSTApplication.XmlInOut(lstrXMLInOutInput)
                
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Amortization Data retrieved"
                
                'Load the Super Trump Response XML
                Call lobjSTResponseXMLDOM.loadXML(lstrReturnXMLInOut)
                
                Set lobjNodeList1 = lobjSTResponseXMLDOM.selectNodes("//Stream[@Label='Payment amount']/Amounts/Amount")
                Set lobjNodeList2 = lobjSTResponseXMLDOM.selectNodes("//Stream[@Label='Principal amount']/Amounts/Amount")
                Set lobjNodeList3 = lobjSTResponseXMLDOM.selectNodes("//Stream[@Label='Interest amount']/Amounts/Amount")
                
                              
                AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild, _
                                "PAYMENT_LIST", _
                               ""
                licount = 0
                
                If Not lobjNodeList1 Is Nothing Then
                    For licount = 0 To lobjNodeList1.length - 1
                    
                    Set lobjRefrenceNode = lobjNodeList1(licount)
                    
                    'For Loans, the last row contains Amt = $0 for balloon type.
                    'which is not required. Hence ignore the last row.
                    If lobjSTTransaction.Mode = ST_Mode_Lender _
                        And licount = lobjNodeList1.length - 1 _
                        And lobjRefrenceNode.Text = 0 Then

                        Exit For
                    End If
                        
                        
                    AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild, _
                                "PAYMENT", _
                                ""
                    
                    'Payment number
                    AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild.lastChild, _
                                "PAYMENT_NUMBER", _
                                licount + 1
                                                
                    lstrPaymentStartDate = Format(lobjRefrenceNode.getAttribute("date"), "mm/dd/yyyy")
                    'Payment Start Date
                    AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild.lastChild, _
                                "PAYMENT_START_DATE", _
                                lstrPaymentStartDate
                    
                    'Payment Amount
                    AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild.lastChild, _
                                "PAYMENT_AMOUNT", _
                                lobjRefrenceNode.Text
                    
                    'Lease Factor
                    If lobjRefrenceNode.Text > 0 Then
                        ldLeaseFactor = lobjRefrenceNode.Text / lobjSTTransaction.Amount
                    Else
                        ldLeaseFactor = 0
                    End If

                    AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                        lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild.lastChild, _
                                        "LEASE_FACTOR", _
                                        ldLeaseFactor

                    'For loans add the Principal & Interest amounts as well
                    If lobjSTTransaction.Mode = ST_Mode_Lender Then
                        
                        
                        If Not lobjNodeList2 Is Nothing Then
                            Set lobjRefrenceNodePrincipal = lobjNodeList2(licount)
                            'Principal Amount
                            AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                            lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild.lastChild, _
                                            "PRINCIPAL_AMOUNT", _
                                            lobjRefrenceNodePrincipal.Text
                        End If
                                       
                                    
                        If Not lobjNodeList3 Is Nothing Then
                            Set lobjRefrenceNodeInterest = lobjNodeList3(licount)
                            'Interest Amount
                            AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                            lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild.lastChild, _
                                            "INTEREST_AMOUNT", _
                                            lobjRefrenceNodeInterest.Text
                        End If
                    End If
                     
                    Set lobjRefrenceNode = Nothing
                    Set lobjRefrenceNodePrincipal = Nothing
                    Set lobjRefrenceNodeInterest = Nothing
                    
                    Next
                End If
                
                If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Data added to the output XML"
                
            'Else if the PRM file can't be read
            Else
                
                'Add the <ERROR> node to the Return XML for the PRM file
                AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild, _
                                "ERROR", _
                                ""
                    
                AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild, _
                                "ERROR_NBR", _
                                gcINVALID_PRM_FILE
                
                AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                                lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild, _
                                "ERROR_DESC", _
                                "Invalid binary PRM file."
            End If
            
NextRec:
            Set lobjSTTransaction = Nothing
            Set lobjSTQuick = Nothing
            Set lobjSTStream = Nothing
            Set lobjSTStreamPrincipal = Nothing
            Set lobjSTStreamInterest = Nothing
            
            Set lobjSTApplication = Nothing
            Set lobjNodeList1 = Nothing
            Set lobjNodeList2 = Nothing
            Set lobjNodeList3 = Nothing
        Next
        
    'Else if input XML is invalid
    Else
        llErrNbr = lobjPRMFileListXMLDOM.parseError.errorCode
        lstrErrDesc = lobjPRMFileListXMLDOM.parseError.reason
        
        'Add the <ERROR> node to the Return XML
        AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                    lobjReturnAmortSchedLstXMLDOM.documentElement, _
                    "ERROR", _
                    ""
                    
        AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                    lobjReturnAmortSchedLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_NBR", _
                    llErrNbr
                
        AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                    lobjReturnAmortSchedLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_DESC", _
                    "Error on line number " & lobjPRMFileListXMLDOM.parseError.Line & " of the XML. " & lstrErrDesc
    End If

    'Return the final PRM list
    GetAmortizationSchedule = lobjReturnAmortSchedLstXMLDOM.xml
            
    'Delete the saved PRM files from disk
    Call DeletePRMBinaryFile(lobjPRMFileListXMLDOM.getElementsByTagName("FILE_NAME"), lstrPRMFilePath)
        
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): PRM files deleted from temp location."
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Return value: " & GetAmortizationSchedule
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Exit GetAmortizationSchedule() method"

CleanUp:
    
    Set lobjPRMFileListXMLDOM = Nothing
    Set lobjXMLSchemaSpace = Nothing
    Set lobjReturnAmortSchedLstXMLDOM = Nothing
    
    Exit Function
    
ErrHandler:
    
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description
    
    If lbGetAmort Then
        
        lbGetAmort = False
        
        'Add the <ERROR> node to the Return XML for the PRM file
        AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                        lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild, _
                        "ERROR", _
                        ""
            
        AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                        lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild, _
                        "ERROR_NBR", _
                        llErrNbr
        
        AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                        lobjReturnAmortSchedLstXMLDOM.documentElement.lastChild.lastChild, _
                        "ERROR_DESC", _
                        lstrErrDesc
                        
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Error retrieving Amort Sched - " & lstrErrDesc
        
        lbGetAmort = True
        
        Resume NextRec
    Else
        
        Set lobjReturnAmortSchedLstXMLDOM = Nothing
                
        'Build error XML
        Call lobjReturnAmortSchedLstXMLDOM.loadXML("<AMORTIZATION_SCHEDULE_LIST><ERROR></ERROR></AMORTIZATION_SCHEDULE_LIST>")
            
        AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                    lobjReturnAmortSchedLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_NBR", _
                    llErrNbr
                    
        AddXMLElement lobjReturnAmortSchedLstXMLDOM, _
                    lobjReturnAmortSchedLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_DESC", _
                    lstrErrDesc
        
        'Return the final PRM list
        GetAmortizationSchedule = lobjReturnAmortSchedLstXMLDOM.xml
        
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): General Error : " & GetAmortizationSchedule
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Exit GetAmortizationSchedule() method"
        
        Resume CleanUp
    End If
End Function

'================================================================
'METHOD  : ModifyPRMFiles
'PURPOSE : To modify the parameters contained in the binary PRM
'          files and return the modified binary PRM file and/or
'          XML equivalent for the binary PRM file and/or write
'          to a file location.
'PARMS   :
'          astrModifyPRMFilesXML [String] = XML string containing
'          the PRM file [either binary PRM file(s) or just path
'          to the binary PRM file(s)], Parameters to be modified
'          and the type of output that is expected (modified
'          binary PRM file and/or XML equivalent for the binary PRM file and/or write to a file location).
'RETURN  : String
'================================================================
Public Function ModifyPRMFiles(ByVal astrModifyPRMFilesXML As String) As String
On Error GoTo ErrHandler

    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): In ModifyPRMFiles() method"
    
'Declare Super Trump Variables
Dim lobjSTApplication               As New STApplication
Dim lobjSTTransaction               As New STTransaction

'Declare XML Dom variables
Dim lobjResponseXMLDOM              As New DOMDocument40
Dim lobjRequestXMLDOM               As New DOMDocument40
Dim lobjXMLSchemaSpace              As New XMLSchemaCache40
Dim lobjElem                        As IXMLDOMElement
Dim lobjPRMExceptionXMLDOM          As New DOMDocument40
Dim lobjExeceptionlst               As IXMLDOMNodeList
Dim lobjDocFragment                 As IXMLDOMDocumentFragment

'Other Declarations
Dim lstrSchemaFileLoc               As String
Dim lstrModifyPRMFilesXML           As String
Dim lstrPRMFilePath                 As String
Dim lstrTempPath                    As String
Dim lstrModifiedPRMFilePath         As String
Dim lbProcessingPRMFile             As Boolean
Dim liCnt                           As Integer
Dim lstrModifiedParamsXML           As String
Dim lstrPRMFileName                 As String
Dim lstrSTServerReqXML              As String
Dim lstrSTServerRespXML             As String
Dim liCnt2                          As Integer
Dim lvPRMFileData                   As Variant
Dim lbPRMFileTagAdded               As Boolean

Dim llErrNbr                As Long
Dim lstrErrSrc              As String
Dim lstrErrDesc             As String

Dim lobjErrComment          As IXMLDOMNode
Dim lstrErrComment          As String

    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Input Argument 1:" & astrModifyPRMFilesXML

    'Initialize Response XML
    Call lobjResponseXMLDOM.loadXML("<MODIFY_PRM_RESPONSE/>")
    
'Schema validation code commented for the time being till the schema is ready
'''    'Get the ModifyPRMFilesXML.xsd Schema
'''    lstrSchemaFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcSchemaFilePathKey)
'''
'''    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Schema file path read from registry"
'''
'''    Call lobjXMLSchemaSpace.Add("", lstrSchemaFileLoc & "\" & gcModifyPRMFilesSchemaName)
'''
'''    'Assign Schema to the XML DOM object
'''    Set lobjRequestXMLDOM.schemas = lobjXMLSchemaSpace
'''    lobjRequestXMLDOM.async = False
    
    'Remove "xmlns:dt" & "dt:dt" attributes from the <FILE_DATA> elements in the Input XML
    'as we cannot include these attributes in the schema as they contain special
    'characters(:). We will add them after schema validation.
    lstrModifyPRMFilesXML = Replace(astrModifyPRMFilesXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Validating Input XML"
    
    'Load the Input Request XML into the XML DOM object
    Call lobjRequestXMLDOM.loadXML(lstrModifyPRMFilesXML)
    
    'Check if Request XML is not valid
    If lobjRequestXMLDOM.parseError.errorCode <> 0 Then
        
        llErrNbr = lobjRequestXMLDOM.parseError.errorCode
        lstrErrDesc = lobjRequestXMLDOM.parseError.reason
        
        'Add the <ERROR> node to the Response XML
        AddXMLElement lobjResponseXMLDOM, _
                    lobjResponseXMLDOM.documentElement, _
                    "ERROR", _
                    ""
                    
        AddXMLElement lobjResponseXMLDOM, _
                    lobjResponseXMLDOM.documentElement.lastChild, _
                    "ERROR_NBR", _
                    llErrNbr
                
        AddXMLElement lobjResponseXMLDOM, _
                    lobjResponseXMLDOM.documentElement.lastChild, _
                    "ERROR_DESC", _
                    "Input XML Invalid!!! Error on line number " & lobjRequestXMLDOM.parseError.Line & " of the XML. " & lstrErrDesc
                      
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): " & lobjResponseXMLDOM.xml
         
        'Return the error to the client
        ModifyPRMFiles = lobjResponseXMLDOM.xml
        Exit Function
    End If
            
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Input XML Valid"
        
    'Get the temporary path where PRM files will be copied to disk
    lstrTempPath = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMFilePathKey)
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Read Temporary file path " & lstrTempPath
    
    'Set flag for start of PRM processing
    lbProcessingPRMFile = True
    
    'For each PRM file in Request XML
    For liCnt = 0 To lobjRequestXMLDOM.documentElement.childNodes.length - 1
        
        lbPRMFileTagAdded = False
        lstrPRMFileName = ""
        
        'Check if PRM file path is specified in Request XML
        Set lobjElem = lobjRequestXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE/FILE_PATH")
        If Not (lobjElem Is Nothing) Then
    
            'Get the PRM file path from Request XML
            lstrPRMFilePath = lobjElem.Text
            
        'Otherwise if the PRM file path is not specified in Request XML
        Else
            
            'Use temporary path and save file to disk
            lobjRequestXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE/FILE_DATA").dataType = "bin.base64"
            SavePRMBinaryFile lobjRequestXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE/FILE_NAME").Text, lstrTempPath, lobjRequestXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE/FILE_DATA").nodeTypedValue
            
            lstrPRMFilePath = lstrTempPath & "\" & lobjRequestXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("PRM_FILE/FILE_NAME").Text
        End If
        Set lobjElem = Nothing
                
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Processing " & lstrPRMFilePath
        
        'Add PRM file name to Response XML
        lstrPRMFileName = Mid(lstrPRMFilePath, InStrRev(lstrPRMFilePath, "\") + 1)
            
        AddXMLElement lobjResponseXMLDOM, _
                        lobjResponseXMLDOM.documentElement, _
                        "PRM_FILE", _
                        ""
        
        AddXMLElement lobjResponseXMLDOM, _
                        lobjResponseXMLDOM.documentElement.lastChild, _
                        "FILE_NAME", _
                        lstrPRMFileName
                        
        lbPRMFileTagAdded = True
        
        'Get the data to be modified from Request XML
        lstrModifiedParamsXML = lobjRequestXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("MODIFY_PARAMS").xml
        
        'Strip off the outer <MODIFY_PARAMS> tag
        lstrModifiedParamsXML = Replace(lstrModifiedParamsXML, "<MODIFY_PARAMS>", "")
        lstrModifiedParamsXML = Trim(Replace(lstrModifiedParamsXML, "</MODIFY_PARAMS>", ""))
        lstrModifiedParamsXML = Trim(Replace(lstrModifiedParamsXML, "<MODIFY_PARAMS/>", ""))
        lstrModifiedParamsXML = Trim(Replace(lstrModifiedParamsXML, vbCrLf, ""))
        
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): got modified params "
        
        'Check if Modified PRM file has to be copied to a specified location
        Set lobjElem = lobjRequestXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("OUTPUT_DETAILS/PRM_PATH")
        If Not (lobjElem Is Nothing) Then
        
            'Get the location from Request XML
            lstrModifiedPRMFilePath = lobjElem.Text
            
        'Otherwise if no location is specified
        Else
            
            'Use original PRM file path (PRM file will be overwritten)
            lstrModifiedPRMFilePath = lstrPRMFilePath
        End If
        
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): got modified prm path " & lstrModifiedPRMFilePath
        
        'Build the Input XML for the XMLInOut() method
        lstrSTServerReqXML = "<SuperTRUMP>" & _
                                "<Transaction query=""" & lstrPRMFileName & """>" & _
                                    "<Initialize/>" & _
                                    "<ReadFile filename=""" & lstrPRMFilePath & """ />" & _
                                    lstrModifiedParamsXML & _
                                    "<WriteFile filename=""" & lstrModifiedPRMFilePath & """ />" & _
                                "</Transaction>" & _
                            "</SuperTRUMP>"
        
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Input XML for the XMLInOut() method: " & lstrSTServerReqXML
        
        'Call the XMLInOut() method to modify the PRM file and return the XML equivalent
        lstrSTServerRespXML = lobjSTApplication.XmlInOut(lstrSTServerReqXML)
        
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Response XML for the XMLInOut() method: " & lstrSTServerRespXML
        
        'If cannot Load the XML representation into DOM
        If Not (lobjPRMExceptionXMLDOM.loadXML(lstrSTServerRespXML)) Then
            
            llErrNbr = lobjPRMExceptionXMLDOM.parseError.errorCode
            lstrErrDesc = lobjPRMExceptionXMLDOM.parseError.reason
        
            'Add the <ERROR> node for the PRM file to the Response XML
            AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild, _
                            "ERROR", _
                            ""
            
            AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild.lastChild, _
                            "ERROR_NBR", _
                            llErrNbr
            
            AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild.lastChild, _
                            "ERROR_DESC", _
                            "Error loading STServer response XML!!! " & lstrErrDesc
                                
            If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Error loading STServer response XML!!! " & lobjPRMExceptionXMLDOM.xml
            
        'Otherwise Check for exception returned from STServer
        ElseIf (lobjPRMExceptionXMLDOM.getElementsByTagName("Exception").length) > 0 Then
            
            'Add the <ERROR> node for the PRM file to the Response XML
            Set lobjExeceptionlst = lobjPRMExceptionXMLDOM.getElementsByTagName("Exception")
                
            AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild, _
                            "ERROR", _
                            ""
            
            AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild.lastChild, _
                            "ERROR_NBR", _
                            lobjExeceptionlst.Item(0).selectSingleNode("Number").Text
            
            Set lobjErrComment = lobjExeceptionlst.Item(0).selectSingleNode("Comment")
            lstrErrComment = ""
            If Not (lobjErrComment Is Nothing) Then lstrErrComment = lobjErrComment.Text
                
            AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild.lastChild, _
                            "ERROR_DESC", _
                            "Exception returned from STServer!!! " & lobjExeceptionlst.Item(0).selectSingleNode("Description").Text & " " & lstrErrComment
                            
            'lobjResponseXMLDOM.selectSingleNode("//ERROR").appendChild (lobjPRMExceptionXMLDOM.selectSingleNode("SuperTRUMP"))
            Set lobjDocFragment = lobjResponseXMLDOM.createDocumentFragment
            lobjDocFragment.appendChild lobjResponseXMLDOM.createElement("PRM_XML")
            lobjDocFragment.selectSingleNode("PRM_XML").appendChild lobjPRMExceptionXMLDOM.selectSingleNode("SuperTRUMP")
            'lobjResponseXMLDOM.documentElement.appendChild lobjDocFragment
            lobjResponseXMLDOM.selectSingleNode("//MODIFY_PRM_RESPONSE/PRM_FILE/ERROR").appendChild lobjDocFragment
            
'            AddXMLElement lobjResponseXMLDOM, _
'                            lobjResponseXMLDOM.selectSingleNode("//ERROR"), _
'                            "PRM_XML", lobjPRMExceptionXMLDOM.selectSingleNode("SuperTRUMP").xml
            
                           
            If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Exception returned from STServer!!! " & lobjPRMExceptionXMLDOM.xml
            
        'Otherwise if no exception
        Else
        
            'For each output specification
            For liCnt2 = 0 To lobjRequestXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("OUTPUT_DETAILS").childNodes.length - 1
                
                Select Case lobjRequestXMLDOM.documentElement.childNodes(liCnt).selectSingleNode("OUTPUT_DETAILS").childNodes(liCnt2).nodeName
                
                    'If XML equivalent of Modified PRM file has to be returned
                    Case "PRM_XML"
                    
                        'Add XML equivalent to Response XML
                        AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild, _
                            "PRM_XML", _
                            ""
                        
                        lobjResponseXMLDOM.documentElement.lastChild.lastChild.appendChild lobjPRMExceptionXMLDOM.documentElement
                        
                    'If Modified PRM file has to be returned
                    Case "PRM_FILE"
                    
                        
                        'Read the modified PRM File from disk
                        lvPRMFileData = GetBinaryFileData(lstrModifiedPRMFilePath)
                                                    
                        'Add Modified PRM file to Response XML
                        AddBinaryXMLElement lobjResponseXMLDOM, _
                                            lobjResponseXMLDOM.documentElement.lastChild, _
                                            "FILE_DATA", _
                                            lvPRMFileData
                                            
                    'If Modified PRM file has to be copied to a specified location
                    Case "PRM_PATH"
                    
                        'Add Success status to Response XML
                        AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild, _
                            "STATUS", _
                            "Successfully copied to " & lstrModifiedPRMFilePath
                            
                End Select
                
            Next liCnt2
                
        End If

ProcessNextPRMFile:

    Next liCnt
    
    'Set flag for end of PRM processing
    lbProcessingPRMFile = False
    
    ModifyPRMFiles = lobjResponseXMLDOM.xml
    
    'Delete the saved PRM files from disk
    Call DeletePRMBinaryFile(lobjResponseXMLDOM.getElementsByTagName("FILE_NAME"), lstrTempPath)
        
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): PRM files deleted from temp location."
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Output:" & ModifyPRMFiles
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Exit ModifyPRMFiles"
    
    If Not (lobjDocFragment Is Nothing) Then
        Set lobjDocFragment = Nothing
    End If
    
    Exit Function
    
'In Error handler
ErrHandler:
    
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description
            
    'If PRM file is being processed
    If lbProcessingPRMFile Then
        
        If lstrPRMFileName <> "" Then
            
            lbProcessingPRMFile = False
            
            'Add <PRM_FILE> & <FILE_NAME> tag if it has not been added
            If Not (lbPRMFileTagAdded) Then

                AddXMLElement lobjResponseXMLDOM, _
                                lobjResponseXMLDOM.documentElement, _
                                "PRM_FILE", _
                                ""
        
                AddXMLElement lobjResponseXMLDOM, _
                                lobjResponseXMLDOM.documentElement.lastChild, _
                                "FILE_NAME", _
                                lstrPRMFileName
            End If
            
            'Add the <ERROR> node for the PRM file to the Response XML
            AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild, _
                            "ERROR", _
                            ""
            
            AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild.lastChild, _
                            "ERROR_NBR", _
                            llErrNbr
            
            AddXMLElement lobjResponseXMLDOM, _
                            lobjResponseXMLDOM.documentElement.lastChild.lastChild, _
                            "ERROR_DESC", _
                            "Error Processing PRM file!!! " & lstrErrDesc
                                    
            If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Error Processing PRM file" & lstrPRMFileName & "!!! " & lstrErrDesc
            
            lbProcessingPRMFile = True
            
            'Process next PRM file
            Resume ProcessNextPRMFile
        End If
    End If
    
    'If any other error
    Call lobjResponseXMLDOM.loadXML("<MODIFY_PRM_RESPONSE/>")
    
    'Add the <ERROR> node to the Response XML
    AddXMLElement lobjResponseXMLDOM, _
                lobjResponseXMLDOM.documentElement, _
                "ERROR", _
                ""
                
    AddXMLElement lobjResponseXMLDOM, _
                lobjResponseXMLDOM.documentElement.lastChild, _
                "ERROR_NBR", _
                llErrNbr
            
    AddXMLElement lobjResponseXMLDOM, _
                lobjResponseXMLDOM.documentElement.lastChild, _
                "ERROR_DESC", _
                "General Failure!!! " & lstrErrDesc
    
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): " & lobjResponseXMLDOM.xml
    
    ModifyPRMFiles = lobjResponseXMLDOM.xml
    
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Exit ModifyPRMFiles"
End Function

'================================================================
'METHOD  : GeneratePRMFilesForPmtStruct
'PURPOSE : To generate binary PRM file for each set of PRM
'          parameters and Meta data. The PRM parameters contains
'          the payment structure. This method is solving for
'          payments.
'PARMS   :
'          astrPRMInfoXML [String] = XML string containing the
'          PRM Parameters and Meta data required to generate the
'          binary PRM file(s).
'
'          Sample Input Parameter structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_INFO>
'                <PRM_FILE>
'                    <PRM_META_DATA>
'                        <FILE_NAME>MyPRMFile.prm</FILE_NAME>
'                        <TEMPLATE_NAME>USA 5 MACRS.TEM</TEMPLATE_NAME>
'                        <MODE>Lessor</MODE>
'                    </PRM_META_DATA>
'                    <PRM_PARAMS>
'                        <TRANSACTIONAMOUNT>25000000</TRANSACTIONAMOUNT>
'                        <TRANSACTIONSTARTDATE>2002-08-20</TRANSACTIONSTARTDATE>
'                        <PERIODICITY>Monthly</PERIODICITY>
'                        <PAYMENTTIMING>Advance</PAYMENTTIMING>
'                        <STRUCTURE>Level</STRUCTURE>
'                        ...
'                    </PRM_PARAMS>
'                </PRM_FILE>
'                <PRM_FILE>
'                    <PRM_META_DATA>
'                        <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
'                        ...
'                    </PRM_META_DATA>
'                    …
'                </PRM_FILE>
'                …
'            </PRM_INFO>
'RETURN  : String= XML string containing, the binary PRM File or
'          <ERROR> tag, for each set of PRM Input Parameters.
'          It may also return an <ERROR> tag for any general
'          failure condition.
'
'            Sample Return XML structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_FILE_LIST>
'
'                <!-- Sucessful generation of PRM file -->
'                <PRM_FILE>
'                    <FILE_NAME>MyPRMFile.prm</FILE_NAME>
'                    <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
'                </PRM_FILE>
'
'                <!-- Error generating PRM file -->
'                <PRM_FILE>
'                    <FILE_NAME>ErrorPRMFile.prm </FILE_NAME>
'                    <ERROR>
'                        <ERROR_NBR>-1072896682</ERROR_NBR>
'                        <ERROR_DESC>Error!!!...</ERROR_DESC>
'                    </ERROR>
'                </PRM_FILE>
'                …
'            </PRM_FILE_LIST>
'
'            OR In case of general failure:
'
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_FILE_LIST>
'                <ERROR>
'                    <ERROR_NBR>-1072896682</ERROR_NBR>
'                    <ERROR_DESC>Error!!!...</ERROR_DESC>
'                </ERROR>
'            </PRM_FILE_LIST>
'================================================================
Public Function GeneratePRMFilesForPmtStruct(ByVal astrPRMInfoXML As String) As String

    GeneratePRMFilesForPmtStruct = GeneratePRMFilesForPmtStructure(astrPRMInfoXML, _
                                                                        ecSolveForPayments)
End Function

'================================================================
'METHOD  : GeneratePRMFilesForPmtStruct2
'PURPOSE : To generate binary PRM file for each set of PRM
'          parameters and Meta data. The PRM parameters contains
'          the payment structure. This method is solving for
'          rate.
'PARMS   :
'          astrPRMInfoXML [String] = XML string containing the
'          PRM Parameters and Meta data required to generate the
'          binary PRM file(s).
'
'          Sample Input Parameter structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_INFO>
'                <PRM_FILE>
'                    <PRM_META_DATA>
'                        <FILE_NAME>MyPRMFile.prm</FILE_NAME>
'                        <TEMPLATE_NAME>USA 5 MACRS.TEM</TEMPLATE_NAME>
'                        <MODE>Lessor</MODE>
'                    </PRM_META_DATA>
'                    <PRM_PARAMS>
'                        <TRANSACTIONAMOUNT>25000000</TRANSACTIONAMOUNT>
'                        <TRANSACTIONSTARTDATE>2002-08-20</TRANSACTIONSTARTDATE>
'                        <PERIODICITY>Monthly</PERIODICITY>
'                        <PAYMENTTIMING>Advance</PAYMENTTIMING>
'                        <STRUCTURE>Level</STRUCTURE>
'                        ...
'                    </PRM_PARAMS>
'                </PRM_FILE>
'                <PRM_FILE>
'                    <PRM_META_DATA>
'                        <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
'                        ...
'                    </PRM_META_DATA>
'                    …
'                </PRM_FILE>
'                …
'            </PRM_INFO>
'RETURN  : String= XML string containing, the binary PRM File or
'          <ERROR> tag, for each set of PRM Input Parameters.
'          It may also return an <ERROR> tag for any general
'          failure condition.
'
'            Sample Return XML structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_FILE_LIST>
'
'                <!-- Sucessful generation of PRM file -->
'                <PRM_FILE>
'                    <FILE_NAME>MyPRMFile.prm</FILE_NAME>
'                    <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
'                </PRM_FILE>
'
'                <!-- Error generating PRM file -->
'                <PRM_FILE>
'                    <FILE_NAME>ErrorPRMFile.prm </FILE_NAME>
'                    <ERROR>
'                        <ERROR_NBR>-1072896682</ERROR_NBR>
'                        <ERROR_DESC>Error!!!...</ERROR_DESC>
'                    </ERROR>
'                </PRM_FILE>
'                …
'            </PRM_FILE_LIST>
'
'            OR In case of general failure:
'
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_FILE_LIST>
'                <ERROR>
'                    <ERROR_NBR>-1072896682</ERROR_NBR>
'                    <ERROR_DESC>Error!!!...</ERROR_DESC>
'                </ERROR>
'            </PRM_FILE_LIST>
'================================================================
Public Function GeneratePRMFilesForPmtStruct2(ByVal astrPRMInfoXML As String) As String

    GeneratePRMFilesForPmtStruct2 = GeneratePRMFilesForPmtStructure(astrPRMInfoXML, _
                                                                        ecSolveForRate)
End Function

'================================================================
'METHOD  : GeneratePRMFilesForPmtStruct
'PURPOSE : To generate binary PRM file for each set of PRM
'          parameters and Meta data. The PRM parameters contains
'          the payment structure.
'PARMS   :
'          astrPRMInfoXML [String] = XML string containing the
'          PRM Parameters and Meta data required to generate the
'          binary PRM file(s).
'
'          Sample Input Parameter structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_INFO>
'                <PRM_FILE>
'                    <PRM_META_DATA>
'                        <FILE_NAME>MyPRMFile.prm</FILE_NAME>
'                        <TEMPLATE_NAME>USA 5 MACRS.TEM</TEMPLATE_NAME>
'                        <MODE>Lessor</MODE>
'                    </PRM_META_DATA>
'                    <PRM_PARAMS>
'                        <TRANSACTIONAMOUNT>25000000</TRANSACTIONAMOUNT>
'                        <TRANSACTIONSTARTDATE>2002-08-20</TRANSACTIONSTARTDATE>
'                        <PERIODICITY>Monthly</PERIODICITY>
'                        <PAYMENTTIMING>Advance</PAYMENTTIMING>
'                        <STRUCTURE>Level</STRUCTURE>
'                        ...
'                    </PRM_PARAMS>
'                </PRM_FILE>
'                <PRM_FILE>
'                    <PRM_META_DATA>
'                        <FILE_NAME>ErrorPRMFile.prm</FILE_NAME>
'                        ...
'                    </PRM_META_DATA>
'                    …
'                </PRM_FILE>
'                …
'            </PRM_INFO>
'
'            astrMethodType [eSolveMethod] = Method to be used to balance
'            payments -- Solve for Payments or Solve for Rate.
'RETURN  : String= XML string containing, the binary PRM File or
'          <ERROR> tag, for each set of PRM Input Parameters.
'          It may also return an <ERROR> tag for any general
'          failure condition.
'
'            Sample Return XML structure:
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_FILE_LIST>
'
'                <!-- Sucessful generation of PRM file -->
'                <PRM_FILE>
'                    <FILE_NAME>MyPRMFile.prm</FILE_NAME>
'                    <FILE_DATA>/CQAGAAAAAAAAAAAAAAACAAAA3AAAAAAA…</FILE_DATA>
'                </PRM_FILE>
'
'                <!-- Error generating PRM file -->
'                <PRM_FILE>
'                    <FILE_NAME>ErrorPRMFile.prm </FILE_NAME>
'                    <ERROR>
'                        <ERROR_NBR>-1072896682</ERROR_NBR>
'                        <ERROR_DESC>Error!!!...</ERROR_DESC>
'                    </ERROR>
'                </PRM_FILE>
'                …
'            </PRM_FILE_LIST>
'
'            OR In case of general failure:
'
'            <?xml version="1.0" encoding="UTF-8"?>
'            <PRM_FILE_LIST>
'                <ERROR>
'                    <ERROR_NBR>-1072896682</ERROR_NBR>
'                    <ERROR_DESC>Error!!!...</ERROR_DESC>
'                </ERROR>
'            </PRM_FILE_LIST>
'================================================================
Private Function GeneratePRMFilesForPmtStructure(ByVal astrPRMInfoXML As String, _
                                                ByVal astrMethodType As eSolveMethod) As String
On Error GoTo ErrHandler
    
'Declare Super Trump Variables
Dim lobjSTApplication       As New STApplication
Dim lobjSTTransaction       As New STTransaction
Dim lobjSTQuick             As STQuick
Dim lobjSTCashflows         As STCashFlows
Dim lobjSTCashflow          As STCashFlow

'Declare XML Dom variables
Dim lobjPRMInfoXMLDOM       As New DOMDocument40
Dim lobjXMLSchemaSpace      As New XMLSchemaCache40
Dim lobjReturnPRMLstXMLDOM  As New DOMDocument40
Dim lobjSTQueryXMLDOM       As New DOMDocument40
Dim lobjSTResponseXMLDOM    As New DOMDocument40
Dim lobjExeceptionlst       As IXMLDOMNodeList
Dim lobjPaymentNodes        As IXMLDOMNodeList

'Other Declarations
Dim lstrFileLoc             As String
Dim lstrPRMInfoXML          As String
Dim liPRMParamsCnt          As Integer
Dim liPaymentsCnt           As Integer
Dim lstrPRMFilePath         As String
Dim lstrPRMTemplatePath     As String
Dim lstrPRMMode             As String
Dim lstrReturnXML           As String
Dim lstrPRMFileName         As String
Dim lvPRMFileData           As Variant
Dim lbGenPRM                As Boolean

Dim llErrNbr                As Long
Dim lstrErrSrc              As String
Dim lstrErrDesc             As String

Dim lstrTransactionAmt      As String
Dim lstrLendingRate         As String
Dim lstrCommencementDt      As String
Dim lstrPeriodicity         As String
Dim licount                 As Integer
Dim lstrMoneyCostDate       As String
Dim lstrResidualAmt         As String
Dim lstrBalloonAmt          As String
Dim lstrGEBusiness          As String
Dim lstrGEProduct           As String

'New variables declared to force the aniequity error in the output in 9549
'which is supressing it: 5th march 2007
Dim liExceptionCount        As Integer
Dim lobjExcptionXMLDOM      As New DOMDocument40
Dim liPRMFile               As Integer

Dim lobjFeesNodeList As MSXML2.IXMLDOMNodeList
    
    lbGenPRM = False
    liPRMFile = 0
    
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): In GeneratePRMFilesForPmtStructure() method"
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Input Argument 1:" & astrPRMInfoXML
    
    'Load Return XML
    Call lobjReturnPRMLstXMLDOM.loadXML("<PRM_FILE_LIST></PRM_FILE_LIST>")
    
    'Get the GeneratePRMForPmtStructXML.xsd Schema
    lstrFileLoc = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcSchemaFilePathKey)

    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Schema file read from registry"

    Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcGenPRMForPmtStructSchemaName)

    'Assign Schema to the XML DOM object
    Set lobjPRMInfoXMLDOM.schemas = lobjXMLSchemaSpace
    lobjPRMInfoXMLDOM.async = False
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Validating Input XML"
    
    'Load the Input XML into the XML DOM object
    Call lobjPRMInfoXMLDOM.loadXML(astrPRMInfoXML)
    
    'Check if Input XML is valid
    If lobjPRMInfoXMLDOM.parseError.errorCode = 0 Then
        
        If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Input XML Valid"
        
        'Get File Paths
        lstrPRMFilePath = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMFilePathKey)
        lstrPRMTemplatePath = ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcPRMTemplatePathKey)
        
        'For Each set of PRM Parameters in the Input XML
         For liPRMParamsCnt = 0 To lobjPRMInfoXMLDOM.documentElement.childNodes.length - 1
            
            lbGenPRM = True
            
            'Build the Input XML for the XMLInOut() method
            Call lobjSTQueryXMLDOM.loadXML("<SuperTRUMP>" & _
                                            "<Transaction id='TRANS_ID_GEN_PRM' query='true'/>" & _
                                        "</SuperTRUMP>")
            
            'Mode
            lstrPRMMode = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_META_DATA/MODE")
            
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "MODE", _
                        lstrPRMMode
            
            'Initialize
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "INITIALIZE", _
                        ""
            
            'Read Template
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "READFILE", _
                        ""
            
            AddXMLElementAttribute lobjSTQueryXMLDOM, _
                                    lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                                    "filename", _
                                    lstrPRMTemplatePath & "\" & GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                                                        "PRM_META_DATA/TEMPLATE_NAME")
            
            'Transaction Amount
            lstrTransactionAmt = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/TRANSACTIONAMOUNT")
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "TRANSACTIONAMOUNT", _
                        lstrTransactionAmt
                                                
            'Transaction Date
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "TRANSACTIONSTARTDATE", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/TRANSACTIONSTARTDATE")
            
            'Structure
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "STRUCTURE", _
                        "Other"
            
            'Periodicity
            lstrPeriodicity = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/PERIODICITY")
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "PERIODICITY", _
                        lstrPeriodicity
                                    
            'Payment timing
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "PAYMENTTIMING", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/PAYMENTTIMING")
                        
            
            'Commencement Date
            lstrCommencementDt = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/COMMENCEMENTDATE")
            If lstrCommencementDt <> "" Then
                AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                            "COMMENCEMENTDATE", _
                            lstrCommencementDt
            End If
            
            'added by Neil to send the Expense break up of GE and Dealer Fees.
            '---------------------------------------------------------------------
            If lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt).selectNodes("PRM_PARAMS/FEES").length > 0 Then
                AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "FEES", _
                        ""
                AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                        "DELETE", _
                        ""
                     
                AddXMLElementAttribute lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.firstChild, _
                        "INDEX", _
                        "*"
                      
                Set lobjFeesNodeList = lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt).selectNodes("PRM_PARAMS/FEES/FEE")
                
                For licount = 0 To lobjFeesNodeList.length - 1
                                          
                        AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                        "FEE", _
                        ""
                                       
                        AddXMLElementAttribute lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.lastChild, _
                        "INDEX", _
                        licount
                                       
                                             
                        AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.lastChild, _
                        "KeptAsAPercent", _
                        "false"
                        
                        AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.lastChild, _
                        "IsAnExpense", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt).selectSingleNode("PRM_PARAMS/FEES").childNodes(licount), _
                                                            "ISANEXPENSE")
                                               
                        AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.lastChild, _
                        "Amount", _
                        GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt).selectSingleNode("PRM_PARAMS/FEES").childNodes(licount), _
                                                "AMOUNT")
                                    
                        AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.lastChild, _
                        "FederalDepreciation", _
                        ""
                        
                        AddXMLElement lobjSTQueryXMLDOM, _
                                lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.lastChild.lastChild, _
                                "Method", _
                                "Expensed"
                  
                Next
                
            End If
            
            '--------------------------------------------------------------------
                        
            
                                                
            'for Lease
            If UCase(lstrPRMMode) = "LESSOR" Then
                    
                'Residual Amout
                lstrResidualAmt = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                    "PRM_PARAMS/RESIDUALAMOUNT")
                                              
                If lstrResidualAmt <> "" Then
                    AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                            "ASSETS", _
                            ""
                    
                    AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                            "ASSET", _
                            ""
                                            
                    AddXMLElementAttribute lobjSTQueryXMLDOM, _
                                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0), _
                                        "index", _
                                        0
                                        
                     AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0), _
                            "ResidualKeptAsAPercent", _
                            "false"
                                                         
                    AddXMLElement lobjSTQueryXMLDOM, _
                                lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0), _
                                "RESIDUAL", _
                                lstrResidualAmt
                End If
                
            'For loan
            ElseIf UCase(lstrPRMMode) = "LENDER" Then
            
                'Lending Rate
                lstrLendingRate = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                    "PRM_PARAMS/LENDINGRATE")
                
                'Balloon Amount
                lstrBalloonAmt = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                    "PRM_PARAMS/BALLOON")
                                                    
                If lstrBalloonAmt <> "" Then
                    AddXMLElement lobjSTQueryXMLDOM, _
                                lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                                "BALLOON", _
                                lstrBalloonAmt
                End If
                
                '===========Start of Payment Structure =======================================================
                AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                            "LENDINGLOANS", _
                            ""
                
                AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                        "LENDINGLOAN", _
                        ""
                
                AddXMLElementAttribute lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0), _
                            "INDEX", _
                            "0"
                
                AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0), _
                        "CASHFLOWSTEPS", _
                        ""
                
                'Delete all old cashflow steps
                AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0), _
                        "DELETE", _
                        ""
                
                AddXMLElementAttribute lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                        "INDEX", _
                        "*"
                
                'Add funding payment
                AddXMLElement lobjSTQueryXMLDOM, _
                    lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0), _
                    "CASHFLOWSTEP", _
                    ""
                    
                AddXMLElementAttribute lobjSTQueryXMLDOM, _
                    lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                    "INDEX", _
                    "0"
                    
                AddXMLElement lobjSTQueryXMLDOM, _
                    lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                    "TYPE", _
                    "Funding"
                        
                AddXMLElement lobjSTQueryXMLDOM, _
                    lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                    "AMOUNT", _
                    -1 * Val(lstrTransactionAmt)
                    
                AddXMLElement lobjSTQueryXMLDOM, _
                    lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                    "AMOUNTLOCKED", _
                    "true"
                                    
                'Add cashflow step for each payment
                Set lobjPaymentNodes = lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt).selectNodes("PRM_PARAMS/PAYMENTS/PAYMENT")
                For liPaymentsCnt = 0 To lobjPaymentNodes.length - 1
                    
                    AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0), _
                        "CASHFLOWSTEP", _
                        ""
                    
                    AddXMLElementAttribute lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                        "INDEX", _
                        liPaymentsCnt + 1
                    
                    AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                        "NUMBEROFPAYMENTS", _
                        GetXMLElementValue(lobjPaymentNodes(liPaymentsCnt), _
                                                    "NUMBEROFPAYMENTS")
                    
                    AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                        "TYPE", _
                        "Payment"
                    
                    'For stub payment
                    If UCase(GetXMLElementValue(lobjPaymentNodes(liPaymentsCnt), _
                                                    "PAYMENT_TYPE")) = "STUB" Then
                        
                        AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                            "PERIODICITY", _
                            "Stub"
                        
                        AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                            "ENDACCRUAL", _
                            lstrCommencementDt
                    
                    'For regular payments we have to specify periodicity otherwise it will inherit from the stub payment
                    'which might cause an issue.
                    Else
                        AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                            "PERIODICITY", _
                            lstrPeriodicity
                    End If
                    
                    AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                        "RATE", _
                        lstrLendingRate
                        
                    AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                        "AMOUNT", _
                        GetXMLElementValue(lobjPaymentNodes(liPaymentsCnt), _
                                                    "AMOUNT")
                    
                    AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild.childNodes(0).childNodes(0).lastChild, _
                        "AMOUNTLOCKED", _
                        GetXMLElementValue(lobjPaymentNodes(liPaymentsCnt), _
                                                    "AMOUNTLOCKED")
                                                    
                Next liPaymentsCnt
                '===========End of Payment Structure =======================================================
            End If
            
            'Business
            AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                            "GEDATA", _
                            ""
            
            lstrGEBusiness = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                    "PRM_PARAMS/GEBUSINESS")
            
            If lstrGEBusiness <> "" Then
                AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                            "GEBUSINESS", _
                            lstrGEBusiness
            End If
            
            'Product
            lstrGEProduct = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                    "PRM_PARAMS/GEPRODUCT")
                                                    
            If lstrGEProduct <> "" Then
                AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                            "GEPRODUCT", _
                            lstrGEProduct
            End If
                                                             
            'add the moneyCost date....
            lstrMoneyCostDate = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                "PRM_PARAMS/MONEYCOSTDATE")
                      
                                                 
            If lstrMoneyCostDate <> "" Then
                AddXMLElement lobjSTQueryXMLDOM, _
                            lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                            "MONEYCOSTDATE", _
                            lstrMoneyCostDate
            End If
            
            'Write PRM file
            AddXMLElement lobjSTQueryXMLDOM, _
                        lobjSTQueryXMLDOM.documentElement.childNodes(0), _
                        "WRITEFILE", _
                        ""
                                    
            lstrPRMFileName = GetXMLElementValue(lobjPRMInfoXMLDOM.documentElement.childNodes(liPRMParamsCnt), _
                                                    "PRM_META_DATA/FILE_NAME")
            
            AddXMLElementAttribute lobjSTQueryXMLDOM, _
                                    lobjSTQueryXMLDOM.documentElement.childNodes(0).lastChild, _
                                    "filename", _
                                    lstrPRMFilePath & "\" & lstrPRMFileName
                                                
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Calling XMLInOut() method with input - " & lobjSTQueryXMLDOM.xml
            
            'Generate the PRM file.
            lstrReturnXML = lobjSTApplication.XmlInOut(lobjSTQueryXMLDOM.xml)
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): " & lstrPRMFileName & " file generated."
            
            AddXMLElement lobjReturnPRMLstXMLDOM, _
                                lobjReturnPRMLstXMLDOM.documentElement, _
                                "PRM_FILE", _
                                ""
            
            AddXMLElement lobjReturnPRMLstXMLDOM, _
                                lobjReturnPRMLstXMLDOM.documentElement.lastChild, _
                                "FILE_NAME", _
                                lstrPRMFileName
                                
            'Load the super Trump response XML
            
            ReDim Preserve gliPRMFilearr(liPRMFile + 1)
            gliPRMFilearr(liPRMFile) = 0
            
            ReDim Preserve gstrExcptionXMLDOMarr(liPRMFile + 1)
            gstrExcptionXMLDOMarr(liPRMFile) = ""
            
            Call lobjSTResponseXMLDOM.loadXML(lstrReturnXML)
            
            'Check for any Exception.
            Set lobjExeceptionlst = lobjSTResponseXMLDOM.getElementsByTagName("Exception")
                        
            If (lobjExeceptionlst.length) > 0 Then
                   For liExceptionCount = 0 To lobjExeceptionlst.length - 1
                        llErrNbr = lobjExeceptionlst.Item(liExceptionCount).childNodes(0).Text
                        lstrErrDesc = lobjExeceptionlst.Item(liExceptionCount).childNodes(2).Text
                        
                        '9549 is supressing all the errors in the output. BTO while processing checks
                        'only the aniequity error so that ICE can recheck the credit ratings for particular
                        'business and product. For any other error that is coming , its not captured.
                        'The new version anyway was supressing all the logical errors.
                        '5th March 2007
                        If InStr(UCase(lstrErrDesc), "ANIEQUITY") > 0 Then
                           gstrExceptionFlag = "TRUE"
                           gliPRMFilearr(liPRMFile) = liPRMFile + 1
                           Call lobjExcptionXMLDOM.loadXML(lobjExeceptionlst.Item(liExceptionCount).xml)
                           gstrExcptionXMLDOMarr(liPRMFile) = lobjExcptionXMLDOM.xml
                           Exit For ' no more errors are reqd after aniequity error is found.so exiting for.
                        End If
                    Next

                'Since we haven't solve for payments there will be exception which we need to ignore.
                'Just writing to log file for debugging purposes.
                If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Exception from STServer : " & lobjSTResponseXMLDOM.selectSingleNode("//Exceptions").xml
            End If
            
            'Open the PRM file
            lobjSTTransaction.OpenFile (lstrPRMFilePath & "\" & lstrPRMFileName)
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): PRM file read from " & lstrPRMFilePath
            
            lobjSTTransaction.Calculate
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): PRM file values calculated"
            
            'Solve for Payments or Rate
            If UCase(lstrPRMMode) = "LESSOR" Then   'for Lease
                Set lobjSTCashflows = lobjSTTransaction.GetCashFlows(ST_CF_Rent)
            ElseIf UCase(lstrPRMMode) = "LENDER" Then   'for Loan
                Set lobjSTCashflows = lobjSTTransaction.GetCashFlows(ST_CF_LoansLend)
            End If
            Set lobjSTCashflow = lobjSTCashflows.GetAt(0)
            
            If astrMethodType = ecSolveForPayments Then
                lobjSTCashflow.BalancePayments
            Else
                lobjSTCashflow.BalanceRate
            End If
        
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): PRM solved for payments"
            
            'Save changes
            lobjSTTransaction.SaveFile lstrPRMFilePath & "\" & lstrPRMFileName
            
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): PRM file changes updated to disk"
            
            Set lobjSTCashflow = Nothing
            Set lobjSTCashflows = Nothing
            Set lobjSTQuick = Nothing
            Set lobjSTTransaction = Nothing
                        
            'Read the generated PRM File from disk
            lvPRMFileData = GetBinaryFileData(lstrPRMFilePath & "\" & lstrPRMFileName)
                                            
            'Add it to the Return XML
            AddBinaryXMLElement lobjReturnPRMLstXMLDOM, _
                                    lobjReturnPRMLstXMLDOM.documentElement.lastChild, _
                                    "FILE_DATA", _
                                    lvPRMFileData
                                    
            If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Final PRM file read from disk and added to output XML"
            
NextRec:

            Set lobjSTResponseXMLDOM = Nothing
            Set lobjExeceptionlst = Nothing
            Set lobjSTQueryXMLDOM = Nothing
            liPRMFile = liPRMFile + 1
            
        Next
    
    'Else if input XML is invalid
    Else
        llErrNbr = lobjPRMInfoXMLDOM.parseError.errorCode
        lstrErrDesc = lobjPRMInfoXMLDOM.parseError.reason
        
        'Add the <ERROR> node to the Return XML
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                    lobjReturnPRMLstXMLDOM.documentElement, _
                    "ERROR", _
                    ""
                    
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                    lobjReturnPRMLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_NBR", _
                    llErrNbr
                
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                    lobjReturnPRMLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_DESC", _
                    "Error on line number " & lobjPRMInfoXMLDOM.parseError.Line & " of the XML. " & lstrErrDesc
    End If
    
    'Return the final PRM list
    GeneratePRMFilesForPmtStructure = lobjReturnPRMLstXMLDOM.xml
                
    'Delete the saved PRM files from disk
    Call DeletePRMBinaryFile(lobjReturnPRMLstXMLDOM.getElementsByTagName("FILE_NAME"), lstrPRMFilePath)
    
    If giDebugLevel > 2 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): PRM files deleted from temp location."
    If giDebugLevel > 1 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Return value: " & GeneratePRMFilesForPmtStructure
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Exit GeneratePRMFilesForPmtStructure() method"

CleanUp:
    
    Set lobjSTApplication = Nothing
    Set lobjPRMInfoXMLDOM = Nothing
    Set lobjXMLSchemaSpace = Nothing
    Set lobjReturnPRMLstXMLDOM = Nothing
    Set lobjExcptionXMLDOM = Nothing
    
    Exit Function
    
ErrHandler:
    
    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description

    If lbGenPRM Then
        
        lbGenPRM = False
        
        'Add the <ERROR> node to the Return XML for the PRM file
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                        lobjReturnPRMLstXMLDOM.documentElement.lastChild, _
                        "ERROR", _
                        ""
            
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                        lobjReturnPRMLstXMLDOM.documentElement.lastChild.lastChild, _
                        "ERROR_NBR", _
                        llErrNbr
        
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                        lobjReturnPRMLstXMLDOM.documentElement.lastChild.lastChild, _
                        "ERROR_DESC", _
                        lstrErrDesc
            
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): PRM file Generation error - " & lstrErrDesc
        
        lbGenPRM = True
        
        Resume NextRec
    Else
    
        Set lobjReturnPRMLstXMLDOM = Nothing
        
        'Build the Error XML
        Call lobjReturnPRMLstXMLDOM.loadXML("<PRM_FILE_LIST><ERROR></ERROR></PRM_FILE_LIST>")
            
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                    lobjReturnPRMLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_NBR", _
                    llErrNbr
                    
        AddXMLElement lobjReturnPRMLstXMLDOM, _
                    lobjReturnPRMLstXMLDOM.documentElement.childNodes(0), _
                    "ERROR_DESC", _
                    lstrErrDesc
                        
        'Return error XML
        GeneratePRMFilesForPmtStructure = lobjReturnPRMLstXMLDOM.xml
        
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): General error : " & GeneratePRMFilesForPmtStructure
        If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Exit GeneratePRMFilesForPmtStructure() method"
        
        Resume CleanUp
    End If
End Function

'================================================================
'METHOD  : RunAdHocXMLInOutQuery
'PURPOSE : to allow adhoc XML queries to be submitted via the
'           XMLInOut method in STServer. Values that are not
'           available in the ConvertPRMToXML xml structure can
'           be set/received this way. Files can be read,
'           modified and new versions written to disk
'PARMS   :
'          astrXMLInOutQuery [String] = XML string containing the
'          query to be executed and the file to be read
'
'          Sample Input Parameter structure:
'            <PRM_INFO>
'                <PRM_FILE>
'                    <AD_HOC_QUERY>
'                        <SuperTRUMP>
'                            <Transaction id="TRAN4">
'                                <ReadFile filename="\\ce213043914auct\Pricing$\test.prm"/>
'                                <TransactionAmount query="true"/>
'                            </Transaction>
'                        </SuperTRUMP>
'                    </AD_HOC_QUERY>
'                </PRM_FILE>
'            </PRM_INFO>
'
'RETURN  : String= XML string containing, the PRM query result or
'          <ERROR> tag, for each set of PRM Input Parameters.
'          It may also return an <ERROR> tag for any general
'          failure condition.
'================================================================

Public Function RunAdHocXMLInOutQuery(ByVal astrXMLInOutQuery As String) As String

'Super Trump Server Objects Declarations
Dim lobjSTApplication As STApplication
Dim lobjXMLDoc As MSXML2.DOMDocument40
Dim lobjXMLResultDoc As MSXML2.DOMDocument40
Dim lobjXMLNode As MSXML2.IXMLDOMNode

Dim llErrNbr As Long
Dim lstrErrSrc As String
Dim lstrErrDesc As String

On Error GoTo Err_Hand

    Set lobjXMLDoc = New MSXML2.DOMDocument40
    lobjXMLDoc.loadXML astrXMLInOutQuery
    
    'check for parser errors
    If lobjXMLDoc.parseError.errorCode <> 0 Then
        Err.Raise lobjXMLDoc.parseError.errorCode, lobjXMLDoc.parseError.srcText, "Error loading XML query: " & lobjXMLDoc.parseError.reason
    End If
    
    Set lobjSTApplication = New STApplication
    lobjXMLDoc.loadXML lobjSTApplication.XmlInOut(lobjXMLDoc.selectSingleNode("//PRM_INFO/PRM_FILE/AD_HOC_QUERY").childNodes.Item(0).xml)
    
    Set lobjXMLResultDoc = New MSXML2.DOMDocument40
    lobjXMLResultDoc.loadXML cADHOC_QUERY_RESULT_XML
        
    If InStr(lobjXMLDoc.childNodes(0).xml, "?xml") > 0 Then
        Set lobjXMLNode = lobjXMLDoc.childNodes(1)
    Else
        Set lobjXMLNode = lobjXMLDoc.childNodes(0)
    End If
    
    Set lobjXMLNode = lobjXMLResultDoc.selectSingleNode("//PRM_INFO/PRM_FILE/AD_HOC_QUERY").appendChild(lobjXMLNode)
    
    RunAdHocXMLInOutQuery = lobjXMLResultDoc.xml
    

'clean up
If Not (lobjSTApplication Is Nothing) Then
    Set lobjSTApplication = Nothing
End If
If Not (lobjXMLDoc Is Nothing) Then
    Set lobjXMLDoc = Nothing
End If
If Not (lobjXMLResultDoc Is Nothing) Then
    Set lobjXMLResultDoc = Nothing
End If

Exit Function
Err_Hand:

    llErrNbr = Err.Number
    lstrErrSrc = Err.Source
    lstrErrDesc = Err.Description
    
    'Return the Final XML with <ERROR> node specifying the application error
    RunAdHocXMLInOutQuery = "<PRM_INFO><PRM_FILE><AD_HOC_QUERY>" & _
                             "<ERROR>" & _
                                 "<ERROR_NBR>" & llErrNbr & "</ERROR_NBR>" & _
                                 "<ERROR_DESC><![CDATA[" & lstrErrDesc & "]]></ERROR_DESC>" & _
                             "</ERROR>" & _
                        "</AD_HOC_QUERY></PRM_FILE></PRM_INFO>"
    
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): General Error : " & RunAdHocXMLInOutQuery
    If giDebugLevel > 0 Then WriteToTextDebugFile gstrDebugFile, "BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): Exit RunAdHocXMLInOutQuery() method"
        
    
End Function
