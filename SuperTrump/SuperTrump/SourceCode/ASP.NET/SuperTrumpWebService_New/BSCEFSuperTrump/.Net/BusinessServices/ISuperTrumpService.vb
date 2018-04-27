Imports System.EnterpriseServices
Imports System.ComponentModel
Imports System.Configuration
Imports System.Runtime.InteropServices
Imports System.Xml
Imports System.IO
Imports Interop
<JustInTimeActivation(), _
 EventTrackingEnabled(), _
 Transaction(TransactionOption.Supported, Timeout:=120), _
 ComponentAccessControl(True)> _
Public Class ISuperTrumpService
    Inherits ServicedComponent
    Private Shared STLogger As log4net.ILog
    Dim obj As New Object
#Region "ISuperTrumpService Code Area"
    Private Const cADHOC_QUERY_RESULT_XML As String = "<PRM_INFO><PRM_FILE><AD_HOC_QUERY></AD_HOC_QUERY></PRM_FILE></PRM_INFO>"
    Private Enum eSolveMethod
        ecSolveForPayments = 0
        ecSolveForRate = 1
    End Enum
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
    <STAThreadAttribute()> _
    Function ConvertPRMToXML(ByVal astrPRMFileListXML As String) As String
        'Super Trump Server Objects Declarations
        Dim lobjSTApplication As STSERVER.STApplication
        Dim lobjPRMXMLDOM As New Xml.XmlDocument
        Dim lobjPRMExceptionXMLDOM As New Xml.XmlDocument
        Dim lobjFileNameList As Xml.XmlNodeList = Nothing
        Dim lobjFileDataList As Xml.XmlNodeList = Nothing
        Dim lobjExecpList As Xml.XmlNodeList = Nothing
        Dim lobjXMLSchemaSpace As New Xml.Schema.XmlSchemaSet
        Dim lobjPRMFileNameDOM As New Xml.XmlDocument
        Dim lobjElem As Xml.XmlNode
        'Other Declarations
        Dim lintCtrLoop As Short
        Dim lstrPRM2XML As String
        Dim lstrFileName As String
        Dim lstrPRMFileName As String
        Dim lstrFilePath As String
        Dim lstrSupertrumpQuery As String
        Dim lstrPRMBIN2XML As String
        Dim lstrFileLoc As String
        Dim lstrPRMFileListXML As String
        Dim llErrNbr As Integer
        Dim lstrErrSrc As String
        Dim lstrErrDesc As String
        Dim lobjErrComment As Xml.XmlNode
        Dim lstrErrComment As String

        Try
            SetLog4Net()
            lstrPRM2XML = "" ' Initilize variable to avoid Warning            
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): In ConvertPRMToXML() method")
            'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Input Argument 1:" & astrPRMFileListXML)

            'Get the PRMFileListXML.xsd Schema
            lstrFileLoc = GetConfigurationKey("SchemaFilePath")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Schema file read from registry")
            Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPRMFileLstXMLSchemaName)

            'Assign Schema to the XML DOM object
            lobjPRMXMLDOM.Schemas = lobjXMLSchemaSpace
            lstrPRMFileListXML = Replace(astrPRMFileListXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Validating Input XML")
            SyncLock obj
                Dim _ProcessID As String = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
                Dim _ThreadID As String = System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString()
                Dim MyGuid As Guid = Guid.NewGuid()

                'Load the Input XML into the XML DOM object
                Try
                    lobjPRMXMLDOM.LoadXml(lstrPRMFileListXML)
                Catch ex As Exception
                    'Raise Error                            
                    STLogger.Error(Err.Number & "BSCEFSuperTrump_ISuperTrumpService_ConvertPRMToXML():" & Err.Source & Err.Description)
                    lstrPRM2XML = "<PRM_FILE_LIST>" & "<ERROR>" & "<ERROR_NBR>" & Err.Number & "</ERROR_NBR>" & "<ERROR_DESC><![CDATA[Error  of the XML." & Err.Description & "]]></ERROR_DESC>" & "</ERROR>" & "</PRM_FILE_LIST>"
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Return value : " & lstrPRM2XML)
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Exit ConvertPRMToXML() method")
                    Return lstrPRM2XML
                End Try
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Input XML Valid")
                'Get the FileName and FileData list from DOM
                lobjFileNameList = lobjPRMXMLDOM.GetElementsByTagName("FILE_NAME")
                lobjFileDataList = lobjPRMXMLDOM.GetElementsByTagName("FILE_DATA")
                lstrFileLoc = GetConfigurationKey("PRMFilePath")

                'Traverse Each PRM File.]
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock Starts here")
                lobjSTApplication = New STSERVER.STApplication
                For lintCtrLoop = 0 To (lobjFileNameList.Count - 1)
                    Dim lstrFilePathChangedByProcID As String = ""
                    lstrFileName = lobjFileNameList.Item(lintCtrLoop).InnerText
                    If UCase(Right(lstrFileName, 4)) <> ".PRM" Then
                        lstrPRMFileName = lstrFileName & ".PRM"
                    Else
                        lstrPRMFileName = lstrFileName
                    End If
                    lstrFilePath = ""
                    'Create Supertrump XML Query.
                    lstrSupertrumpQuery = GetSuperTrumpQuery("TRANS_ID_" & Mid(lstrPRMFileName, 1, InStrRev(UCase(lstrPRMFileName), ".PRM") - 1), lstrFilePath, lobjFileDataList.Item(lintCtrLoop).InnerXml)
                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Calling XMLInOut() method")

                    'Get XML representation for the PRM file.                     
                    lstrPRMBIN2XML = lobjSTApplication.XmlInOut(lstrSupertrumpQuery)

                    ''**''STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Output from  XMLINOUT- " & lstrPRMBIN2XML)

                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): XML representation for BINARY PRM DATA")

                    'Load the XML representation into DOM
                    Call lobjPRMExceptionXMLDOM.LoadXml(lstrPRMBIN2XML)

                    'Check for any Exception.
                    If (lobjPRMExceptionXMLDOM.GetElementsByTagName("Exception").Count) > 0 Then

                        'Create the <ERROR> node for the exception
                        lobjExecpList = lobjPRMExceptionXMLDOM.GetElementsByTagName("Exception")
                        lobjErrComment = lobjExecpList.Item(0).SelectSingleNode("Comment")
                        lstrErrComment = ""
                        If Not (lobjErrComment Is Nothing) Then lstrErrComment = lobjErrComment.InnerText
                        lstrPRMBIN2XML = "<ERROR>" & "<ERROR_NBR>" & lobjExecpList.Item(0).SelectSingleNode("Number").InnerText & "</ERROR_NBR>" & "<ERROR_DESC><![CDATA[" & lobjExecpList.Item(0).SelectSingleNode("Description").InnerText & " " & lstrErrComment & "]]></ERROR_DESC>" & "<PRM_XML>" & lobjPRMExceptionXMLDOM.DocumentElement.OuterXml & "</PRM_XML>" & "</ERROR>"
                        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Exception from STServer :" & lobjExecpList.Item(0).SelectSingleNode("Description").InnerText)

                        'Else if no exception
                    Else
                        'Create <PRM_XML> node
                        lstrPRMBIN2XML = "<PRM_XML>" & lobjPRMExceptionXMLDOM.DocumentElement.OuterXml & "</PRM_XML>"
                    End If

                    'Create <PRM_FILE> node
                    Try
                        lobjPRMFileNameDOM.LoadXml("<PRM_FILE>" & lstrPRMBIN2XML & "</PRM_FILE>")
                        lobjElem = lobjPRMFileNameDOM.CreateElement("FILE_NAME")
                        lobjElem.InnerText = lstrFileName
                        lobjPRMFileNameDOM.DocumentElement.InsertBefore(lobjElem, lobjPRMFileNameDOM.DocumentElement.ChildNodes(0))
                        lobjElem = Nothing
                        lstrPRM2XML = lstrPRM2XML & lobjPRMFileNameDOM.OuterXml
                    Catch ex As Exception
                    End Try
                    lobjPRMFileNameDOM.RemoveAll() '  Changed By Sanjay lobjPRMFileNameDOM=nothing
                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Data added to the output XML.")
                Next

                'Create <PRM_FILE_LIST> node
                lstrPRM2XML = "<PRM_FILE_LIST>" & lstrPRM2XML & "</PRM_FILE_LIST>"
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML():")
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Exit ConvertPRMToXML() method")
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock End here")
                'Return the Final XML
                Return lstrPRM2XML.Replace(IIf(_ProcessID = "", "", _ProcessID & "_") & IIf(_ThreadID = "", "", _ThreadID & "_") & IIf(MyGuid.ToString() = "", "", MyGuid.ToString() & "_"), "")

            End SyncLock
        Catch ex As Exception
            llErrNbr = Err.Number
            lstrErrSrc = Err.Source
            lstrErrDesc = Err.Description
            ConvertPRMToXML = "<PRM_FILE_LIST>" & "<ERROR>" & "<ERROR_NBR>" & llErrNbr & "</ERROR_NBR>" & "<ERROR_DESC><![CDATA[" & lstrErrDesc & "]]></ERROR_DESC>" & "</ERROR>" & "</PRM_FILE_LIST>"
            STLogger.Error("BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): General Error : " & ConvertPRMToXML)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_ConvertPRMToXML(): Exit ConvertPRMToXML() method")
            'Return the Final XML with <ERROR> node specifying the application error
            Return ConvertPRMToXML
        Finally
            If Not (lobjXMLSchemaSpace Is Nothing) Then
                lobjXMLSchemaSpace = Nothing
            End If
            If Not (lobjExecpList Is Nothing) Then
                lobjExecpList = Nothing
            End If
            If Not (lobjFileDataList Is Nothing) Then
                lobjFileDataList = Nothing
            End If
            If Not (lobjFileNameList Is Nothing) Then
                lobjFileNameList = Nothing
            End If
            If Not (lobjPRMExceptionXMLDOM Is Nothing) Then
                lobjPRMExceptionXMLDOM = Nothing
            End If
            If Not (lobjPRMXMLDOM Is Nothing) Then
                lobjPRMXMLDOM = Nothing
            End If
            If Not (lobjSTApplication Is Nothing) Then
                lobjSTApplication = Nothing
            End If
            If Not (lobjPRMFileNameDOM Is Nothing) Then
                lobjPRMFileNameDOM = Nothing
            End If
        End Try
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
    <STAThreadAttribute()> _
    Public Function GeneratePRMFiles(ByVal astrPRMInfoXML As String) As String
        'Declare XML Dom variables
        Dim lobjPRMInfoXMLDOM As New Xml.XmlDocument
        Dim lobjXMLSchemaSpace As New Xml.Schema.XmlSchemaSet
        Dim lobjReturnPRMLstXMLDOM As New Xml.XmlDocument
        Dim lobjSTQueryXMLDOM As New Xml.XmlDocument
        Dim lobjSTResponseXMLDOM As New Xml.XmlDocument
        Dim lobjExeceptionlst As Xml.XmlNodeList = Nothing
        Dim lobjBinarylst As Xml.XmlNodeList = Nothing

        ' Added By Sanjay for Fees to edfs
        Dim lobjFeesNodeList As Xml.XmlNodeList = Nothing
        Dim licount As Integer

        'Other Declarations
        Dim lstrFileLoc As String
        Dim liPRMParamsCnt As Short
        Dim lstrPRMFilePath As String
        Dim lstrPRMTemplatePath As String
        Dim lstrPRMMode As String
        Dim lstrReturnXML As String
        Dim lstrPRMFileName As String
        Dim lvPRMFileData As Xml.XmlNode
        Dim lbGenPRM As Boolean
        Dim llErrNbr As Integer
        Dim lstrErrSrc As String
        Dim lstrErrDesc As String
        Dim lobjErrComment As Xml.XmlNode
        Dim lstrErrComment As String
        lbGenPRM = False

        'Declare Super Trump Variables
        Dim lobjSTApplication As STSERVER.STApplication
        Dim lstrSTServerReqXML As String
        Dim lstrTRANSACTIONSTARTDATE As String
        Try
            SetLog4Net()
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): In GeneratePRMFiles() method")
            'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Input Argument 1:" & astrPRMInfoXML)

            'Load Return XML
            Call lobjReturnPRMLstXMLDOM.LoadXml("<PRM_FILE_LIST></PRM_FILE_LIST>")

            'Get the PRMInfoXML.xsd Schema
            lstrFileLoc = GetConfigurationKey("SchemaFilePath")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Schema file read from registry")
            Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPRMInfoXMLSchemaName)

            'Assign Schema to the XML DOM object
            lobjPRMInfoXMLDOM.Schemas = lobjXMLSchemaSpace
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Validating Input XML")
            lstrPRMFilePath = GetConfigurationKey("PRMFilePath")

            SyncLock obj
                Dim _ProcessID As String = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
                Dim _ThreadID As String = System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString()
                Dim MyGuid As Guid = Guid.NewGuid()
                'Load the Input XML into the XML DOM object & Check if Input XML is valid
                Try
                    lobjPRMInfoXMLDOM.LoadXml(astrPRMInfoXML)
                Catch ex As Exception
                    llErrNbr = Err.Number
                    lstrErrDesc = Err.Description
                    'Add the <ERROR> node to the Return XML                
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): PRM file Generation error - " & lstrErrDesc)
                    AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement, "ERROR", "")
                    AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
                    AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", " of the XML. " & lstrErrDesc)
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Return value: " & lobjReturnPRMLstXMLDOM.OuterXml)
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Exit GeneratePRMFiles() method")

                    'Return the final PRM list
                    Return lobjReturnPRMLstXMLDOM.OuterXml
                End Try
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Input XML Valid")
                lstrPRMTemplatePath = GetConfigurationKey("PRMTemplatePath")

                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock Starts here")
                'Declare Super Trump Variables
                lobjSTApplication = New STSERVER.STApplication

                'For Each set of PRM Parameters in the Input XML
                For liPRMParamsCnt = 0 To lobjPRMInfoXMLDOM.DocumentElement.ChildNodes.Count - 1
                    Try
                        Dim ORGFileName As String = ""
                        lbGenPRM = True

                        'Build the Input XML for the XMLInOut() method
                        Call lobjSTQueryXMLDOM.LoadXml("<SuperTRUMP>" & "<Transaction id='TRANS_ID_GEN_PRM' query='true'/>" & "</SuperTRUMP>")
                        lstrPRMMode = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_META_DATA/MODE")
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "MODE", lstrPRMMode)

                        'Initialize                        
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "INITIALIZE", "")

                        'Read Template                        
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "READFILE", "")

                        AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "filename", lstrPRMTemplatePath & "\" & GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_META_DATA/TEMPLATE_NAME"))


                        'Transaction Amount                        
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "TRANSACTIONAMOUNT", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TRANSACTIONAMOUNT"))

                        'Transaction Date                        
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "TransactionStartDate", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TRANSACTIONSTARTDATE"))
                        lstrTRANSACTIONSTARTDATE = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TRANSACTIONSTARTDATE")

                        'Residual Amout for Lease
                        If UCase(lstrPRMMode) = "LESSOR" Then

                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "ASSETS", "")
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "ASSET", "")
                            AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0), "index", CStr(0))
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0), "ResidualKeptAsAPercent", "false")
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0), "RESIDUAL", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/RESIDUALAMOUNT"))
                        End If

                        'Term                        
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "NUMBEROFPAYMENTS", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/NUMBEROFPAYMENTS"))

                        'Periodicity                        
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "PERIODICITY", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/PERIODICITY"))

                        'Payment timing                        
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "PAYMENTTIMING", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/PAYMENTTIMING"))

                        'Structure                        
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "STRUCTURE", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/STRUCTURE"))

                        'Yield Data                        
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "TARGETDATA", "")
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "TYPEOFSTATISTIC", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TARGETDATA/TYPEOFSTATISTIC"))
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "STATISTICINDEX", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TARGETDATA/STATISTICINDEX"))
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "NEPA", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TARGETDATA/NEPA"))
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "TARGETVALUE", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TARGETDATA/TARGETVALUE"))
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "ADJUST", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TARGETDATA/ADJUST"))
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "ADJUSTMENTMETHOD", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TARGETDATA/ADJUSTMENTMETHOD"))

                        'Added by Nizar - For fees
                        '-------------------------------------------------------------------------------------------------
                        If lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectNodes("PRM_PARAMS/FEES").Count > 0 Then
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "FEES", "")
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "DELETE", "")
                            AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.FirstChild, "INDEX", "*")
                            lobjFeesNodeList = lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectNodes("PRM_PARAMS/FEES/FEE")
                            For licount = 0 To lobjFeesNodeList.Count - 1
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "FEE", "")
                                AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "INDEX", CStr(licount))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "Description", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectSingleNode("PRM_PARAMS/FEES").ChildNodes(licount), "DESCRIPTION"))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "KeptAsAPercent", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectSingleNode("PRM_PARAMS/FEES").ChildNodes(licount), "KEPTASAPERCENT"))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "Amount", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectSingleNode("PRM_PARAMS/FEES").ChildNodes(licount), "AMOUNT"))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "FeeDate", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectSingleNode("PRM_PARAMS/FEES").ChildNodes(licount), "FEEDATE"))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "IsAnExpense", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectSingleNode("PRM_PARAMS/FEES").ChildNodes(licount), "ISANEXPENSE"))
                                'AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "FederalDepreciation", "")
                                'AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild.LastChild, "Method", "Expensed")
                            Next
                        End If
                        '-----------------------------------------------------------------------------------------------------------

                        'Write PRM file                        
                        ORGFileName = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_META_DATA/FILE_NAME")
                        lstrPRMFileName = _ProcessID & "_" & _ThreadID & "_" & MyGuid.ToString & "_" & GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_META_DATA/FILE_NAME")
                        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Calling XMLInOut() method")
                        ''''Added By lalit June 14th, 2010 Starts  
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "CALCULATE", "")
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "ISTEMPLATE", "false")
                        'Transaction State will return Binary Data                        
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "TRANSACTIONSTATE", "")
                        AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "query", "true")
                        ''''Added By lalit June 14th, 2010 Ends

                        'Generate the PRM file.                        
                        lstrReturnXML = lobjSTApplication.XmlInOut(lobjSTQueryXMLDOM.OuterXml)
                        ''**''STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Output from  XMLINOUT- " & lstrReturnXML)
                        'Load the super Trump response XML
                        Call lobjSTResponseXMLDOM.LoadXml(lstrReturnXML)
                        lobjBinarylst = lobjSTResponseXMLDOM.GetElementsByTagName("Transaction")
                        lvPRMFileData = lobjBinarylst.Item(0).SelectSingleNode("TRANSACTIONSTATE")
                        lstrSTServerReqXML = "<SuperTRUMP>" & "<Transaction query=""true"">" & "<Initialize/>" & "<TRANSACTIONSTATE>" & lvPRMFileData.InnerText & "</TRANSACTIONSTATE><TRANSACTIONSTARTDATE>" & lstrTRANSACTIONSTARTDATE & "</TRANSACTIONSTARTDATE><Target /><Calculate /><TRANSACTIONSTATE query='true'></TRANSACTIONSTATE></Transaction>" & "</SuperTRUMP>"
                        lstrReturnXML = ""
                        lstrReturnXML = lobjSTApplication.XmlInOut(lstrSTServerReqXML)
                        ''**''STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Output from  XMLINOUT- " & lstrReturnXML)
                        Call lobjSTResponseXMLDOM.LoadXml(lstrReturnXML)

                        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): " & lstrPRMFileName & " file generated.")
                        AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement, "PRM_FILE", "")
                        AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild, "FILE_NAME", ORGFileName)

                        'Check for any Exception.
                        lobjExeceptionlst = lobjSTResponseXMLDOM.GetElementsByTagName("Exception")

                        If (lobjExeceptionlst.Count) > 0 Then

                            'Add the <ERROR> node to the Return XML for the PRM file                            
                            AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild, "ERROR", "")
                            AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_NBR", lobjExeceptionlst.Item(0).SelectSingleNode("Number").InnerText)
                            lobjErrComment = lobjExeceptionlst.Item(0).SelectSingleNode("Comment")
                            lstrErrComment = ""
                            If Not (lobjErrComment Is Nothing) Then lstrErrComment = lobjErrComment.InnerText
                            AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_DESC", lobjExeceptionlst.Item(0).SelectSingleNode("Description").InnerText & " " & lstrErrComment)
                            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Exception from STServer : " & lobjExeceptionlst.Item(0).SelectSingleNode("Description").InnerText)

                            'Else if no exception
                        Else

                            'Read the generated PRM File from disk 
                            ''Added by Lalit to read Transaction State for Binary data                          
                            lobjBinarylst = lobjSTResponseXMLDOM.GetElementsByTagName("Transaction")
                            lvPRMFileData = lobjBinarylst.Item(0).SelectSingleNode("TRANSACTIONSTATE")

                            'Add it to the Return XML                            
                            AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild, "FILE_DATA", lvPRMFileData.InnerText)
                        End If
                    Catch ex As Exception
                        llErrNbr = Err.Number
                        lstrErrSrc = Err.Source
                        lstrErrDesc = Err.Description
                        lbGenPRM = False

                        'Add the <ERROR> node to the Return XML for the PRM file                        
                        AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild, "ERROR", "")
                        AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_NBR", CStr(llErrNbr))
                        AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_DESC", lstrErrDesc)
                        lbGenPRM = True
                    End Try
                    lobjSTResponseXMLDOM.RemoveAll()
                    lobjExeceptionlst = Nothing
                    lobjSTQueryXMLDOM.RemoveAll()

                Next
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles():")
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Exit GeneratePRMFiles() method")
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock End here")
                'Return the final PRM list
                Return lobjReturnPRMLstXMLDOM.OuterXml

            End SyncLock
        Catch ex As Exception
            llErrNbr = Err.Number
            lstrErrSrc = Err.Source
            lstrErrDesc = Err.Description
            lobjReturnPRMLstXMLDOM.RemoveAll()

            'Build the Error XML
            Call lobjReturnPRMLstXMLDOM.LoadXml("<PRM_FILE_LIST><ERROR></ERROR></PRM_FILE_LIST>")
            AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): General error : " & lobjReturnPRMLstXMLDOM.OuterXml)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_GeneratePRMFiles(): Exit GeneratePRMFiles() method")

            'Return error XML
            Return lobjReturnPRMLstXMLDOM.OuterXml
        Finally
            If Not (lobjSTQueryXMLDOM Is Nothing) Then
                lobjSTQueryXMLDOM = Nothing
            End If
            If Not (lobjFeesNodeList Is Nothing) Then
                lobjFeesNodeList = Nothing
            End If
            If Not (lobjExeceptionlst Is Nothing) Then
                lobjExeceptionlst = Nothing
            End If
            If Not (lobjSTResponseXMLDOM Is Nothing) Then
                lobjSTResponseXMLDOM = Nothing
            End If
            If Not (lobjReturnPRMLstXMLDOM Is Nothing) Then
                lobjReturnPRMLstXMLDOM = Nothing
            End If
            If Not (lobjXMLSchemaSpace Is Nothing) Then
                lobjXMLSchemaSpace = Nothing
            End If
            If Not (lobjPRMInfoXMLDOM Is Nothing) Then
                lobjPRMInfoXMLDOM = Nothing
            End If
            If Not (lobjSTApplication Is Nothing) Then
                lobjSTApplication = Nothing
            End If
        End Try

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
    <STAThreadAttribute()> _
    Public Function GetAmortizationSchedule(ByVal astrPRMFileListXML As String) As String

        'Declare Super Trump Variables

        'Declare XML Dom variables
        Dim lobjXMLSchemaSpace As New Xml.Schema.XmlSchemaSet
        Dim lobjReturnAmortSchedLstXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileListXMLDOM As New Xml.XmlDocument
        Dim lobjPRMFileBinDataElem As Xml.XmlElement

        'Other Declarations
        Dim lstrFileLoc As String
        Dim lstrPRMFileListXML As String
        Dim lstrPRMFilePath As String
        Dim liPRMFileLstCnt As Short
        Dim lstrPRMFileName As String
        Dim ldLeaseFactor As Double
        Dim lbGetAmort As Boolean
        Dim llErrNbr As Integer
        Dim lstrErrSrc As String
        Dim lstrErrDesc As String
        Dim lstrXMLInOutInput As String
        Dim lstrReturnXMLInOut As String
        Dim lobjSTResponseXMLDOM As New Xml.XmlDocument
        Dim lobjPRMInfoXMLDOM As New Xml.XmlDocument
        Dim lobjNodeList1 As Object = Nothing
        Dim lobjNodeList2 As Object = Nothing
        Dim lobjNodeList3 As Object = Nothing
        Dim lobjRefrenceNode As Xml.XmlElement
        Dim lobjRefrenceNodePrincipal As Xml.XmlElement
        Dim lobjRefrenceNodeInterest As Xml.XmlElement
        Dim licount As Short
        Dim lstrPaymentStartDate As String
        lbGetAmort = False
        Dim lobjSTApplication As STSERVER.STApplication
        Dim lstrPRMMode As String
        Dim lstrPRMAmount As String
        Try
            SetLog4Net()
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): In GetAmortizationSchedule() method")
            'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Input Argument 1:" & astrPRMFileListXML)

            'Load Return XML
            Call lobjReturnAmortSchedLstXMLDOM.LoadXml("<AMORTIZATION_SCHEDULE_LIST></AMORTIZATION_SCHEDULE_LIST>")

            'Get the PRMFileLstXML.xsd Schema
            lstrFileLoc = GetConfigurationKey("SchemaFilePath")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Schema file read from registry")
            Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPRMFileLstXMLSchemaName)

            'Assign Schema to the XML DOM object
            lobjPRMFileListXMLDOM.Schemas = lobjXMLSchemaSpace
            lstrPRMFileListXML = Replace(astrPRMFileListXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Validating Input XML")

            SyncLock obj
                Dim _ProcessID As String = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
                Dim _ThreadID As String = System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString()
                Dim MyGuid As Guid = Guid.NewGuid()

                'Load the Input XML into the XML DOM object &   Check if Input XML is valid
                Try
                    lobjPRMFileListXMLDOM.LoadXml(lstrPRMFileListXML)
                Catch ex As Exception
                    llErrNbr = Err.Number
                    lstrErrDesc = Err.Description

                    'Add the <ERROR> node to the Return XML                
                    AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement, "ERROR", "")
                    AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
                    AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", "Error of the XML. " & lstrErrDesc)
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Return value: " & GetAmortizationSchedule)
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Exit GetAmortizationSchedule() method")

                    'Return the final PRM list
                    Return lobjReturnAmortSchedLstXMLDOM.OuterXml
                End Try

                lstrPRMFilePath = GetConfigurationKey("PRMFilePath") ' Placed Here
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Input XML Valid")

                'For each binary PRM file in the Input XML  

                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock Starts here")
                lobjSTApplication = New STSERVER.STApplication
                For liPRMFileLstCnt = 0 To lobjPRMFileListXMLDOM.DocumentElement.ChildNodes.Count - 1
                    Try
                        lbGetAmort = True

                        'Save the PRM Binary data to a physical file.
                        lstrPRMFileName = GetXMLElementValue(lobjPRMFileListXMLDOM.DocumentElement.ChildNodes(liPRMFileLstCnt), "FILE_NAME")
                        lobjPRMFileBinDataElem = lobjPRMFileListXMLDOM.DocumentElement.ChildNodes(liPRMFileLstCnt).SelectSingleNode("FILE_DATA")

                        AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement, "AMORTIZATION_SCHEDULE", "")
                        AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild, "PRM_FILE_NAME", lstrPRMFileName)
                        '''''Added By lalit June, 2010
                        lstrXMLInOutInput = "<SuperTRUMP>" & "<Transaction query='true'><TRANSACTIONSTATE>" & lobjPRMFileBinDataElem.InnerXml & "</TRANSACTIONSTATE><Calculate /></Transaction>" & "</SuperTRUMP>"

                        lstrReturnXMLInOut = lobjSTApplication.XmlInOut(lstrXMLInOutInput)
                        ''**''STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Output from  XMLINOUT- " & lstrReturnXMLInOut)
                        lobjPRMInfoXMLDOM.LoadXml(lstrReturnXMLInOut)
                        If (lobjPRMInfoXMLDOM.GetElementsByTagName("Exception").Count) > 0 Then

                            'Create the <ERROR> node for the exception
                            Dim lobjExecpList As XmlNodeList = lobjPRMInfoXMLDOM.GetElementsByTagName("Exception")
                            Dim lobjErrComment1 As XmlNode = lobjExecpList.Item(0).SelectSingleNode("Comment")
                            Dim lstrErrComment1 As String = ""
                            If Not (lobjErrComment1 Is Nothing) Then lstrErrComment1 = lobjErrComment1.InnerText
                            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Exception from STServer : ")
                        End If

                        ''lalit
                        lstrPRMMode = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(0), "Mode")
                        lstrPRMAmount = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(0), "TransactionAmount")
                        If lstrPRMAmount Is Nothing Then
                            lstrPRMAmount = "0.00"
                        End If

                        'Get the Amortization Data
                        'lobjSTTransaction.GetFreeStream
                        If UCase(lstrPRMMode) = "LENDER" Then
                            'For Loans , Payment Amount , Principal Amount and Interest Amount is fetched.
                            'Build the Input XML for loans ..
                            lstrXMLInOutInput = "<SuperTRUMP>" & "<Transaction>" & "<Results>" & "<Stream name=""Lending Loans Debt Service"" query=""true"" Label=""Payment amount"" />" & "<Stream name=""Lending Loans Principal"" query=""true"" Label=""Principal amount"" />" & "<Stream name=""Lending Loans Interest"" query=""true"" Label=""Interest amount"" />" & "</Results>" & "</Transaction>" & "</SuperTRUMP>"
                        Else
                            'For Lease.....
                            lstrXMLInOutInput = "<SuperTRUMP>" & "<Transaction>" & "<Results>" & "<Stream name=""Rent"" query=""true"" Label=""Payment amount"" />" & "</Results>" & "</Transaction>" & "</SuperTRUMP>"

                        End If

                        ' SyncLock lobjSTApplication
                        lstrReturnXMLInOut = lobjSTApplication.XmlInOut(lstrXMLInOutInput)
                        ''**''STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Output from  XMLINOUT- " & lstrReturnXMLInOut)
                        ' End SyncLock

                        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Amortization Data retrieved")

                        'Load the Super Trump Response XML
                        Call lobjSTResponseXMLDOM.LoadXml(lstrReturnXMLInOut)

                        If (lobjSTResponseXMLDOM.GetElementsByTagName("Exception").Count) > 0 Then
                            'Create the <ERROR> node for the exception
                            Dim lobjExecpList As XmlNodeList = lobjSTResponseXMLDOM.GetElementsByTagName("Exception")
                            Dim lobjErrComment As XmlNode = lobjExecpList.Item(0).SelectSingleNode("Comment")
                            Dim lstrErrComment As String = ""
                            If Not (lobjErrComment Is Nothing) Then lstrErrComment = lobjErrComment.InnerText
                            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Exception from STServer : ")
                        End If

                        lobjNodeList1 = lobjSTResponseXMLDOM.SelectNodes("//Stream[@Label='Payment amount']/Amounts/Amount")
                        lobjNodeList2 = lobjSTResponseXMLDOM.SelectNodes("//Stream[@Label='Principal amount']/Amounts/Amount")
                        lobjNodeList3 = lobjSTResponseXMLDOM.SelectNodes("//Stream[@Label='Interest amount']/Amounts/Amount")


                        AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild, "PAYMENT_LIST", "")
                        licount = 0
                        If Not lobjNodeList1 Is Nothing Then
                            For licount = 0 To lobjNodeList1.count - 1
                                lobjRefrenceNode = lobjNodeList1(licount)

                                'For Loans, the last row contains Amt = $0 for balloon type.
                                'which is not required. Hence ignore the last row.                                    
                                If UCase(lstrPRMMode) = "LENDER" And licount = lobjNodeList1.count - 1 And CDbl(lobjRefrenceNode.InnerText) = 0 Then
                                    '''''If lobjSTTransaction.Mode = STSERVER.STEnumMode.ST_Mode_Lender And licount = lobjNodeList1.count - 1 And CDbl(lobjRefrenceNode.InnerText) = 0 Then

                                    Exit For
                                End If
                                AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild.LastChild, "PAYMENT", "")

                                'Payment number                                    
                                AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild.LastChild.LastChild, "PAYMENT_NUMBER", CStr(licount + 1))

                                'Change To New Statement to remove vb6 lstrPaymentStartDate = VB6.Format(lobjRefrenceNode.getAttribute("date"), "mm/dd/yyyy")
                                lstrPaymentStartDate = String.Format("{0:d}", Convert.ToDateTime(lobjRefrenceNode.GetAttribute("date")))

                                'Payment Start Date                                    
                                AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild.LastChild.LastChild, "PAYMENT_START_DATE", lstrPaymentStartDate)

                                'Payment Amount                                    
                                AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild.LastChild.LastChild, "PAYMENT_AMOUNT", lobjRefrenceNode.InnerText)

                                'Lease Factor
                                If CDbl(lobjRefrenceNode.InnerText) > 0 Then
                                    ''''ldLeaseFactor = lobjRefrenceNode.InnerText / lobjSTTransaction.Amount
                                    ldLeaseFactor = lobjRefrenceNode.InnerText / CInt(lstrPRMAmount)
                                Else
                                    ldLeaseFactor = 0
                                End If
                                AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild.LastChild.LastChild, "LEASE_FACTOR", CStr(ldLeaseFactor))

                                'For loans add the Principal & Interest amounts as well
                                ''''If lobjSTTransaction.Mode = STSERVER.STEnumMode.ST_Mode_Lender Then
                                If UCase(lstrPRMMode) = "LENDER" Then
                                    If Not lobjNodeList2 Is Nothing Then
                                        lobjRefrenceNodePrincipal = lobjNodeList2(licount)
                                        'Principal Amount                                            
                                        AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild.LastChild.LastChild, "PRINCIPAL_AMOUNT", lobjRefrenceNodePrincipal.InnerText)
                                    End If
                                    If Not lobjNodeList3 Is Nothing Then
                                        lobjRefrenceNodeInterest = lobjNodeList3(licount)
                                        'Interest Amount                                            
                                        AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild.LastChild.LastChild, "INTEREST_AMOUNT", lobjRefrenceNodeInterest.InnerText)
                                    End If
                                End If
                                lobjRefrenceNode = Nothing
                                lobjRefrenceNodePrincipal = Nothing
                                lobjRefrenceNodeInterest = Nothing

                            Next
                        End If


                        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Data added to the output XML")
                    Catch ex As Exception
                        llErrNbr = Err.Number() 'Err.Number
                        lstrErrSrc = Err.Source() 'Err.Source
                        lstrErrDesc = Err.Description()
                        lbGetAmort = False

                        'Add the <ERROR> node to the Return XML for the PRM file
                        AddXMLElement(lobjReturnAmortSchedLstXMLDOM, _
                                        lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild, _
                                        "ERROR", _
                                        "")

                        AddXMLElement(lobjReturnAmortSchedLstXMLDOM, _
                                        lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild.LastChild, _
                                        "ERROR_NBR", _
                                        llErrNbr)

                        AddXMLElement(lobjReturnAmortSchedLstXMLDOM, _
                                        lobjReturnAmortSchedLstXMLDOM.DocumentElement.LastChild.LastChild, _
                                        "ERROR_DESC", _
                                        lstrErrDesc)

                        STLogger.Error("BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Error retrieving Amort Sched - " & lstrErrDesc)
                        lbGetAmort = True
                    End Try
                    lobjSTApplication = Nothing
                    lobjSTApplication = New STSERVER.STApplication()
                    lobjNodeList1 = Nothing
                    lobjNodeList2 = Nothing
                    lobjNodeList3 = Nothing
                Next
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock End here")
            End SyncLock
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule()")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Exit GetAmortizationSchedule() method")

            'Return the final PRM list
            Return lobjReturnAmortSchedLstXMLDOM.OuterXml
        Catch ex As Exception
            llErrNbr = Err.Number() 'Err.Number
            lstrErrSrc = Err.Source() 'Err.Source
            lstrErrDesc = Err.Description()
            lobjReturnAmortSchedLstXMLDOM.RemoveAll()

            'Build error XML
            Call lobjReturnAmortSchedLstXMLDOM.LoadXml("<AMORTIZATION_SCHEDULE_LIST><ERROR></ERROR></AMORTIZATION_SCHEDULE_LIST>")
            AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjReturnAmortSchedLstXMLDOM, lobjReturnAmortSchedLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): General Error : " & lobjReturnAmortSchedLstXMLDOM.OuterXml)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_GetAmortizationSchedule(): Exit GetAmortizationSchedule() method")
            'Return the final PRM list
            Return lobjReturnAmortSchedLstXMLDOM.OuterXml
        Finally
            If Not (lobjNodeList3 Is Nothing) Then
                lobjNodeList3 = Nothing
            End If
            If Not (lobjPRMFileListXMLDOM Is Nothing) Then
                lobjPRMFileListXMLDOM = Nothing
            End If
            If Not (lobjXMLSchemaSpace Is Nothing) Then
                lobjXMLSchemaSpace = Nothing
            End If
            If Not (lobjReturnAmortSchedLstXMLDOM Is Nothing) Then
                lobjReturnAmortSchedLstXMLDOM = Nothing
            End If
            If Not (lobjNodeList2 Is Nothing) Then
                lobjNodeList2 = Nothing
            End If
            If Not (lobjNodeList1 Is Nothing) Then
                lobjNodeList1 = Nothing
            End If
            If Not (lobjSTApplication Is Nothing) Then
                lobjSTApplication = Nothing
            End If
        End Try
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
    <STAThreadAttribute()> _
    Public Function GetPricingReports(ByVal astrPricingRepInfoXML As String) As String
        'Declare Super Trump Variables
        Dim lobjReportSTTrans As STSERVER.STTransaction
        Dim lobjReportSTResults As STSERVER.STResults = Nothing
        Dim lobjSTApp As STSERVER.STApplication

        'Declare XML Dom variables
        Dim lobjPRICING_REPORT_INFO_DOM As New Xml.XmlDocument
        Dim lobjPRMBIN2XML_DOM As New Xml.XmlDocument
        Dim lobjPRICING_REPORTS As Xml.XmlNodeList = Nothing
        Dim lobjPRICING_REPORT As Xml.XmlNode
        Dim lobjFILE_NAME As Xml.XmlNode = Nothing
        Dim lobjFILE_DATA As Xml.XmlNode = Nothing
        Dim lobjREPORT_TYPES As Xml.XmlNodeList = Nothing
        Dim lobjREPORT_TYPE As Xml.XmlNode = Nothing
        Dim lobjXMLSchemaSpace As New Xml.Schema.XmlSchemaSet
        Dim lobjDOM As New Xml.XmlDocument
        Dim lobjElem As Xml.XmlElement = Nothing

        'Other Declarations
        Dim lstrReport As String
        Dim lbReportProcessFlag As Boolean
        Dim lstrPRM2ReportXML As String
        Dim lstrPRICING_REPORT As String
        Dim lstrFilePath As String
        Dim lstrFileLoc As String
        Dim lstrReportTemplateLoc As String
        Dim lstrFileName As String
        Dim lstrPRMFileName As String
        Dim lstrPricingRepInfoXML As String
        Dim llErrNbr As Integer
        Dim lstrErrSrc As String
        Dim lstrErrDesc As String
        Dim LoadXmlErrFlag As Boolean
        Dim lobjPRMdata As String
        Dim lobjPRMresponse As String
        Try
            SetLog4Net()
            lstrPRM2ReportXML = "" ' Added By Sanjay to avoid Warning Used before inilization
            lstrPRICING_REPORT = ""  ' Added By Sanjay to avoid Warning Used before inilization
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): In GetPricingReports() method")
            'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): Input Argument 1:" & astrPricingRepInfoXML)

            'Get the PricingRepInfoXML.xsd Schema
            lstrFileLoc = GetConfigurationKey("SchemaFilePath")
            'ReadRegistry(gcFacilityConfigPath & gcFacilityID & "\" & gcSchemaFilePathKey)

            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): Schema file read from registry")
            Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPricingRepInfoXMLSchemaName)

            'Assign Schema to the XML DOM object
            lobjPRICING_REPORT_INFO_DOM.Schemas = lobjXMLSchemaSpace
            lstrPricingRepInfoXML = Replace(astrPricingRepInfoXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): Validating Input XML")
            SyncLock obj
                Dim _ProcessID As String = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
                Dim _ThreadID As String = System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString()
                Dim MyGuid As Guid = Guid.NewGuid()

                'Load the Input XML into the XML DOM object
                Try
                    lobjPRICING_REPORT_INFO_DOM.LoadXml(lstrPricingRepInfoXML)
                Catch ex As Exception
                    lstrPRICING_REPORT = "<PRICING_REPORT_LIST>" & "<ERROR>" & "<ERROR_NBR>" & Err.Number & "</ERROR_NBR>" & " of the XML." & Err.Description & "]]></ERROR_DESC>" & "</ERROR>" & "</PRICING_REPORT_LIST>"
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GetPricingReports(): Return value: " & lstrPRICING_REPORT)
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GetPricingReports(): Exit GetPricingReports() method")

                    'Return the final XML
                    Return lstrPRICING_REPORT
                End Try
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): Input XML Valid")

                'Get the Pricing Report List
                lobjPRICING_REPORTS = lobjPRICING_REPORT_INFO_DOM.GetElementsByTagName("PRICING_REPORT")
                lstrFileLoc = GetConfigurationKey("PRMFilePath")
                lstrReportTemplateLoc = GetConfigurationKey("ReportTemplatePath")
                'Traverse Each Pricing Report
                For Each lobjPRICING_REPORT In lobjPRICING_REPORTS
                    'Get the File Name & File data for Each Pricing Report
                    lobjFILE_NAME = lobjPRICING_REPORT.SelectSingleNode("PRM_FILE/FILE_NAME")
                    lobjFILE_DATA = lobjPRICING_REPORT.SelectSingleNode("PRM_FILE/FILE_DATA")

                    lobjPRMdata = "<SuperTRUMP><Transaction><TransactionState>"
                    lobjPRMdata = lobjPRMdata & lobjFILE_DATA.InnerXml
                    lobjPRMdata = lobjPRMdata & "</TransactionState></Transaction></SuperTRUMP>"
                    lstrFileName = lobjFILE_NAME.InnerText

                    'Get the Report List for the PRM File
                    lobjREPORT_TYPES = lobjPRICING_REPORT.SelectNodes("REPORT_LIST/REPORT_TYPE")

                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock Starts here")
                    lobjReportSTTrans = New STSERVER.STTransaction
                    lobjSTApp = New STSERVER.STApplication
                    lobjPRMresponse = lobjSTApp.XmlInOut(lobjPRMdata)
                    ''**''STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): Output from  XMLINOUT- " & lobjPRMresponse)

                    lobjReportSTTrans.Calculate()
                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): PRM file recalculated.")
                    lobjReportSTResults = lobjReportSTTrans.Results
                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): STResult object created.")

                    'Traverse Each Report List.
                    For Each lobjREPORT_TYPE In lobjREPORT_TYPES
                        Try
                            lbReportProcessFlag = True
                            LoadXmlErrFlag = False
                            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): report template name - " & lobjREPORT_TYPE.InnerText)

                            'Assign the Report File Name
                            lobjReportSTResults.ReportFileName = lstrReportTemplateLoc & "\" & lobjREPORT_TYPE.InnerText

                            'Retrieve the Report
                            lstrReport = lobjReportSTResults.PrintBuffer

                            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): report retrieved.")

                            'Create the <REPORT> Node.

                            Try
                                lobjDOM.LoadXml("<REPORT></REPORT>")
                            Catch ex As Exception
                                LoadXmlErrFlag = True
                            End Try

                            If LoadXmlErrFlag = False Then
                                lobjElem = lobjDOM.CreateElement("REPORT_TYPE")
                                lobjElem.InnerText = lobjREPORT_TYPE.InnerText
                                lobjDOM.DocumentElement.AppendChild(lobjElem)
                                lobjElem = Nothing
                                lobjElem = lobjDOM.CreateElement("TEXT_REPORT")
                                lobjElem.InnerText = lstrReport
                                lobjDOM.DocumentElement.AppendChild(lobjElem)
                                lobjElem = Nothing
                                lstrPRM2ReportXML = lstrPRM2ReportXML & lobjDOM.OuterXml
                            End If
                            lobjDOM.RemoveAll()
                            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): Report Data added to the output XML.")
                        Catch ex As Exception
                            llErrNbr = Err.Number
                            lstrErrSrc = Err.Source
                            lstrErrDesc = Err.Description
                            lbReportProcessFlag = False

                            'Create an <ERROR> node for that <REPORT> node and resume processing of the next report
                            lstrPRM2ReportXML = lstrPRM2ReportXML & "<REPORT>" & "<REPORT_TYPE>" & lobjREPORT_TYPE.InnerText & "</REPORT_TYPE>" & "<ERROR>" & "<ERROR_NBR>" & llErrNbr & "</ERROR_NBR>" & "<ERROR_DESC><![CDATA[" & lstrErrDesc & "]]></ERROR_DESC>" & "</ERROR>" & "</REPORT>"
                            STLogger.Error("BSSuperTrump.ISuperTrumpService_GetPricingReports(): Report Generation Error - " & lstrErrDesc)
                            lbReportProcessFlag = True

                            'Else if general application error
                        End Try

                    Next lobjREPORT_TYPE
                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock End here")


                    lobjReportSTResults = Nothing
                    lobjReportSTTrans = Nothing
                    lobjReportSTTrans = New STSERVER.STTransaction
                    LoadXmlErrFlag = False
                    Try
                        lobjDOM.LoadXml("<PRICING_REPORT><REPORT_LIST>" & lstrPRM2ReportXML & "</REPORT_LIST></PRICING_REPORT>")
                    Catch ex As Exception
                        LoadXmlErrFlag = True
                    End Try

                    If LoadXmlErrFlag = False Then
                        lobjElem = lobjDOM.CreateElement("PRM_FILE_NAME")
                        lobjElem.InnerText = lobjFILE_NAME.InnerText
                        lobjDOM.DocumentElement.InsertBefore(lobjElem, lobjDOM.DocumentElement.ChildNodes(0))
                        lobjElem = Nothing
                        lstrPRICING_REPORT = lstrPRICING_REPORT & lobjDOM.OuterXml
                    End If
                    lobjDOM.RemoveAll() ' Changed By Sanjay lobjDOM=nothing
                    lstrPRM2ReportXML = ""
                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): PRM added to the output XML.")
                Next lobjPRICING_REPORT

                'Create the <PRICING_REPORT_LIST> node
                lstrPRICING_REPORT = "<PRICING_REPORT_LIST>" & lstrPRICING_REPORT & "</PRICING_REPORT_LIST>"

                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports()")
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPricingReports(): Exit GetPricingReports() method")
            End SyncLock

            'Return the final XML
            Return lstrPRICING_REPORT

        Catch ex As Exception
            llErrNbr = Err.Number
            lstrErrSrc = Err.Source
            lstrErrDesc = Err.Description
            STLogger.Error("BSSuperTrump.ISuperTrumpService_GetPricingReports(): General Error : " & lstrPRICING_REPORT)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_GetPricingReports(): Exit GetPricingReports() method")
            'Return the Final XML with <ERROR> node specifying the application error
            Return "<PRICING_REPORT_LIST>" & "<ERROR>" & "<ERROR_NBR>" & llErrNbr & "</ERROR_NBR>" & "<ERROR_DESC><![CDATA[" & lstrErrDesc & "]]></ERROR_DESC>" & "</ERROR>" & "</PRICING_REPORT_LIST>"
        Finally
            If Not (lobjReportSTResults Is Nothing) Then
                lobjReportSTResults = Nothing
            End If
            If Not (lobjReportSTTrans Is Nothing) Then
                lobjReportSTTrans = Nothing
            End If
            If Not (lobjElem Is Nothing) Then
                lobjElem = Nothing
            End If
            If Not (lobjDOM Is Nothing) Then
                lobjDOM = Nothing
            End If
            If Not (lobjFILE_DATA Is Nothing) Then
                lobjFILE_DATA = Nothing
            End If
            If Not (lobjREPORT_TYPES Is Nothing) Then
                lobjREPORT_TYPES = Nothing
            End If
            If Not (lobjREPORT_TYPE Is Nothing) Then
                lobjREPORT_TYPE = Nothing
            End If
            If Not (lobjFILE_NAME Is Nothing) Then
                lobjFILE_NAME = Nothing
            End If
            If Not (lobjPRICING_REPORTS Is Nothing) Then
                lobjPRICING_REPORTS = Nothing
            End If
            If Not (lobjPRMBIN2XML_DOM Is Nothing) Then
                lobjPRMBIN2XML_DOM = Nothing
            End If
            If Not (lobjPRICING_REPORT_INFO_DOM Is Nothing) Then
                lobjPRICING_REPORT_INFO_DOM = Nothing
            End If
            If Not (lobjSTApp Is Nothing) Then
                lobjSTApp = Nothing
            End If
            If Not (lobjReportSTResults Is Nothing) Then
                lobjReportSTResults = Nothing
            End If
            If Not (lobjReportSTTrans Is Nothing) Then
                lobjReportSTTrans = Nothing
            End If
        End Try
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
    <STAThreadAttribute()> _
    Public Function GetPRMParams(ByVal astrPRMParamsInfoXML As String) As String

        'Declare Super Trump Variables


        'Declare XML Dom variables
        Dim lobjReturnPRMParamLstXMLDOM As New Xml.XmlDocument
        Dim lobjXMLSchemaSpace As New Xml.Schema.XmlSchemaSet
        Dim lobjPRMParamsInfoXMLDOM As New Xml.XmlDocument
        Dim lobjSTQueryXMLDOM As New Xml.XmlDocument
        Dim lobjSTResponseXMLDOM As New Xml.XmlDocument
        Dim lobjExeceptionlst As Xml.XmlNodeList = Nothing
        Dim lobjPRMFileBinDataElem As Xml.XmlElement = Nothing
        Dim lobjPRMParamElem As Xml.XmlElement = Nothing
        Dim lobjCloneNode As Xml.XmlNode = Nothing

        'Other Declarations
        Dim lstrPRMFilePath As String
        Dim liPRMParamsSpecCnt As Short
        Dim lstrPRMParamQuery As String
        Dim lstrPRMFileName As String
        Dim lstrReturnXML As String
        Dim lstrSTQuery As String
        Dim liPRMParamsCnt As Short
        Dim lstrFileLoc As String
        Dim lstrPRMParamsInfoXML As String
        Dim lbGetPRMParam As Boolean
        Dim liStart As Short
        Dim liEnd As Short
        Dim lstrResdElem As String
        Dim llErrNbr As Integer
        Dim lstrErrSrc As String
        Dim lstrErrDesc As String
        Dim lobjErrComment As Xml.XmlNode
        Dim lstrErrComment As String
        Dim liPRMFile As Short
        Dim liCntPRMs As Short
        lbGetPRMParam = False
        liPRMFile = 1
        Dim lobjSTApplication As STSERVER.STApplication
        Dim lstrXMLInOutInput As String
        Dim lstrReturnXMLInOut As String
        Dim lobjPRMInfoXMLDOM As New Xml.XmlDataDocument
        Dim lstrPRMMode As String

        Try
            SetLog4Net()
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): In GetPRMParams() method")
            'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Input Argument 1:" & astrPRMParamsInfoXML)

            'Load Return XML
            Call lobjReturnPRMParamLstXMLDOM.LoadXml("<PRM_PARAMS_LIST></PRM_PARAMS_LIST>")

            'Get the PRMParamsInfoXML.xsd Schema
            lstrFileLoc = GetConfigurationKey("SchemaFilePath")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Schema file read from registry")
            Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcPRMParamsInfoXMLSchemaName)

            'Assign Schema to the XML DOM object
            lobjPRMParamsInfoXMLDOM.Schemas = lobjXMLSchemaSpace
            lstrPRMParamsInfoXML = Replace(astrPRMParamsInfoXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Validating Input XML")
            'Load the Input XML into the XML DOM object
            SyncLock obj
                Dim _ProcessID As String = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
                Dim _ThreadID As String = System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString()
                Dim MyGuid As Guid = Guid.NewGuid()

                Try
                    lobjPRMParamsInfoXMLDOM.LoadXml(lstrPRMParamsInfoXML)
                Catch ex As Exception
                    llErrNbr = Err.Number
                    lstrErrDesc = Err.Description

                    'Add the <ERROR> node to the Return XML
                    AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement, "ERROR", "")
                    AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
                    AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", " of the XML. " & lstrErrDesc)


                    lstrPRMFilePath = GetConfigurationKey("PRMFilePath")
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GetPRMParams(): Return value: " & lobjReturnPRMParamLstXMLDOM.OuterXml)
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GetPRMParams(): Exit GetPRMParams() method")
                    'Return the final PRM list
                    Return lobjReturnPRMParamLstXMLDOM.OuterXml
                End Try
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Input XML Valid")

                'Get File Paths
                lstrPRMFilePath = GetConfigurationKey("PRMFilePath")

                'For Each set of PRM Parameters Specs in the Input XML

                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock Starts here")
                lobjSTApplication = New STSERVER.STApplication
                For liPRMParamsSpecCnt = 0 To lobjPRMParamsInfoXMLDOM.DocumentElement.ChildNodes.Count - 1
                    Try
                        lbGetPRMParam = True

                        'Get Parameter Query for Super Trump
                        lstrPRMParamQuery = lobjPRMParamsInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsSpecCnt).SelectSingleNode("PRM_PARAMS_SPECS").OuterXml
                        lstrPRMParamQuery = Replace(Replace(lstrPRMParamQuery, "<PRM_PARAMS_SPECS>", ""), "</PRM_PARAMS_SPECS>", "")
                        With lobjPRMParamsInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsSpecCnt)

                            'Get PRM File name
                            lstrPRMFileName = GetXMLElementValue(.SelectSingleNode("PRM_FILE"), "FILE_NAME")



                            lobjPRMFileBinDataElem = .SelectSingleNode("PRM_FILE").SelectSingleNode("FILE_DATA")

                            If Not (lobjPRMFileBinDataElem Is Nothing) Then


                                '''''Added By lalit June, 2010
                                lstrXMLInOutInput = "<SuperTRUMP>" & "<Transaction query='true'><TRANSACTIONSTATE>" & lobjPRMFileBinDataElem.InnerXml & "</TRANSACTIONSTATE><Calculate /></Transaction>" & "</SuperTRUMP>"
                                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Calling XMLInOut() method with input to get binary trnsaction state")
                                lstrReturnXMLInOut = lobjSTApplication.XmlInOut(lstrXMLInOutInput)
                                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Parmeter values returned as XML string (to get binary trnsaction state)")
                                ''**''STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams: Output from  XMLINOUT- " & lstrReturnXMLInOut)
                                lobjPRMInfoXMLDOM.LoadXml(lstrReturnXMLInOut)
                                If (lobjPRMInfoXMLDOM.GetElementsByTagName("Exception").Count) > 0 Then

                                    'Create the <ERROR> node for the exception
                                    Dim lobjExecpList As XmlNodeList = lobjPRMInfoXMLDOM.GetElementsByTagName("Exception")
                                    Dim lobjErrComment1 As XmlNode = lobjExecpList.Item(0).SelectSingleNode("Comment")
                                    Dim lstrErrComment1 As String = ""
                                    If Not (lobjErrComment1 Is Nothing) Then lstrErrComment1 = lobjErrComment1.InnerText
                                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Exception from STServer : ")
                                End If

                                ''Checking the mode for further calculations  - Added by Lalit June, 2010
                                lstrPRMMode = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(0), "Mode")

                                If UCase(lstrPRMMode) = "LENDER" Then
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
                                            lstrResdElem = Mid(lstrPRMParamQuery, liStart, liEnd - liStart)

                                            'Remove it
                                            lstrPRMParamQuery = Replace(lstrPRMParamQuery, lstrResdElem, "")
                                        End If
                                    End If
                                End If

                                'Build the Input XML for the XMLInOut() method
                                Call lobjSTQueryXMLDOM.LoadXml("<SuperTRUMP>" & "<Transaction id='TRANS_ID_GET_PRM_PARAMS'></Transaction>" & "</SuperTRUMP>")
                                lstrSTQuery = lobjSTQueryXMLDOM.OuterXml
                                lstrSTQuery = Replace(lstrSTQuery, "</Transaction>", "<TRANSACTIONSTATE>" & lobjPRMFileBinDataElem.InnerXml & "</TRANSACTIONSTATE>" & lstrPRMParamQuery & "</Transaction>")
                                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Calling XMLInOut() method")
                                lobjPRMFileBinDataElem = Nothing

                                'Get the PRM Parameters.                                
                                lstrReturnXML = lobjSTApplication.XmlInOut(lstrSTQuery)
                                ''**''STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams: Output from  XMLINOUT- " & lstrReturnXML)
                                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Parmeter values returned as XML string")
                                AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement, "PRM_PARAMS", "")
                                AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.LastChild, "PRM_FILE_NAME", lstrPRMFileName)

                                'Load the Super Trump response XML
                                Call lobjSTResponseXMLDOM.LoadXml(lstrReturnXML)
                                If UCase(Trim(gstrExceptionFlag)) = "TRUE" Then
                                    For liCntPRMs = 0 To UBound(gliPRMFilearr)
                                        If liPRMFile = gliPRMFilearr(liCntPRMs) Then

                                            'Add the <ERRORS> node to the Return XML                                        
                                            AddXMLElement(lobjSTResponseXMLDOM, lobjSTResponseXMLDOM.DocumentElement.LastChild, "Exceptions", gstrExcptionXMLDOMarr(liCntPRMs))
                                        End If
                                    Next
                                End If
                                'This replace is reqd as AddXMLElement adds values to the tags as .text.
                                'So the XMl tags in it get converted to text values , which need to be re-converted.
                                Call lobjSTResponseXMLDOM.LoadXml(Replace(Replace(lobjSTResponseXMLDOM.OuterXml, "&lt;", "<"), "&gt;", ">"))
                                'Check for any Exception.
                                lobjExeceptionlst = lobjSTResponseXMLDOM.GetElementsByTagName("Exception")
                                If (lobjExeceptionlst.Count) > 0 Then

                                    'Add the <ERROR> node to the Return XML for the PRM file                                
                                    AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.LastChild, "ERROR", "")
                                    AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_NBR", lobjExeceptionlst.Item(0).SelectSingleNode("Number").InnerText)
                                    lobjErrComment = lobjExeceptionlst.Item(0).SelectSingleNode("Comment")
                                    lstrErrComment = ""
                                    If Not (lobjErrComment Is Nothing) Then lstrErrComment = lobjErrComment.InnerText
                                    AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_DESC", lobjExeceptionlst.Item(0).SelectSingleNode("Description").InnerText & " " & lstrErrComment)
                                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Exception from STServer : ")

                                    'Else if no exception
                                Else
                                    'Add the response to the Return XML
                                    For liPRMParamsCnt = 0 To lobjSTResponseXMLDOM.DocumentElement.ChildNodes(0).ChildNodes.Count - 1
                                        lobjPRMParamElem = lobjSTResponseXMLDOM.DocumentElement.ChildNodes(0).ChildNodes(liPRMParamsCnt)
                                        If Not (lobjPRMParamElem Is Nothing) Then
                                            lobjCloneNode = lobjPRMParamElem.CloneNode(True)
                                            lobjReturnPRMParamLstXMLDOM.DocumentElement.LastChild.AppendChild(lobjReturnPRMParamLstXMLDOM.ImportNode(lobjCloneNode, True))
                                            lobjCloneNode = Nothing
                                        End If
                                        lobjPRMParamElem = Nothing
                                    Next
                                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Data added to output XML.")
                                End If
                                lobjSTResponseXMLDOM.RemoveAll()
                                lobjExeceptionlst = Nothing
                                lobjSTQueryXMLDOM.RemoveAll()
                                lobjPRMInfoXMLDOM.RemoveAll()
                            End If
                        End With
                        liPRMFile = liPRMFile + 1
                    Catch ex As Exception
                        llErrNbr = Err.Number
                        lstrErrSrc = Err.Source
                        lstrErrDesc = Err.Description
                        lbGetPRMParam = False

                        'Add the <ERROR> node to the Return XML for the PRM file
                        AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.LastChild, "ERROR", "")
                        AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_NBR", CStr(llErrNbr))
                        AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_DESC", lstrErrDesc)
                        STLogger.Error("BSSuperTrump.ISuperTrumpService_GetPRMParams(): PRM file parameter retrieval error - " & lstrErrDesc)

                        lbGetPRMParam = True
                    End Try
                Next
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock End here")


                lstrPRMFilePath = GetConfigurationKey("PRMFilePath")
            End SyncLock
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams()")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GetPRMParams(): Exit GetPRMParams() method")
            'Return the final PRM list
            Return lobjReturnPRMParamLstXMLDOM.OuterXml
        Catch ex As Exception
            llErrNbr = Err.Number
            lstrErrSrc = Err.Source
            lstrErrDesc = Err.Description
            lobjReturnPRMParamLstXMLDOM.RemoveAll()

            'Build Error XML
            Call lobjReturnPRMParamLstXMLDOM.LoadXml("<PRM_PARAMS_LIST><ERROR></ERROR></PRM_PARAMS_LIST>")
            AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjReturnPRMParamLstXMLDOM, lobjReturnPRMParamLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_GetPRMParams(): General error : " & lobjReturnPRMParamLstXMLDOM.OuterXml)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_GetPRMParams(): Exit GetPRMParams() method")
            'Return error XML
            Return lobjReturnPRMParamLstXMLDOM.OuterXml
        Finally
            If Not (lobjCloneNode Is Nothing) Then
                lobjCloneNode = Nothing
            End If
            If Not (lobjPRMFileBinDataElem Is Nothing) Then
                lobjPRMFileBinDataElem = Nothing
            End If
            If Not (lobjPRMParamElem Is Nothing) Then
                lobjPRMParamElem = Nothing
            End If
            If Not (lobjSTQueryXMLDOM Is Nothing) Then
                lobjSTQueryXMLDOM = Nothing
            End If
            If Not (lobjExeceptionlst Is Nothing) Then
                lobjExeceptionlst = Nothing
            End If
            If Not (lobjSTResponseXMLDOM Is Nothing) Then
                lobjSTResponseXMLDOM = Nothing
            End If
            If Not (lobjReturnPRMParamLstXMLDOM Is Nothing) Then
                lobjReturnPRMParamLstXMLDOM = Nothing
            End If
            If Not (lobjXMLSchemaSpace Is Nothing) Then
                lobjXMLSchemaSpace = Nothing
            End If
            If Not (lobjPRMParamsInfoXMLDOM Is Nothing) Then
                lobjPRMParamsInfoXMLDOM = Nothing
            End If
            If Not (lobjSTApplication Is Nothing) Then
                lobjSTApplication = Nothing
            End If
        End Try

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
    <STAThreadAttribute()> _
    Public Function ModifyPRMFiles(ByVal astrModifyPRMFilesXML As String) As String
        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): In ModifyPRMFiles() method")

        'Declare Super Trump Variables
        Dim lobjSTApplication As STSERVER.STApplication

        'Declare XML Dom variables
        Dim lobjResponseXMLDOM As New Xml.XmlDocument
        Dim lobjRequestXMLDOM As New Xml.XmlDocument
        Dim lobjXMLSchemaSpace As New Xml.Schema.XmlSchemaSet
        Dim lobjElem As Xml.XmlElement = Nothing
        Dim lobjPRMExceptionXMLDOM As New Xml.XmlDocument
        Dim lobjExeceptionlst As Xml.XmlNodeList = Nothing
        Dim lobjDocFragment As Xml.XmlDocumentFragment

        'Other Declarations
        Dim lstrModifyPRMFilesXML As String
        Dim lstrPRMFilePath As String
        Dim lstrTempPath As String
        Dim lstrModifiedPRMFilePath As String
        Dim lbProcessingPRMFile As Boolean
        Dim liCnt As Short
        Dim lstrModifiedParamsXML As String
        Dim lstrPRMFileName As String
        Dim lstrSTServerReqXML As String
        Dim lstrSTServerRespXML As String
        Dim liCnt2 As Short
        Dim lvPRMFileData As Object
        Dim lbPRMFileTagAdded As Boolean
        Dim llErrNbr As Integer
        Dim lstrErrSrc As String
        Dim lstrErrDesc As String
        Dim lobjErrComment As Xml.XmlNode
        Dim lstrErrComment As String
        Dim lobjPRMExceptionXMLDOM_Flag As Boolean
        Dim TempDoc As New Xml.XmlDocument
        Dim ORGFileName As String = ""
        Try
            SetLog4Net()
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): In ModifyPRMFiles()")
            'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Input Argument 1:" & astrModifyPRMFilesXML)

            'Initialize Response XML
            Call lobjResponseXMLDOM.LoadXml("<MODIFY_PRM_RESPONSE/>")
            lstrModifyPRMFilesXML = Replace(astrModifyPRMFilesXML, " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""", "")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Validating Input XML")

            SyncLock obj
                Dim _ProcessID As String = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
                Dim _ThreadID As String = System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString()
                Dim MyGuid As Guid = Guid.NewGuid()

                'Load the Input Request XML into the XML DOM object & Check if Request XML is not valid
                Try
                    lobjRequestXMLDOM.LoadXml(lstrModifyPRMFilesXML)
                Catch ex As Exception
                    llErrNbr = Err.Number
                    lstrErrDesc = Err.Description

                    'Add the <ERROR> node to the Response XML                
                    AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement, "ERROR", "")
                    AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "ERROR_NBR", CStr(llErrNbr))
                    AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "ERROR_DESC", "Input XML Invalid!!! of the XML. " & lstrErrDesc)
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): " & lobjResponseXMLDOM.OuterXml)

                    'Return the error to the client
                    ModifyPRMFiles = lobjResponseXMLDOM.OuterXml
                    Exit Function
                End Try
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Input XML Valid")

                'Get the temporary path where PRM files will be copied to disk
                lstrTempPath = GetConfigurationKey("PRMFilePath")
                'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Read Temporary file path " & lstrTempPath)

                'Set flag for start of PRM processing
                lbProcessingPRMFile = True
                lstrPRMFileName = "" 'Added By sanjay to avoid Used before inilization            
                'For each PRM file in Request XML

                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock Starts here")
                lobjSTApplication = New STSERVER.STApplication

                For liCnt = 0 To lobjRequestXMLDOM.DocumentElement.ChildNodes.Count - 1
                    Try
                        ORGFileName = ""
                        lbPRMFileTagAdded = False
                        lstrPRMFileName = ""

                        'Check if PRM file path is specified in Request XML
                        lobjElem = lobjRequestXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE/FILE_PATH")
                        If Not (lobjElem Is Nothing) Then

                            'Get the PRM file path from Request XML
                            lstrPRMFilePath = lobjElem.InnerText

                            'Otherwise if the PRM file path is not specified in Request XML
                        Else


                            ORGFileName = lobjRequestXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE/FILE_NAME").InnerText
                            lstrPRMFilePath = lstrTempPath & "\" & IIf(_ProcessID = "", "", _ProcessID & "_") & IIf(_ThreadID = "", "", _ThreadID & "_") & IIf(MyGuid.ToString() = "", "", MyGuid.ToString() & "_") & lobjRequestXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE/FILE_NAME").InnerText
                        End If
                        lobjElem = Nothing
                        'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Processing " & lstrPRMFilePath)

                        'Add PRM file name to Response XML
                        lstrPRMFileName = Mid(lstrPRMFilePath, InStrRev(lstrPRMFilePath, "\") + 1)
                        AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement, "PRM_FILE", "")
                        AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "FILE_NAME", ORGFileName)
                        lbPRMFileTagAdded = True
                        'Get the data to be modified from Request XML
                        lstrModifiedParamsXML = lobjRequestXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("MODIFY_PARAMS").OuterXml

                        'Strip off the outer <MODIFY_PARAMS> tag
                        lstrModifiedParamsXML = Replace(lstrModifiedParamsXML, "<MODIFY_PARAMS>", "")
                        lstrModifiedParamsXML = Trim(Replace(lstrModifiedParamsXML, "</MODIFY_PARAMS>", ""))
                        lstrModifiedParamsXML = Trim(Replace(lstrModifiedParamsXML, "<MODIFY_PARAMS/>", ""))
                        lstrModifiedParamsXML = Trim(Replace(lstrModifiedParamsXML, vbCrLf, ""))
                        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): got modified params ")

                        'Check if Modified PRM file has to be copied to a specified location
                        lobjElem = lobjRequestXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("OUTPUT_DETAILS/PRM_PATH")
                        If Not (lobjElem Is Nothing) Then

                            'Get the location from Request XML
                            lstrModifiedPRMFilePath = lobjElem.InnerText

                            'Otherwise if no location is specified
                        Else

                            'Use original PRM file path (PRM file will be overwritten)
                            lstrModifiedPRMFilePath = lstrPRMFilePath
                        End If
                        'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): got modified prm path " & lstrModifiedPRMFilePath)

                        'Build the Input XML for the XMLInOut() method                        
                        lstrSTServerReqXML = "<SuperTRUMP>" & "<Transaction query=""" & lstrPRMFileName & """>" & "<Initialize/>" & "<TransactionState>" & lobjRequestXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("PRM_FILE/FILE_DATA").InnerXml & "</TransactionState>" & lstrModifiedParamsXML & "</Transaction>" & "</SuperTRUMP>"
                        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Input XML for the XMLInOut() method:")
                        lstrSTServerRespXML = lobjSTApplication.XmlInOut(lstrSTServerReqXML)

                        ''**''STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles: Output from  XMLINOUT- " & lstrSTServerRespXML)
                        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Response XML for the XMLInOut() method")

                        lobjPRMExceptionXMLDOM_Flag = False
                        Try
                            lobjPRMExceptionXMLDOM.LoadXml(lstrSTServerRespXML)
                        Catch ex As Exception
                            lobjPRMExceptionXMLDOM_Flag = True
                        End Try

                        If (lobjPRMExceptionXMLDOM.GetElementsByTagName("Exception").Count) > 0 Then

                            'Create the <ERROR> node for the exception
                            Dim lobjExecpList As XmlNodeList = lobjPRMExceptionXMLDOM.GetElementsByTagName("Exception")
                            Dim lobjErrComment1 As XmlNode = lobjExecpList.Item(0).SelectSingleNode("Comment")
                            Dim lstrErrComment1 As String = ""
                            If Not (lobjErrComment1 Is Nothing) Then lstrErrComment1 = lobjErrComment1.InnerText
                            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Exception from STServer : ")
                        End If

                        'If cannot Load the XML representation into XML Object
                        If lobjPRMExceptionXMLDOM_Flag = True Then
                            llErrNbr = Err.Number
                            lstrErrDesc = Err.Description

                            'Add the <ERROR> node for the PRM file to the Response XML                        
                            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "ERROR", "")
                            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_NBR", CStr(llErrNbr))
                            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_DESC", "Error loading STServer response XML!!! " & lstrErrDesc)
                            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Error loading STServer response XML!!! ")

                            'Otherwise Check for exception returned from STServer
                        ElseIf (lobjPRMExceptionXMLDOM.GetElementsByTagName("Exception").Count) > 0 Then
                            'Add the <ERROR> node for the PRM file to the Response XML
                            lobjExeceptionlst = lobjPRMExceptionXMLDOM.GetElementsByTagName("Exception")
                            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "ERROR", "")
                            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_NBR", lobjExeceptionlst.Item(0).SelectSingleNode("Number").InnerText)
                            lobjErrComment = lobjExeceptionlst.Item(0).SelectSingleNode("Comment")
                            lstrErrComment = ""
                            If Not (lobjErrComment Is Nothing) Then lstrErrComment = lobjErrComment.InnerText
                            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_DESC", "Exception returned from STServer!!! " & lobjExeceptionlst.Item(0).SelectSingleNode("Description").InnerText & " " & lstrErrComment)
                            lobjDocFragment = TempDoc.CreateDocumentFragment
                            TempDoc.AppendChild(TempDoc.ImportNode(lobjResponseXMLDOM.CreateElement("PRM_XML"), True))
                            lobjDocFragment.AppendChild(TempDoc.CreateElement("PRM_XML"))
                            TempDoc.RemoveAll()
                            TempDoc.AppendChild(TempDoc.ImportNode(lobjPRMExceptionXMLDOM.SelectSingleNode("SuperTRUMP"), True))
                            lobjDocFragment.SelectSingleNode("PRM_XML").AppendChild(TempDoc.SelectSingleNode("SuperTRUMP"))
                            lobjResponseXMLDOM.SelectSingleNode("//MODIFY_PRM_RESPONSE/PRM_FILE/ERROR").AppendChild(lobjResponseXMLDOM.ImportNode(lobjDocFragment, True))
                            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Exception returned from STServer!!! ")

                            'Otherwise if no exception
                        Else
                            'For each output specification
                            For liCnt2 = 0 To lobjRequestXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("OUTPUT_DETAILS").ChildNodes.Count - 1
                                Select Case lobjRequestXMLDOM.DocumentElement.ChildNodes(liCnt).SelectSingleNode("OUTPUT_DETAILS").ChildNodes(liCnt2).Name

                                    'If XML equivalent of Modified PRM file has to be returned
                                    Case "PRM_XML"
                                        'Add XML equivalent to Response XML                                    
                                        AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "PRM_XML", "")
                                        lobjResponseXMLDOM.DocumentElement.LastChild.LastChild.AppendChild(lobjResponseXMLDOM.ImportNode(lobjPRMExceptionXMLDOM.DocumentElement, True))
                                        'If Modified PRM file has to be returned
                                    Case "PRM_FILE"
                                        'Read the modified PRM File from disk                                    
                                        lvPRMFileData = GetBinaryFileData(lstrModifiedPRMFilePath)
                                        'Add Modified PRM file to Response XML
                                        AddBinaryXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "FILE_DATA", lvPRMFileData)
                                        'If Modified PRM file has to be copied to a specified location
                                    Case "PRM_PATH"
                                        'Add Success status to Response XML                                    
                                        AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "STATUS", "Successfully copied to " & lstrModifiedPRMFilePath)
                                End Select
                            Next liCnt2
                        End If
                    Catch ex As Exception
                        llErrNbr = Err.Number
                        lstrErrSrc = Err.Source
                        lstrErrDesc = Err.Description
                        If lstrPRMFileName <> "" Then
                            lbProcessingPRMFile = False

                            'Add <PRM_FILE> & <FILE_NAME> tag if it has not been added
                            If Not (lbPRMFileTagAdded) Then
                                AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement, "PRM_FILE", "")
                                AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "FILE_NAME", ORGFileName)
                            End If

                            'Add the <ERROR> node for the PRM file to the Response XML                        
                            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "ERROR", "")
                            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_NBR", CStr(llErrNbr))
                            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_DESC", "Error Processing PRM file!!! " & lstrErrDesc)
                            STLogger.Error("BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Error Processing PRM file" & lstrPRMFileName & "!!! " & lstrErrDesc)
                            lbProcessingPRMFile = True
                        End If
                    End Try
                Next liCnt
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock End here")

                'Set flag for end of PRM processing
                lbProcessingPRMFile = False
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Output")
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Exit ModifyPRMFiles")
                Return lobjResponseXMLDOM.OuterXml.Replace(IIf(_ProcessID = "", "", _ProcessID & "_") & IIf(_ThreadID = "", "", _ThreadID & "_") & IIf(MyGuid.ToString() = "", "", MyGuid.ToString() & "_"), "")
            End SyncLock
        Catch ex As Exception
            llErrNbr = Err.Number
            lstrErrSrc = Err.Source
            lstrErrDesc = Err.Description

            'If any other error
            Call lobjResponseXMLDOM.LoadXml("<MODIFY_PRM_RESPONSE/>")

            'Add the <ERROR> node to the Response XML            
            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement, "ERROR", "")
            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjResponseXMLDOM, lobjResponseXMLDOM.DocumentElement.LastChild, "ERROR_DESC", "General Failure!!! " & lstrErrDesc)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): " & lobjResponseXMLDOM.OuterXml)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_ModifyPRMFiles(): Exit ModifyPRMFiles")
            Return lobjResponseXMLDOM.OuterXml
        Finally
            If Not (lobjExeceptionlst Is Nothing) Then
                lobjExeceptionlst = Nothing
            End If
            If Not (lobjDocFragment Is Nothing) Then
                lobjDocFragment = Nothing
            End If
            If Not (lobjElem Is Nothing) Then
                lobjElem = Nothing
            End If
            If Not (lobjSTApplication Is Nothing) Then

            End If
            If Not (lobjErrComment Is Nothing) Then
                lobjErrComment = Nothing
            End If
            If Not (lobjPRMExceptionXMLDOM Is Nothing) Then
                lobjPRMExceptionXMLDOM = Nothing
            End If
            If Not (lobjXMLSchemaSpace Is Nothing) Then
                lobjXMLSchemaSpace = Nothing
            End If
            If Not (lobjRequestXMLDOM Is Nothing) Then
                lobjRequestXMLDOM = Nothing
            End If
            If Not (lobjResponseXMLDOM Is Nothing) Then
                lobjResponseXMLDOM = Nothing
            End If
            If Not (lobjSTApplication Is Nothing) Then
                lobjSTApplication = Nothing
            End If
            If Not (TempDoc Is Nothing) Then
                TempDoc = Nothing
            End If
        End Try
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
    <STAThreadAttribute()> _
    Public Function GeneratePRMFilesForPmtStruct(ByVal astrPRMInfoXML As String) As String

        Return GeneratePRMFilesForPmtStructure(astrPRMInfoXML, eSolveMethod.ecSolveForPayments)
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
        Return GeneratePRMFilesForPmtStructure(astrPRMInfoXML, eSolveMethod.ecSolveForRate)
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
    <STAThreadAttribute()> _
    Private Function GeneratePRMFilesForPmtStructure(ByVal astrPRMInfoXML As String, ByVal astrMethodType As eSolveMethod) As String
        'Declare Super Trump Variables



        'Declare XML Dom variables
        Dim lobjPRMInfoXMLDOM As New Xml.XmlDocument
        Dim lobjXMLSchemaSpace As New Xml.Schema.XmlSchemaSet
        Dim lobjReturnPRMLstXMLDOM As New Xml.XmlDocument
        Dim lobjSTQueryXMLDOM As New Xml.XmlDocument
        Dim lobjSTResponseXMLDOM As New Xml.XmlDocument
        Dim lobjExeceptionlst As Xml.XmlNodeList = Nothing
        Dim lobjPaymentNodes As Xml.XmlNodeList

        'Other Declarations
        Dim lstrFileLoc As String
        Dim liPRMParamsCnt As Short
        Dim liPaymentsCnt As Short
        Dim lstrPRMFilePath As String
        Dim lstrPRMTemplatePath As String
        Dim lstrPRMMode As String
        Dim lstrReturnXML As String
        Dim lstrPRMFileName As String
        Dim lvPRMFileData As Object
        Dim lbGenPRM As Boolean
        Dim llErrNbr As Integer
        Dim lstrErrSrc As String
        Dim lstrErrDesc As String
        Dim lstrTransactionAmt As String
        Dim lstrLendingRate As String
        Dim lstrCommencementDt As String
        Dim lstrPeriodicity As String
        Dim licount As Short
        Dim lstrMoneyCostDate As String
        Dim lstrResidualAmt As String
        Dim lstrBalloonAmt As String
        Dim lstrGEBusiness As String
        Dim lstrGEProduct As String
        Dim liExceptionCount As Short
        Dim lobjExcptionXMLDOM As New Xml.XmlDocument
        Dim liPRMFile As Short
        Dim lobjFeesNodeList As Xml.XmlNodeList
        Dim lobjBinarylst As Xml.XmlNodeList
        Dim lobjSTApplication As STSERVER.STApplication
        Try
            SetLog4Net()
            lbGenPRM = False
            liPRMFile = 0
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): In GeneratePRMFilesForPmtStructure() method")
            'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Input Argument 1:" & astrPRMInfoXML)

            'Load Return XML
            Call lobjReturnPRMLstXMLDOM.LoadXml("<PRM_FILE_LIST></PRM_FILE_LIST>")

            'Get the GeneratePRMForPmtStructXML.xsd Schema
            lstrFileLoc = GetConfigurationKey("SchemaFilePath")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Schema file read from registry")
            Call lobjXMLSchemaSpace.Add("", lstrFileLoc & "\" & gcGenPRMForPmtStructSchemaName)

            'Assign Schema to the XML DOM object
            lobjPRMInfoXMLDOM.Schemas = lobjXMLSchemaSpace
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Validating Input XML")
            SyncLock obj
                Dim _ProcessID As String = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
                Dim _ThreadID As String = System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString()
                Dim MyGuid As Guid = Guid.NewGuid()

                'Load the Input XML into the XML DOM object & Check if Input XML is valid
                Try
                    lobjPRMInfoXMLDOM.LoadXml(astrPRMInfoXML)
                Catch ex As Exception
                    llErrNbr = Err.Number
                    lstrErrDesc = Err.Description

                    'Add the <ERROR> node to the Return XML            
                    AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement, "ERROR", "")
                    AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
                    AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", " of the XML. " & lstrErrDesc)

                    lstrPRMFilePath = GetConfigurationKey("PRMFilePath")
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Return value: " & GeneratePRMFilesForPmtStructure)
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Exit GeneratePRMFilesForPmtStructure() method")

                    'Return the final PRM list
                    Return lobjReturnPRMLstXMLDOM.OuterXml
                End Try
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Input XML Valid")

                'Get File Paths
                lstrPRMFilePath = GetConfigurationKey("PRMFilePath")
                lstrPRMTemplatePath = GetConfigurationKey("PRMTemplatePath")

                Try

                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock Starts here")
                    lobjSTApplication = New STSERVER.STApplication
                    For liPRMParamsCnt = 0 To lobjPRMInfoXMLDOM.DocumentElement.ChildNodes.Count - 1
                        lbGenPRM = True

                        'Build the Input XML for the XMLInOut() method
                        Call lobjSTQueryXMLDOM.LoadXml("<SuperTRUMP>" & "<Transaction id='TRANS_ID_GEN_PRM' query='true'/>" & "</SuperTRUMP>")

                        'Mode
                        lstrPRMMode = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_META_DATA/MODE")
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "MODE", lstrPRMMode)

                        'Initialize                
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "INITIALIZE", "")

                        'Read Template                
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "READFILE", "")
                        AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "filename", lstrPRMTemplatePath & "\" & GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_META_DATA/TEMPLATE_NAME"))

                        'Transaction Amount
                        lstrTransactionAmt = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TRANSACTIONAMOUNT")
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "TRANSACTIONAMOUNT", lstrTransactionAmt)

                        'Transaction Date                
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "TRANSACTIONSTARTDATE", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/TRANSACTIONSTARTDATE"))

                        'Structure                
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "STRUCTURE", "Other")

                        'Periodicity
                        lstrPeriodicity = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/PERIODICITY")
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "PERIODICITY", lstrPeriodicity)

                        'Payment timing                
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "PAYMENTTIMING", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/PAYMENTTIMING"))

                        'Commencement Date
                        lstrCommencementDt = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/COMMENCEMENTDATE")
                        If lstrCommencementDt <> "" Then
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "COMMENCEMENTDATE", lstrCommencementDt)
                        End If

                        'added by Neil to send the Expense break up of GE and Dealer Fees.
                        '---------------------------------------------------------------------
                        If lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectNodes("PRM_PARAMS/FEES").Count > 0 Then
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "FEES", "")
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "DELETE", "")
                            AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.FirstChild, "INDEX", "*")
                            lobjFeesNodeList = lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectNodes("PRM_PARAMS/FEES/FEE")
                            For licount = 0 To lobjFeesNodeList.Count - 1
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "FEE", "")
                                AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "INDEX", CStr(licount))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "KeptAsAPercent", "false")
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "IsAnExpense", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectSingleNode("PRM_PARAMS/FEES").ChildNodes(licount), "ISANEXPENSE"))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "Amount", GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectSingleNode("PRM_PARAMS/FEES").ChildNodes(licount), "AMOUNT"))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild, "FederalDepreciation", "")
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.LastChild.LastChild, "Method", "Expensed")
                            Next

                        End If

                        'for Lease
                        If UCase(lstrPRMMode) = "LESSOR" Then

                            'Residual Amout
                            lstrResidualAmt = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/RESIDUALAMOUNT")

                            If lstrResidualAmt <> "" Then
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "ASSETS", "")
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "ASSET", "")
                                AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0), "index", CStr(0))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0), "ResidualKeptAsAPercent", "false")
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0), "RESIDUAL", lstrResidualAmt)
                            End If

                            'For loan
                        ElseIf UCase(lstrPRMMode) = "LENDER" Then

                            'Lending Rate
                            lstrLendingRate = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/LENDINGRATE")

                            'Balloon Amount
                            lstrBalloonAmt = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/BALLOON")

                            If lstrBalloonAmt <> "" Then
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "BALLOON", lstrBalloonAmt)
                            End If

                            '===========Start of Payment Structure =======================================================                    
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "LENDINGLOANS", "")
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "LENDINGLOAN", "")
                            AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0), "INDEX", "0")
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0), "CASHFLOWSTEPS", "")
                            'Delete all old cashflow steps                    
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0), "DELETE", "")
                            AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "INDEX", "*")

                            'Add funding payment

                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0), "CASHFLOWSTEP", "")
                            AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "INDEX", "0")
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "TYPE", "Funding")
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "AMOUNT", CStr(-1 * Val(lstrTransactionAmt)))
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "AMOUNTLOCKED", "true")
                            If astrMethodType = eSolveMethod.ecSolveForPayments Then
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0), "BalancePayments", "")
                                ' lobjSTCashflow.BalancePayments()
                            Else
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0), "BalanceRate", "")
                                'lobjSTCashflow.BalanceRate()
                            End If

                            'Add cashflow step for each payment
                            lobjPaymentNodes = lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt).SelectNodes("PRM_PARAMS/PAYMENTS/PAYMENT")
                            For liPaymentsCnt = 0 To lobjPaymentNodes.Count - 1
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0), "CASHFLOWSTEP", "")
                                AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "INDEX", CStr(liPaymentsCnt + 1))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "NUMBEROFPAYMENTS", GetXMLElementValue(lobjPaymentNodes(liPaymentsCnt), "NUMBEROFPAYMENTS"))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "TYPE", "Payment")

                                'For stub payment
                                If UCase(GetXMLElementValue(lobjPaymentNodes(liPaymentsCnt), "PAYMENT_TYPE")) = "STUB" Then
                                    AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "PERIODICITY", "Stub")
                                    AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "ENDACCRUAL", lstrCommencementDt)

                                    'For regular payments we have to specify periodicity otherwise it will inherit from the stub payment
                                    'which might cause an issue.
                                Else
                                    AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "PERIODICITY", lstrPeriodicity)
                                End If
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "RATE", lstrLendingRate)
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "AMOUNT", GetXMLElementValue(lobjPaymentNodes(liPaymentsCnt), "AMOUNT"))
                                AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild.ChildNodes(0).ChildNodes(0).LastChild, "AMOUNTLOCKED", GetXMLElementValue(lobjPaymentNodes(liPaymentsCnt), "AMOUNTLOCKED"))

                            Next liPaymentsCnt
                            '===========End of Payment Structure =======================================================
                        End If

                        'Business                
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "GEDATA", "")
                        lstrGEBusiness = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/GEBUSINESS")
                        If lstrGEBusiness <> "" Then
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "GEBUSINESS", lstrGEBusiness)

                        End If


                        'Product
                        lstrGEProduct = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/GEPRODUCT")
                        If lstrGEProduct <> "" Then
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "GEPRODUCT", lstrGEProduct)
                        End If

                        'add the moneyCost date....
                        lstrMoneyCostDate = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_PARAMS/MONEYCOSTDATE")
                        If lstrMoneyCostDate <> "" Then
                            AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "MONEYCOSTDATE", lstrMoneyCostDate)
                        End If


                        lstrPRMFileName = GetXMLElementValue(lobjPRMInfoXMLDOM.DocumentElement.ChildNodes(liPRMParamsCnt), "PRM_META_DATA/FILE_NAME")
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "CALCULATE", "")
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "ISTEMPLATE", "false")
                        AddXMLElement(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0), "TRANSACTIONSTATE", "")
                        AddXMLElementAttribute(lobjSTQueryXMLDOM, lobjSTQueryXMLDOM.DocumentElement.ChildNodes(0).LastChild, "query", "true")


                        'Generate the PRM file.
                        lstrReturnXML = lobjSTApplication.XmlInOut(lobjSTQueryXMLDOM.OuterXml)
                        ''**''STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure: Output from  XMLINOUT- " & lstrReturnXML)

                        STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): XMLINOUT Called for Binary Data")
                        AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement, "PRM_FILE", "")

                        AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild, "FILE_NAME", lstrPRMFileName)
                        'Load the super Trump response XML

                        ReDim Preserve gliPRMFilearr(liPRMFile + 1)
                        gliPRMFilearr(liPRMFile) = 0
                        ReDim Preserve gstrExcptionXMLDOMarr(liPRMFile + 1)
                        gstrExcptionXMLDOMarr(liPRMFile) = ""
                        Call lobjSTResponseXMLDOM.LoadXml(lstrReturnXML)

                        'Check for any Exception.
                        lobjExeceptionlst = lobjSTResponseXMLDOM.GetElementsByTagName("Exception")

                        If (lobjExeceptionlst.Count) > 0 Then
                            For liExceptionCount = 0 To lobjExeceptionlst.Count - 1
                                llErrNbr = CInt(lobjExeceptionlst.Item(liExceptionCount).ChildNodes(0).InnerText)
                                lstrErrDesc = lobjExeceptionlst.Item(liExceptionCount).ChildNodes(2).InnerText
                                If InStr(UCase(lstrErrDesc), "ANIEQUITY") > 0 Then
                                    gstrExceptionFlag = "TRUE"
                                    gliPRMFilearr(liPRMFile) = liPRMFile + 1
                                    Call lobjExcptionXMLDOM.LoadXml(lobjExeceptionlst.Item(liExceptionCount).OuterXml)
                                    gstrExcptionXMLDOMarr(liPRMFile) = lobjExcptionXMLDOM.OuterXml
                                    Exit For ' no more errors are reqd after aniequity error is found.so exiting for.
                                End If
                            Next

                            'Since we haven't solve for payments there will be exception which we need to ignore.
                            'Just writing to log file for debugging purposes.
                            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Exception from STServer : ")
                        End If
                        'Add it to the Return XML    
                        lobjBinarylst = lobjSTResponseXMLDOM.GetElementsByTagName("Transaction")
                        lvPRMFileData = lobjBinarylst.Item(0).SelectSingleNode("TRANSACTIONSTATE")

                        'Add it to the Return XML                            
                        AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild, "FILE_DATA", lvPRMFileData.InnerText)

                        '''AddBinaryXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild, "FILE_DATA", lvPRMFileData)                        
                        lobjSTResponseXMLDOM.RemoveAll()
                        lobjExeceptionlst = Nothing
                        lobjSTQueryXMLDOM.RemoveAll()
                        liPRMFile = liPRMFile + 1
                    Next
                    STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock End here")
                Catch ex As Exception
                    llErrNbr = Err.Number
                    lstrErrSrc = Err.Source
                    lstrErrDesc = Err.Description
                    lbGenPRM = False

                    'Add the <ERROR> node to the Return XML for the PRM file            
                    AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild, "ERROR", "")
                    AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_NBR", CStr(llErrNbr))
                    AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.LastChild.LastChild, "ERROR_DESC", lstrErrDesc)
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): PRM file Generation error - " & lstrErrDesc)
                    lbGenPRM = True
                Finally
                End Try
                'For Each set of PRM Parameters in the Input XML            
                lstrPRMFilePath = GetConfigurationKey("PRMFilePath")
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Return value: ")
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Exit GeneratePRMFilesForPmtStructure() method")
            End SyncLock
            'Return the final PRM list
            Return lobjReturnPRMLstXMLDOM.OuterXml
        Catch ex As Exception
            llErrNbr = Err.Number
            lstrErrSrc = Err.Source
            lstrErrDesc = Err.Description
            lobjReturnPRMLstXMLDOM.RemoveAll()

            'Build the Error XML
            Call lobjReturnPRMLstXMLDOM.LoadXml("<PRM_FILE_LIST><ERROR></ERROR></PRM_FILE_LIST>")
            AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_NBR", CStr(llErrNbr))
            AddXMLElement(lobjReturnPRMLstXMLDOM, lobjReturnPRMLstXMLDOM.DocumentElement.ChildNodes(0), "ERROR_DESC", lstrErrDesc)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): General error : " & lobjReturnPRMLstXMLDOM.OuterXml)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_GeneratePRMFilesForPmtStructure(): Exit GeneratePRMFilesForPmtStructure() method")

            'Return error XML
            Return lobjReturnPRMLstXMLDOM.OuterXml
        Finally
            If Not (lobjSTApplication Is Nothing) Then
                lobjSTApplication = Nothing
            End If
            If Not (lobjPRMInfoXMLDOM Is Nothing) Then
                lobjPRMInfoXMLDOM = Nothing
            End If
            If Not (lobjXMLSchemaSpace Is Nothing) Then
                lobjXMLSchemaSpace = Nothing
            End If
            If Not (lobjReturnPRMLstXMLDOM Is Nothing) Then
                lobjReturnPRMLstXMLDOM = Nothing
            End If
            If Not (lobjExcptionXMLDOM Is Nothing) Then
                lobjExcptionXMLDOM = Nothing
            End If
            If Not (lobjSTResponseXMLDOM Is Nothing) Then
                lobjSTResponseXMLDOM = Nothing
            End If
            If Not (lobjExeceptionlst Is Nothing) Then
                lobjExeceptionlst = Nothing
            End If
            If Not (lobjSTQueryXMLDOM Is Nothing) Then
                lobjSTQueryXMLDOM = Nothing
            End If
        End Try
    End Function

    '================================================================
    'METHOD  : Test
    'PURPOSE : Returns a string that this component is able to invoke
    '          STServer (Ivory's SuperTrump Server component).
    'PARMS   : NONE
    'RETURN  : String
    '================================================================  
    <STAThreadAttribute()> _
    Public Function Test() As String
        Dim lobjSTApp As STSERVER.STApplication
        Try
            SetLog4Net()
            SyncLock obj
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock Starts here")
                lobjSTApp = New STSERVER.STApplication
                Return "Test to invoke Ivory's SuperTrump Server component - " & lobjSTApp.Version & " (" & lobjSTApp.BuildInfo & ") successful."
                If Not (lobjSTApp Is Nothing) Then
                    lobjSTApp = Nothing
                End If
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock End here")
            End SyncLock
        Catch ex As Exception
            Return "Error occured while invoking STServer - " & Err.Description
        Finally
            If Not (lobjSTApp Is Nothing) Then
                lobjSTApp = Nothing
            End If
        End Try
        Exit Function
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
    <STAThreadAttribute()> _
    Public Function RunAdHocXMLInOutQuery(ByVal astrXMLInOutQuery As String) As String

        'Super Trump Server Objects Declarations
        Dim lobjFileNameList As Xml.XmlNodeList = Nothing
        Dim lobjSTApplication As STSERVER.STApplication = Nothing
        Dim lobjXMLDoc As Xml.XmlDocument = Nothing
        Dim lobjXMLResultDoc As Xml.XmlDocument = Nothing
        Dim lobjXMLNode As Xml.XmlNode = Nothing
        Dim llErrNbr As Integer
        Dim lstrErrSrc As String
        Dim lstrErrDesc As String
        Dim lintCtrLoop As Integer
        Dim lstrPrmFileData As String
        Try
            SetLog4Net()
            lobjXMLDoc = New Xml.XmlDocument

            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): In RunAdHocXMLInOutQuery() method")
            'check for parser errors
            Try
                lobjXMLDoc.LoadXml(astrXMLInOutQuery)
            Catch ex As Exception
                STLogger.Error("BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): " & Err.Number & Err.Description)
            End Try
            SyncLock obj
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock Starts here")
                lobjSTApplication = New STSERVER.STApplication

                lobjFileNameList = lobjXMLDoc.GetElementsByTagName("AD_HOC_QUERY")

                'cADHOC_QUERY_RESULT_XML = ""
                Dim Result As String = ""
                lobjXMLResultDoc = New Xml.XmlDocument
                Result = "<PRM_INFO>"

                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): Loop for traverse all prm files Start")
                'Loop for traverse all prm files
                For lintCtrLoop = 0 To (lobjFileNameList.Count - 1)
                    Result = Result & "<PRM_FILE><AD_HOC_QUERY><SuperTRUMP>"
                    lstrPrmFileData = lobjFileNameList.Item(lintCtrLoop).InnerXml
                    'SyncLock lobjSTApplication
                    lobjXMLDoc.LoadXml(lobjSTApplication.XmlInOut(lstrPrmFileData))

                    If InStr(lobjXMLDoc.ChildNodes(0).OuterXml, "?xml") > 0 Then
                        lobjXMLNode = lobjXMLDoc.ChildNodes(1)
                    Else
                        lobjXMLNode = lobjXMLDoc.ChildNodes(0)
                    End If
                    Result = Result & lobjXMLNode.InnerXml & "</SuperTRUMP></AD_HOC_QUERY></PRM_FILE>"
                Next

                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): Loop for traverse all prm files End")

                Result = Result & "</PRM_INFO>"

                Try
                    lobjXMLResultDoc.LoadXml(Result)
                Catch ex As Exception
                    STLogger.Error("BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): " & Err.Number & Err.Description)
                End Try
                'STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): XML OutPut Of Function:-" & lobjXMLResultDoc.OuterXml)
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): End  RunAdHocXMLInOutQuery() method")
                STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock End here")
                Return lobjXMLResultDoc.OuterXml
            End SyncLock
        Catch ex As Exception
            llErrNbr = Err.Number
            lstrErrSrc = Err.Source
            lstrErrDesc = Err.Description

            'Return the Final XML with <ERROR> node specifying the application error
            'STLogger.Error("BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): General Error : " & lobjXMLResultDoc.OuterXml)
            STLogger.Error("BSSuperTrump.ISuperTrumpService_RunAdHocXMLInOutQuery(): Exit RunAdHocXMLInOutQuery() method")
            Return "<PRM_INFO><PRM_FILE><AD_HOC_QUERY>" & "<ERROR>" & "<ERROR_NBR>" & llErrNbr & "</ERROR_NBR>" & "<ERROR_DESC><![CDATA[" & lstrErrDesc & "]]></ERROR_DESC>" & "</ERROR>" & "</AD_HOC_QUERY></PRM_FILE></PRM_INFO>"
        Finally
            If Not (lobjSTApplication Is Nothing) Then
                lobjSTApplication = Nothing
            End If
            If Not (lobjXMLDoc Is Nothing) Then
                lobjXMLDoc = Nothing
            End If
            If Not (lobjXMLResultDoc Is Nothing) Then
                lobjXMLResultDoc = Nothing
            End If
            If Not (lobjXMLNode Is Nothing) Then
                lobjXMLNode = Nothing
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
    <STAThreadAttribute()> _
    Public Function Ping() As String
        Try
            SetLog4Net()
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & " BSSuperTrump.ISuperTrumpService_Ping(): In Ping() method")
            STLogger.Debug("ProcessId:" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString() & " ThreadId:" & System.AppDomain.CurrentDomain.GetCurrentThreadId.ToString() & "  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Synclock End here")
            Return "Ping request to BSCEFSuperTrump.ISuperTrumpService returned at " & String.Format("{0:G}", System.DateTime.Now()) & " server time."
        Catch ex As Exception
            Return "Error occured while invoking Ping - " & Err.Description
        Finally
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
    Private Sub SetLog4Net()
        Try
            If log4net.LogManager.GetRepository.Configured = False Then
                'log4net.Config.XmlConfigurator.ConfigureAndWatch(New System.IO.FileInfo(System.Configuration.ConfigurationManager.AppSettings("SuperTRUMPLogForNetConfigPath").ToString()))
                log4net.Config.XmlConfigurator.ConfigureAndWatch(New System.IO.FileInfo("E:\internalsites\SuperTrump\Components_NET\bin\log4net_DLL.config"))
            End If
            STLogger = log4net.LogManager.GetLogger("SuperTRUMP")
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Shared Function for Sigleton"
    Private Shared _instanceISuperTrump As ISuperTrumpService
    Public Shared Function Instance() As ISuperTrumpService
        If _instanceISuperTrump Is Nothing Then
            _instanceISuperTrump = New ISuperTrumpService
        End If
        Return _instanceISuperTrump
    End Function
#End Region

End Class
