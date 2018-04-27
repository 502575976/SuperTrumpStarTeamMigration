Option Strict Off
Option Explicit On
Imports System
Imports System.IO
Imports System.text
Imports System.Data
Imports System.Xml
Imports Microsoft.Win32
Imports System.Configuration


Module modBSCEFSuperTrump
    '================================================================
    'MODULE  : modBSCEFSuperTrump
    'PURPOSE : Contains library functions, global consts and vars.
    '================================================================   
    '=Other Constants============================================================================
    Public Const gcPRMFileLstXMLSchemaName As String = "PRMFileListXML.xsd"
    Public Const gcPricingRepInfoXMLSchemaName As String = "PricingRepInfoXML.xsd"
    Public Const gcPRMInfoXMLSchemaName As String = "PRMInfoXML.xsd"
    Public Const gcPRMParamsInfoXMLSchemaName As String = "PRMParamsInfoXML.xsd"
    Public Const gcPricingReqInfoXMLSchemaName As String = "PricingRequestInfoXML.xsd"
    Public Const gcMQMsgInfoXMLSchemaName As String = "MQMessageInfo.xsd"
    Public Const gcModifyPRMFilesSchemaName As String = "ModifyPRMFilesXML.xsd"
    Public Const gcGenPRMForPmtStructSchemaName As String = "GeneratePRMForPmtStructXML.xsd"
    '============================================================================================

    '=== Error Constants ========================================================================
    Public Const gcHIGHEST_ERROR As Decimal = vbObjectError + 256
    Public Const gcINVALID_PRM_FILE As Decimal = gcHIGHEST_ERROR + 1000
    '============================================================================================

        'Variables defined to  force the aniequity error in the output : 5th march 2007
    Public gstrExceptionFlag As String
    Public gliPRMFilearr() As Short
    Public gstrExcptionXMLDOMarr() As String


    '============================================================================================

    '================================================================
    'METHOD  : DeletePRMBinaryFile
    'PURPOSE : Deletes .PRM file(s)
    'PARMS   :
    '          alobjFilenameList [XMLNodeList] = List of file
    '          names.
    '          E.g.
    '          <FILE_NAME>file1.prm</FILE_NAME>
    '          <FILE_NAME>file2.prm</FILE_NAME>
    '          ...
    '          astrPath [String] = Complete Path to the .PRM File
    'RETURN  : None
    '================================================================
    Public Sub DeletePRMBinaryFile(ByVal alobjFilenameList As Xml.XmlNodeList, ByVal astrPath As String, ByVal _ProcessID As String, ByVal _ThreadID As String, ByVal Guid As String)
        Dim lobjFileNode As Xml.XmlNode = Nothing
        Dim lstrFileName As String
        Try
            'For each file in the file list
            For Each lobjFileNode In alobjFilenameList
                lstrFileName = lobjFileNode.InnerText
                If UCase(Right(lstrFileName, 4)) <> ".PRM" Then lstrFileName = lstrFileName & ".PRM"
                'Check if file exists
                lstrFileName = IIf(_ProcessID = "", "", _ProcessID & "_") & IIf(_ThreadID = "", "", _ThreadID & "_") & IIf(Guid = "", "", Guid & "_") & lstrFileName
                If System.IO.File.Exists(astrPath & "\" & lstrFileName) Then
                    'Delete the PRM file
                    Call System.IO.File.Delete(astrPath & "\" & lstrFileName)
                End If
            Next lobjFileNode
        Catch ex As Exception
            Err.Source = "modBSCEFSuperTrump:DeletePRMBinaryFile/" & Err.Source
            Err.Description = "modBSCEFSuperTrump:DeletePRMBinaryFile/" & Err.Description
            Throw ex
        Finally
            If Not (lobjFileNode Is Nothing) Then
                lobjFileNode = Nothing
            End If
        End Try
    End Sub

    '================================================================
    'METHOD  : GetSuperTrumpQuery
    'PURPOSE : Generates the SuperTrump XML query form the XMLInOut()
    '          method.
    'PARMS   :
    '          astrTransID [String] = Transaction Id for the
    '          current Query
    '          astrFileName [String] = Path to the .PRM file used
    '          in the current query
    'RETURN  : String = Super Trump Query
    '================================================================
    Public Function GetSuperTrumpQuery(ByVal astrTransID As String, ByVal astrFileName As String, Optional ByVal astrFileBinaryData As String = "") As String
        Dim lobjSTDOM As New Xml.XmlDocument
        Dim lobjElem As Xml.XmlElement
        Dim lobjAttr As Xml.XmlAttribute
        Dim lstrXML As String
        Try
            lstrXML = "<SuperTRUMP></SuperTRUMP>"
            Try
                lobjSTDOM.LoadXml(lstrXML)
            Catch ex As Exception
                GetSuperTrumpQuery = ""
                Err.Raise(Err.Number, "modBSCEFSuperTrump:GetSuperTrumpQuery", Err.Description)
            End Try
            lobjElem = lobjSTDOM.CreateElement("Transaction")
            lobjAttr = lobjSTDOM.CreateAttribute("id")
            lobjAttr.InnerText = astrTransID
            lobjElem.Attributes.SetNamedItem(lobjAttr)
            lobjAttr = Nothing
            lobjAttr = lobjSTDOM.CreateAttribute("query")
            lobjAttr.InnerText = "true"
            lobjElem.Attributes.SetNamedItem(lobjAttr)
            lobjAttr = Nothing
            lobjSTDOM.DocumentElement.AppendChild(lobjElem)
            lobjElem = Nothing
            'lobjElem = lobjSTDOM.CreateElement("ReadFile")
            lobjElem = lobjSTDOM.CreateElement("TransactionState")
            lobjElem.InnerText = astrFileBinaryData
            'lobjAttr = lobjSTDOM.CreateAttribute("filename")
            'lobjAttr.InnerText = astrFileName
            'lobjElem.Attributes.SetNamedItem(lobjAttr)            
            'lobjAttr = Nothing
            lobjSTDOM.DocumentElement.ChildNodes(0).AppendChild(lobjElem)
            lobjElem = Nothing
            Return lobjSTDOM.OuterXml
        Catch ex As Exception
            Err.Source = "modBSCEFSuperTrump:GetSuperTrumpQuery/" & Err.Source
            Err.Description = "modBSCEFSuperTrump:GetSuperTrumpQuery/" & Err.Description
            Throw ex
        Finally
            If Not (lobjSTDOM Is Nothing) Then
                lobjSTDOM = Nothing
            End If
        End Try
    End Function

    '================================================================
    'METHOD  : SavePRMBinaryFile
    'PURPOSE : Saves the .PRM file using the given Path & File Name
    'PARMS   :
    '          astrFileName [String] = Name of the .PRM file to
    '          be saved
    '          astrFilePath [String] = Path where .PRM file to be
    '          saved
    '          aarrFileData [Byte] = Data to be saved
    'RETURN  : String = Complete Path to the saved .PRM file
    '================================================================
    Public Function SavePRMBinaryFile(ByVal astrFileName As String, ByVal astrFilePath As String, ByRef aarrFileData() As Byte, ByVal _ProcessID As String, ByVal _ThreadID As String, ByVal Guid As String) As String
        Dim liFileDes As Short
        Try
            liFileDes = FreeFile()
            astrFileName = astrFilePath & "\" & IIf(_ProcessID = "", "", _ProcessID & "_") & IIf(_ThreadID = "", "", _ThreadID & "_") & IIf(Guid = "", "", Guid & "_") & UCase(astrFileName)
            FileOpen(liFileDes, astrFileName, OpenMode.Binary, OpenAccess.Write)
            FilePut(liFileDes, aarrFileData)
            FileClose(liFileDes)
            Return astrFileName
        Catch ex As Exception
            Err.Source = "modBSCEFSuperTrump:SavePRMBinaryFile/" & Err.Source
            Err.Description = "modBSCEFSuperTrump:SavePRMBinaryFile/" & Err.Description
            Throw ex
        End Try
    End Function

    'Created By Sanjay Srivastava [03-04-08]
    'METHOD  : WriteToTextDebugFile with .NetLog
    'PURPOSE : To help in debugging errors in the compiled component.
    'PARMS   :
    '          astrFileName [String] = Debug file name with complete
    '          path.
    '          astrData [String] = Data to be written to the debug
    '          file.
    'RETURN  : Boolean = True if data is written successfully & false
    '          otherwise.
    '================================================================
    Public Sub WriteToTextDebugFile(ByVal astrFileName As String, ByVal astrMessage As String, Optional ByVal astrClassName As String = "", Optional ByVal astrFunctionName As String = "")
        Dim lstrErrorMessage As New StringBuilder
        Dim mobjLogger2 As log4net.ILog = Nothing
        Try
            log4net.Config.XmlConfigurator.ConfigureAndWatch(New System.IO.FileInfo(GetConfigurationKey("DebugLogFilePath_LogForNet")))            
            mobjLogger2 = log4net.LogManager.GetLogger("SUPER_TRUMP")
            'Genarate Error Message String
            If astrClassName <> "" Then lstrErrorMessage.Append("ClassName: " + astrClassName)
            If astrFunctionName <> "" Then lstrErrorMessage.Append("FunctionName: " + astrFunctionName + "()")            
            lstrErrorMessage.Append("""" & astrMessage & """")
            mobjLogger2.Error(lstrErrorMessage.ToString)
        Catch ex As Exception
            Err.Source = "modBSCEFSuperTrump:WriteToTextDebugFile/" & Err.Source
            Err.Description = "modBSCEFSuperTrump:WriteToTextDebugFile/" & Err.Description
            Throw ex
        Finally
            If Not (mobjLogger2 Is Nothing) Then
                mobjLogger2 = Nothing
            End If
        End Try
    End Sub
    '================================================================
    'METHOD  : AddXMLElement
    'PURPOSE : To add an element to the specified XML DOM Document.
    'PARMS   :
    '          aobjXMLDOM [XMLDOMDocument] = XML DOM Document.
    '          aobjParentNode [XMLNode] = Parent Node reference.
    '          astrElementName [String] = Name of the element to add.
    '          astrElementContent [String] = Value of the element.
    'RETURN  : [XMLNode] = Inputed XML DOM with the new
    '          element.
    '================================================================
    Public Function AddXMLElement(ByRef aobjXMLDOM As Xml.XmlDocument, ByVal aobjParentNode As Xml.XmlNode, ByVal astrElementName As String, ByVal astrElementContent As String) As Xml.XmlNode
        Dim lobjNewNode As Xml.XmlNode = Nothing
        Dim lobjElementNode As Xml.XmlElement = Nothing
        Try
            Select Case aobjParentNode.NodeType
                Case Xml.XmlNodeType.Document, Xml.XmlNodeType.DocumentFragment, Xml.XmlNodeType.EntityReference, Xml.XmlNodeType.Element
                    lobjElementNode = aobjXMLDOM.CreateElement(astrElementName)                    
                    lobjNewNode = aobjParentNode.AppendChild(lobjElementNode)
                    If (Len(astrElementContent)) Then
                        lobjNewNode.InnerText = astrElementContent
                    End If
                    Return lobjNewNode
                Case Else                    
                    Return Nothing
            End Select           
        Catch ex As Exception
            Err.Source = "modBSCEFSuperTrump:AddXMLElement/" & Err.Source
            Err.Description = "modBSCEFSuperTrump:AddXMLElement/" & Err.Description
            Throw ex                       
        Finally
            If Not (lobjNewNode Is Nothing) Then
                lobjNewNode = Nothing
            End If
            If Not (lobjElementNode Is Nothing) Then
                lobjElementNode = Nothing
            End If                     
        End Try
    End Function

    '================================================================
    'METHOD  : AddXMLElementAttribute
    'PURPOSE : To add an attribute to the specified element in the
    '          XML DOM Document.
    'PARMS   :
    '          aobjXMLDOM [XMLDocument] = XML DOM Document.
    '          aobjElementNode [XMLNode] = Element Node
    '          reference.
    '          astrAttributeName [String] = Name of the attribute to
    '          be added to the element.
    '          astrAttributeContent [String] = Value of the attribute.
    'RETURN  : [XMLNode] = Element Node with the new attribute.
    '================================================================
    Public Function AddXMLElementAttribute(ByRef aobjXMLDOM As Xml.XmlDocument, ByVal aobjElementNode As Xml.XmlNode, ByVal astrAttributeName As String, ByVal astrAttributeContent As String) As Xml.XmlNode
        Dim lobjAttr As Xml.XmlAttribute = Nothing
        Try
            Select Case aobjElementNode.NodeType
                Case Xml.XmlNodeType.DocumentType, Xml.XmlNodeType.DocumentFragment, Xml.XmlNodeType.EntityReference, Xml.XmlNodeType.Element
                    lobjAttr = aobjXMLDOM.CreateAttribute(astrAttributeName)
                    lobjAttr.InnerText = astrAttributeContent                    
                    aobjElementNode.Attributes.SetNamedItem(lobjAttr)                    
                    lobjAttr = Nothing
                    Return aobjElementNode
                Case Else
                    Return Nothing
            End Select
        Catch ex As Exception            
            Err.Source = "modBSCEFSuperTrump:AddXMLElementAttribute/" & Err.Source
            Err.Description = "modBSCEFSuperTrump:AddXMLElementAttribute/" & Err.Description
            Throw ex           
        Finally
            If Not (lobjAttr Is Nothing) Then
                lobjAttr = Nothing
            End If                      
        End Try
    End Function

    '================================================================
    'METHOD  : GetXMLElementValue
    'PURPOSE : To get the value of an element in the XML DOM Document.
    'PARMS   :
    '          aobjXMLDOM [XMLNode] = XML DOM Document.
    '          astrElementXPath [String] = XPath of the element in the
    '          XML DOM Document.
    'RETURN  : String = Value of the element.
    '================================================================
    Public Function GetXMLElementValue(ByVal aobjXMLDOM As Xml.XmlNode, ByVal astrElementXPath As String) As String
        Dim lobjElement As Xml.XmlElement = Nothing
        Dim lstrElementVal As String
        Try
            lobjElement = aobjXMLDOM.SelectSingleNode(astrElementXPath)
            If Not (lobjElement Is Nothing) Then
                lstrElementVal = lobjElement.InnerText
            Else
                lstrElementVal = ""
            End If
            Return lstrElementVal
        Catch ex As Exception
            Err.Source = "modBSCEFSuperTrump:GetXMLElementValue/" & Err.Source
            Err.Description = "modBSCEFSuperTrump:GetXMLElementValue/" & Err.Description
            Throw ex                        
        Finally
            If Not (lobjElement Is Nothing) Then
                lobjElement = Nothing
            End If            
        End Try
    End Function

    '================================================================
    'METHOD  : AddBinaryXMLElement
    'PURPOSE : To add a binary element to the specified XML DOM
    '          Document.
    'PARMS   :
    '          aobjXMLDOM [XMLDocument] = XML DOM Document.
    '          aobjParentNode [XMLNode] = Parent Node reference.
    '          astrElementName [String] = Name of the element to add.
    '          avElementContent [String] = Value of the element.
    'RETURN  : [XMLNode] = Inputed XML DOM with the new
    '          element.
    '================================================================
    Public Function AddBinaryXMLElement(ByRef aobjXMLDOM As Xml.XmlDocument, ByVal aobjParentNode As Xml.XmlNode, ByVal astrElementName As String, ByVal avElementContent As Object) As Xml.XmlNode
        Dim lobjNewNode As Xml.XmlNode = Nothing
        Dim lobjElementNode As Xml.XmlElement = Nothing
        Try
            Select Case aobjParentNode.NodeType
                Case Xml.XmlNodeType.DocumentType, Xml.XmlNodeType.DocumentFragment, Xml.XmlNodeType.EntityReference, Xml.XmlNodeType.Element
                    lobjElementNode = aobjXMLDOM.CreateElement(astrElementName)                    
                    lobjNewNode = aobjParentNode.AppendChild(lobjElementNode)
                    If (Convert.ToString(avElementContent).Length > 0) Then
                        lobjNewNode.InnerText = Convert.ToBase64String(avElementContent)
                    End If
                    Return lobjNewNode
                Case Else                   
                    Return Nothing
            End Select            
        Catch ex As Exception
            Err.Source = "modBSCEFSuperTrump:AddBinaryXMLElement/" & Err.Source
            Err.Description = "modBSCEFSuperTrump:AddBinaryXMLElement/" & Err.Description
            Throw ex                       
        Finally           
            If Not (lobjNewNode Is Nothing) Then
                lobjNewNode = Nothing
            End If
            If Not (lobjElementNode Is Nothing) Then
                lobjElementNode = Nothing
            End If
        End Try
    End Function

    '================================================================
    'METHOD  : GetBinaryFileData
    'PURPOSE : Retrieve the data contained in a binary file.
    'PARMS   :
    '          astrFileName [String] = File name with the complete
    '          path information.
    'RETURN  : Variant = The data contained in a binary file.
    '================================================================
    Public Function GetBinaryFileData(ByVal astrFileName As String) As Object
        '''''''Changed To New Code ----Dim lobjFile As New ADODB.Stream        
        Dim oFile As System.IO.FileInfo
        oFile = New System.IO.FileInfo(astrFileName)
        Dim oFileStream As System.IO.FileStream = oFile.OpenRead()
        Dim lBytes As Long = oFileStream.Length
        Dim fileData(lBytes) As Byte
        Try
            ' Read the file into a byte array
            oFileStream.Read(fileData, 0, lBytes)
            oFileStream.Close()            
            Return fileData                       
        Catch ex As Exception
            Err.Source = "modBSCEFSuperTrump:GetBinaryFileData/" & Err.Source
            Err.Description = "modBSCEFSuperTrump:GetBinaryFileData/" & Err.Description
            Throw ex                      
        Finally
            oFileStream.Dispose()
            oFileStream = Nothing
        End Try
    End Function
    Function GetConfigurationKey(ByVal astrKey As String) As String
        Dim ConstantsFilePath As String
        Dim lDocXmlFile As New XmlDocument
        Dim regKey As RegistryKey
        Dim aName As String
        Try
            'regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\FacilitySettings\1000\FilePath", True)
            'ConstantsFilePath = regKey.GetValue("ConfigFilePath")
            'ConstantsFilePath = ConstantsFilePath + "\Constant.XML"  
            ConstantsFilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetCallingAssembly.GetName.CodeBase)
            'aName = System.Reflection.Assembly.GetEntryAssembly.GetModules()(0).FullyQualifiedName            
            'ConstantsFilePath = System.IO.Path.GetDirectoryName(aName)
            ConstantsFilePath = ConstantsFilePath + "\Constant.XML"
            lDocXmlFile.Load(ConstantsFilePath)
            Return lDocXmlFile.GetElementsByTagName(astrKey).Item(0).InnerText                  
        Catch ex As Exception
            Err.Source = "modBSCEFSuperTrump:GetConfigurationKey/" & Err.Source
            Err.Description = "modBSCEFSuperTrump:GetConfigurationKey/" & Err.Description
            Throw ex                       
        Finally            
            If Not (lDocXmlFile Is Nothing) Then
                lDocXmlFile = Nothing
            End If
        End Try
    End Function
End Module