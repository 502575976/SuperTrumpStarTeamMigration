Attribute VB_Name = "modBSCEFSuperTrump"
'================================================================
'MODULE  : modBSCEFSuperTrump
'PURPOSE : Contains library functions, global consts and vars.
'================================================================
Option Explicit

'=Registry Constants=========================================================================
'FACILITY ID which is specific to this Component. This is used to pull out all information
'specific to this component from the registry.
Public Const gcFacilityID                   As Integer = 1000
Public Const gcFacilityConfigPath           As String = "HKEY_LOCAL_MACHINE\Software\FacilitySettings\"

Public Const gcPRMFilePathKey               As String = "FilePath\PRMFilePath"
Public Const gcReportTemplatePathKey        As String = "FilePath\ReportTemplatePath"
Public Const gcSchemaFilePathKey            As String = "FilePath\SchemaFilePath"
Public Const gcPRMTemplatePathKey           As String = "FilePath\PRMTemplatePath"

Public Const gcResponseQMgrKey              As String = "QInfo\ResponseQueueManager"
Public Const gcResponseQKey                 As String = "QInfo\ResponseQueue"
Public Const gcRequestQMgrKey               As String = "QInfo\RequestQueueManager"
Public Const gcRequestQKey                  As String = "QInfo\RequestQueue"
'============================================================================================

'=Other Constants============================================================================
Public Const gcPRMFileLstXMLSchemaName      As String = "PRMFileListXML.xsd"
Public Const gcPricingRepInfoXMLSchemaName  As String = "PricingRepInfoXML.xsd"
Public Const gcPRMInfoXMLSchemaName         As String = "PRMInfoXML.xsd"
Public Const gcPRMParamsInfoXMLSchemaName   As String = "PRMParamsInfoXML.xsd"
Public Const gcPricingReqInfoXMLSchemaName  As String = "PricingRequestInfoXML.xsd"
Public Const gcMQMsgInfoXMLSchemaName       As String = "MQMessageInfo.xsd"
Public Const gcModifyPRMFilesSchemaName     As String = "ModifyPRMFilesXML.xsd"
Public Const gcGenPRMForPmtStructSchemaName As String = "GeneratePRMForPmtStructXML.xsd"
'============================================================================================

'=== Error Constants ========================================================================
Public Const gcHIGHEST_ERROR                As Variant = vbObjectError + 256
Public Const gcINVALID_PRM_FILE             As Variant = gcHIGHEST_ERROR + 1000
'============================================================================================

'=== Debug Constants ========================================================================
Public giDebugLevel             As Integer
Public gstrDebugFile            As String
Public glMaxDebugFileSize       As Long

Public Const gcAPPEND_IO_MODE = 8
Public Const gcWRITE_IO_MODE = 2
'============================================================================================

'=== Debug Constants ========================================================================
'Variables defined to  force the aniequity error in the output : 5th march 2007
Public gstrExceptionFlag        As String
Public gliPRMFilearr()          As Integer
Public gstrExcptionXMLDOMarr()  As String


'============================================================================================

'================================================================
'METHOD  : DeletePRMBinaryFile
'PURPOSE : Deletes .PRM file(s)
'PARMS   :
'          alobjFilenameList [IXMLDOMNodeList] = List of file
'          names.
'          E.g.
'          <FILE_NAME>file1.prm</FILE_NAME>
'          <FILE_NAME>file2.prm</FILE_NAME>
'          ...
'          astrPath [String] = Complete Path to the .PRM File
'RETURN  : None
'================================================================
Public Sub DeletePRMBinaryFile(ByVal alobjFilenameList As IXMLDOMNodeList, _
                                ByVal astrPath As String)
On Error GoTo ErrHandler

Dim lobjFileNode    As IXMLDOMNode
Dim lobjFileSysObj  As New Scripting.FileSystemObject
Dim lstrFileName    As String

    
    'For each file in the file list
    For Each lobjFileNode In alobjFilenameList
            
        lstrFileName = lobjFileNode.Text
        If UCase(Right(lstrFileName, 4)) <> ".PRM" Then lstrFileName = lstrFileName & ".PRM"
        
        'Check if file exists
        If lobjFileSysObj.FileExists(astrPath & "\" & lstrFileName) Then
            
            'Delete the PRM file
            Call lobjFileSysObj.DeleteFile(astrPath & "\" & lstrFileName)
        End If
    Next
    
    Set lobjFileNode = Nothing
    Set lobjFileSysObj = Nothing
    
    Exit Sub
    
ErrHandler:
    Set lobjFileNode = Nothing
    Set lobjFileSysObj = Nothing
    Err.Raise Err.Number, "modBSCEFSuperTrump:DeletePRMBinaryFile", Err.Description
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
Public Function GetSuperTrumpQuery(ByVal astrTransID As String, _
                                    ByVal astrFileName As String) As String
On Error GoTo ErrHandler

Dim lobjSTDOM   As New DOMDocument40
Dim lobjElem    As IXMLDOMElement
Dim lobjAttr    As IXMLDOMAttribute
Dim lstrXML     As String


    lstrXML = "<SuperTRUMP></SuperTRUMP>"
                        
    If lobjSTDOM.loadXML(lstrXML) Then
        Set lobjElem = lobjSTDOM.createElement("Transaction")
        
        Set lobjAttr = lobjSTDOM.createAttribute("id")
        lobjAttr.Text = astrTransID
        lobjElem.Attributes.setNamedItem lobjAttr
        Set lobjAttr = Nothing
        
        Set lobjAttr = lobjSTDOM.createAttribute("query")
        lobjAttr.Text = "true"
        lobjElem.Attributes.setNamedItem lobjAttr
        Set lobjAttr = Nothing
        
        lobjSTDOM.documentElement.appendChild lobjElem
        Set lobjElem = Nothing
        
        Set lobjElem = lobjSTDOM.createElement("ReadFile")
        
        Set lobjAttr = lobjSTDOM.createAttribute("filename")
        lobjAttr.Text = astrFileName
        lobjElem.Attributes.setNamedItem lobjAttr
        Set lobjAttr = Nothing
        
        lobjSTDOM.documentElement.childNodes(0).appendChild lobjElem
        Set lobjElem = Nothing
        
        GetSuperTrumpQuery = lobjSTDOM.xml
    Else
        GetSuperTrumpQuery = ""
        Err.Raise lobjSTDOM.parseError.errorCode, "modBSCEFSuperTrump:GetSuperTrumpQuery", "Error on line " & lobjSTDOM.parseError.Line & " of XML. " & lobjSTDOM.parseError.reason
    End If
    
    Set lobjSTDOM = Nothing
    Exit Function
    
ErrHandler:
    GetSuperTrumpQuery = ""
    Err.Raise Err.Number, "modBSCEFSuperTrump:GetSuperTrumpQuery", Err.Description
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
Public Function SavePRMBinaryFile(ByVal astrFileName As String, _
                                    ByVal astrFilePath As String, _
                                    ByRef aarrFileData() As Byte) As String
On Error GoTo ErrHandler

Dim liFileDes As Integer

    
    liFileDes = FreeFile()
    astrFileName = astrFilePath & "\" & UCase(astrFileName)
    Open astrFileName For Binary Access Write As liFileDes
    Put liFileDes, , aarrFileData
    Close liFileDes
    
    SavePRMBinaryFile = astrFileName
    Exit Function
    
ErrHandler:
    SavePRMBinaryFile = ""
    Err.Raise Err.Number, "modBSCEFSuperTrump:SavePRMBinaryFile", Err.Description
End Function

'================================================================
'METHOD  : ReadRegistry
'PURPOSE : Retrives the spcified registry Key value.
'PARMS   :
'          astrRegKey [String] = Registry Key Name along with
'          it's complete path.
'RETURN  : String = Value of the Registry Key.
'================================================================
Public Function ReadRegistry(ByVal astrRegKey As String) As String
On Error GoTo ErrHandler

Dim lobjReg As New IWshRuntimeLibrary.IWshShell_Class   '.WshShell
            
    'Read the key value from the registry
    ReadRegistry = lobjReg.RegRead(astrRegKey)
    Set lobjReg = Nothing
    Exit Function
        
ErrHandler:
    Set lobjReg = Nothing
    Err.Raise Err.Number, "modBSCEFSuperTrump:ReadRegistry", Err.Description
End Function

'================================================================
'METHOD  : WriteToTextDebugFile
'PURPOSE : To help in debugging errors in the compiled component.
'PARMS   :
'          astrFileName [String] = Debug file name with complete
'          path.
'          astrData [String] = Data to be written to the debug
'          file.
'RETURN  : Boolean = True if data is written successfully & false
'          otherwise.
'================================================================
Public Function WriteToTextDebugFile(ByVal astrFileName As String, _
                                    ByVal astrData As String) As Boolean
On Error GoTo ErrHandler:

Dim lobjFileSystem  As New Scripting.FileSystemObject
Dim lobjFile        As Scripting.File
Dim lobjTxtStream   As Scripting.TextStream
Dim liIOMode        As Integer

    
    'Check if debug file exists
    If lobjFileSystem.FileExists(astrFileName) Then
        
        'Retrieve the debug file
        Set lobjFile = lobjFileSystem.GetFile(astrFileName)
        
        'If debug file size exceeds the max. defined size
        If lobjFile.Size > glMaxDebugFileSize Then
            
            'Copy the existing file as a .bak file and set file mode to overwrite
            liIOMode = gcWRITE_IO_MODE
            lobjFile.Copy Replace(astrFileName, ".txt", ".bak"), True
        
        'Else If debug file size doesn't exceeds the max. defined size
        Else
            
            'Set file mode to append
            liIOMode = gcAPPEND_IO_MODE
        End If
        
        'Open the file in the appropriate mode
        Set lobjTxtStream = lobjFile.OpenAsTextStream(liIOMode, TristateUseDefault)
    
    'Else if debug file doesn't exists
    Else
        
        'Create and open the file
        Set lobjTxtStream = lobjFileSystem.CreateTextFile(astrFileName)
        liIOMode = gcWRITE_IO_MODE
    End If
    
    'Write data to the debug file
    lobjTxtStream.WriteLine Now & " " & astrData
    
    'Close the file
    lobjTxtStream.Close
    
    Set lobjFile = Nothing
    Set lobjTxtStream = Nothing
    Set lobjFileSystem = Nothing
    WriteToTextDebugFile = True
    
    Exit Function

ErrHandler:
    WriteToTextDebugFile = False
    Set lobjFile = Nothing
    Set lobjTxtStream = Nothing
    Set lobjFileSystem = Nothing
End Function

'================================================================
'METHOD  : AddXMLElement
'PURPOSE : To add an element to the specified XML DOM Document.
'PARMS   :
'          aobjXMLDOM [DOMDocument] = XML DOM Document.
'          aobjParentNode [IXMLDOMNode] = Parent Node reference.
'          astrElementName [String] = Name of the element to add.
'          astrElementContent [String] = Value of the element.
'RETURN  : [IXMLDOMNode] = Inputed XML DOM with the new
'          element.
'================================================================
Public Function AddXMLElement(ByRef aobjXMLDOM As DOMDocument, _
                            ByVal aobjParentNode As IXMLDOMNode, _
                            ByVal astrElementName As String, _
                            ByVal astrElementContent As String) As IXMLDOMNode
On Error GoTo ErrHand
    
Dim lobjNewNode     As IXMLDOMNode
Dim lobjElementNode As IXMLDOMElement
    
    Select Case aobjParentNode.nodeType
        Case NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT, _
            NODE_ENTITY_REFERENCE, NODE_ELEMENT:
            
            Set lobjElementNode = aobjXMLDOM.createElement(astrElementName)
            Set lobjNewNode = aobjParentNode.appendChild(lobjElementNode)
            If (Len(astrElementContent)) Then
                lobjNewNode.Text = astrElementContent
            End If
            Set AddXMLElement = lobjNewNode
        
        Case Else
            Set AddXMLElement = Nothing
    End Select
    
    Set lobjNewNode = Nothing
    Set lobjElementNode = Nothing
    
    Exit Function

ErrHand:
    Set AddXMLElement = Nothing
    Set lobjNewNode = Nothing
    Set lobjElementNode = Nothing
    
    Err.Source = "modBSCEFSuperTrump:AddXMLElement/" & Err.Source
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'================================================================
'METHOD  : AddXMLElementAttribute
'PURPOSE : To add an attribute to the specified element in the
'          XML DOM Document.
'PARMS   :
'          aobjXMLDOM [DOMDocument] = XML DOM Document.
'          aobjElementNode [IXMLDOMNode] = Element Node
'          reference.
'          astrAttributeName [String] = Name of the attribute to
'          be added to the element.
'          astrAttributeContent [String] = Value of the attribute.
'RETURN  : [IXMLDOMNode] = Element Node with the new attribute.
'================================================================
Public Function AddXMLElementAttribute(ByRef aobjXMLDOM As DOMDocument, _
                            ByVal aobjElementNode As IXMLDOMNode, _
                            ByVal astrAttributeName As String, _
                            ByVal astrAttributeContent As String) As IXMLDOMNode
On Error GoTo ErrHand
    
Dim lobjAttr As IXMLDOMAttribute

    Select Case aobjElementNode.nodeType
        Case NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT, _
            NODE_ENTITY_REFERENCE, NODE_ELEMENT:
            
            Set lobjAttr = aobjXMLDOM.createAttribute(astrAttributeName)
            lobjAttr.Text = astrAttributeContent
            aobjElementNode.Attributes.setNamedItem lobjAttr
            Set lobjAttr = Nothing
        
            Set AddXMLElementAttribute = aobjElementNode
            
        Case Else
            Set AddXMLElementAttribute = Nothing
    End Select
        
    Set lobjAttr = Nothing
    
    Exit Function

ErrHand:
    Set AddXMLElementAttribute = Nothing
    Set lobjAttr = Nothing
    
    Err.Source = "modBSCEFSuperTrump:AddXMLElementAttribute/" & Err.Source
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'================================================================
'METHOD  : GetXMLElementValue
'PURPOSE : To get the value of an element in the XML DOM Document.
'PARMS   :
'          aobjXMLDOM [IXMLDOMNode] = XML DOM Document.
'          astrElementXPath [String] = XPath of the element in the
'          XML DOM Document.
'RETURN  : String = Value of the element.
'================================================================
Public Function GetXMLElementValue(ByVal aobjXMLDOM As IXMLDOMNode, _
                                    ByVal astrElementXPath As String) As String
On Error GoTo ErrHand
    
Dim lobjElement     As IXMLDOMElement
Dim lstrElementVal  As String

    Set lobjElement = aobjXMLDOM.selectSingleNode(astrElementXPath)
    If Not (lobjElement Is Nothing) Then
        lstrElementVal = lobjElement.Text
    Else
        lstrElementVal = ""
    End If
    
    GetXMLElementValue = lstrElementVal
    
    Set lobjElement = Nothing
    
    Exit Function

ErrHand:
    GetXMLElementValue = ""
    Set lobjElement = Nothing
    
    Err.Source = "modBSCEFSuperTrump:GetXMLElementValue/" & Err.Source
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'================================================================
'METHOD  : AddBinaryXMLElement
'PURPOSE : To add a binary element to the specified XML DOM
'          Document.
'PARMS   :
'          aobjXMLDOM [DOMDocument] = XML DOM Document.
'          aobjParentNode [IXMLDOMNode] = Parent Node reference.
'          astrElementName [String] = Name of the element to add.
'          avElementContent [String] = Value of the element.
'RETURN  : [IXMLDOMNode] = Inputed XML DOM with the new
'          element.
'================================================================
Public Function AddBinaryXMLElement(ByRef aobjXMLDOM As DOMDocument, _
                            ByVal aobjParentNode As IXMLDOMNode, _
                            ByVal astrElementName As String, _
                            ByVal avElementContent As Variant) As IXMLDOMNode
On Error GoTo ErrHand
    
Dim lobjNewNode     As IXMLDOMNode
Dim lobjElementNode As IXMLDOMElement
    
    Select Case aobjParentNode.nodeType
        Case NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT, _
            NODE_ENTITY_REFERENCE, NODE_ELEMENT:
            
            Set lobjElementNode = aobjXMLDOM.createElement(astrElementName)
            lobjElementNode.dataType = "bin.base64"
            Set lobjNewNode = aobjParentNode.appendChild(lobjElementNode)
            If (Len(avElementContent)) Then
                lobjNewNode.nodeTypedValue = avElementContent
            End If
            Set AddBinaryXMLElement = lobjNewNode
        
        Case Else
            Set AddBinaryXMLElement = Nothing
    End Select
    
    Set lobjNewNode = Nothing
    Set lobjElementNode = Nothing
    
    Exit Function

ErrHand:
    Set AddBinaryXMLElement = Nothing
    Set lobjNewNode = Nothing
    Set lobjElementNode = Nothing
    
    Err.Source = "modBSCEFSuperTrump:AddBinaryXMLElement/" & Err.Source
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'================================================================
'METHOD  : GetBinaryFileData
'PURPOSE : Retrieve the data contained in a binary file.
'PARMS   :
'          astrFileName [String] = File name with the complete
'          path information.
'RETURN  : Variant = The data contained in a binary file.
'================================================================
Public Function GetBinaryFileData(ByVal astrFileName As String) As Variant
Dim lobjFile As New ADODB.Stream

On Error GoTo ErrHandler
    
    'Open the pricing report text file for binary read.
    lobjFile.Type = adTypeBinary
    lobjFile.Open
    lobjFile.LoadFromFile astrFileName
    
    'Return the binary pricing report data.
    GetBinaryFileData = lobjFile.Read
    
    Set lobjFile = Nothing
    Exit Function
    
ErrHandler:
    GetBinaryFileData = ""
    Set lobjFile = Nothing
    Err.Source = "modBSCEFSuperTrump:GetBinaryFileData/" & Err.Source
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
