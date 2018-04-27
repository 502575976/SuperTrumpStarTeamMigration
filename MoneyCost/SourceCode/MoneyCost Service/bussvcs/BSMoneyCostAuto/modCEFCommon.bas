Attribute VB_Name = "modCEFCommon"
'================================================================
' GE Capital Proprietary and Confidential
' Copyright (c) 2001-2002 by GE Capital - All rights reserved.
'
' This code may not be reproduced in any way without express
' permission from GE Capital.
'================================================================

Option Explicit

'=== Constant for module name ===================================
Private Const cMODULE_NAME  As String = "modCEFCommon"
'================================================================

'=== Error Constants ============================================
Public Const cERR_ROOT_TAG  As String = "ERROR_DETAILS"
Public Const cERR_NBR       As String = "ERROR_NUMBER"
Public Const cERR_DESC      As String = "ERROR_DESCRIPTION"
Public Const cERR_SRC       As String = "ERROR_SOURCE"
Public Const cERR_ShowUser  As String = "ERROR_SHOW_USER"
'================================================================

'=== Debug Constants & variables ================================
Public Const cAPPEND_IO_MODE    As Integer = 8
Public Const cWRITE_IO_MODE     As Integer = 2

Public giDebugLevel             As Integer
Public gstrDebugFile            As String
Public glMaxDebugFileSize       As Long

Public gstrErrMailBox           As String
Public gstrEmailOverride        As String
Public gstrDeveloperEmail       As String
Public gstrClarifySiteId        As String
Public gstrClarifyEmailSub      As String
Public gstrClarifyPriority      As String
Public gstrClarifyContactFname  As String
Public gstrClarifyContactLname  As String
Public gstrClarifyContactPhone  As String
Public gstrFrom                 As String



'================================================================

'================================================================
'METHOD  : BuildErrXML
'PURPOSE : Builds the Error XML.
'PARMS   :
'          astrErrNbr [String] = Error number
'          astrErrSource [String] = Error Source
'          astrErrDesc [String] = Error description
'RETURN  : String = Error XML
'================================================================
Public Function BuildErrXML(ByVal astrErrNbr As String, _
                                ByVal astrErrSource As String, _
                                ByVal astrErrDesc As String) As String
On Error GoTo Errhandler

Dim lobjErrXMLDOM As New DOMDocument

    'Load output XML
    lobjErrXMLDOM.loadXML ("<" & cERR_ROOT_TAG & "/>")

    'Add the error number tag to the output XML
    AddXMLElement lobjErrXMLDOM, _
                    lobjErrXMLDOM.documentElement, _
                    cERR_NBR, _
                    astrErrNbr

    'Add the error source tag to the output XML
    AddXMLElement lobjErrXMLDOM, _
                    lobjErrXMLDOM.documentElement, _
                    cERR_SRC, _
                    astrErrSource

    'Add the error description tag to the output XML
    AddXMLElement lobjErrXMLDOM, _
                    lobjErrXMLDOM.documentElement, _
                    cERR_DESC, _
                    astrErrDesc

    'Return the error XML
    BuildErrXML = lobjErrXMLDOM.xml

    Set lobjErrXMLDOM = Nothing
    Exit Function

Errhandler:
    BuildErrXML = vbNullString
    Set lobjErrXMLDOM = Nothing

    Err.Raise Err.Number, cMODULE_NAME & ":BuildErrXML()", Err.Description
End Function

'================================================================
'METHOD  : BuildErrorXML
'PURPOSE : Builds the Error XML with User friendly error messages.
'PARMS   :
'          astrErrNbr [String] = Error number
'          astrErrSource [String] = Error Source
'          astrErrDesc [String] = Error description
'          astrErrorShowUser [String] = Error description for User
'RETURN  : String = Error XML
'================================================================
Public Function BuildErrorXML(ByVal astrErrNbr As String, _
                                ByVal astrErrSource As String, _
                                ByVal astrErrDesc As String, _
                                ByVal astrErrorShowUser As String) As String
On Error GoTo Errhandler

Dim lobjErrXMLDOM As New DOMDocument

    'Load output XML
    lobjErrXMLDOM.loadXML ("<" & cERR_ROOT_TAG & "/>")

    'Add the error number tag to the output XML
    AddXMLElement lobjErrXMLDOM, _
                    lobjErrXMLDOM.documentElement, _
                    cERR_NBR, _
                    astrErrNbr

    'Add the error source tag to the output XML
    AddXMLElement lobjErrXMLDOM, _
                    lobjErrXMLDOM.documentElement, _
                    cERR_SRC, _
                    astrErrSource

    'Add the error description tag to the output XML
    AddXMLElement lobjErrXMLDOM, _
                    lobjErrXMLDOM.documentElement, _
                    cERR_DESC, _
                    astrErrDesc

    'Add the error description for user tag to the output XML
    AddXMLElement lobjErrXMLDOM, _
                    lobjErrXMLDOM.documentElement, _
                    cERR_ShowUser, _
                    astrErrorShowUser

    'Return the error XML
    BuildErrorXML = lobjErrXMLDOM.xml

    Set lobjErrXMLDOM = Nothing
    Exit Function

Errhandler:
    BuildErrorXML = vbNullString
    Set lobjErrXMLDOM = Nothing

    Err.Raise Err.Number, cMODULE_NAME & ":BuildErrorXML()", Err.Description
End Function

'================================================================
'METHOD :   ReadRegistry.
'PURPOSE:   This method will use the Windows Scripting Host
'           Object Library's - IWshRuntimeLibrary.IWshShell_Class
'           component to query the registry to get the key value
'           based on the supplied Full Key name.
'PARMS  :   astrRegFullKeyName [String] = The Key name along with
'           its full registry path.
'           Eg:
'           HKEY_LOCAL_MACHINE\Software\FacilitySettings\BSMyDll
'           \ConnectStrings\MyDbConn
'RETURN :   String.
'================================================================
Public Function ReadRegistry(ByVal astrRegFullKeyName As String) As String
On Error GoTo Errhandler

Dim lobjReg     As New IWshRuntimeLibrary.IWshShell_Class

    'Read the key value from the registry
    ReadRegistry = lobjReg.RegRead(astrRegFullKeyName)
    Set lobjReg = Nothing
    Exit Function

Errhandler:
    ReadRegistry = vbNullString
    Set lobjReg = Nothing
    Err.Raise Err.Number, cMODULE_NAME & ":ReadRegistry()", Err.Description
End Function

'================================================================
'METHOD  : IsXMLElementPresent
'PURPOSE : To check if an element is present in the XML DOM
'          Document.
'PARMS   :
'          aobjXMLDOM [IXMLDOMNode] = XML DOM Document.
'          astrElementXPath [String] = XPath of the element in the
'          XML DOM Document.
'RETURN  : Boolean = True if present Else false.
'================================================================
Public Function IsXMLElementPresent(ByVal aobjXMLDOM As IXMLDOMNode, _
                                    ByVal astrElementXPath As String) As Boolean
On Error GoTo Errhandler

Dim lobjElement         As IXMLDOMElement
Dim lbElementPresent    As Boolean

    lbElementPresent = False
    Set lobjElement = aobjXMLDOM.selectSingleNode(astrElementXPath)
    If Not (lobjElement Is Nothing) Then
        lbElementPresent = True
    End If

    IsXMLElementPresent = lbElementPresent

    Set lobjElement = Nothing

    Exit Function

Errhandler:
    IsXMLElementPresent = False
    Set lobjElement = Nothing

    Err.Raise Err.Number, cMODULE_NAME & ":IsXMLElementPresent()", Err.Description
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
On Error GoTo Errhandler

Dim lobjNewNode     As IXMLDOMNode
Dim lobjElementNode As IXMLDOMElement

    'Check if the Parent node type is NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT,
    'NODE_ENTITY_REFERENCE, NODE_ELEMENT
    Select Case aobjParentNode.nodeType
        Case NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT, _
            NODE_ENTITY_REFERENCE, NODE_ELEMENT:

            'Add new element to the Parent node of the XML DOM
            Set lobjElementNode = aobjXMLDOM.createElement(astrElementName)
            Set lobjNewNode = aobjParentNode.appendChild(lobjElementNode)
            If (Len(astrElementContent)) Then
                lobjNewNode.Text = astrElementContent
            End If

            'Return the new XML DOM
            Set AddXMLElement = lobjNewNode

    'Otherwise, return nothing
        Case Else
            Set AddXMLElement = Nothing
    End Select

    Set lobjNewNode = Nothing
    Set lobjElementNode = Nothing

    Exit Function

Errhandler:
    Set AddXMLElement = Nothing
    Set lobjNewNode = Nothing
    Set lobjElementNode = Nothing

    Err.Raise Err.Number, cMODULE_NAME & ":AddXMLElement()", Err.Description
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
On Error GoTo Errhandler

Dim lobjAttr As IXMLDOMAttribute

    'Check if the element node type is NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT,
    'NODE_ENTITY_REFERENCE, NODE_ELEMENT
    Select Case aobjElementNode.nodeType
        Case NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT, _
            NODE_ENTITY_REFERENCE, NODE_ELEMENT:

            'Add new attribute to the element node of the XML DOM
            Set lobjAttr = aobjXMLDOM.createAttribute(astrAttributeName)
            lobjAttr.Text = astrAttributeContent
            aobjElementNode.Attributes.setNamedItem lobjAttr
            Set lobjAttr = Nothing

            'Return the new XML DOM
            Set AddXMLElementAttribute = aobjElementNode

    'Otherwise, return nothing
        Case Else
            Set AddXMLElementAttribute = Nothing
    End Select

    Set lobjAttr = Nothing

    Exit Function

Errhandler:
    Set AddXMLElementAttribute = Nothing
    Set lobjAttr = Nothing

    Err.Raise Err.Number, cMODULE_NAME & ":AddXMLElementAttribute()", Err.Description
End Function

'================================================================
'METHOD  : WriteToTextDebugFile
'PURPOSE : To help in debugging errors in the compiled component.
'PARMS   :
'          astrFileName [String] = Debug file name with complete
'          path.
'          astrData [String] = Data to be written to the debug
'          file.
'          abFileOverwrite [Boolean] = Flag which indicating
'          whether file needs to be overwritten or not.
'RETURN  : Boolean = True if data is written successfully & false
'          otherwise.
'================================================================
Public Function WriteToTextDebugFile(ByVal astrFileName As String, _
                                    ByVal astrData As String, _
                                    Optional abFileOverwrite As Boolean = False) As Boolean
On Error GoTo Errhandler

Dim lobjFileSystem  As New FileSystemObject
Dim lobjFile        As File
Dim lobjTxtStream   As TextStream
Dim liIOMode        As Integer



    'Check if debug file exists
    If lobjFileSystem.FileExists(astrFileName) Then

        'Retrieve the debug file
        Set lobjFile = lobjFileSystem.GetFile(astrFileName)

        'If flag set to overwrite file
        If abFileOverwrite Then

            'Set file mode to overwrite
            liIOMode = cWRITE_IO_MODE

        'Else if the overwrite flag is not set
        Else

            'If debug file size exceeds the max. defined size
            If lobjFile.Size > glMaxDebugFileSize Then

                'Copy the existing file as a .bak file and set file mode to overwrite
                liIOMode = cWRITE_IO_MODE
                lobjFile.Copy Replace(astrFileName, ".txt", ".bak"), True

            'Else If debug file size doesn't exceeds the max. defined size
            Else

                'Set file mode to append
                liIOMode = cAPPEND_IO_MODE
            End If
        End If

        'Open the file in the appropriate mode
        Set lobjTxtStream = lobjFile.OpenAsTextStream(liIOMode, TristateUseDefault)

    'Else if debug file doesn't exists
    Else

        'Create and open the file
        Set lobjTxtStream = lobjFileSystem.CreateTextFile(astrFileName)
        liIOMode = cWRITE_IO_MODE
    End If

    'Write data to the debug file
    lobjTxtStream.WriteLine Now & " " & astrData & vbCrLf

    'Close the file
    lobjTxtStream.Close

    Set lobjFile = Nothing
    Set lobjTxtStream = Nothing
    Set lobjFileSystem = Nothing
    WriteToTextDebugFile = True

    Exit Function

Errhandler:
    WriteToTextDebugFile = False
    Set lobjFile = Nothing
    Set lobjTxtStream = Nothing
    Set lobjFileSystem = Nothing
End Function
