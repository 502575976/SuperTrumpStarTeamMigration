Imports BSMoneyCostEntity
Imports System.Reflection
Imports Microsoft.Win32
Imports System.EnterpriseServices
Imports System.Xml
Module modCEFCommon
    '=== Constant for module name ===================================
    Private Const cMODULE_NAME As String = "modCEFCommon"
    '================================================================   
    '=== Error Constants ============================================
    Public Const cERR_ROOT_TAG As String = "ERROR_DETAILS"
    Public Const cERR_NBR As String = "ERROR_NUMBER"
    Public Const cERR_DESC As String = "ERROR_DESCRIPTION"
    Public Const cERR_SRC As String = "ERROR_SOURCE"
    Public Const cERR_ShowUser As String = "ERROR_SHOW_USER"
    '================================================================

    '=== Debug Constants & variables ================================
    Public gstrErrMailBox As String
    Public gstrEmailOverride As String
    Public gstrDeveloperEmail As String
    Public gstrClarifySiteId As String
    Public gstrClarifyEmailSub As String
    Public gstrClarifyPriority As String
    Public gstrClarifyContactFname As String
    Public gstrClarifyContactLname As String
    Public gstrClarifyContactPhone As String
    Public gstrFrom As String
    Public NotificationMailFlag As Boolean

    Public Function EncrptFile(ByVal ExeFileName As String, ByVal FileName As String)
        Dim process As New System.Diagnostics.Process
        Dim processInfo As New System.Diagnostics.ProcessStartInfo
        Dim outputEncryptFile As String
        Dim strReader As IO.StreamReader
        Try
            processInfo.FileName = ExeFileName
            processInfo.Arguments = FileName
            processInfo.RedirectStandardOutput = True
            processInfo.UseShellExecute = False
            strReader = process.Start(processInfo).StandardOutput
            outputEncryptFile = strReader.ReadLine.Trim
        Catch ex As Exception
            Throw
        Finally
            process.Dispose()
        End Try
        Return outputEncryptFile.Substring(outputEncryptFile.Length - 9, 9)
    End Function


    Public Function BuildErrXMLauto(ByVal objXmlErrEntity As XmlErrEntity) As cDataEntity
        Dim cdataEntity As New cDataEntity
        Dim lobjErrXMLDOM As New Xml.XmlDocument
        Try
            'Load output XML
            lobjErrXMLDOM.LoadXml("<" & cERR_ROOT_TAG & "/>")

            'Add the error number tag to the output XML
            AddXMLElement(lobjErrXMLDOM, _
                            lobjErrXMLDOM.DocumentElement, _
                            cERR_NBR, _
                            objXmlErrEntity.ErrNbr)

            'Add the error source tag to the output XML
            AddXMLElement(lobjErrXMLDOM, _
                            lobjErrXMLDOM.DocumentElement, _
                            cERR_SRC, _
                            objXmlErrEntity.ErrSource)

            'Add the error description tag to the output XML
            AddXMLElement(lobjErrXMLDOM, _
                            lobjErrXMLDOM.DocumentElement, _
                            cERR_DESC, _
                            objXmlErrEntity.ErrDesc)

            'Return the error XML
            cdataEntity.OutputString = lobjErrXMLDOM.OuterXml
            Return cdataEntity

            lobjErrXMLDOM = Nothing
        Catch ex As Exception
            cdataEntity.OutputString = vbNullString
            Return cdataEntity
            Throw
        Finally
            If Not IsNothing(lobjErrXMLDOM) Then
                lobjErrXMLDOM = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
            If Not IsNothing(objXmlErrEntity) Then
                objXmlErrEntity = Nothing
            End If

        End Try
    End Function
    Function GetConfigurationKey(ByVal astrKey As String) As String

        Dim ConstantsFilePath As String
        Dim lDocXmlFile As New XmlDocument
        Dim lobjreg As RegistryKey = Registry.LocalMachine
        Dim val As Object

        Try
            lobjreg = lobjreg.OpenSubKey("SOFTWARE\FacilitySettings\MoneyCost\FilePath", False)
            val = lobjreg.GetValue("ConfigFilePath")
            ConstantsFilePath = val
            ConstantsFilePath = ConstantsFilePath + "\Constant.XML"
            lDocXmlFile.Load(ConstantsFilePath)

            Return lDocXmlFile.GetElementsByTagName(astrKey).Item(0).InnerText

        Catch ex As Exception
            Throw ex
        Finally
            lDocXmlFile = Nothing
        End Try
    End Function
    Public Function ReadRegistry(ByVal astrRegFullKeyName As String) As String

        Dim regKey As RegistryKey
        Try
            regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\FacilitySettings\1000\FilePath", True)
            Return regKey.GetValue(astrRegFullKeyName)

        Catch ex As Exception
            Err.Raise(Err.Number, cMODULE_NAME & ":ReadRegistry()", Err.Description)
            Return vbNullString
            Throw
        End Try

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
    Public Function IsXMLElementPresent(ByVal objXmlErrEntity As XmlErrEntity) As cDataEntity
        Dim cdataEntity As New cDataEntity
        Dim lobjElement As Xml.XmlElement = Nothing
        Dim lbElementPresent As Boolean

        Try
            lbElementPresent = False
            lobjElement = objXmlErrEntity.IXMLDOMNode.SelectSingleNode(objXmlErrEntity.ElementXPath)
            If Not (lobjElement Is Nothing) Then
                lbElementPresent = True
            End If
            cdataEntity.OutputString = lbElementPresent
            Return cdataEntity

            lobjElement = Nothing
        Catch ex As Exception
            cdataEntity.OutputString = "False"
            Return cdataEntity
            Throw
        Finally
            If Not IsNothing(lobjElement) Then
                lobjElement = Nothing
            End If
            If Not IsNothing(objXmlErrEntity) Then
                objXmlErrEntity = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try

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
    Public Function AddXMLElement(ByRef aobjXMLDOM As Xml.XmlDocument, _
                                    ByVal aobjParentNode As Xml.XmlNode, _
                                    ByVal astrElementName As String, _
                                    ByVal astrElementContent As String) As Xml.XmlNode

        Dim lobjNewNode As Xml.XmlNode = Nothing
        Dim lobjElementNode As Xml.XmlElement = Nothing

        Try
            'Check if the Parent node type is NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT,
            'NODE_ENTITY_REFERENCE, NODE_ELEMENT
            Select Case aobjParentNode.NodeType

                Case Xml.XmlNodeType.Document, Xml.XmlNodeType.DocumentFragment, _
                    Xml.XmlNodeType.EntityReference, Xml.XmlNodeType.Element

                    'Add new element to the Parent node of the XML DOM
                    lobjElementNode = aobjXMLDOM.CreateElement(astrElementName)
                    lobjNewNode = aobjParentNode.AppendChild(lobjElementNode)
                    If (Len(astrElementContent)) Then
                        lobjNewNode.InnerText = astrElementContent
                    End If

                    'Return the new XML DOM
                    Return lobjNewNode

                    'Otherwise, return nothing
                Case Else
                    Return Nothing
            End Select
        Catch ex As Exception
            Return Nothing
            Throw
        Finally
            If Not IsNothing(lobjNewNode) Then
                lobjNewNode = Nothing
            End If
            If Not IsNothing(lobjElementNode) Then
                lobjElementNode = Nothing
            End If           
            If Not IsNothing(lobjNewNode) Then
                lobjNewNode = Nothing
            End If
            If Not IsNothing(lobjElementNode) Then
                lobjElementNode = Nothing
            End If
        End Try
    End Function
End Module
