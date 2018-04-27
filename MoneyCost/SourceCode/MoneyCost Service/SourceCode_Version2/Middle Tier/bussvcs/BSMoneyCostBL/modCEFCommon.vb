Imports Microsoft.Win32
Imports System.text
Imports System.Xml
Imports System.Reflection
Public Module modCEFCommon
    Private Const cMODULE_NAME As String = "modCEFCommon"
    '================================================================
    '=== Error Constants ============================================
    Public Const cERR_ROOT_TAG As String = "ERROR_DETAILS"
    Public Const cERR_NBR As String = "ERROR_NUMBER"
    Public Const cERR_DESC As String = "ERROR_DESCRIPTION"
    Public Const cERR_SRC As String = "ERROR_SOURCE"
    Public Const cERR_ShowUser As String = "ERROR_SHOW_USER"
    '================================================================
    

    '=== Function Convert DS To XML ================================
    Public Function DsToXML(ByVal astrDS As DataSet) As String
        Dim iCount, iRow, iCol As Integer
        Dim vXML As String

        vXML = ""
        Try
            If astrDS.Tables.Count > 0 Then
                vXML = "<" & astrDS.DataSetName & ">"
                For iCount = 0 To astrDS.Tables.Count - 1  '' Loop for Table
                    For iRow = 0 To astrDS.Tables(iCount).Rows.Count - 1 'Loop for Record
                        vXML = vXML & "<" & astrDS.Tables(iCount).TableName & ">"
                        For iCol = 0 To astrDS.Tables(iCount).Columns.Count - 1  ''Loop for each column
                            vXML = vXML & "<" & astrDS.Tables(iCount).Columns(iCol).ColumnName & ">"
                            vXML = vXML & astrDS.Tables(iCount).Rows(iRow).Item(iCol)
                            vXML = vXML & "</" & astrDS.Tables(iCount).Columns(iCol).ColumnName & ">"
                        Next
                        vXML = vXML & "</" & astrDS.Tables(iCount).TableName & ">"
                    Next
                Next
                vXML = vXML & "</" & astrDS.DataSetName & ">"
            End If

            vXML = Replace(vXML, "&", "&amp;")
            Return vXML
        Catch ex As Exception
            Throw
        End Try        
    End Function


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


        Dim lobjErrXMLDOM As New Xml.XmlDataDocument

        Try
            'Load output XML
            lobjErrXMLDOM.LoadXml("<" & cERR_ROOT_TAG & "/>")

            'Add the error number tag to the output XML
            AddXMLElement(lobjErrXMLDOM, _
                            lobjErrXMLDOM.DocumentElement, _
                            cERR_NBR, _
                            astrErrNbr)

            'Add the error source tag to the output XML
            AddXMLElement(lobjErrXMLDOM, _
                            lobjErrXMLDOM.DocumentElement, _
                            cERR_SRC, _
                            astrErrSource)

            'Add the error description tag to the output XML
            AddXMLElement(lobjErrXMLDOM, _
                            lobjErrXMLDOM.DocumentElement, _
                            cERR_DESC, _
                            astrErrDesc)

            'Return the error XML
            Return lobjErrXMLDOM.InnerXml

            lobjErrXMLDOM = Nothing

        Catch ex As Exception
            Return vbNullString
            lobjErrXMLDOM = Nothing
            Throw
        Finally
            If Not IsNothing(lobjErrXMLDOM) Then
                lobjErrXMLDOM = Nothing
            End If
        End Try


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


        Dim lobjErrXMLDOM As New Xml.XmlDataDocument

        Try
            'Load output XML
            lobjErrXMLDOM.LoadXml("<" & cERR_ROOT_TAG & "/>")

            'Add the error number tag to the output XML
            AddXMLElement(lobjErrXMLDOM, _
                            lobjErrXMLDOM.DocumentElement, _
                            cERR_NBR, _
                            astrErrNbr)

            'Add the error source tag to the output XML
            AddXMLElement(lobjErrXMLDOM, _
                            lobjErrXMLDOM.DocumentElement, _
                            cERR_SRC, _
                            astrErrSource)

            'Add the error description tag to the output XML
            AddXMLElement(lobjErrXMLDOM, _
                            lobjErrXMLDOM.DocumentElement, _
                            cERR_DESC, _
                            astrErrDesc)

            'Add the error description for user tag to the output XML
            AddXMLElement(lobjErrXMLDOM, _
                            lobjErrXMLDOM.DocumentElement, _
                            cERR_ShowUser, _
                            astrErrorShowUser)

            'Return the error XML
            Return lobjErrXMLDOM.InnerXml

            lobjErrXMLDOM = Nothing
        Catch ex As Exception
            Return vbNullString
            lobjErrXMLDOM = Nothing
            Throw
        Finally
            If Not IsNothing(lobjErrXMLDOM) Then
                lobjErrXMLDOM = Nothing
            End If
        End Try

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

        Dim regKey As RegistryKey
        Try
            regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\FacilitySettings\1000\FilePath", True)
            Return regKey.GetValue(astrRegFullKeyName)

        Catch ex As Exception
            Return (vbNullString)
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
    Public Function IsXMLElementPresent(ByVal aobjXMLDOM As Xml.XmlNode, _
                                        ByVal astrElementXPath As String) As Boolean

        Dim lobjElement As Xml.XmlElement = Nothing
        Dim lbElementPresent As Boolean

        Try
            lbElementPresent = False
            lobjElement = aobjXMLDOM.SelectSingleNode(astrElementXPath)
            If Not (lobjElement Is Nothing) Then
                lbElementPresent = True
            End If

            Return lbElementPresent

            lobjElement = Nothing
        Catch ex As Exception
            Return False
            Throw
        Finally
            If Not IsNothing(lobjElement) Then
                lobjElement = Nothing
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
    Public Function AddXMLElement(ByRef aobjXMLDOM As Xml.XmlDataDocument, _
                                    ByVal aobjParentNode As Xml.XmlNode, _
                                    ByVal astrElementName As String, _
                                    ByVal astrElementContent As String) As Xml.XmlNode

        Dim lobjNewNode As Xml.XmlNode = Nothing
        Dim lobjElementNode As Xml.XmlElement = Nothing

        Try
            'Check if the Parent node type is NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT,
            'NODE_ENTITY_REFERENCE , NODE_ELEMENT
            Select Case aobjParentNode.NodeType
                Case "NODE_DOCUMENT", "NODE_DOCUMENT_FRAGMENT", _
                    "NODE_ENTITY_REFERENCE", "NODE_ELEMENT"

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

            lobjNewNode = Nothing
            lobjElementNode = Nothing

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
        End Try

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
    Public Function AddXMLElementAttribute(ByRef aobjXMLDOM As Xml.XmlDataDocument, _
                                            ByVal aobjElementNode As Xml.XmlNode, _
                                            ByVal astrAttributeName As String, _
                                            ByVal astrAttributeContent As String) As Xml.XmlNode

        Dim lobjAttr As Xml.XmlAttribute = Nothing

        Try
            'Check if the element node type is NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT,
            'NODE_ENTITY_REFERENCE , NODE_ELEMENT
            Select Case aobjElementNode.NodeType
                Case "NODE_DOCUMENT", "NODE_DOCUMENT_FRAGMENT", _
                    "NODE_ENTITY_REFERENCE", "NODE_ELEMENT"

                    'Add new attribute to the element node of the XML DOM
                    lobjAttr = aobjXMLDOM.CreateAttribute(astrAttributeName)
                    lobjAttr.InnerText = astrAttributeContent
                    aobjElementNode.Attributes.SetNamedItem(lobjAttr)
                    lobjAttr = Nothing

                    'Return the new XML DOM
                    Return aobjElementNode

                    'Otherwise, return nothing
                Case Else
                    Return Nothing
            End Select

            lobjAttr = Nothing
        Catch ex As Exception
            Return Nothing
            lobjAttr = Nothing
            Throw
        Finally
            If Not IsNothing(lobjAttr) Then
                lobjAttr = Nothing
            End If
        End Try

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
    Optional ByVal abFileOverwrite As Boolean = False, Optional ByVal astrClassName As String = "", Optional ByVal astrFunctionName As String = "") As Boolean
        Dim lstrErrorMessage As New StringBuilder
        Dim mobjLogger2 As log4net.ILog = Nothing

        Try

            log4net.Config.XmlConfigurator.ConfigureAndWatch(New System.IO.FileInfo(GetConfigurationKey("LogForNetDebugFilePath")))
            'log4net.Config.XmlConfigurator.ConfigureAndWatch(New System.IO.FileInfo("c:\\LogFiles\\log4net_DLL.config"))
            mobjLogger2 = log4net.LogManager.GetLogger("MoneyCost")

            'Genarate Error Message String
            If astrClassName <> "" Then lstrErrorMessage.Append("ClassName: " + astrClassName)
            If astrFunctionName <> "" Then lstrErrorMessage.Append("FunctionName: " + astrFunctionName + "()")

            'lstrErrorMessage.Append("ErrorMessageDescription: ")
            lstrErrorMessage.Append("""" & astrData & """")

            mobjLogger2.Error(lstrErrorMessage.ToString)


            'LogError(lstrErrorMessage.ToString)
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(mobjLogger2) Then
                mobjLogger2 = Nothing
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
            'ConstantsFilePath = "D:\StarTeam\MoneyCost\QA-ConstantFile" + "\Constant.XML"
            lDocXmlFile.Load(ConstantsFilePath)

            Return lDocXmlFile.GetElementsByTagName(astrKey).Item(0).InnerText

        Catch ex As Exception
            Throw ex
        Finally
            lDocXmlFile = Nothing
        End Try
    End Function
End Module
