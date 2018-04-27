Imports System.Xml
Imports System.DirectoryServices
Imports System.Data

Public Class cMoneyCostLDAPSvc
    Dim STLogger As log4net.ILog
    Private Const cIService_SearchSSOUser As String = "<LDAP://~1/ou=~2,o=~3>;(~4);~5"
    Private _CommandSQL As String
    Private Const cDELIMITER As String = "|~^"
    Private _ReplaceQuote As Boolean = True
    Private mstrLDAPServer As String
    Private mstrLDAPAdminUID As String
    Private mstrLDAPAdminPwd As String
    Private mstrOrganization As String
    Private mstrExternalUser_ou As String
    Private mstrExternalUserGroups As String
    Private mstrInternalUser_ou As String
    Public Function GetUserDetailsByAttributes(ByVal astrXMLSearchCriteria As String) As String
        Dim objdsLDAP As DataSet
        Dim objXmlDoc As New XmlDocument
        Dim objRootNode As XmlElement
        Dim lobjXmlReturnDoc As New XmlDocument
        Dim larrMultiValAttr() As Object
        Dim lintCOunt As Integer
        Dim lintArrCount As Integer
        Static lintErrCount As Integer
        Dim lstrSearchFilter As String
        Dim lstrAttributes As String
        Dim arrDNList(0, 0) As Object
        Dim lstrOrgUnit As String = String.Empty
        Dim lobjOUAttrNode As XmlNode
        On Error GoTo GetUserDetailsByAttributesErr

        'Variables for Error Count used in Error Handler
        lintErrCount = 0

        'Load the Search Criteria XML
        objXmlDoc.LoadXml(astrXMLSearchCriteria)

        'Get the root node
        objRootNode = objXmlDoc.DocumentElement()

        'Check if root node is present
        If Not objRootNode Is Nothing Then

            'If root node contains no attributes - then set the default
            'LDAP attributes (uid,givenname,sn,mail,initials,gessouid), that
            'will be retrieved from the LDAP directory.
            If objRootNode.Attributes.Count = 0 Then
                lstrAttributes = "uid,givenname,sn,mail,initials,gessouid"

                'Else extract the LDAP attributes specified in the
                'attribute of the Root node, that will be retrieved from
                'the LDAP directory.
            Else
                If UCase(objRootNode.Attributes(0).Name) <> "OU" Then
                    lstrAttributes = objRootNode.Attributes(0).Value
                Else
                    lstrAttributes = "uid,givenname,sn,mail,initials,gessouid"
                End If

                'Get the Organization Unit attribute if present
                lobjOUAttrNode = objRootNode.GetAttributeNode("ou")
                If Not (lobjOUAttrNode Is Nothing) Then lstrOrgUnit = lobjOUAttrNode.Value
                lobjOUAttrNode = Nothing
            End If
        Else
            Err.Source = "BSLDAP.IService2.GetUserDetailsByAttributes"
            objXmlDoc = Nothing
            lobjXmlReturnDoc = Nothing
            objdsLDAP = Nothing
            GetUserDetailsByAttributes = "Input Parameter does not have a valid XML Format."
            Exit Function
        End If

        'If the ObjectClass is not present in the list of LDAP attributes to
        'be retrieved from the LDAP directory - then add it.
        'This LDAP attribute will help in determining to which group
        'the user belongs.
        If InStr(1, UCase(lstrAttributes), UCase("objectClass"), vbTextCompare) <= 0 Then
            lstrAttributes = lstrAttributes + ",objectClass"
        End If

        'Build the Search Filter for the LDAP Query using the Input Search Criteria XML
        lstrSearchFilter = GetLDAPFilterFromXML(astrXMLSearchCriteria)

        'Get LDAP Settings
        GetLDAPConnectionSettings()

        'If organisation unit is not supplied, use default SSO Businesses
        If lstrOrgUnit = "" Then lstrOrgUnit = mstrExternalUser_ou

        'Start Creating the return XML DOM Structure
        lobjXmlReturnDoc.LoadXml(CStr("<LDAPSEARCH><RECORDS></RECORDS></LDAPSEARCH>"))

        'The following check will enable this method to return quicker results even when you query on gessouid. If this
        'check is removed and ADO query is used even if the filter is gessouid then query response time will go up
        'to 1 minute as compared to 1 sec with this check.
        If objXmlDoc.GetElementsByTagName("LDAP_ATTRIB").Count >= 1 Then 'Make sure there is at least one search criteria
            'Now check to make sure that criteria is gessouid. Also, make sure there is no wild card search
            If (UCase(objXmlDoc.GetElementsByTagName("LDAP_ATTRIB").Item(0).Attributes.GetNamedItem("NAME").Value) = "GESSOUID" And InStr(objXmlDoc.GetElementsByTagName("LDAP_ATTRIB").Item(0).Value, "*") = 0 And objXmlDoc.GetElementsByTagName("LDAP_ATTRIB").Count = 1) Then
                ReDim arrDNList(0, 0)
                arrDNList(0, 0) = "gessouid=" & Trim(objXmlDoc.GetElementsByTagName("LDAP_ATTRIB").Item(0).InnerText) & ",ou=" & lstrOrgUnit & ",o=ge.com"

                GetLDAPDataXML(lobjXmlReturnDoc, lstrAttributes, arrDNList, "RECORDS", "RECORD")
            Else
                'Get SSO User Details using ADO
                objdsLDAP = DSExecute(mstrLDAPServer, lstrOrgUnit, mstrOrganization, lstrSearchFilter, "adspath")
                'If records found
                If objdsLDAP.Tables(0).Rows.Count > 0 Then
                    'arrDNList = objdsLDAP.Tables(0).Rows(0).ItemArray
                    'Dim strDetailIDList As List(Of String) = New List(Of String)
                    'Dim iRow As DataRow
                    'For Each iRow In objdsLDAP.Tables(0).Rows
                    '    strDetailIDList.Add(iRow(0).ToString)
                    'Next
                    'arrDNList = strDetailIDList.ToArray()
                    GetLDAPDataXML(lobjXmlReturnDoc, lstrAttributes, arrDNList, "RECORDS", "RECORD")    'Mid(objRs.Fields("adspath").Value, InStr(objRs.Fields("adspath").Value, "gessouid"))
                End If
            End If 'End GESSOUID check
            'Return the Final XML
            GetUserDetailsByAttributes = lobjXmlReturnDoc.InnerXml
        Else
            Err.Raise(-1, "BSLDAP.IService2.GetUserDetailsByAttributes", "Invalid Input XML")
            Return String.Empty
        End If 'End LDAP_ATTRIB check      


        objXmlDoc = Nothing
        lobjXmlReturnDoc = Nothing
        Exit Function

        'In Error Handler
GetUserDetailsByAttributesErr:
        GetUserDetailsByAttributes = ""
        Err.Source = "BSLDAP.IService2.GetUserDetailsByAttributes"

        Select Case Err.Number
            'Weird error "Specified table does not exist" which occurs
            'very infrequently even if the records exist in LDAP. SO re-run the code for 5 times
            Case -2147217865
                'Variables for Error Count used in Error Handler
                lintErrCount = lintErrCount + 1
                Do While lintErrCount < 5
                    Resume
                Loop

            Case Else
                objXmlDoc = Nothing
                lobjXmlReturnDoc = Nothing
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Function
    Function GetLDAPFilterFromXML(ByVal XMLDoc As String) As String
        Dim objXML As New XmlDocument
        Dim objRootNode As XmlElement = Nothing
        Dim iCounter As Integer
        Dim strResult As String = String.Empty
        Dim str2 As String
        Dim k As Integer
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            strResult = ""

            'Load the Input Search Criteria XML
            objXML.LoadXml(XMLDoc)

            'Get the root node
            objRootNode = objXML.DocumentElement

            'Check if root node sub nodes
            If objRootNode.HasChildNodes Then

                'For each sub nodes
                For iCounter = 0 To objRootNode.ChildNodes.Count - 1

                    'Check if node type is NODE_TEXT or NODE_CDATA_SECTION
                    If objRootNode.ChildNodes(iCounter).NodeType = 3 Or objRootNode.ChildNodes(iCounter).NodeType = 4 Then

                        'Check if node contains attributes
                        If (objRootNode.Attributes.Count > 0) Then
                            str2 = ""

                            'For each attribute
                            For k = 0 To objRootNode.Attributes.Count - 1
                                Select Case UCase(objRootNode.Attributes.Item(k).Name)

                                    'if attribute is NAME
                                    Case "NAME"

                                        'store the LDAP attribute name
                                        str2 = objRootNode.Attributes.Item(k).Value

                                        'if attribute is OPERATOR
                                    Case "OPERATOR"

                                        'Translate the value of the attribute to the appropriate Comparison Operator.
                                        'i.e EQ translates to =, GE to >=, LE to <= and APPROX to ~=
                                        'Add the operator to LDAP attribute name
                                        Select Case UCase(objRootNode.Attributes.Item(k).Value)

                                            Case "EQ"
                                                str2 = str2 & "="
                                            Case "GE"
                                                str2 = str2 & ">="
                                            Case "LE"
                                                str2 = str2 & "<="
                                            Case "APPROX"
                                                str2 = str2 & "~="
                                        End Select
                                End Select
                            Next

                            'Extract the LDAP attribute value from the Node and complete search filter statement
                            strResult = strResult & "(" & str2 & objRootNode.InnerText & ")"
                        End If

                        'Else if node type is NEITHER NODE_TEXT NOR NODE_CDATA_SECTION
                    Else

                        'Check if the node contains CDATA section and doesn't have child nodes
                        If InStr(1, objRootNode.ChildNodes(iCounter).OuterXml, "CDATA", vbTextCompare) And (Not (objRootNode.ChildNodes(iCounter).HasChildNodes)) Then

                            'Call the function recursively by passing the value of the node
                            strResult = strResult & GetLDAPFilterFromXML(objRootNode.ChildNodes(iCounter).Value)

                            'Else if the node does have child nodes
                        Else

                            'Call the function recursively by passing the node along with it's sub nodes
                            strResult = strResult & GetLDAPFilterFromXML(objRootNode.ChildNodes(iCounter).OuterXml)
                        End If
                    End If
                Next iCounter

                'if Node NAME is "OR"/"NOT"/"AND" convet it to the appropriate logical operator
                Select Case objRootNode.Name
                    Case "OR"
                        strResult = "(|" & strResult & ")"
                    Case "NOT"
                        strResult = "(!" & strResult & ")"
                    Case "AND"
                        strResult = "(&" & strResult & ")"
                End Select
            End If
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            Return strResult
            objXML = Nothing
            objRootNode = Nothing

        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Throw
        Finally
            If Not objXML Is Nothing Then
                objXML = Nothing
            End If
            If Not objRootNode Is Nothing Then
                objRootNode = Nothing
            End If
        End Try
    End Function
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
    Public Sub GetLDAPConnectionSettings()
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            mstrLDAPServer = GetConfigurationKey("LDAPServer")
            mstrLDAPAdminUID = GetConfigurationKey("LDAPAdminUID")
            mstrLDAPAdminPwd = GetConfigurationKey("LDAPAdminPwd")
            mstrOrganization = GetConfigurationKey("Organization")
            mstrExternalUser_ou = GetConfigurationKey("ExternalUser_ou")
            mstrExternalUserGroups = GetConfigurationKey("ExternalUserGroup")
            mstrInternalUser_ou = GetConfigurationKey("InternalUser_ou")
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
        End Try
    End Sub
    Private Function GetLDAPDataXML(ByRef aobjXMLDOM As XmlDocument, _
                                ByVal astrAttributes As String, _
                                ByRef arrDNList As Array, _
                                ByVal astrParentNodeName As String, _
                                ByVal astrElementName As String, _
    Optional ByVal astrElementValue As String = "") As String

        On Error GoTo GetLDAPDataXMLErr

        'Declarations related to ADSI
        Dim objLDAPCon As DirectoryEntry

        Dim objLDAPns As Object
        'objLDAPns = GetObject("LDAP")
        Dim objLDAPPropertyList As System.DirectoryServices.PropertyCollection
        Dim objLDAPPropertyEntry As PropertyValueCollection
        Dim objLDAPPropertyValue As Object
        Dim lPropertyValue As Object

        'Declarations related to XML
        Dim objChildNode As XmlElement
        Dim objCurrNode As XmlNode

        'Misc Declaration statements
        Dim arrAttributes() As String 'Array which holds the name of all attributes, whose values are to be fetched for a given user
        Dim ictrAttrib As Integer 'Total number of attributes desired
        Dim iCtrDNList As Integer 'This variable is used as a counter to process multiple DN's that may be sent to this method for data fetch.
        SetLog4Net()
        arrAttributes = Split(astrAttributes, ",")

        Do While iCtrDNList <= UBound(arrDNList)
            'Bind an ADSI object to the LDAP directory store
            'objLDAPns = GetObject("LDAP")

            'Supply full credentials along with the LDAP query to initiate a connection to a server
            'Note: External user details are stored under ou=SSO Businesses,o=ge.com
            On Error Resume Next
            objLDAPCon = New DirectoryEntry("LDAP://" & mstrLDAPServer & "/" & Mid(arrDNList(0, iCtrDNList), InStrRev(arrDNList(0, iCtrDNList), "/") + 1), mstrLDAPAdminUID, mstrLDAPAdminPwd, AuthenticationTypes.Anonymous)
            objLDAPPropertyList = objLDAPCon.Properties
            If (Err.Number = "-2147016656") Or (objLDAPPropertyList Is Nothing) Then
                GoTo ResumeNextDNRetrieval
            End If

            On Error GoTo GetLDAPDataXMLErr

            'Retrieve LDAP attributes from the underlying directory storage.
            'objLDAP.GetInfoEx Array(strAttributes), 0
            'objLDAPPropertyList.GetInfo()
            'Create a new node which contains the whole record
            objCurrNode = aobjXMLDOM.CreateNode(XmlNodeType.Element, astrElementName, "")
            objCurrNode.InnerText = astrElementValue
            'For each attribute
            For ictrAttrib = LBound(arrAttributes) To UBound(arrAttributes)

                'Get the PropEntry from LDAP
                'If the attribute value is null (e.g. initials) then ADSI throws an error,
                'so to avoid that - use Resume next
                On Error Resume Next
                objLDAPPropertyEntry = objLDAPPropertyList.Item(arrAttributes(ictrAttrib))
                If (Err.Number = "-2147463155") Or (objLDAPPropertyEntry Is Nothing) Then
                    objChildNode = aobjXMLDOM.CreateNode(XmlNodeType.Element, arrAttributes(ictrAttrib), "")
                    objChildNode.InnerText = vbNullString
                    objCurrNode.AppendChild(objChildNode)
                    objLDAPPropertyValue = Nothing
                    objLDAPPropertyEntry = Nothing
                    Err.Clear()
                    GoTo ResumePropertyRetrieval
                End If

                On Error GoTo GetLDAPDataXMLErr

                'For each value in PropEntry -get the data.
                'In case of MulitValued attributes (e.g. objectclass) - there is more than one value in Prop Entry.
                For Each lPropertyValue In objLDAPPropertyEntry
                    'Get current property attribute
                    objLDAPPropertyValue = lPropertyValue
                    If UCase(arrAttributes(ictrAttrib)) = "OBJECTCLASS" Then
                        Select Case UCase(objLDAPPropertyValue.ToString)
                            Case "TOP", "PERSON", "ORGANIZATIONALPERSON", "INETORGPERSON", "GESSOPERSON", "GESSOCONTACT", "GECEFSSOPERSON"
                            Case Else
                                'Create the element to return
                                objChildNode = aobjXMLDOM.CreateNode(XmlNodeType.Element, arrAttributes(ictrAttrib), "")

                                'In case of NULL atrribute value
                                If Not objLDAPPropertyValue Is Nothing Then
                                    objChildNode.InnerText = objLDAPPropertyValue.ToString
                                Else
                                    objChildNode.InnerText = vbNullString
                                End If
                                'append the child to return dom
                                objCurrNode.AppendChild(objChildNode)
                        End Select
                    Else
                        'Create the element to return
                        objChildNode = aobjXMLDOM.CreateNode(XmlNodeType.Element, arrAttributes(ictrAttrib), "")

                        'In case of NULL atrribute value
                        If Not objLDAPPropertyValue Is Nothing Then
                            objChildNode.InnerText = objLDAPPropertyValue.ToString
                        Else
                            objChildNode.InnerText = vbNullString
                        End If
                        'append the child to return dom
                        objCurrNode.AppendChild(objChildNode)
                    End If
                Next lPropertyValue 'loop for each attribute

                'Destroy the objects
                objLDAPPropertyValue = Nothing
                objLDAPPropertyEntry = Nothing

ResumePropertyRetrieval:

            Next ictrAttrib

            aobjXMLDOM.GetElementsByTagName(astrParentNodeName).Item(0).AppendChild(objCurrNode)

ResumeNextDNRetrieval:

            iCtrDNList = iCtrDNList + 1
        Loop
        Return String.Empty
GetLDAPDataXMLErr:
        Select Case Err.Number
            Case "-2147016656" 'User not found, dn does not exist on LDAP server.
                Resume ResumeNextDNRetrieval
            Case "-2147463155" 'The Active Directory property cannot be found in the cache. It occurs for only those properties whose value is not set for a given user. So we can safely resume to next statement
                objChildNode = aobjXMLDOM.CreateNode(XmlNodeType.Element, arrAttributes(ictrAttrib), "")
                objChildNode.Value = vbNullString
                objCurrNode.AppendChild(objChildNode)
                'Resume ResumePropertyRetrieval
            Case Else
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Function
    Public Function DSExecute(ByVal astrLDAPServer As String, ByVal astrOrgUnit As String, ByVal astrOrganization As String, ByVal astrSearchFilter As String, ByVal astrADSPath As String) As DataSet
        Dim dsLDAP As DataSet = Nothing
        Dim objoConnection As OleDb.OleDbConnection = Nothing
        Dim lobjSQLDA As OleDb.OleDbDataAdapter = Nothing
        Dim lobjSQL As String
        Try
            objoConnection = New OleDb.OleDbConnection
            objoConnection.ConnectionString = GetConfigurationKey("ConnectStrings_LDAP")
            lobjSQL = GetSQLStatement(astrLDAPServer, astrOrgUnit, astrOrganization, astrSearchFilter, astrADSPath)
            lobjSQLDA = New OleDb.OleDbDataAdapter(lobjSQL, objoConnection)

            lobjSQLDA.SelectCommand.CommandTimeout = 0
            lobjSQLDA.Fill(dsLDAP)
            Return dsLDAP
        Catch ex As Exception
            Return dsLDAP
        Finally
            If Not IsNothing(dsLDAP) Then
                dsLDAP.Dispose()
                dsLDAP = Nothing
            End If
            If Not IsNothing(objoConnection) Then
                objoConnection.Dispose()
                objoConnection = Nothing
            End If
            If Not IsNothing(lobjSQLDA) Then
                lobjSQLDA.Dispose()
                lobjSQLDA = Nothing
            End If
        End Try
    End Function
    Private Function GetSQLStatement(ByVal ParamArray aobjSQLParams As Object()) As String

        Dim lstrSQL As String
        Dim liPos As Integer
        Dim liLoopCtr As Integer
        Dim liUBound As Integer
        Try

            lstrSQL = cIService_SearchSSOUser
            'If there are no substitution parms, we're done
            liPos = InStr(1, lstrSQL, cDELIMITER)

            If liPos = 0 Then
                GetSQLStatement = lstrSQL
                Exit Function
            End If

            If IsArray(aobjSQLParams) Then
                If UBound(aobjSQLParams) = -1 Then
                    Err.Raise(cINVALID_PARMS)
                End If
            End If

            'Loop through each of the param
            liUBound = UBound(aobjSQLParams)
            For liLoopCtr = liUBound + 1 To 1 Step -1

                'First, replace any apostrophes (') in the parm with two ('') for SQL Server.
                If _ReplaceQuote Then aobjSQLParams(liLoopCtr - 1) = Replace(aobjSQLParams(liLoopCtr - 1), "'", "''")

                'Then Insert the parm into the SQL statement.
                lstrSQL = Replace(lstrSQL, cDELIMITER & CStr(liLoopCtr), aobjSQLParams(liLoopCtr - 1))
            Next liLoopCtr

            'REPLACE ANY STRINGS THAT CAME UP 'NULL' WITH NULL
            lstrSQL = Replace(lstrSQL, "'NULL'", "NULL")

            'Make sure all substitution parms have been swapped out
            liPos = InStr(1, lstrSQL, cDELIMITER)
            If liPos > 0 Then Err.Raise(cINVALID_PARMS)

        Catch ex As Exception
            Throw
        End Try
        'Return Completed SQL statement
        GetSQLStatement = lstrSQL

        Exit Function
    End Function
End Class
