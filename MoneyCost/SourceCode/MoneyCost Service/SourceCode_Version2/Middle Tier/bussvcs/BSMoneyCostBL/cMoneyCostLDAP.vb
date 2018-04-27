Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.data.Linq
Imports System.Xml.Linq
Imports System.DirectoryServices
Imports System.DirectoryServices.Protocols
Imports System.Net
Imports System.Collections.Generic
Imports System.Xml
Imports System.Configuration
Public Class cMoneyCostLDAP
    Dim STLogger As log4net.ILog
    Dim mstrLDAPServer As String = String.Empty
    Dim mstrLDAPUserId As String = String.Empty
    Dim mstrLDAPPassword As String = String.Empty
    Dim mstrOrganization As String = String.Empty
    Dim mstrXMLResult As String = String.Empty
    Public Function GetUserDetailsByAttributes(ByVal astrSearchId As String) As String
        Dim dicLDAP As Dictionary(Of String, String) = Nothing
        Dim tsLDAP As TimeSpan = Nothing
        Dim credentialLDAP As NetworkCredential = Nothing
        Dim identifierLDAP As LdapDirectoryIdentifier = Nothing
        Dim connectionLDAP As LdapConnection = Nothing
        Dim searchRequest As SearchRequest = Nothing
        Dim srLDAP As SearchResponse = Nothing
        Dim srEntry As SearchResultEntry = Nothing
        Dim srAttributesLDAP As SearchResultAttributeCollection = Nothing
        Dim daLDAP As DirectoryAttribute = Nothing
        Dim intI As Integer
        Dim intJ As Integer
        Dim outElm As New XElement("LDAP")
        Dim objXMLDOM As New XmlDocument
        Dim objCurrNode As XmlNode = Nothing
        Dim objChildNode As XmlElement = Nothing
        Try
            SetLog4Net()
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            GetLDAPConnectionSettings()
            dicLDAP = New Dictionary(Of String, String)()
            tsLDAP = New TimeSpan(10, 30, 60)
            STLogger.Debug("Enter: " & "Object of NetworkCredential")
            credentialLDAP = New NetworkCredential(mstrLDAPUserId, mstrLDAPPassword)
            STLogger.Debug("End: " & "Object of NetworkCredential")
            STLogger.Debug("Enter: " & "Object of LdapDirectoryIdentifier")
            identifierLDAP = New LdapDirectoryIdentifier(mstrLDAPServer)
            STLogger.Debug("End: " & "Object of LdapDirectoryIdentifier")
            STLogger.Debug("Enter: " & "Object of LdapConnection")
            connectionLDAP = New LdapConnection(identifierLDAP, credentialLDAP, AuthType.Basic)
            STLogger.Debug("End: " & "Object of LdapConnection")
            connectionLDAP.Timeout = tsLDAP
            STLogger.Debug("Enter: " & "Object of SearchRequest")
            searchRequest = New SearchRequest(mstrOrganization, astrSearchId, System.DirectoryServices.Protocols.SearchScope.Subtree)

            searchRequest.TimeLimit = tsLDAP
            STLogger.Debug("End: " & "Object of SearchRequest")
            Dim intLDAPCount As Integer = 0
            STLogger.Debug("Enter: " & "Object of connectionLDAP.SendRequest")
            srLDAP = CType(connectionLDAP.SendRequest(searchRequest, tsLDAP), SearchResponse)
            STLogger.Debug("End: " & "Object of connectionLDAP.SendRequest")
            STLogger.Debug("Enter: " & "Loop for LDAP entries/Attributes and put into dictionary")
            For Each srEntry In srLDAP.Entries
                srAttributesLDAP = srEntry.Attributes
                For Each daLDAP In srAttributesLDAP.Values
                    intLDAPCount = daLDAP.Count
                    For intI = 0 To intLDAPCount - 1
                        For intJ = 0 To dicLDAP.Count
                            If Not dicLDAP.ContainsKey(daLDAP.Name) Then
                                dicLDAP.Add(daLDAP.Name, daLDAP(intI).ToString())
                            End If
                        Next
                    Next
                Next
            Next
            STLogger.Debug("Enter: " & "Loop for LDAP entries/Attributes and put into dictionary")

            Dim keys As Dictionary(Of String, String).KeyCollection = dicLDAP.Keys



            STLogger.Debug("Enter: " & "Convert Dictionary into XML")
            objCurrNode = objXMLDOM.CreateNode(XmlNodeType.Element, "RECORDS", "")
            For Each key As String In keys
                objChildNode = objXMLDOM.CreateNode(XmlNodeType.Element, key, "")
                If dicLDAP(key) <> String.Empty Then
                    objChildNode.InnerText = dicLDAP(key)
                Else
                    objChildNode.InnerText = vbNullString
                End If
                objCurrNode.AppendChild(objChildNode)
            Next
            objXMLDOM.AppendChild(objCurrNode)
            STLogger.Debug("End: " & "Convert Dictionary into XML")
            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
            Return objXMLDOM.InnerXml()
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
            Return String.Empty
        Finally
            If Not IsNothing(dicLDAP) Then
                dicLDAP = Nothing
            End If
            If Not IsNothing(tsLDAP) Then
                tsLDAP = Nothing
            End If
            If Not IsNothing(credentialLDAP) Then
                credentialLDAP = Nothing
            End If

            If Not IsNothing(identifierLDAP) Then
                identifierLDAP = Nothing
            End If
            If Not IsNothing(connectionLDAP) Then
                connectionLDAP = Nothing
            End If
            If Not IsNothing(searchRequest) Then
                searchRequest = Nothing
            End If
            If Not IsNothing(srLDAP) Then
                srLDAP = Nothing
            End If
            If Not IsNothing(srEntry) Then
                srEntry = Nothing
            End If
            If Not IsNothing(srAttributesLDAP) Then
                srAttributesLDAP = Nothing
            End If
            If Not IsNothing(daLDAP) Then
                daLDAP = Nothing
            End If
            If Not IsNothing(outElm) Then
                outElm = Nothing
            End If
            If Not IsNothing(objXMLDOM) Then
                objXMLDOM = Nothing
            End If
            If Not IsNothing(objCurrNode) Then
                objCurrNode = Nothing
            End If
            If Not IsNothing(objChildNode) Then
                objChildNode = Nothing
            End If
        End Try
    End Function

    Public Sub GetLDAPConnectionSettings()
        SetLog4Net()
        Try
            STLogger.Debug("Enter: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            mstrLDAPServer = GetConfigurationKey("LDAPServer")
            mstrLDAPUserId = GetConfigurationKey("LDAPAdminUID")
            mstrLDAPPassword = GetConfigurationKey("LDAPAdminPwd")
            mstrOrganization = GetConfigurationKey("Organization")

            STLogger.Debug("Exit: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
        End Try
    End Sub
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
End Class


'Namespace LDAPSearchWebService
'{
'    /// <summary>
'    /// Summary description for Service1
'    /// </summary>
'    [WebService(Namespace = "http://tempuri.org/")]
'    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
'    [ToolboxItem(false)]
'    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
'    // [System.Web.Script.Services.ScriptService]
'    public class LDAPService : System.Web.Services.WebService
'    {
'        String strLDAPServer = String.Empty;
'        String strLDAPUserId = String.Empty;
'        String strLDAPPassword = String.Empty;
'        String strOrganization = String.Empty;
'        String strXMLResult = String.Empty;
'        [WebMethod]
'        public string GetUserDetailsByAttributes(string strSearchId)
'        {
'            // Fetch the values from Web.Config
'            strLDAPServer = ConfigurationManager.AppSettings["LDAPServer"].ToString();
'            strLDAPUserId = ConfigurationManager.AppSettings["LDAPUserId"].ToString();
'            strLDAPPassword = ConfigurationManager.AppSettings["LDAPPassword"].ToString();
'            strOrganization = ConfigurationManager.AppSettings["Organization"].ToString();
'            Dictionary<String, String> dicLDAP = new Dictionary<string, string>();

'            // LDAP  Connection details start here
'            TimeSpan tsLDAP = new TimeSpan(10, 30, 60);
'            NetworkCredential credentialLDAP = new NetworkCredential(strLDAPUserId, strLDAPPassword);
'            LdapDirectoryIdentifier identifierLDAP = new LdapDirectoryIdentifier(strLDAPServer);
'            LdapConnection connectionLDAP = new LdapConnection(identifierLDAP, credentialLDAP, AuthType.Basic);
'            connectionLDAP.Timeout = tsLDAP;
'            // Note strSearchId should like  "(uid=" + txtTest.Text + ")"
'            SearchRequest searchRequest = new SearchRequest(strOrganization, strSearchId, System.DirectoryServices.Protocols.SearchScope.Subtree);

'            try
'            {
'                searchRequest.TimeLimit = tsLDAP;
'                int intLDAPCount = 0;
'                SearchResponse srLDAP = (SearchResponse)connectionLDAP.SendRequest(searchRequest, tsLDAP);
'                foreach (SearchResultEntry srEntry in srLDAP.Entries)
'                {
'                    SearchResultAttributeCollection srAttributesLDAP = srEntry.Attributes;
'                    foreach (DirectoryAttribute daLDAP in srAttributesLDAP.Values)
'                    {
'                        intLDAPCount = daLDAP.Count;
'                        for (int i = 0; i < intLDAPCount; i++)
'                        {
'                            // Add the attributes and it's value in Dictionary
'                            for (int j = 0; j <= dicLDAP.Count; j++)
'                            {
'                                if (!dicLDAP.ContainsKey(daLDAP.Name))
'                                {
'                                    dicLDAP.Add(daLDAP.Name, daLDAP[i].ToString());
'                                }
'                            }
'                        }
'                    }
'                }

'                var result = new XDocument(new XElement("LDAP", dicLDAP.Select(i => new XElement(i.Key, i.Value))));
'                strXMLResult = result.ToString();
'            }
'            catch (Exception ex)
'            {
'                return ex.Message.ToString();
'            }
'            finally
'            {
'                connectionLDAP.Dispose();

'            }
'            return strXMLResult;
'        }
'    }
'}


