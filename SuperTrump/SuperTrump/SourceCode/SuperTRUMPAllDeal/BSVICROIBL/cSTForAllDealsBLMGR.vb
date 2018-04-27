Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports BSVICROIEntity.BSVICROIEntity
Imports BSVICROIBL.MCCommon
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System
Imports System.Xml
Imports Microsoft.Win32
Imports BSVICROIEntity
Imports System.Net.Mail
Imports System.EnterpriseServices
Imports SuperTRUMPCommon


Namespace BSVICROIBL
    Public Interface ISTForAllDealsBLMGR
        Function LoadDWData() As cSTForAllDealsEntity
        Function UpdateDataVICDB() As cSTForAllDealsEntity
    End Interface
    Public Class cSTForAllDealsBLMGR
        Dim HTMLReport As String = String.Empty
        Dim iMatchCount As Integer = 0


        ''' <summary>
        ''' Used to export DAT file data into database *****
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ExportDATFileDataInDatabase() As String
            Dim STLogger As log4net.ILog
            Dim strMileStone As String = "1"
            Dim objCommon As BSVICROICommon = New BSVICROICommon()
            STLogger = objCommon.SetLog4Net()
            Dim arrFiles As String()
            Dim strFilePath As String = String.Empty

            STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method: " + MethodInfo.GetCurrentMethod.Name())
            Try
                arrFiles = objCommon.ReadConfigurationFileValue("DATControlFileList").Split(",")
                For i As Integer = 0 To arrFiles.Length - 1
                    strFilePath = objCommon.ReadConfigurationFileValue("DATControlFilePath") + arrFiles(i).ToString()
                    If File.Exists(strFilePath) Then
                        BulkCopyInsertInDB(strFilePath)
                        STLogger.Debug("Bulk Copy Insert in Database for Control file:- " + strFilePath)
                        strFilePath = String.Empty
                    End If
                Next

                Dim objCDataCls As BSVICROIDAL.cDataClass
                objCDataCls = New BSVICROIDAL.cDataClass

                Dim iFlag As Integer = objCDataCls.InsertDataInMainStreamDetailTable(objCommon.ReadConfigurationFileValue("ConnectionstringOLEDBProvider"))
                STLogger.Debug("END:  " + DateTime.Now + " Tracing process within method: " + MethodInfo.GetCurrentMethod.Name())
                Return iFlag.ToString()
            Catch ex As Exception
                STLogger.Error("MileStone:- " & strMileStone & " Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
                Return "ERROR"
            End Try

        End Function


        Public Function MapLocation(ByVal strLocation As String, ByVal strUName As String, ByVal strPassword As String) As Integer
            Dim objFtp As New FTPEntity
            objFtp.FTPLocation = strLocation
            objFtp.FTPUser = strUName
            objFtp.FTPPassword = strPassword
            If Convert.ToBoolean(IsFTPLocationMapped(objFtp)) Then
                Return 1
            End If
            Return 0
        End Function

        Private Function IsFTPLocationMapped(ByVal objFtpEntity As FTPEntity) As String
            Dim lstrUNCPath As String
            Dim STLogger As log4net.ILog
            Dim objCommon As BSVICROICommon = New BSVICROICommon()
            STLogger = objCommon.SetLog4Net()
            Try
                lstrUNCPath = objFtpEntity.FTPLocation
                If Directory.Exists(lstrUNCPath) Then
                    Return "True"

                Else
                    If Convert.ToBoolean(MapDrive(objFtpEntity)) Then
                        Return "True"

                    Else
                        Return "False"

                    End If
                End If
            Catch ex As Exception
                STLogger.Error(" Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
                Return "False"
                Throw
            Finally
                If Not IsNothing(objFtpEntity) Then
                    objFtpEntity = Nothing
                End If
                objCommon = Nothing
                STLogger = Nothing
            End Try

        End Function

        Public Function UnMapDrive(ByVal DriveLetter As String) As Boolean
            Dim rc As Integer
            rc = WNetCancelConnection2(DriveLetter & ":", 0, 1)

            If rc = 0 Or rc = 2250 Then
                Return True
            Else
                MsgBox("rc= " & rc & vbLf & "Drive letter " & DriveLetter & " NOT DISCONNECTED!")
                Return False
            End If

        End Function

        Public Function MapDrive(ByVal objFtpEntity As FTPEntity) As String
            Dim nr As NETRESOURCE
            Dim STLogger As log4net.ILog
            Dim objCommon As BSVICROICommon = New BSVICROICommon()
            STLogger = objCommon.SetLog4Net()
            Dim strUsername As String
            Dim strPassword As String
            Try
                nr = New NETRESOURCE
                nr.lpRemoteName = objFtpEntity.FTPLocation
                nr.lpLocalName = ""
                strUsername = objFtpEntity.FTPUser
                strPassword = objFtpEntity.FTPPassword
                nr.dwType = RESOURCETYPE_DISK

                Dim result As Integer
                result = WNetAddConnection2(nr, strPassword, strUsername, 0)

                If result = 0 Then
                    Return "True"

                Else
                    Return "False"

                End If
            Catch ex As Exception
                STLogger.Error(" Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
                Return "False"
                Throw
            Finally
                objCommon = Nothing
                objFtpEntity = Nothing
                STLogger = Nothing
            End Try

        End Function
        Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" _
   (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, _
   ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer

        Public Declare Function WNetCancelConnection2 Lib "mpr" Alias "WNetCancelConnection2A" _
  (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer



        Public Const RESOURCETYPE_DISK As Long = &H1

        Public Structure NETRESOURCE
            Public dwScope As Integer
            Public dwType As Integer
            Public dwDisplayType As Integer
            Public dwUsage As Integer
            Public lpLocalName As String
            Public lpRemoteName As String
            Public lpComment As String
            Public lpProvider As String
        End Structure

        Public Function TestServiceinBL() As String
            Dim objDL As BSVICROIDAL.cDataClass = New BSVICROIDAL.cDataClass
            Return "BL Successfully - " + objDL.TestServiceinDAL()
        End Function
        ''' <summary>
        ''' Used to send the notification mail
        ''' </summary>
        ''' <param name="strBody"></param>
        ''' <param name="boolAttachmentRequired"></param>
        ''' <remarks></remarks>
        Public Sub SendNotificationByMail(ByVal strBody As String, ByVal boolAttachmentRequired As Boolean)
            Dim objCommon As BSVICROICommon = New BSVICROICommon()
            Dim STLogger As log4net.ILog
            STLogger = objCommon.SetLog4Net()
            Dim strMileStone As String = "5"
            Dim lstrMailBody As String
            Dim objMail As System.Net.Mail.MailMessage
            Dim lobjSMTP As New System.Net.Mail.SmtpClient(objCommon.ReadConfigurationFileValue("SmtpClient"))
            Try
                strMileStone = "5.1"
                lstrMailBody = "Dear SuperTRUMP For AllDeal User," & vbCrLf & vbCrLf
                lstrMailBody = lstrMailBody & strBody & vbCrLf & vbCrLf
                lstrMailBody = lstrMailBody & "Note: This is an auto-generated email from  SuperTRUMPAllDeal Service. Please do not reply to this email."
                objMail = New MailMessage()
                strMileStone = "5.2"
                Dim strMailAToArr() As String = objCommon.ReadConfigurationFileValue("MailTo").ToString().Split(";")
                strMileStone = "5.3"
                For Each Item As String In strMailAToArr
                    objMail.To.Add(Item)
                Next
                strMileStone = "5.4"
                objMail.Body = lstrMailBody
                objMail.Subject = objCommon.ReadConfigurationFileValue("EmailMessageTitle").ToString()
                strMileStone = "5.5"
                objMail.From = New Net.Mail.MailAddress(objCommon.ReadConfigurationFileValue("MailFrom").ToString())

                If (String.Compare(objCommon.ReadConfigurationFileValue("AttachmentRequiredinMailFlag"), "True", True) = 0 And boolAttachmentRequired = True) Then
                    File.Copy(objCommon.ReadConfigurationFileValue("LogFilePath"), objCommon.ReadConfigurationFileValue("TempLogFilePath"), True)
                    Dim objAttach As Attachment = New Attachment(objCommon.ReadConfigurationFileValue("TempLogFilePath"))
                    objMail.Attachments.Add(objAttach)
                End If
                strMileStone = "5.6"
                lobjSMTP.Send(objMail)
                strMileStone = "5.7"
                STLogger.Debug("Mail Send Successfully")
            Catch ex As Exception
                STLogger.Error("MileStone:- " & strMileStone & " Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
            Finally
                If Not IsNothing(objMail) Then
                    objMail = Nothing
                End If
                If Not IsNothing(lobjSMTP) Then
                    lobjSMTP = Nothing
                End If
                STLogger = Nothing
                objCommon = Nothing
            End Try
        End Sub


        ''' <summary>
        ''' Used to Get Input Files from FTP *****
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFTPInputFiles() As String
            Dim ftpClient As cFTPClient
            Dim arrFiles As String()
            Dim arrFilesAll As String()
            Dim dirInfo As DirectoryInfo
            Dim arrFileInfo As FileInfo()
            Dim strDestinationPath As String
            Dim objCommon As BSVICROICommon = New BSVICROICommon()
            Dim STLogger As log4net.ILog
            STLogger = objCommon.SetLog4Net()
            STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method: " + MethodInfo.GetCurrentMethod.Name())
            Try
                ftpClient = New cFTPClient
                strDestinationPath = objCommon.ReadConfigurationFileValue("DATDestinationPath")
                STLogger.Debug("Read DATDestinationPath")
                ftpClient.RemoteHostFTPServer = objCommon.ReadConfigurationFileValue("DATRemoteHostFTPServer")
                STLogger.Debug("Read DATRemoteHostFTPServer")
                ftpClient.RemotePort = objCommon.ReadConfigurationFileValue("DATRemotePort")
                STLogger.Debug("Read DATRemotePort")
                ftpClient.RemoteUser = objCommon.ReadConfigurationFileValue("DATRemoteUser")
                STLogger.Debug("Read DATRemoteUser")
                ftpClient.RemotePassword = objCommon.ReadConfigurationFileValue("DATRemotePassword")
                STLogger.Debug("Read DATRemotePassword")
                ftpClient.ChangeDirectory(objCommon.ReadConfigurationFileValue("DATSourcePath"))
                STLogger.Debug("Read DATSourcePath")
                arrFilesAll = ftpClient.GetFileList("*.dat")
                arrFiles = objCommon.ReadConfigurationFileValue("DATFileList").Split(",")
                STLogger.Debug("Get All DAT file list")

                'If Directory.Exists(strDestinationPath) Then
                '    Directory.Delete(strDestinationPath, True)
                '    STLogger.Debug("Delete Directory Path :- " + strDestinationPath)
                'End If

                If Not Directory.Exists(strDestinationPath) Then
                    Directory.CreateDirectory(strDestinationPath)
                    STLogger.Debug("Create Directory Path :- " + strDestinationPath)
                End If

                STLogger.Debug("ArrFile Length -: " + arrFiles.Length.ToString())
                For i As Integer = 0 To arrFiles.Length - 1
                    If CheckFileExistance(arrFilesAll, arrFiles(i)) Then
                        ftpClient.DownloadFile(arrFiles(i), strDestinationPath & arrFiles(i).Substring(0, arrFiles(i).Length))
                        STLogger.Debug("File No:- " + i.ToString() + "   Source File Name :- " + arrFiles(i).ToString() + " Download")
                    Else
                        STLogger.Debug("File No:- " + i.ToString() + "   Source File Name :- " + arrFiles(i).ToString() + " Not Exist")
                    End If
                Next
                STLogger.Debug("Download all files")
                STLogger.Debug("END:  " + DateTime.Now + " Tracing process within method: " + MethodInfo.GetCurrentMethod.Name())
                Return "done"
            Catch ex As Exception
                STLogger.Error(" Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
                Return "Error"
            Finally
                If Not IsNothing(ftpClient) Then
                    If ftpClient.LoggedIn Then
                        ftpClient.CloseConnection()
                    End If
                    ftpClient = Nothing
                End If
                If Not IsNothing(arrFileInfo) Then
                    arrFileInfo = Nothing
                End If
                If Not IsNothing(dirInfo) Then
                    dirInfo = Nothing
                End If
                If Not IsNothing(arrFiles) Then
                    arrFiles = Nothing
                End If
                STLogger = Nothing
                objCommon = Nothing
            End Try
        End Function
        ''' <summary>
        ''' Used to check the existance of file at desired location
        ''' </summary>
        ''' <param name="arrFilesAll"></param>
        ''' <param name="strFileName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CheckFileExistance(ByVal arrFilesAll As String(), ByVal strFileName As String) As Boolean
            Try

                For index As Integer = 0 To arrFilesAll.Length - 2
                    If String.Compare(strFileName.Trim(), arrFilesAll(index).Trim(), True) = 0 Then
                        Return True
                    End If
                Next
                Return False
            Catch ex As Exception
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Used to check the Exceptions in XML comming from SuperTRUMP web services
        ''' </summary>
        ''' <param name="aobjSTXMLDoc"></param>
        ''' <remarks></remarks>
        Private Sub CheckSuperTrumpExceptions(ByRef aobjSTXMLDoc As Xml.XmlDocument)
            Dim objCommon As BSVICROICommon = New BSVICROICommon()
            Dim STLogger As log4net.ILog
            STLogger = objCommon.SetLog4Net()
            STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method: " + MethodInfo.GetCurrentMethod.Name())
            Try
                If Not IsNothing(aobjSTXMLDoc.DocumentElement.SelectSingleNode("//PRM_INFO/PRM_FILE/AD_HOC_QUERY/SuperTRUMP/Transaction/Exceptions")) Then
                    Dim ndLstException As XmlNodeList = aobjSTXMLDoc.SelectNodes("//PRM_INFO/PRM_FILE/AD_HOC_QUERY/SuperTRUMP/Transaction/Exceptions/Exception")
                    Dim iCount As Int32 = 1
                    For Each Item As XmlNode In ndLstException
                        Dim ndComment As XmlNode = Item.SelectSingleNode("Comment")
                        STLogger.Debug("Exception Number:- " + Item.SelectSingleNode("Number").InnerText + Environment.NewLine + "Description :- " + Item.SelectSingleNode("Description").InnerText + Environment.NewLine + "XPath :- " + Item.SelectSingleNode("XPath").InnerText)

                        If Not ndComment Is Nothing Then
                            STLogger.Debug(Environment.NewLine + "Exception Number:- " + Item.SelectSingleNode("Comment").InnerText)
                        End If
                    Next
                End If
                STLogger.Debug("END:  " + DateTime.Now + " Tracing Exceptions in superTRUMP output within method: " + MethodInfo.GetCurrentMethod.Name())

            Catch lobjSysEx As System.Exception
                Throw lobjSysEx
            End Try
        End Sub

        Public Sub BulkCopyInsertInDB(ByVal strControlFileName As String)
            Dim STLogger As log4net.ILog
            Dim objCommon As BSVICROICommon = New BSVICROICommon()
            STLogger = objCommon.SetLog4Net()
            Dim proc As New Process()
            Try
                STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method in BL: " + MethodInfo.GetCurrentMethod.Name())
                Dim myCommand As String = "CMD.EXE"
                proc.StartInfo = New ProcessStartInfo(myCommand)
                proc.StartInfo.Arguments = "/c SQLLDR vic_roi/vic_roi@mmfd01 CONTROL=" + strControlFileName
                proc.StartInfo.RedirectStandardOutput = True
                proc.StartInfo.RedirectStandardError = True
                proc.StartInfo.UseShellExecute = False
                proc.StartInfo.WorkingDirectory = objCommon.ReadConfigurationFileValue("DATDestinationPath") '"E:\logfiles\STALLDEAL\"
                proc.Start()
                proc.WaitForExit()
                STLogger.Debug("END:  " + DateTime.Now + " Tracing process within method in BL: " + MethodInfo.GetCurrentMethod.Name())
            Catch ex As Exception
                STLogger.Error(" Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + ex.Message)
            Finally
                STLogger = Nothing
                objCommon = Nothing
                proc = Nothing
            End Try
        End Sub

        Public Function FillDataInXmlFormat() As String
           
            Dim objCommon As BSVICROICommon = New BSVICROICommon()
            Dim STLogger As log4net.ILog
            STLogger = objCommon.SetLog4Net()
            Dim strMileStone As String = "4"
            Dim strSuperTRUMPOutXml As String = String.Empty
            Dim strCon As String = String.Empty
            Dim objCDataCls As BSVICROIDAL.cDataClass
            Dim dsData As New DataSet
            Dim dsCapAdder As New DataSet
            strMileStone = "4.1"
            Try
                Dim iMapOutput As Integer = MapLocation(objCommon.ReadConfigurationFileValue("PRMPath"), objCommon.ReadConfigurationFileValue("PRMPathAccessUserName"), objCommon.ReadConfigurationFileValue("PRMPathAccessPassword"))
                STLogger.Debug("Map the drive with return no" + iMapOutput.ToString())
                HTMLReport = "<Table width=640px cellpadding=0 cellspacing=0>"
                HTMLReport += "<TR><TD width=160px style='background-color:#E6FFCC'><Strong>A S Number</Strong></TD><TD width=160px style='background-color:#E6FFCC'><Strong>Yield in DB</Strong></TD><TD width=160px style='background-color:#E6FFCC'><Strong>Book-Yield</Strong></TD><TD width=160px style='background-color:#E6FFCC'><Strong>Eco-Yield</Strong></TD><TD width=160px style='background-color:#E6FFCC'><Strong>PRM Status</Strong></TD><TD width=160px style='background-color:#E6FFCC'><Strong>Lease/Loan</Strong></TD><TD width=160px style='background-color:#E6FFCC'><Strong>Difference</Strong></TD></TR><TR><TD colspan=7></TD></TR>"

                STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method in BL: " + MethodInfo.GetCurrentMethod.Name())
                objCDataCls = New BSVICROIDAL.cDataClass
                strCon = objCommon.ReadConfigurationFileValue("ConnectionstringOLEDBProvider")
                STLogger.Debug("Read ConnectionstringOLEDBProvider from Config file")

                Dim strInputXML As String = objCommon.ReadConfigurationFileValue("InputXmlPath")
                Dim strOutputXML As String = objCommon.ReadConfigurationFileValue("OutputXmlPath")
                Dim strLeaseXSLTPath As String = objCommon.ReadConfigurationFileValue("LeaseXSLTPath")

                STLogger.Debug("Read Input Output - LeaseXslt Path")
                strMileStone = "4.2"

                'If Directory.Exists(strInputXML) Then
                '    Directory.Delete(strInputXML, True)
                '    STLogger.Debug("InputXmlPath Directory Deleted")
                'End If

                'If Directory.Exists(strOutputXML) Then
                '    Directory.Delete(strOutputXML, True)
                '    STLogger.Debug("OutputXmlPath Directory Deleted")
                'End If
                If Not Directory.Exists(strInputXML) Then
                    Directory.CreateDirectory(strInputXML)
                    STLogger.Debug("InputXmlPath Directory created")
                End If
                If Not Directory.Exists(strOutputXML) Then
                    Directory.CreateDirectory(strOutputXML)
                    STLogger.Debug("OutputXmlPath Directory created")
                End If

                strMileStone = "4.3"

                dsData = objCDataCls.GetAllRecordsWithAccountScheduleNo(objCommon.ReadConfigurationFileValue("ConnectionstringOracleClientProvider"))
                STLogger.Debug("Get All Records from database in dataset for No of records :- " + (dsData.Tables(0).Rows.Count).ToString())
                strMileStone = "4.4"

                For asnum As Integer = 0 To dsData.Tables(0).Rows.Count - 1
                    Try
                        Dim dsTOConvertXml As New DataSet
                        Dim dtRow() As DataRow
                        strMileStone = "4.5"
                        dsTOConvertXml.Tables.Add(dsData.Tables(0).Clone())
                        dsTOConvertXml.Tables.Add(dsData.Tables(1).Clone())
                        dsTOConvertXml.Tables.Add(dsData.Tables(2).Clone())
                        dsTOConvertXml.Tables.Add(dsData.Tables(3).Clone())
                        dsTOConvertXml.Tables.Add(dsData.Tables(4).Clone())
                        dsTOConvertXml.Tables.Add(dsData.Tables(5).Clone())

                        strMileStone = "4.6"

                        dsTOConvertXml.Tables(0).TableName = "AccountScheduleFeed"
                        dsTOConvertXml.Tables(1).TableName = "StreamFeed"
                        dsTOConvertXml.Tables(2).TableName = "AssetLevelFeed"
                        dsTOConvertXml.Tables(3).TableName = "ProductMapping"
                        dsTOConvertXml.Tables(4).TableName = "TemplateMapping"
                        dsTOConvertXml.Tables(5).TableName = "Depriciation"

                        strMileStone = "4.7"

                        dtRow = dsData.Tables(0).Select("ACCOUNT_SCHEDULE_NBR ='" & dsData.Tables(0).Rows(asnum)(0).ToString() & "'")
                        dtRow.CopyToDataTable(dsTOConvertXml.Tables(0), LoadOption.Upsert)
                        strMileStone = "4.8"

                        dtRow = dsData.Tables(1).Select("ACCOUNT_SCHEDULE_NBR ='" & dsData.Tables(0).Rows(asnum)(0).ToString() & "'")
                        dtRow.CopyToDataTable(dsTOConvertXml.Tables(1), LoadOption.Upsert)
                        strMileStone = "4.9"

                        dtRow = dsData.Tables(2).Select("ACCOUNT_SCHEDULE_NBR ='" & dsData.Tables(0).Rows(asnum)(0).ToString() & "'")
                        dtRow.CopyToDataTable(dsTOConvertXml.Tables(2), LoadOption.Upsert)
                        strMileStone = "4.10"

                        dtRow = dsData.Tables(3).Select("PMS_Location ='" & dsData.Tables(0).Rows(asnum)("Location").ToString() & "'")
                        dtRow.CopyToDataTable(dsTOConvertXml.Tables(3), LoadOption.Upsert)
                        strMileStone = "4.11"

                        dtRow = dsData.Tables(4).Select("Product ='" & dsData.Tables(0).Rows(asnum)("Product").ToString() & "'" & " and TERM_MIN <='" & dsData.Tables(0).Rows(asnum)("Term").ToString() & "'" & " and TERM_MAX >='" & dsData.Tables(0).Rows(asnum)("Term").ToString() & "'")
                        dtRow.CopyToDataTable(dsTOConvertXml.Tables(4), LoadOption.Upsert)
                        strMileStone = "4.12"

                        If IsDBNull(dsTOConvertXml.Tables(2).Rows(0)("Depreciation_Type")) Then
                            dtRow = dsData.Tables(5).Select("Depreciation_Type =0")
                        Else
                            dtRow = dsData.Tables(5).Select("Depreciation_Type ='" & dsTOConvertXml.Tables(2).Rows(0)("Depreciation_Type").ToString() & "'")
                        End If
                        dtRow.CopyToDataTable(dsTOConvertXml.Tables(5), LoadOption.Upsert)
                        strMileStone = "4.13"

                        Dim columns As DataColumnCollection = dsTOConvertXml.Tables(5).Columns
                        If columns.Contains("Depreciation_Type") Then
                            columns.Remove("Depreciation_Type")
                        End If
                        strMileStone = "4.14"


                        dsCapAdder = objCDataCls.GetCapMarketAdder(dsData.Tables(0).Rows(asnum)(0).ToString(), dsData.Tables(0).Rows(asnum)("Term").ToString(), dsData.Tables(0).Rows(asnum)("Product").ToString(), dsData.Tables(3).Rows(asnum)("ST_PRODUCT_NAME").ToString(), strCon)
                        dsTOConvertXml.Merge(dsCapAdder)
                        strMileStone = "4.15"

                        dsTOConvertXml.WriteXml(strInputXML + dsData.Tables(0).Rows(asnum)(0).ToString() + ".xml")
                        strMileStone = "4.16"
                        Dim xslt As System.Xml.Xsl.XslCompiledTransform
                        xslt = New System.Xml.Xsl.XslCompiledTransform()
                        If File.Exists(strLeaseXSLTPath) Then
                            xslt.Load(strLeaseXSLTPath)
                        Else
                            STLogger.Debug("LeaseXSLTPath file not present at location:- " + strLeaseXSLTPath)
                        End If

                        xslt.Transform(strInputXML + dsData.Tables(0).Rows(asnum)(0).ToString() + ".xml", strOutputXML + dsData.Tables(0).Rows(asnum)(0).ToString() + ".xml")
                        STLogger.Debug(dsData.Tables(0).Rows(asnum)(0).ToString() + " Output XML Created.")
                        strMileStone = "4.17"
                    Catch exOutputXml As Exception
                        STLogger.Error(" Error to create Output XML for Account Schedule Number :- " & dsData.Tables(0).Rows(asnum)(0).ToString() & " Error Desc:- " + exOutputXml.Message)
                    End Try
                Next



                STLogger.Debug("Creation of All OutputXML Completed.....")
                'PRM Creation Function


                STLogger.Debug("Prm Creation Start")
                Dim aobjSTXMLDoc As New Xml.XmlDocument
                strSuperTRUMPOutXml = String.Empty
                Dim sbQuery As New StringBuilder(" BEGIN ")
                Dim dInfo As New DirectoryInfo(strOutputXML)
                If Directory.Exists(strOutputXML) Then
                    Dim fInfo As FileInfo() = dInfo.GetFiles("*.xml")
                    STLogger.Debug("Get All Xml Files name in array")
                    For Each fi As FileInfo In fInfo
                        Try
                            strMileStone = "4.18"
                            Dim STService As SuperTRUMPService.ISuperTrumpServiceSoapPort = New SuperTRUMPService.ISuperTrumpServiceSoapPort()
                            Dim xmlDoc As New XmlDocument()
                            xmlDoc.Load(strOutputXML + fi.Name)
                            strMileStone = "4.19"
                            Dim strOutXml As String = xmlDoc.OuterXml

                            If strOutXml.IndexOf("<PRM") > -1 Then
                                strOutXml = strOutXml.Substring(strOutXml.IndexOf("<PRM"))
                            End If
                            strMileStone = "4.20"
                            strSuperTRUMPOutXml = STService.RunAdHocXMLInOutQuery(strOutXml.Trim())


                            Dim foundRows() As DataRow = dsData.Tables(0).Select("ACCOUNT_SCHEDULE_NBR ='" & fi.Name.ToString().Substring(0, fi.Name.IndexOf(".xml")) & "'")
                            CreateHTMLReportForPRM(foundRows(0)("Book_Yield").ToString(), strSuperTRUMPOutXml, fi.Name.ToString().Substring(0, fi.Name.IndexOf(".xml")), foundRows(0)("PRODUCT").ToString())


                            aobjSTXMLDoc.LoadXml(strSuperTRUMPOutXml)
                            strMileStone = "4.21"
                            If String.Compare(objCommon.ReadConfigurationFileValue("CheckSuperTrumpExceptionFlag"), "True", True) = 0 Then
                                CheckSuperTrumpExceptions(aobjSTXMLDoc)
                                File.WriteAllText(objCommon.ReadConfigurationFileValue("ST_ExceptionPath") + fi.Name, aobjSTXMLDoc.ToString())
                            End If
                            strMileStone = "4.22"
                            sbQuery.Append(" update TBL_ACCOUNTSCHEDULE_DETAIL set Process_Flag = 1 where ACCOUNT_SCHEDULE_NBR = '" + fi.Name.Substring(0, fi.Name.IndexOf(".")) + "';")
                            STLogger.Debug(fi.Name.Substring(0, fi.Name.IndexOf(".")) + " PRM Created.")
                            STService.Dispose()
                        Catch ex As Exception
                            STLogger.Error(" Error to create PRM for Account Schedule Number :- " & fi.Name.Substring(0, fi.Name.IndexOf(".")) & " Error Desc:- " + ex.Message)
                        End Try
                    Next
                    HTMLReport += "<TR><TD colspan=7> Total Match Found : - " + iMatchCount.ToString() + "</TD></TR>"
                    HTMLReport += "</TABLE>"
                    File.WriteAllText(objCommon.ReadConfigurationFileValue("PRMMatchingStatusReportPath"), HTMLReport)

                    sbQuery.Append(" END;")
                    objCDataCls = New BSVICROIDAL.cDataClass
                    Dim strDBOutput As String = String.Empty

                    If String.Compare(sbQuery.ToString(), " BEGIN  END;", True) <> 0 Then
                        strDBOutput = objCDataCls.ExportDATFileByQuery(sbQuery.ToString(), objCommon.ReadConfigurationFileValue("ConnectionstringOracleClientProvider"))
                        If strDBOutput = 1 Then
                            STLogger.Debug("Set Process_Flag to 1 for all Processing Account Schedule No.")
                        End If
                    End If

                    STLogger.Debug("PRM Creation End.....")
                    Dim bUnMapped As Boolean = UnMapDrive(objCommon.ReadConfigurationFileValue("PRMPath"))
                    STLogger.Debug("UnMap Drive with ReturnType " + bUnMapped.ToString())
                End If
                If String.Compare(objCommon.ReadConfigurationFileValue("InputOutputFolderDeleted"), "yes", True) = 0 Then

                    If Directory.Exists(strInputXML) Then
                        Directory.Delete(strInputXML, True)
                        STLogger.Debug("InputXmlPath Directory Deleted")
                    End If

                    If Directory.Exists(strOutputXML) Then
                        Directory.Delete(strOutputXML, True)
                        STLogger.Debug("OutputXmlPath Directory Deleted")
                    End If
                End If

            Catch exPrm As Exception
                STLogger.Error("MileStone:- " & strMileStone & " Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + exPrm.Message)
            Finally
                dsData = Nothing
                objCommon = Nothing
                STLogger = Nothing
                objCDataCls = Nothing
                dsCapAdder = Nothing
            End Try
            Return "DONE SUCCESSFULLY"
        End Function

        Public Sub CreateHTMLReportForPRM(ByVal strYield As String, ByVal strInput As String, ByVal strAcNUm As String, ByVal strProduct As String)
            'loan - book yield   ,   lease - Economic yield
            Dim xmlDoc As New XmlDocument()

            xmlDoc.LoadXml(strInput)
            HTMLReport += "<TR><TD>" + strAcNUm + "</TD>"
            HTMLReport += "<TD>" + strYield + "</TD>"
            HTMLReport += "<TD>" + Math.Round(Val((xmlDoc.SelectSingleNode("//NPTBookYield").InnerText.ToString())), 2).ToString() + "</TD>"
            HTMLReport += "<TD>" + Math.Round(Val((xmlDoc.SelectSingleNode("//NPTEconomicYield").InnerText.ToString())), 2).ToString() + "</TD>"

            If String.Compare(strProduct, "MEREG", True) = 0 Or String.Compare(strProduct, "MEQMUN", True) = 0 Or String.Compare(strProduct, "MEOQSI", True) = 0 Or String.Compare(strProduct, "MEGMUN", True) = 0 Then   'Loan 
                'If Math.Round(Val((xmlDoc.SelectSingleNode("//NPTBookYield").InnerText.ToString())), 2) = Convert.ToDouble(strYield) Then
                If xmlDoc.SelectSingleNode("//NPTBookYield").InnerText.ToString().Substring(0, xmlDoc.SelectSingleNode("//NPTBookYield").InnerText.ToString().LastIndexOf(".") + 2) = strYield.Substring(0, strYield.LastIndexOf(".") + 2) Then
                    HTMLReport += "<TD style='background-color:#CCFFCC'>" + "Matched" + "</TD>"
                    iMatchCount += 1
                Else
                    HTMLReport += "<TD style='background-color:#FFCCCC'>" + "Unmatched" + "</TD>"
                End If
                HTMLReport += "<TD style='background-color:#CCFFFF'>" + "Loan" + "</TD>"
                HTMLReport += "<TD style='background-color:#FFFFCC'>" + (Math.Round(Val((xmlDoc.SelectSingleNode("//NPTBookYield").InnerText.ToString())), 2) - Convert.ToDouble(strYield)).ToString() + "</TD>"
            Else   'Lease
                'If Math.Round(Val((xmlDoc.SelectSingleNode("//NPTEconomicYield").InnerText.ToString())), 2) = Convert.ToDouble(strYield) Then
                If xmlDoc.SelectSingleNode("//NPTEconomicYield").InnerText.ToString().Substring(0, xmlDoc.SelectSingleNode("//NPTEconomicYield").InnerText.ToString().LastIndexOf(".") + 2) = strYield.Substring(0, strYield.LastIndexOf(".") + 2) Then
                    HTMLReport += "<TD style='background-color:#CCFFCC'>" + "Matched" + "</TD>"
                    iMatchCount += 1
                Else
                    HTMLReport += "<TD style='background-color:#FFCCCC'>" + "Unmatched" + "</TD>"
                End If
                HTMLReport += "<TD style='background-color:#E6CCFF'>" + "Lease" + "</TD>"
                HTMLReport += "<TD style='background-color:#FFFFCC'>" + (Math.Round(Val((xmlDoc.SelectSingleNode("//NPTEconomicYield").InnerText.ToString())), 2) - Convert.ToDouble(strYield)).ToString() + "</TD>"
            End If

            HTMLReport += "</TR>"
        End Sub

        '''''''UNUSED FUNCTIONS'''''''''''

        'Private Function GetAmountForSchedulerNumber(ByVal objAList As ArrayList) As ArrayList

        '    Dim objRowData As String()
        '    Dim STLogger As log4net.ILog
        '    Dim strMileStone As String = "3"
        '    Dim strRow As String = String.Empty
        '    Dim alOutputValue As New ArrayList
        '    Dim strMonth As String = String.Empty
        '    Dim strAmount As String = String.Empty
        '    Dim iOccurrance As Integer = 0
        '    Dim objHeading As String() = Nothing
        '    Dim objCommon As BSVICROICommon = New BSVICROICommon()
        '    STLogger = objCommon.SetLog4Net()
        '    Try
        '        STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method: " + MethodInfo.GetCurrentMethod.Name())
        '        strMileStone = "3.1"
        '        If objAList.Count > 0 Then
        '            'Get all heading of months 
        '            objHeading = objAList.Item(0).ToString().Split(",")
        '        End If

        '        strMileStone = "3.3"

        '        Dim STDate As String = String.Empty
        '        Dim blnFlag As Boolean = False
        '        'Iterate the loop on arraylist with Items(Start from 2nd item. Firts Item is months Heading)

        '        For index As Integer = 0 To objAList.Count - 1
        '            strRow = objAList.Item(index).ToString() + ",*"
        '            strMileStone = "3.4"
        '            objRowData = strRow.Split(",")
        '            strMileStone = "3.5"
        '            iOccurrance = 0
        '            strAmount = objRowData(1)

        '            If objRowData(1) = "0.00" Then
        '                blnFlag = True
        '            End If


        '            strMileStone = "3.6"
        '            'Iterate loop on each arraylist item that fill in other array by split data with "Comma"
        '            For iItem As Integer = 1 To objRowData.Length - 2
        '                'Used to find out the starting month of each calculation
        '                STDate = objRowData(objRowData.Length - 2)

        '                strMileStone = "3.7"

        '                'Used to increase the count in case of continious matching.
        '                If strAmount = objRowData(iItem) Then
        '                    iOccurrance = iOccurrance + 1
        '                Else

        '                    If iItem = objRowData.Length - 2 And objRowData(iItem - 1) = "0.00" Then  ' Check for last occurance of '0.00'
        '                        Exit For
        '                    ElseIf blnFlag = True Then  ' Check for first occurance of '0.00'
        '                        blnFlag = False
        '                        strAmount = objRowData(iItem)
        '                    Else
        '                        alOutputValue.Add(objRowData(0) + "," + strAmount + "," + iOccurrance.ToString() + "," + STDate)
        '                        strAmount = objRowData(iItem)
        '                    End If

        '                    iOccurrance = 1
        '                End If
        '                strMileStone = "3.8"
        '            Next

        '            objRowData = Nothing
        '        Next
        '        STLogger.Debug("END:  " + DateTime.Now + " Tracing process within method: " + MethodInfo.GetCurrentMethod.Name())
        '        Return alOutputValue
        '    Catch ex As Exception
        '        STLogger.Error("MileStone:- " & strMileStone & " Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
        '        Throw
        '        Return Nothing
        '    Finally
        '        alOutputValue = Nothing
        '        objCommon = Nothing
        '        STLogger = Nothing
        '    End Try

        'End Function



        'Public Function GetDATFileDataInArrayList(ByVal strDATFilename As String) As ArrayList
        '    Dim strMileStone As String = "2"
        '    Dim STLogger As log4net.ILog
        '    Dim ALDAT As New ArrayList
        '    Dim tmpStream As StreamReader
        '    Dim objCommon As BSVICROICommon = New BSVICROICommon()
        '    STLogger = objCommon.SetLog4Net()
        '    Try
        '        STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method: " + MethodInfo.GetCurrentMethod.Name())
        '        If File.Exists(strDATFilename) Then
        '            strMileStone = "2.1"

        '            STLogger.Debug("Reading and storing full text from DAT file to stream")
        '            tmpStream = File.OpenText(strDATFilename)
        '            Dim strContent As String = tmpStream.ReadToEnd().ToString()
        '            Dim strLines() As String
        '            Dim strCompositKeyValue As String = String.Empty
        '            STLogger.Debug("Splitting stream on newline and store it into strLinesarray")
        '            strLines = strContent.Split(Environment.NewLine)
        '            strMileStone = "2.2"
        '            ALDAT.Clear()

        '            For iItem As Integer = 1 To strLines.Length - 1
        '                ALDAT.Add(strLines(iItem).Trim())
        '                strMileStone = "2.3"
        '            Next
        '            STLogger.Debug("Adding all rows data from DAT File to ArrayList")

        '        End If
        '        STLogger.Debug("END:  " + DateTime.Now + " Tracing process within method: " + MethodInfo.GetCurrentMethod.Name())
        '    Catch ex As Exception
        '        STLogger.Error("MileStone:- " & strMileStone & " Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
        '        Throw ex
        '    Finally
        '        objCommon = Nothing
        '        STLogger = Nothing
        '        tmpStream = Nothing
        '    End Try
        '    Return ALDAT
        'End Function

    End Class
End Namespace

