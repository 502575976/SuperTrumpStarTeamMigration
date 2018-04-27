Imports BSMoneyCostEntity
Imports System.Reflection
Imports BSMoneyCostAuto.modCEFCommon
Imports System.Net.mail
Imports BSMoneyCostAuto.modBSMoneyCostAuto
Imports System.EnterpriseServices
Module modBSMoneyCostAuto
    Dim STLogger As log4net.ILog
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
    '=== Constant for module name ===================================
    Private Const cMODULE_NAME As String = "modBSMoneyCostAuto"
    'Public gstrErrorLogFile As String
    Public Const cCOMPONENT_NAME As String = "BSMoneyCostAuto"
    '================================================================

    '=== Registry Constants =========================================
    Public Const cFACILITY_CONFIG_REG_PATH As String = "HKEY_LOCAL_MACHINE\SOFTWARE\FacilitySettings\"
    Public Const cErrorMailBoxKey As String = "ErrorMailBox"
    Public Const cEmailOverrideKey As String = "EmailOverride"
    Public Const cDeveloperEmailKey As String = "DeveloperEmail"
    Public Const cEmailFromKey As String = "EmailFrom"
    Public Const cClarifyPriorityKey As String = "ClarifyPriority"
    Public Const cClarifyContactFNameKey As String = "ClarifyContact_Fname"
    Public Const cClarifyContactLNameKey As String = "ClarifyContact_Lname"
    Public Const cClarifyContactPhoneKey As String = "ClarifyContact_Phone"
    Public Const cClarifyQNameKey As String = "ClarifyQueueName"
    Public Const cClarifySiteIdKey As String = "SiteID"
    Public Const cClarifyEmailSubject As String = "EmailSubject"
    '================================================================

    Public Const cWORKING_DIRECTORY_PATH As String = "WorkingDirectory"
    Public Const cBACKUP_LOCATION As String = "Backup_Location"
    Public Const cNETWORK_LOCATION As String = "Network_Location"
    Public Const cFTP_LOCATION As String = "FTP_Location"
    Public Const cFTP_DIRECTORY As String = "FTP_Directory"
    Public Const cFTP_USER As String = "FTP_User"
    Public Const cFTP_PASSWORD As String = "FTP_Password"

    Public Const cFTP_LOCATION_NEWDATEFORMAT As String = "FTP_LocationNewDateFormat"
    Public Const cFTP_USER_NEWDATEFORMAT As String = "FTP_UserNewDateFormat"
    Public Const cFTP_PASSWORD_NEWDATEFORMAT As String = "FTP_PasswordNewDateFormat"
    Public Const cFTP_DIRECTORY_NEWDATEFORMAT As String = "FTP_DirectoryNewDateFormat"
    Public Const cEncrptFileFlag As String = "EncrptFileFlag"

    Public Enum eDebugLevel
        ecDebugDoNotLog = 0
        ecDebugFatalError = 3
        ecDebugCriticalError = 6
        ecDebugWarning = 9
        ecDebugInputTrace = 12
        ecDebugOutputTrace = 15
        ecDebugLogData = 18
        ecDebugLoglargeData = 21
    End Enum

    ''MapNetwork path
    Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, ByVal cbRemoteName As Long) As Long
    Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByVal lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
    Public Structure NETRESOURCE
        Dim dwScope As Long
        Dim dwType As Long
        Dim dwDisplayType As Long
        Dim dwUsage As Long
        Dim lpLocalName As String
        Dim lpRemoteName As String
        Dim lpComment As String
        Dim lpProvider As String
    End Structure

    '============================================================
    'METHOD  : CopyFiles
    'PURPOSE : Copy/Move files to location specified
    'PARMS   : [argSource] From where the files have to be moved
    '          [argDestination] Location where the files have to be moved or copied.
    '          [argMove] Boolean argument specifying copy or move
    'RETURN  : XML Error String if any
    '============================================================
    Public Function CopyFiles(ByVal objFileEntity As FileInfoEntity) As cDataEntity
        Dim cdataEntity As New cDataEntity
        Dim lstrReturnString As String = ""
        SetLog4Net()
        Try
            If objFileEntity.Move = "" Then
                objFileEntity.Move = "False"
            End If
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): In CopyFiles() method")
            ' -------------------------------------------
            ' Check move(true) or copy argument
            ' -------------------------------------------
            If Convert.ToBoolean(objFileEntity.Move) Then
                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): Deleting Source file - " & objFileEntity.Source)

                IO.File.Delete(objFileEntity.Source)

                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): Move Source file - " & objFileEntity.Source & " to destination - " & objFileEntity.Destination)

                IO.File.Move(objFileEntity.Source, objFileEntity.Destination)
            Else
                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): Copying Source file - " & objFileEntity.Source & " to destination - " & objFileEntity.Destination)

                IO.File.Copy(objFileEntity.Source, objFileEntity.Destination, True)
            End If
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): Exit CopyFiles() method")
            cdataEntity.OutputString = lstrReturnString
            Return cdataEntity
        Catch ex As Exception
            lstrReturnString = "<ERROR_DETAILS>" & _
                                    "<ERROR_NUMBER>" & Err.Number & "</ERROR_NUMBER>" & _
                                    "<ERROR_DESCRIPTION>" & Err.Description & "</ERROR_DESCRIPTION>" & _
                                    "<ERROR_SOURCE>" & Err.Source & "::CopyFiles()</ERROR_SOURCE>" & _
                                "</ERROR_DETAILS>"
            STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_CopyFiles(): Error occured - " & Err.Description)
            Throw
        Finally
            If Not IsNothing(objFileEntity) Then
                objFileEntity = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try

    End Function

    '================================================================
    'METHOD  : GetDateFormat
    'PURPOSE : To get particular date in mmddyy format
    'PARMS   : astrDate [String] = date, which needs to convert in mmddyy format
    'RETURN  : [String] = date formatted into mmddyy string
    '================================================================
    Public Function GetDateFormat(ByVal objcdataEntity As cDataEntity) As cDataEntity
        Dim lstrDay As String
        Dim lstrMonth As String
        Dim lstrYear As String

        Try
            lstrDay = Day(objcdataEntity.ProcessDate)
            If Len(lstrDay) = 1 Then
                lstrDay = "0" & lstrDay
            End If

            lstrMonth = Month(objcdataEntity.ProcessDate)
            If Len(lstrMonth) = 1 Then
                lstrMonth = "0" & lstrMonth
            End If

            lstrYear = Year(objcdataEntity.ProcessDate)
            If Len(lstrYear) > 2 Then
                lstrYear = Right(lstrYear, 2)
            End If
            objcdataEntity.OutputString = lstrMonth & lstrDay & lstrYear
            Return objcdataEntity
        Catch ex As Exception
            objcdataEntity.OutputString = Nothing
            Return objcdataEntity
            Throw
        Finally
            If Not IsNothing(objcdataEntity) Then
                objcdataEntity = Nothing
            End If

        End Try


    End Function
    '================================================================
    'METHOD  : SendNotification
    'PURPOSE : To send Sucess notification to the designated error
    '          mail box.
    'PARMS   :
    '          astrBody [String] = Sucess Message.
    'RETURN  : None.
    '================================================================
    Public Function SendNotification(ByVal objFileInfoEntity As FileInfoEntity) As cDataEntity

        Dim cdataEntity As New cDataEntity
        Dim lstrMailBody As String
        Dim lstrTo As String = ""
        Dim lstrReturnString As String = ""
        Dim mailFrom As String
        Dim Subject As String = ""
        Dim Body As String = ""
        Dim objMail As New System.Web.Mail.MailMessage
        Dim lobjSMTP As System.Web.Mail.SmtpMail
        SetLog4Net()
        Try
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendNotification(): In SendNotification() method")

            If objFileInfoEntity.SendNotification Then
                If objFileInfoEntity.Body.Trim = "" Then
                    lstrMailBody = "Dear Money Cost User," & vbCrLf & vbCrLf & _
                              "MoneyCost process is succesfully completed, no MoneyCost file is updated.  " & vbCrLf & objFileInfoEntity.Body & vbCrLf
                Else
                    lstrMailBody = "Dear Money Cost User," & vbCrLf & vbCrLf & _
                                                  "Following Money Cost Files are succesfully updated " & vbCrLf & objFileInfoEntity.Body & vbCrLf
                End If

                lstrMailBody = lstrMailBody & "Note: This is an auto-generated email from Money Cost Service. Please do not reply to this email."


                mailFrom = gstrFrom
                'mailTo = lstrTo
                If gstrDeveloperEmail <> "" Then objMail.Cc = gstrDeveloperEmail 'mailTo = gstrDeveloperEmail
                objMail.To = objFileInfoEntity.BusinessContact 'mailTo
                objMail.From = mailFrom
                objMail.Subject = "MoneyCostService: Succesful Notification "
                objMail.BodyFormat = Web.Mail.MailFormat.Html
                'objMail.BodyEncoding = System.Text.Encoding.UTF8
                objMail.Body = "<pre>" & lstrMailBody
                lobjSMTP.SmtpServer = GetConfigurationKey("SmtpServer")


                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendNotification(): Sending email to - " & objFileInfoEntity.BusinessContact & " and cc - " & gstrDeveloperEmail)                
                'Send email
                lobjSMTP.Send(objMail)
                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendNotification(): Mail Content - " & lstrMailBody)
            End If
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendNotification(): Exit SendNotification() method")
            cdataEntity.OutputString = vbNullString
            Return cdataEntity
        Catch ex As Exception
            lstrReturnString = "<ERROR_DETAILS><ERROR_NUMBER>" & Err.Number & "</ERROR_NUMBER>" & _
                            "<ERROR_DESCRIPTION>" & Err.Description & "</ERROR_DESCRIPTION>" & _
                            "<ERROR_SOURCE>" & Err.Source & "::SendErrNotification()</ERROR_SOURCE></ERROR_DETAILS>"
            STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): Error occured - " & Err.Description & vbCrLf & lstrReturnString)
            cdataEntity.OutputString = vbNullString
            Return cdataEntity
        Finally
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
            If Not IsNothing(objMail) Then
                objMail = Nothing
            End If
            If Not IsNothing(lobjSMTP) Then
                lobjSMTP = Nothing
            End If

        End Try

    End Function
    '================================================================
    'METHOD  : SendErrNotification
    'PURPOSE : To send error notification to the designated error
    '          mail box.
    'PARMS   :
    '          astrBody [String] = Error Message.
    'RETURN  : None.
    '================================================================
    Public Function SendErrNotification(ByVal objFileInfoEntity As FileInfoEntity) As cDataEntity

        Dim cdataEntity As New cDataEntity
        Dim lstrMailBody As String
        Dim lstrTo As String = ""
        Dim lPos As Integer
        Dim lstrReturnString As String = ""
        Dim mailFrom As String
        Dim mailTo As String
        Dim Subject As String = ""
        Dim Body As String = ""
        Dim objMail As New System.Web.Mail.MailMessage
        Dim lobjSMTP As System.Web.Mail.SmtpMail
        SetLog4Net()
        Try
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): In SendErrNotification() method")

            'If objFileInfoEntity.CutTicket Then

            '    If gstrEmailOverride <> "" Then
            '        lstrTo = gstrEmailOverride
            '    Else
            '        lstrTo = gstrErrMailBox
            '    End If

            '    'Create the email body as per the format required to open a clarify case.
            '    lstrMailBody = "CASE_START" & vbCrLf & _
            '                    "CALL_TYPE: Problem" & vbCrLf & _
            '                    "SEVERITY: High" & vbCrLf & _
            '                    "PRIORITY: " & gstrClarifyPriority & vbCrLf & _
            '                    "SITE_ID: " & gstrClarifySiteId & vbCrLf & _
            '                    "CONTACT_LAST_NAME: " & gstrClarifyContactLname & vbCrLf & _
            '                    "CONTACT_FIRST_NAME: " & objFileInfoEntity.QueueName & vbCrLf & _
            '                    "CONTACT_PHONE: " & gstrClarifyContactPhone & vbCrLf & _
            '                    "CASE_SUMMARY: " & gstrClarifyEmailSub & vbCrLf & _
            '                    "CASE_DESCRIPTION:" & vbCrLf & _
            '                        objFileInfoEntity.Body & vbCrLf & _
            '                    "CASE_END"

            '    lstrMailBody = "<pre>" & lstrMailBody
            '    mailFrom = gstrDeveloperEmail.Trim
            '    mailTo = lstrTo
            '    If gstrDeveloperEmail <> "" Then objMail.Cc = gstrDeveloperEmail
            '    objMail.To = mailTo
            '    objMail.From = mailFrom
            '    objMail.Subject = "Case Request"
            '    objMail.BodyFormat = Web.Mail.MailFormat.Html
            '    objMail.Body = "<pre>" & lstrMailBody
            '    'objMail.BodyEncoding = System.Text.Encoding.UTF8
            '    lobjSMTP.SmtpServer = "mail.ad.ge.com"
            '    STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): Sending email to - " & lstrTo & " and cc - " & gstrDeveloperEmail)

            '    'Send email           
            '    lobjSMTP.Send(objMail)

            'End If

            If objFileInfoEntity.SendNotification Then

                lPos = InStr(1, objFileInfoEntity.Body, "<", vbBinaryCompare)

                lstrMailBody = "Dear Money Cost User," & vbCrLf & vbCrLf & _
                               "An error was reported by service while updating the money cost file." & IIf(objFileInfoEntity.CutTicket, " A clarify case has been created and dispatched to IT support team." & vbCrLf, vbCrLf) & vbCrLf & _
                               "Money Cost Support Team." & vbCrLf & vbCrLf

                If lPos <= 0 Then
                    lstrMailBody = lstrMailBody & "Error Reported:" & vbCrLf & objFileInfoEntity.Body & vbCrLf & vbCrLf
                Else
                    lstrMailBody = lstrMailBody & "Error Reported:" & vbCrLf & Left(objFileInfoEntity.Body, lPos - 1) & vbCrLf & vbCrLf
                End If

                '************************
                'Changes made on 08 Nov 2006 by Nizar
                'Changes made to put an message as auto generated email.
                lstrMailBody = lstrMailBody & "Note: This is an auto-generated email from Money Cost Service. Please do not reply to this email."
                mailFrom = gstrFrom
                'mailTo = lstrTo
                If gstrDeveloperEmail <> "" Then objMail.Cc = gstrDeveloperEmail 'mailTo = gstrDeveloperEmail
                objMail.To = objFileInfoEntity.BusinessContact 'mailTo
                objMail.From = mailFrom
                objMail.Subject = "MoneyCostService: Error reported while updating " & objFileInfoEntity.MCCode & " for " & objFileInfoEntity.ProcessDates
                objMail.BodyFormat = Web.Mail.MailFormat.Html
                'objMail.BodyEncoding = System.Text.Encoding.UTF8
                objMail.Body = "<pre>" & lstrMailBody
                lobjSMTP.SmtpServer = GetConfigurationKey("SmtpServer")


                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): Sending email to - " & objFileInfoEntity.BusinessContact & " and cc - " & gstrDeveloperEmail)
                'Send email
                lobjSMTP.Send(objMail)
                STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendNotification(): Mail Content - " & lstrMailBody)
            End If
            STLogger.Debug("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): Exit SendErrNotification() method")
            cdataEntity.OutputString = vbNullString
            Return cdataEntity
        Catch ex As Exception
            lstrReturnString = "<ERROR_DETAILS><ERROR_NUMBER>" & Err.Number & "</ERROR_NUMBER>" & _
                            "<ERROR_DESCRIPTION>" & Err.Description & "</ERROR_DESCRIPTION>" & _
                            "<ERROR_SOURCE>" & Err.Source & "::SendErrNotification()</ERROR_SOURCE></ERROR_DETAILS>"
            STLogger.Error("BSSTMoneyCostAuto.modBSMoneyCostAuto_SendErrNotification(): Error occured - " & Err.Description & vbCrLf & lstrReturnString)
            cdataEntity.OutputString = vbNullString
            Return cdataEntity
        Finally
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
            If Not IsNothing(objMail) Then
                objMail = Nothing
            End If
            If Not IsNothing(lobjSMTP) Then
                lobjSMTP = Nothing
            End If

        End Try

    End Function

    '============================================================================================
    'METHOD :   fnSortXmlData
    'PURPOSE:   This Function will use to sort the XML records depending upon the column.
    'PARMS  :   astrXmlDoc      [String] = XML string.
    '           astrXslFileName [String] = Path of XSL file.
    'RETURN :   Sorted HTML output.
    '============================================================================================
    Public Function fnSortXmlData(ByVal xmlErrEntity As XmlErrEntity)

        Dim lstrMethodName As String    'to store method name                
        Dim Stylesheet As System.Xml.Xsl.XslCompiledTransform = New System.Xml.Xsl.XslCompiledTransform()
        Dim clsStream As System.IO.MemoryStream = New System.IO.MemoryStream()
        Dim clsTransform As System.Xml.XmlWriter = New System.Xml.XmlTextWriter(clsStream, System.Text.Encoding.ASCII)
        Dim lobjXMLDoc As New Xml.XmlDocument    'to load XML Document
        Dim lobjXSLDoc As New Xml.XmlDocument      'to load XSL Document         

        Try
            lstrMethodName = "fnSortXmlData"


            'to get result output           

            lobjXMLDoc.LoadXml(xmlErrEntity.XmlDoc)
            lobjXSLDoc.Load(xmlErrEntity.XslFilePath)

            Stylesheet.Load(lobjXSLDoc)
            ' apply the transformation to the specified xml... 

            Stylesheet.Transform(lobjXMLDoc, clsTransform)
            clsStream.Position = 0
            ' extract content... 
            Dim bValue As Byte() = DirectCast(Array.CreateInstance(GetType(Byte), clsStream.Length), Byte())
            clsStream.Read(bValue, 0, Convert.ToInt32(clsStream.Length))

            Return System.Text.UnicodeEncoding.ASCII.GetString(bValue)


        Catch ex As Exception
            fnSortXmlData = vbNullString
            Throw
        Finally
            If Not IsNothing(lobjXMLDoc) Then
                lobjXMLDoc = Nothing
            End If
            If Not IsNothing(lobjXSLDoc) Then
                lobjXSLDoc = Nothing
            End If
            If Not IsNothing(Stylesheet) Then
                Stylesheet = Nothing
            End If
            If Not IsNothing(clsTransform) Then
                clsTransform = Nothing
            End If
            If Not IsNothing(clsStream) Then
                clsStream = Nothing
            End If
        End Try

    End Function
End Module
