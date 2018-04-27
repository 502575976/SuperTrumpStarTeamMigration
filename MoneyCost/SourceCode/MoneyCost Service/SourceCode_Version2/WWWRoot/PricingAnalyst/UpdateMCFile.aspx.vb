Imports System
Imports System.text
Imports Microsoft.Win32
Imports BSMoneyCostEntity
Imports BSMoneyCostBL
Partial Public Class UpdateMCFile
    Inherits System.Web.UI.Page
    Private mViewSearchResults As DataView
    Private mstrRegistryHive As String = "SOFTWARE\\FacilitySettings\\MoneyCost\\FilePath"
    Private mstrRegistryKey As String = "SSOAdminRole"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim lstrUSER_Role, lstrSSOAdminURL As String
            'Response.Write("lalit")

            lblErrorMessage.ForeColor = Drawing.Color.Red
            If Not Page.IsPostBack Then
                'cefmoneycostrole

                lstrUSER_Role = Trim(Request.ServerVariables("HTTP_CEFMONEYCOSTROLE"))
         
                If Not UCase(lstrUSER_Role).Contains("ADMIN") Then
                    'Request.Url
                    lstrSSOAdminURL = ReadRegistry(mstrRegistryHive, mstrRegistryKey)
                    Response.Redirect(lstrSSOAdminURL, False)
                End If
                FillGrid()
            End If
        Catch ex As Exception
            LogError(ex.Message)

            XmlErrEntity.ErrorMessage = ex.Message
            lblErrorMessage.Text = ex.Message
            'Response.Redirect("~/ErrorPage.aspx", False)
        End Try
    End Sub

    Function FillGrid()
        Dim lobjcDataEntity As cDataEntity
        Dim lobjcMoneyCostUISvc As cMoneyCostUISvc
        Dim ldsMCDetail As DataSet


        Try
            lblErrorMessage.Text = ""

            lobjcDataEntity = New cDataEntity
            lobjcMoneyCostUISvc = New cMoneyCostUISvc
            lobjcDataEntity.MoneyCostID = -1
         
            lobjcDataEntity = lobjcMoneyCostUISvc.GetMCFileDetails(lobjcDataEntity)
 
            ldsMCDetail = lobjcDataEntity.OutputDataSet
            If ldsMCDetail.Tables(0).Rows.Count > 0 Then
                tblIndexRate.Visible = True
                grdMCDetail.Visible = True
                lblNoMCDetail.Visible = False
                btnSave.Visible = True
                btnCancel.Visible = True

                'Get the Default view and set the Sort by and Sort Order.
                mViewSearchResults = ldsMCDetail.Tables(0).DefaultView
                grdMCDetail.DataSource = mViewSearchResults
                grdMCDetail.DataBind()

            Else
                tblIndexRate.Visible = False
                grdMCDetail.Visible = False
                lblNoMCDetail.Visible = True
                lblNoMCDetail.Text = "No Index Rate Exists."


                btnSave.Visible = False
                btnCancel.Visible = True

            End If

        Catch ex As Exception
            LogError(ex.Message)

            ' XmlErrEntity.ErrorMessage = ex.Message
            lblErrorMessage.Text = ex.Message
            'Response.Redirect("~/ErrorPage.aspx", False)
        Finally
            'If Not IsNothing(lobjcMoneyCostUISvc) Then
            '    'lobjcMoneyCostUISvc.Dispose()
            '    'lobjcMoneyCostUISvc = Nothing
            'End If
            'If Not IsNothing(ldsMCDetail) Then
            '    ldsMCDetail.Dispose()
            '    ldsMCDetail = Nothing
            'End If
            'lobjcDataEntity = Nothing
        End Try
    End Function
    Public Sub LogError(ByVal astrMessage As String)
        Dim lstrErrorMessage As New StringBuilder
        Dim mobjLogger2 As log4net.ILog
        Try

            If log4net.LogManager.GetRepository.Configured = False Then
                log4net.Config.XmlConfigurator.ConfigureAndWatch(New System.IO.FileInfo(GetLog4NetConfigPath()))
            End If

            If IsNothing(mobjLogger2) Then
                mobjLogger2 = log4net.LogManager.GetLogger("MoneyCost_Client")
            End If

            'Genarate Error Message String
            lstrErrorMessage.Append("""" & astrMessage & """")
            mobjLogger2.Error(lstrErrorMessage.ToString)

        Catch ex As Exception
            Throw
        Finally
            log4net.LogManager.GetRepository.Configured = False
            mobjLogger2 = Nothing
        End Try
    End Sub
    Function ReadRegistry(ByVal astrRegistryHive As String, ByVal astrKeyName As String) As String
        Dim lobjreg As RegistryKey = Registry.LocalMachine

        lobjreg = lobjreg.OpenSubKey(astrRegistryHive, False)
        ReadRegistry = lobjreg.GetValue(astrKeyName)

        lobjreg = Nothing
    End Function
    Function GetLog4NetConfigPath() As String
        GetLog4NetConfigPath = ReadRegistry("SOFTWARE\\FacilitySettings\\MoneyCost\\FilePath", "WEBLog4NetConfig")
    End Function

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim lobjcDataEntity As cDataEntity
        Dim lobjcMoneyCostUISvc As cMoneyCostUISvc
        Dim lobjTextBox As New TextBox
        Dim lstrQuery As String
        'IND_Active = txtINDActive
        'Days_To_Skip = txtDaysToSkip
        'Frequency = txtFrequency
        'Frequency_Count = txtFrequencyCount
        'Last_Schedule_Process_Date = txtProcessDate
        'Business_Contact = txtBusinessContact
        'FTP_Directory = txtFTPDirectory
        Dim lstrCurrencyCode As String
        Dim lstrINDActive As String
        Dim lstrDaysToSkip As String
        Dim lstrFrequency As String
        Dim lstrFrequencyCount As String
        Dim lstrProcessDate As String
        Dim lstrBusinessContact As String
        Dim lstrFTPDirectory As String
        Dim lstrMCId As String
        Dim iCounter As Integer
        Dim lstrSSOIdUpdate As String = ""
        Try
            lstrQuery = ""

            For iCounter = 0 To grdMCDetail.Items.Count - 1
                lstrMCId = grdMCDetail.Items(iCounter).Cells(0).Text


                lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtCurrencyCode")
                lstrCurrencyCode = lobjTextBox.Text
                'If lstrCurrencyCode = "" Then                    
                '    lblErrorMessage.Text = "Please enter Currency Code."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'ElseIf lstrCurrencyCode.Length > 4 Then
                '    lblErrorMessage.Text = "Length of Currency Code can not exceed 4 characters."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'End If

                lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtINDActive")
                lstrINDActive = lobjTextBox.Text
                'If lstrINDActive = "" Then
                '    lblErrorMessage.Text = "Please enter IND Active."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'ElseIf IsNumeric(lstrINDActive) = False Then
                '    lblErrorMessage.Text = "IND Active entered is not valid numeric."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'ElseIf CDbl(lstrINDActive) < 0 Or CDbl(lstrINDActive) > 10 Then
                '    lblErrorMessage.Text = "IND Active entered should be between 0 and 2"
                '    lobjTextBox.Focus()
                '    Exit Sub
                'End If

                lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtDaysToSkip")
                lstrDaysToSkip = lobjTextBox.Text

                'If lstrDaysToSkip = "" Then
                '    lblErrorMessage.Text = "Please enter Days To Skip."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'ElseIf Not IsNumeric(lstrDaysToSkip) Then
                '    lblErrorMessage.Text = "Days To Skip entered is not valid numeric."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'ElseIf CDbl(lstrDaysToSkip) < -99.899999 Or CDbl(lstrDaysToSkip) > 99.899999 Then
                '    lblErrorMessage.Text = "Days to Skip entered should be between -99.899999 and +99.899999"
                '    lobjTextBox.Focus()
                '    Exit Sub
                'End If

                lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtFrequency")
                lstrFrequency = lobjTextBox.Text

                'If lstrFrequency = "" Then
                '    lblErrorMessage.Text = "Please enter Frequency."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'ElseIf lstrFrequency <> "d" And lstrFrequency <> "w" And lstrFrequency <> "m" And lstrFrequency <> "y" Then
                '    lblErrorMessage.Text = "Frequency should be between these values : d, w, m, y"
                '    lobjTextBox.Focus()
                '    Exit Sub
                'End If

                lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtFrequencyCount")
                lstrFrequencyCount = lobjTextBox.Text

                'If lstrFrequencyCount = "" Then
                '    lblErrorMessage.Text = "Please enter Frequency Count."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'ElseIf Not IsNumeric(lstrFrequencyCount) Then
                '    lblErrorMessage.Text = "Frequency Count entered is not valid numeric."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'ElseIf CDbl(lstrFrequencyCount) < -99.899999 Or CDbl(lstrFrequencyCount) > 99.899999 Then
                '    lblErrorMessage.Text = "Frequency Count entered should be between -99.899999 and +99.899999"
                '    lobjTextBox.Focus()
                '    Exit Sub
                'End If


                lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtProcessDate")
                lstrProcessDate = lobjTextBox.Text

                'If lstrProcessDate = "" Then
                '    lblErrorMessage.Text = "Please enter Process Date."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'ElseIf Not IsDate(lstrProcessDate) Then
                '    lblErrorMessage.Text = "Process Date entered is not valid."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'End If

                lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtBusinessContact")
                lstrBusinessContact = lobjTextBox.Text


                'If lstrBusinessContact = "" Then
                '    lblErrorMessage.Text = "Please enter Business Contact."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'End If

                lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtFTPDirectory")
                lstrFTPDirectory = lobjTextBox.Text
                'If lstrFTPDirectory = "" Then
                '    lblErrorMessage.Text = "Please enter FTP Directory."
                '    lobjTextBox.Focus()
                '    Exit Sub
                'End If



                ' Build the update query
                lstrQuery = lstrQuery & " UPDATE MC_FILE SET " & _
                            " Currency_Code = '" & lstrCurrencyCode & "'," & _
                             " IND_Active = " & lstrINDActive & "," & _
                            " Days_To_Skip = " & lstrDaysToSkip & "," & _
                            "Frequency = '" & lstrFrequency & "'," & _
                            "Frequency_Count = " & lstrFrequencyCount & "," & _
                            "Last_Schedule_Process_Date = '" & lstrProcessDate & "'," & _
                            "Business_Contact = '" & lstrBusinessContact & "'," & _
                            "FTP_Directory = '" & lstrFTPDirectory & "'," & _
                            " DATE_UPDATE =  GETDATE() " & _
                            " WHERE SQ_MC_ID = " & lstrMCId & "; " & vbCrLf

            Next

            lobjcDataEntity = New cDataEntity
            lobjcMoneyCostUISvc = New cMoneyCostUISvc

            lobjcDataEntity.CommonSQL = lstrQuery
            lobjcDataEntity.ActionID = 2

            Call lobjcMoneyCostUISvc.UpdateMCDetails(lobjcDataEntity)

            lblErrorMessage.ForeColor = Drawing.Color.Green
            lblErrorMessage.Text = "MC Details Updated"

        Catch ex As Exception
            LogError(ex.Message)

            ' XmlErrEntity.ErrorMessage = ex.Message
            lblErrorMessage.Text = ex.Message
            'Response.Redirect("~/ErrorPage.aspx", False)
        Finally
            If Not IsNothing(lobjcMoneyCostUISvc) Then
                'lobjcMoneyCostUISvc.Dispose()
                lobjcMoneyCostUISvc = Nothing
            End If
            lobjcDataEntity = Nothing
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            lblErrorMessage.Text = ""
            lblNoMCDetail.Text = ""
            FillGrid()
        Catch ex As Exception
            LogError(ex.Message)

            ' XmlErrEntity.ErrorMessage = ex.Message
            lblErrorMessage.Text = ex.Message
            ' Response.Redirect("~/ErrorPage.aspx", False)
        End Try
    End Sub
End Class