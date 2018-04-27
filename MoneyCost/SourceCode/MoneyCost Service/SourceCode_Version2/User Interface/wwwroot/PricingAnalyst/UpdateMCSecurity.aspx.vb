Imports BSMoneyCostEntity
Imports BSMoneyCostBL
Imports System
Imports System.text
Imports Microsoft.Win32
Partial Public Class UpdateMCSecurity
    Inherits System.Web.UI.Page
    Private mViewSearchResults As DataView
    Private mstrRegistryHive As String = "SOFTWARE\\FacilitySettings\\MoneyCost\\FilePath"
    Private mstrRegistryKey As String = "SSOAdminRole"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        lblErrorMessage.ForeColor = Drawing.Color.Red
        If Not IsPostBack Then
            Dim lstrUSER_Role, lstrSSOAdminURL As String
            lstrUSER_Role = Trim(Request.ServerVariables("HTTP_CEFMONEYCOSTROLE"))
            If Not UCase(lstrUSER_Role).Contains("ADMIN") Then
                'Request.Url
                lstrSSOAdminURL = ReadRegistry(mstrRegistryHive, mstrRegistryKey)
                Response.Redirect(lstrSSOAdminURL, False)
            End If

            tblButtons.Visible = False
            chkbox.Visible = False
        End If
    End Sub

    Protected Sub btnSSOID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSSOID.Click
        lblErrorMessage.Text = ""
        FillGrid()
    End Sub
    Function FillGrid()
        Dim lobjcDataEntity As cDataEntity
        Dim lobjcMoneyCostUISvc As cMoneyCostUISvc
        Dim ldsMCDetail As DataSet
        Dim lstrSSOID As String

        Try
            lblErrorMessage.Text = ""

            lobjcDataEntity = New cDataEntity
            lobjcMoneyCostUISvc = New cMoneyCostUISvc
            lstrSSOID = txtSSOID.Text
            If lstrSSOID = "" Then
                lblErrorMessage.Text = "Please enter SSOID."
                txtSSOID.Focus()
                Exit Function
            ElseIf Not IsNumeric(lstrSSOID) Then
                lblErrorMessage.Text = "SSOID entered is not valid numeric."
                txtSSOID.Focus()
                Exit Function
            ElseIf lstrSSOID.Length() <> 9 Then
                lblErrorMessage.Text = "SSOID Should be of Nine Nuemric Digits"
                txtSSOID.Focus()
                Exit Function
            End If
            tblButtons.Visible = True
            chkbox.Visible = True
            lobjcDataEntity.SSOID = txtSSOID.Text
            lobjcDataEntity = lobjcMoneyCostUISvc.GetMCSecurity(lobjcDataEntity)
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

            'XmlErrEntity.ErrorMessage = ex.Message
            lblErrorMessage.Text = ex.Message
            Response.Redirect("~/ErrorPage.aspx", False)
        Finally
            If Not IsNothing(lobjcMoneyCostUISvc) Then
                'lobjcMoneyCostUISvc.Dispose()
                lobjcMoneyCostUISvc = Nothing
            End If
            If Not IsNothing(ldsMCDetail) Then
                'ldsMCDetail.Dispose()
                ldsMCDetail = Nothing
            End If
            lobjcDataEntity = Nothing
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

    Private Sub grdMCDetail_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles grdMCDetail.ItemDataBound
        
    End Sub

    Private Sub chkAllFile_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllFile.CheckedChanged
        Dim lbolchk As Boolean
        If chkAllFile.Checked = True Then
            lbolchk = True
        Else
            lbolchk = False
        End If

        Dim iCounter As Integer
        Dim lobjCheckBox As CheckBox

        For iCounter = 0 To grdMCDetail.Items.Count - 1
            lobjCheckBox = grdMCDetail.Items(iCounter).FindControl("chkMCFile")
            lobjCheckBox.Checked = lbolchk
        Next
    End Sub

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

        Dim lstrMCId As String
        Dim iCounter As Integer
        Dim lstrSSOId As String = ""
        Dim lobjChkBox As CheckBox
        Try
            lstrSSOId = txtSSOID.Text
            lstrQuery = "Delete from MC_Security where SSO_ID=" & lstrSSOId & ";"


            For iCounter = 0 To grdMCDetail.Items.Count - 1
                lstrMCId = grdMCDetail.Items(iCounter).Cells(0).Text
                lobjChkBox = grdMCDetail.Items(iCounter).FindControl("chkMCFile")
                If lobjChkBox.Checked Then
                    ' Build the insert query
                    lstrQuery = lstrQuery & " Insert into MC_Security values( " & _
                               "" & lstrMCId & ", " & lstrSSOId & ",GETDATE()); " & vbCrLf
                End If
            Next

            lobjcDataEntity = New cDataEntity
            lobjcMoneyCostUISvc = New cMoneyCostUISvc

            lobjcDataEntity.CommonSQL = lstrQuery
            lobjcDataEntity.ActionID = 3

            Call lobjcMoneyCostUISvc.UpdateMCDetails(lobjcDataEntity)

            lblErrorMessage.ForeColor = Drawing.Color.Green
            lblErrorMessage.Text = "MC Security Updated"

        Catch ex As Exception
            LogError(ex.Message)

            ' XmlErrEntity.ErrorMessage = ex.Message
            lblErrorMessage.Text = ex.Message
            Response.Redirect("~/ErrorPage.aspx", False)
        Finally
            If Not IsNothing(lobjcMoneyCostUISvc) Then
                'lobjcMoneyCostUISvc.Dispose()
                lobjcMoneyCostUISvc = Nothing
            End If
            lobjcDataEntity = Nothing
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        tblButtons.Visible = False
        chkbox.Visible = False
        grdMCDetail.Visible = False
        txtSSOID.Text = ""
        lblErrorMessage.Text = ""
    End Sub
End Class