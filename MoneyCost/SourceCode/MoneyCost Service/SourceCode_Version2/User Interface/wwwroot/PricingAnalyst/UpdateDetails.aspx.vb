Imports System
Imports System.text
Imports Microsoft.Win32
Imports BSMoneyCostEntity
Imports BSMoneyCostBL
Partial Public Class UpdateDetails
    Inherits System.Web.UI.Page
    Private mViewSearchResults As DataView

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            lblErrorMessage.ForeColor = Drawing.Color.Red

            If Not IsPostBack Then
                lblErrorMessage.Text = ""
                lblNoMCDetail.Text = ""
                hSSOId.Value = ""

                tblApplyAll.Visible = False
                tblIndexRate.Visible = False

                btnSave.Visible = False
                btnCancel.Visible = False

                GetMCFiles()

            End If

        Catch ex As Exception
            LogError(ex.Message)

            ' XmlErrEntity.ErrorMessage = ex.Message
            Response.Redirect("~/ErrorPage.aspx", False)
        End Try
    End Sub

    Public Function GetMCFiles()
        Dim lobjcInputDataEntity As cDataEntity
        Dim lobjcOutputDataEntity As cDataEntity
        Dim lobjcMoneyCostUISvc As cMoneyCostUISvc
        Dim lstrPageInfoRequestXML As String
        Dim lstrPageInfoResponseXML As String
        Dim lstrUSER_GESSOUID As String
        Dim lobjPageInfoDOM As New System.Xml.XmlDocument    ''DOM object to load Page Info XML
        Dim lobjMCFileNodeList As System.Xml.XmlNodeList      'Node List object for MC_FILESet
        Dim lobjNode As System.Xml.XmlNode
        Try

            lobjcInputDataEntity = New cDataEntity
            lobjcOutputDataEntity = New cDataEntity
            lobjcMoneyCostUISvc = New cMoneyCostUISvc
            'lstrUSER_GESSOUID = "2ee87ed4-95af-11d6-a612-00d0b785330f"
            'lstrUSER_GESSOUID = "ca2b41ba-95af-11d6-a612-00d0b785330f"
            lstrUSER_GESSOUID = Trim(Request.ServerVariables("HTTP_GESSOUID"))

            'build request XML for fetching all MC File list
            lstrPageInfoRequestXML = "<MC_FILES_REQUEST>" & _
                    "<USER_GESSOUID>" & lstrUSER_GESSOUID & "</USER_GESSOUID>" & _
                   "</MC_FILES_REQUEST>"

            ''Populating Entity with the request XML
            lobjcInputDataEntity.OutputString = lstrPageInfoRequestXML

            ''Calling BL function to get list of MC Files
            lobjcOutputDataEntity = lobjcMoneyCostUISvc.GetMCFiles(lobjcInputDataEntity)

            lstrPageInfoResponseXML = lobjcOutputDataEntity.OutputString

            lobjPageInfoDOM.LoadXml(lstrPageInfoResponseXML)
            lobjMCFileNodeList = lobjPageInfoDOM.SelectNodes("/USER_MC_FILE_RESPONSE/MC_FILE_RESPONSE/MC_FILESet/MC_FILE")

            hSSOId.Value = Trim(lobjPageInfoDOM.GetElementsByTagName("uid").Item(0).InnerText)

            lblUserName.Text = Trim(lobjPageInfoDOM.GetElementsByTagName("givenname").Item(0).InnerText)
            ''Binding MC FIle list with the combo... 

            cmbMoneyCostFiles.Items.Clear()

            For Each lobjNode In lobjMCFileNodeList
                cmbMoneyCostFiles.Items.Add(New ListItem(lobjNode.SelectSingleNode("MONEY_COST_FILE").InnerXml, lobjNode.SelectSingleNode("SQ_MC_ID").InnerXml))
            Next

            cmbMoneyCostFiles.Items.Insert(0, "Select a MoneyCost File")
            cmbMoneyCostFiles.Items(0).Value = 0


        Catch ex As Exception
            LogError(ex.Message)

            'XmlErrEntity.ErrorMessage = ex.Message
            Response.Redirect("~/ErrorPage.aspx", False)
        Finally
            If Not IsNothing(lobjPageInfoDOM) Then
                lobjPageInfoDOM = Nothing
            End If
            If Not IsNothing(lobjcMoneyCostUISvc) Then
                'lobjcMoneyCostUISvc.Dispose()
                lobjcMoneyCostUISvc = Nothing
            End If
            lobjcInputDataEntity = Nothing
            lobjcOutputDataEntity = Nothing

        End Try
    End Function
    Private Sub cmbMoneyCostFiles_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbMoneyCostFiles.SelectedIndexChanged
        Dim lobjcDataEntity As cDataEntity
        Dim lobjcMoneyCostUISvc As cMoneyCostUISvc
        Dim ldsMCDetail As DataSet
        Dim lstrMCID As String

        Try
            lblErrorMessage.Text = ""
            lstrMCID = cmbMoneyCostFiles.SelectedValue
            If lstrMCID = "" Or lstrMCID = "0" Then
                tblApplyAll.Visible = False
                btnSave.Visible = False
                btnCancel.Visible = False
            Else
                lobjcDataEntity = New cDataEntity
                lobjcMoneyCostUISvc = New cMoneyCostUISvc
                lobjcDataEntity.MoneyCostID = lstrMCID
                lobjcDataEntity = lobjcMoneyCostUISvc.GetMCFileDetails(lobjcDataEntity)
                ldsMCDetail = lobjcDataEntity.OutputDataSet
                If ldsMCDetail.Tables(0).Rows.Count > 0 Then
                    tblIndexRate.Visible = True
                    tblApplyAll.Visible = True
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

                    tblApplyAll.Visible = False
                    btnSave.Visible = False
                    btnCancel.Visible = True

                End If
            End If
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
            If Not IsNothing(ldsMCDetail) Then
                'ldsMCDetail.Dispose()
                ldsMCDetail = Nothing
            End If
            lobjcDataEntity = Nothing
        End Try
    End Sub
    Protected Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim lobjcDataEntity As cDataEntity
        Dim lobjcMoneyCostUISvc As cMoneyCostUISvc
        Dim lobjTextBox As New TextBox
        Dim lstrQuery As String
        Dim lstrAdder As String
        Dim lstrDate As String
        Dim iCounter As Integer
        Dim lstrIndexId As String
        Dim lstrSSOIdUpdate As String = ""
        Try
            lstrQuery = ""

            lstrSSOIdUpdate = hSSOId.Value

            For iCounter = 0 To grdMCDetail.Items.Count - 1
                lstrIndexId = grdMCDetail.Items(iCounter).Cells(0).Text

                lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtAdderRate")
                lstrAdder = lobjTextBox.Text
                If lstrAdder = "" Then
                    lblErrorMessage.Text = "Please enter Adder Rate."
                    lobjTextBox.Focus()
                    Exit Sub
                ElseIf IsNumeric(lstrAdder) = False Then
                    lblErrorMessage.Text = "Adder Rate entered is not valid numeric."
                    lobjTextBox.Focus()
                    Exit Sub
                ElseIf CDbl(lstrAdder) < -99.899999 Or CDbl(lstrAdder) > 99.899999 Then
                    lblErrorMessage.Text = "Adder Rate entered should be between -99.899999 and +99.899999"
                    lobjTextBox.Focus()
                    Exit Sub
                End If

                lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtEffectiveDate")
                lstrDate = lobjTextBox.Text

                If lstrDate = "" Then
                    lblErrorMessage.Text = "Please enter Effective Date."
                    lobjTextBox.Focus()
                    Exit Sub
                ElseIf Not IsDate(lstrDate) Then
                    lblErrorMessage.Text = "Effective Date entered is not valid."
                    lobjTextBox.Focus()
                    Exit Sub
                End If

                ' Build the update query
                lstrQuery = lstrQuery & " UPDATE INDEX_RATES SET AMT_ADDER = " & lstrAdder & ", " & _
                                        " SSO_UPDATE = '" & lstrSSOIdUpdate & "' , " & _
                                        " DATE_EFFECTIVE = '" & lstrDate & "', DATE_UPDATE = " & _
                                        " GETDATE() WHERE SQ_INDEX_ID = " & lstrIndexId & "; " & vbCrLf
            Next

            lobjcDataEntity = New cDataEntity
            lobjcMoneyCostUISvc = New cMoneyCostUISvc

            lobjcDataEntity.CommonSQL = lstrQuery
            lobjcDataEntity.ActionID = 1
            Call lobjcMoneyCostUISvc.UpdateMCDetails(lobjcDataEntity)

            lblErrorMessage.ForeColor = Drawing.Color.Green
            lblErrorMessage.Text = "Index Rate updated for " & cmbMoneyCostFiles.SelectedItem.Text

        Catch ex As Exception
            LogError(ex.Message)

            ' XmlErrEntity.ErrorMessage = ex.Message
            Response.Redirect("~/ErrorPage.aspx", False)
        Finally
            If Not IsNothing(lobjcMoneyCostUISvc) Then
                'lobjcMoneyCostUISvc.Dispose()
                lobjcMoneyCostUISvc = Nothing
            End If
            lobjcDataEntity = Nothing
        End Try
    End Sub
    Protected Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            lblErrorMessage.Text = ""
            lblNoMCDetail.Text = ""
            hSSOId.Value = ""

            tblApplyAll.Visible = False
            tblIndexRate.Visible = False

            btnSave.Visible = False
            btnCancel.Visible = False
            GetMCFiles()
        Catch ex As Exception
            LogError(ex.Message)

            ' XmlErrEntity.ErrorMessage = ex.Message
            Response.Redirect("~/ErrorPage.aspx", False)
        End Try
    End Sub

    Protected Sub btnSaveAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveAll.Click
        Dim lobjTextBox As New TextBox
        Dim lstrAdder As String
        Dim lstrDate As String
        Dim iCounter As Integer
        Try
            lblErrorMessage.Text = ""
            lstrAdder = Trim(txtAdderRateCommon.Text)
            lstrDate = Trim(txtEffectiveDateCommon.Text)

            If lstrAdder = "" Then
                lblErrorMessage.Text = "Please enter Common Adder Rate."
            ElseIf IsNumeric(lstrAdder) = False Then
                lblErrorMessage.Text = "Common Adder Rate entered is not valid numeric."
            ElseIf CDbl(lstrAdder) < -99.899999 Or CDbl(lstrAdder) > 99.899999 Then
                lblErrorMessage.Text = "Common Adder Rate entered should be between -99.899999 and +99.899999"
            ElseIf lstrDate = "" Then
                lblErrorMessage.Text = "Please enter Common Effective Date."
            ElseIf Not IsDate(lstrDate) Then
                lblErrorMessage.Text = "Common Effective Date entered is not valid."
            Else
                For iCounter = 0 To grdMCDetail.Items.Count - 1
                    lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtAdderRate")
                    lobjTextBox.Text = lstrAdder
                    lobjTextBox = grdMCDetail.Items(iCounter).FindControl("txtEffectiveDate")
                    lobjTextBox.Text = lstrDate
                Next
            End If

        Catch ex As Exception
            LogError(ex.Message)

            'XmlErrEntity.ErrorMessage = ex.Message
            Response.Redirect("~/ErrorPage.aspx", False)
        Finally
            If Not IsNothing(lobjTextBox) Then
                'lobjTextBox.Dispose()
                lobjTextBox = Nothing
            End If
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
End Class