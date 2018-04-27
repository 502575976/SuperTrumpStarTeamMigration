Public Partial Class UpdateCsvFiles
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        lblErrorMessage.Text = ""
        lblErrorMessage.ForeColor = Drawing.Color.Red
        Try
            If cntrlFileUpload.PostedFile.FileName = "" Then
                lblErrorMessage.Text = "Please Input Source File"
            Else
                If txtDestinationPath.Text = "" Then
                    lblErrorMessage.Text = "Please Input Destination path"
                Else
                    Dim strDestination As String = txtDestinationPath.Text
                    If String.Compare(strDestination.Substring(strDestination.Length), "\") <> 0 Then
                        strDestination = strDestination + "\"
                    End If
                    Dim file As HttpPostedFile = cntrlFileUpload.PostedFile
                    Dim filename As String = GetFileName(file)
                    If System.IO.File.Exists(strDestination & filename) Then
                        System.IO.File.Delete(strDestination & filename)
                    End If
                    file.SaveAs(strDestination & filename)
                    lblErrorMessage.ForeColor = Drawing.Color.Green
                    lblErrorMessage.Text = "File Copied Successfully"
                End If
            End If
        Catch ex As Exception
            lblErrorMessage.Text = "Error:-" & ex.Message        
        End Try
    End Sub
    Private Function GetFileName(ByVal file As HttpPostedFile) As String
        Dim i As Integer = 0, j As Integer = 0
        Dim filename As String

        filename = file.FileName
        Do
            i = filename.IndexOf("\", j + 1)
            If i >= 0 Then
                j = i
            End If
        Loop While i >= 0
        filename = filename.Substring(j + 1, filename.Length - j - 1)

        Return filename
    End Function

   
End Class