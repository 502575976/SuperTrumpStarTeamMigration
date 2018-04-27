Imports System
Imports System.IO
Imports BSMoneyCostEntity
Partial Class ErrorPage
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            'txtErrorDesc.Value = XmlErrEntity.ErrorMessage()

            lblErrorDescription.Text = "It appears that an unexpected event has happened.  Please close your browser and open a new window to log back into the Application."
            btnRetry.Attributes.Add("onclick", "return retry11();")

        Catch ex As Exception

        End Try
    End Sub
End Class
