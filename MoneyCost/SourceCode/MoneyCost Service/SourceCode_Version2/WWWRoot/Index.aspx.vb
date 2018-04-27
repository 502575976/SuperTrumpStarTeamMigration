Imports BSMoneyCostEntity
Partial Public Class Index
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Response.Redirect("~/PricingAnalyst/UpdateDetails.aspx", False)
        Catch ex As Exception
            Response.Redirect("~/ErrorPage.aspx", False)
        End Try
    End Sub

End Class