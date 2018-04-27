Imports System
Imports System.IO
Imports System.text
Imports Microsoft.Win32
Imports BSMoneyCostEntity
Partial Public Class Logout
    Inherits System.Web.UI.Page
    Private mstrRegistryHive As String = "SOFTWARE\\FacilitySettings\\MoneyCost\\FilePath"
    Private mstrRegistryKey As String = "SSOLogOutLink"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim lstrSSOLogOutURL As String
        Try
            Session("UserID") = ""
            Session("ValidUser") = ""

            lstrSSOLogOutURL = ReadRegistry(mstrRegistryHive, mstrRegistryKey)
            Response.Redirect(lstrSSOLogOutURL, False)
        Catch ex As Exception
            'cErrorEntity.ErrorMessage = ex.Message
            Response.Redirect("~/ErrorPage.aspx")
        End Try
    End Sub
    Function ReadRegistry(ByVal astrRegistryHive As String, ByVal astrKeyName As String) As String
        Dim lobjreg As RegistryKey = Registry.LocalMachine

        lobjreg = lobjreg.OpenSubKey(astrRegistryHive, False)
        ReadRegistry = lobjreg.GetValue(astrKeyName)

        lobjreg = Nothing
    End Function
End Class