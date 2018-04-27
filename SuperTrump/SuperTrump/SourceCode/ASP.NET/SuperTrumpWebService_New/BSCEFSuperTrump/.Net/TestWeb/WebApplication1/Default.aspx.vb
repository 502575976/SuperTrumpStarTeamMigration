Imports WebApplication1.SuperTrumpWebService
Partial Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim iobj As IClientServiceSoapPort
        iobj = New IClientServiceSoapPort

        iobj.Url = "https://bscefsupertrump.qa.comfin.ge.com/BSCEFSuperTRUMPNET/SupertrumpService.asmx"

        Dim iobjC As New System.Net.NetworkCredential
        'iobjC.Domain = "comfin"
        iobjC.UserName = "501244408"
        iobjC.Password = "mount23ain"

        'iobjC.UserName = "501269864"
        'iobjC.Password = "Pa55word"

        iobj.UseDefaultCredentials = False
        iobj.Credentials = iobjC

        Response.Write(iobj.Ping())
    End Sub
End Class
