Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim objSuper As SuperTRUMPAllDealService.SuperTRUMPAllDealServiceClient = New SuperTRUMPAllDealService.SuperTRUMPAllDealServiceClient()
            objSuper.ClientCredentials.Windows.ClientCredential.UserName = "992000016"
            objSuper.ClientCredentials.Windows.ClientCredential.Password = "gf6eD8xp"
            objSuper.ClientCredentials.Windows.ClientCredential.Domain = "Comfin"
            objSuper.InnerChannel.OperationTimeout = New TimeSpan(0, 10, 0)
            objSuper.ExecuteServiceFlow()
            MessageBox.Show("SuperTRUMPForAllDeal Process Completed Successfully.")
            Application.Exit()
        Catch ex As Exception
            Application.Exit()
        End Try
    End Sub
End Class
