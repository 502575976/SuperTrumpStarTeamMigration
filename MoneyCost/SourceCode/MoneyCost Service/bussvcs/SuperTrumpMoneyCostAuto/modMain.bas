Attribute VB_Name = "modMain"
Sub Main()
On Error GoTo FormLoad_Error
    Dim objBSMoneyCostAuto As New BSMoneyCostAuto.IBSMoneyCostAutoService
    objBSMoneyCostAuto.ExecuteServiceFlow
    
FormLoad_Cleanup:
        Set objBSMoneyCostAuto = Nothing
    
Exit Sub

FormLoad_Error:
    App.LogEvent (Err.Description)
    
    GoTo FormLoad_Cleanup
End Sub
