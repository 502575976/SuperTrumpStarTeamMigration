Imports BSMoneyCostEntity
Imports System.Reflection
Imports System.EnterpriseServices
Imports System.Runtime.InteropServices
Public Interface IMoneyCostAutoUISvc
    Sub ExecuteServiceFlow()
    Function ping() As String
    Function Test() As String
End Interface
<JustInTimeActivation(), _
EventTrackingEnabled(), _
ClassInterface(ClassInterfaceType.None), _
Transaction(TransactionOption.NotSupported, Isolation:=TransactionIsolationLevel.Serializable, Timeout:=120), _
ComponentAccessControl(True)> _
Public Class cMoneyCostAutoUISvc
    Inherits ServicedComponent
    Implements IMoneyCostAutoUISvc
    Dim STLogger As log4net.ILog
    Public Sub SetLog4Net()
        Try
            If log4net.LogManager.GetRepository.Configured = False Then
                log4net.Config.XmlConfigurator.ConfigureAndWatch(New System.IO.FileInfo(GetConfigurationKey("MoneyCostLog4Net")))
            End If
            STLogger = log4net.LogManager.GetLogger("MoneyCost")
        Catch ex As Exception
            Throw
        End Try

    End Sub
    '================================================================
    'METHOD  : ExecuteServiceFlow
    'PURPOSE : Main Controller procedure for the service flow
    'PARMS   : NONE
    'RETURN  : NONE
    '================================================================
    <AutoComplete()> _
    Public Sub ExecuteServiceFlow() Implements IMoneyCostAutoUISvc.ExecuteServiceFlow
        Dim objMoneyCoStAutoSvc As New cMoneyCostAutoSvc

        SetLog4Net()
        Try
            objMoneyCoStAutoSvc.ExecuteServiceFlow()

        Catch ex As Exception
            STLogger.Error(ex.Message, ex)
        End Try

        Try
            objMoneyCoStAutoSvc.ProcessTreasuryAssessment()
        Catch ex As Exception
            STLogger.Error(ex.Message, ex)

        Finally
            objMoneyCoStAutoSvc = Nothing
        End Try

    End Sub
    <AutoComplete()> _
    Public Function UpdateFile(ByVal lobjEntity As cDataEntity) As cDataEntity
        Dim objMoneyCoStAutoSvc As New cMoneyCostAutoSvc
        Try
            Return objMoneyCoStAutoSvc.UpdateFile(lobjEntity)
        Catch ex As Exception
            Return Nothing
            Throw
        Finally
            UpdateFile = Nothing
            objMoneyCoStAutoSvc.Dispose()
            objMoneyCoStAutoSvc = Nothing
            lobjEntity = Nothing
        End Try
    End Function

    <AutoComplete()> _
    Public Function ping() As String Implements IMoneyCostAutoUISvc.ping
        Dim objMoneyCoStAutoSvc As New cMoneyCostAutoSvc
        Try
            Return objMoneyCoStAutoSvc.Ping()
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(objMoneyCoStAutoSvc) Then
                objMoneyCoStAutoSvc.Dispose()
                objMoneyCoStAutoSvc = Nothing
            End If
        End Try
    End Function
    <AutoComplete()> _
    Public Function Test() As String Implements IMoneyCostAutoUISvc.Test
        Dim objMoneyCoStAutoSvc As New cMoneyCostAutoSvc
        Try
            Test = objMoneyCoStAutoSvc.Test()
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(objMoneyCoStAutoSvc) Then
                objMoneyCoStAutoSvc.Dispose()
                objMoneyCoStAutoSvc = Nothing
            End If

        End Try
    End Function
    <AutoComplete()> _
    Public Function MapDrive(ByVal lobjFtpEntity As FTPEntity) As cDataEntity
        Dim objMoneyCoStAutoSvc As New cMoneyCostAutoSvc
        Try
            Return objMoneyCoStAutoSvc.MapDrive(lobjFtpEntity)
        Catch ex As Exception
            lobjFtpEntity = Nothing
            Return Nothing
            Throw
        Finally
            If Not IsNothing(objMoneyCoStAutoSvc) Then
                objMoneyCoStAutoSvc.Dispose()
                objMoneyCoStAutoSvc = Nothing
            End If
        End Try
    End Function
End Class
