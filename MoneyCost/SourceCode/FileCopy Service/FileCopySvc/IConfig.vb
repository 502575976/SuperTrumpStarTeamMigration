Imports System.Configuration.ConfigurationSettings

Public Class IConfig
    Public Shared ReadOnly Property ConfigFile() As String
        Get
            Return AppSettings("ConfigFile")
        End Get
    End Property
    Public Shared ReadOnly Property ServiceSleepInterval() As Double
        Get
            Return AppSettings("ServiceSleepInterval")
        End Get
    End Property
    Public Shared ReadOnly Property SambhaUID() As String
        Get
            Return AppSettings("SambhaUID")
        End Get
    End Property
    Public Shared ReadOnly Property SambhaPWD() As String
        Get
            Return AppSettings("SambhaPWD")
        End Get
    End Property
    Public Shared ReadOnly Property MappedDrive() As String
        Get
            Return AppSettings("MappedDrive")
        End Get
    End Property
    Public Shared ReadOnly Property InputFileLocation() As String
        Get
            Return AppSettings("InputFileLocation")
        End Get
    End Property
End Class
