Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Namespace BSVICROIConfigEntity
    Public Class ConfigEntity
        Private sLogFilePath As String
        Private sLogger As String

        Public Property LogFilePath() As [String]
            Get
                Return sLogFilePath
            End Get
            Set(ByVal value As [String])
                sLogFilePath = value
            End Set
        End Property
        Public Property LoggerName() As String
            Get
                Return sLogger
            End Get
            Set(ByVal value As String)
                sLogger = value
            End Set
        End Property
    End Class
End Namespace

