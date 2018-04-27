<Serializable()> _
Public Class ServiceEntity

    Private _sUserName As String
    Private _sPassword As String
    Private _sDomain As String


    Public Property UserName() As String
        Get
            Return _sUserName
        End Get
        Set(ByVal value As String)
            _sUserName = value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return _sPassword
        End Get
        Set(ByVal value As String)
            _sPassword = value
        End Set
    End Property

    Public Property Domain() As String
        Get
            Return _sDomain
        End Get
        Set(ByVal value As String)
            _sDomain = value
        End Set
    End Property

End Class
