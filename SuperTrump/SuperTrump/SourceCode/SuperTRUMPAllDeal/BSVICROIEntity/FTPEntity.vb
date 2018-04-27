<Serializable()> _
Public Class FTPEntity

#Region "Private Variables"
    Private _sFTPLocation As String
    Private _sFTPUser As String
    Private _sFTPPassword As String
    Private _sUNCPath As String
    Private _sHOST As String
    Private _Directory As String
    Private _sFilterString As String
    Private _sTransferType As String
#End Region

#Region "Public Property"

    Public Property FTPLocation() As String
        Get
            Return _sFTPLocation
        End Get
        Set(ByVal value As String)
            _sFTPLocation = value
        End Set
    End Property

    Public Property FTPUser() As String
        Get
            Return _sFTPUser
        End Get
        Set(ByVal value As String)
            _sFTPUser = value
        End Set
    End Property
    Public Property FTPPassword() As String
        Get
            Return _sFTPPassword
        End Get
        Set(ByVal value As String)
            _sFTPPassword = value
        End Set
    End Property

    Public Property UNCPath() As String
        Get
            Return _sUNCPath
        End Get
        Set(ByVal value As String)
            _sUNCPath = value
        End Set
    End Property
    Public Property HOST() As String
        Get
            Return _sHOST
        End Get
        Set(ByVal value As String)
            _sHOST = value
        End Set
    End Property
    Public Property Directory() As String
        Get
            Return _Directory
        End Get
        Set(ByVal value As String)
            _Directory = value
        End Set
    End Property
    Public Property FilterString() As String
        Get
            Return _sFilterString
        End Get
        Set(ByVal value As String)
            _sFilterString = value
        End Set
    End Property
    Public Property TransferType() As String
        Get
            Return _sTransferType
        End Get
        Set(ByVal value As String)
            _sTransferType = value
        End Set
    End Property

#End Region

End Class
