<Serializable()> _
Public Class FileInfoEntity

#Region "Private Variables"
    Private _sSource As String
    Private _sDestination As String
    Private _sMove As String
    Private _sQueueName As String
    Private _sBusinessContact As String
    Private _sBody As String
    Private _sCutTicket As String
    Private _sSendNotification As String
    Private _sMCCode As String
    Private _sProcessDates As String
#End Region

#Region "Public Property"

    Public Property Source() As String
        Get
            Return _sSource
        End Get
        Set(ByVal value As String)
            _sSource = value
        End Set
    End Property
    Public Property Destination() As String
        Get
            Return _sDestination
        End Get
        Set(ByVal value As String)
            _sDestination = value
        End Set
    End Property

    Public Property Move() As String
        Get
            Return _sMove
        End Get
        Set(ByVal value As String)
            _sMove = value
        End Set
    End Property

    Public Property QueueName() As String
        Get
            Return _sQueueName
        End Get
        Set(ByVal value As String)
            _sQueueName = value
        End Set
    End Property

    Public Property BusinessContact() As String
        Get
            Return _sBusinessContact
        End Get
        Set(ByVal value As String)
            _sBusinessContact = value
        End Set
    End Property

    Public Property Body() As String
        Get
            Return _sBody
        End Get
        Set(ByVal value As String)
            _sBody = value
        End Set
    End Property


    Public Property CutTicket() As String
        Get
            Return _sCutTicket
        End Get
        Set(ByVal value As String)
            _sCutTicket = value
        End Set
    End Property

    Public Property SendNotification() As String
        Get
            Return _sSendNotification
        End Get
        Set(ByVal value As String)
            _sSendNotification = value
        End Set
    End Property
    Public Property MCCode() As String
        Get
            Return _sMCCode
        End Get
        Set(ByVal value As String)
            _sMCCode = value
        End Set
    End Property
    Public Property ProcessDates() As String
        Get
            Return _sProcessDates
        End Get
        Set(ByVal value As String)
            _sProcessDates = value
        End Set
    End Property

#End Region

End Class
