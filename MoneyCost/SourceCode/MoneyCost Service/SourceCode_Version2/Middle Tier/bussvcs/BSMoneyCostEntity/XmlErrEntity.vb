<Serializable()> _
Public Class XmlErrEntity

#Region "Private Variables"
    Private _sErrNbr As String
    Private _sErrSource As String
    Private _sErrDesc As String
    Private _sErrorShowUser As String
    Private _xIXMLDOMNode As Xml.XmlNode
    Private _xElementXPath As String
    Private _xXmlDoc As String
    Private _xXslFilePath As String
    Private Shared _xErrorMessage As String
#End Region

#Region "Public Property"
    Public Shared Property ErrorMessage() As String
        Get
            Return _xErrorMessage
        End Get
        Set(ByVal Value As String)
            _xErrorMessage = Value
        End Set
    End Property
    Public Property ErrNbr() As String
        Get
            Return _sErrNbr
        End Get
        Set(ByVal value As String)
            _sErrNbr = value
        End Set
    End Property

    Public Property ErrSource() As String
        Get
            Return _sErrSource
        End Get
        Set(ByVal value As String)
            _sErrSource = value
        End Set
    End Property
    Public Property ErrDesc() As String
        Get
            Return _sErrDesc
        End Get
        Set(ByVal value As String)
            _sErrDesc = value
        End Set
    End Property
    Public Property ErrorShowUser() As String
        Get
            Return _sErrorShowUser
        End Get
        Set(ByVal value As String)
            _sErrorShowUser = value
        End Set
    End Property
    Public Property IXMLDOMNode() As Xml.XmlNode
        Get
            Return _xIXMLDOMNode
        End Get
        Set(ByVal value As Xml.XmlNode)
            _xIXMLDOMNode = value
        End Set
    End Property
    Public Property ElementXPath() As String
        Get
            Return _xElementXPath
        End Get
        Set(ByVal value As String)
            _xElementXPath = value
        End Set
    End Property
    Public Property XmlDoc() As String
        Get
            Return _xXmlDoc
        End Get
        Set(ByVal value As String)
            _xXmlDoc = value
        End Set
    End Property
    Public Property XslFilePath() As String
        Get
            Return _xXslFilePath
        End Get
        Set(ByVal value As String)
            _xXslFilePath = value
        End Set
    End Property
#End Region

End Class
