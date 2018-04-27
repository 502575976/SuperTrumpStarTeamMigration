<Serializable()> _
Public Class cDataEntity

#Region "Private Variables"
    Private _sCommonSQL As String
    Private _sWrkDirectory As String
    Private _sUserSSOID As String
    Private _sMoneyCostID As String
    Private _sXmlTagName As String
    Private _dProcessDate As Date
    Private _dstOutputDataSet As DataSet
    Private _sOutputString As String
    'Private _xmlInputXml As Xml.XmlDocument
    Private _sAstrFileName As String
    Private _rdrCsvOutput As DataTable
    Private _sActionID As String
    Private _sCurrencyCode As String
    Private _sYieldCurveType As String
    Private _sTermPeriod As String
    Private _bQuoteReplacement As Boolean

    ' Added for Treasury Assessment
    Private _sCostTypes As String
#End Region

#Region "Public Property"

    Public Property CommonSQL() As String
        Get
            Return _sCommonSQL
        End Get
        Set(ByVal value As String)
            _sCommonSQL = value
        End Set
    End Property

    Public Property UserSSOID() As String
        Get
            Return _sUserSSOID
        End Get
        Set(ByVal value As String)
            _sUserSSOID = value
        End Set
    End Property
    Public Property MoneyCostID() As String
        Get
            Return _sMoneyCostID
        End Get
        Set(ByVal value As String)
            _sMoneyCostID = value
        End Set
    End Property

    Public Property XmlTagName() As String
        Get
            Return _sXmlTagName
        End Get
        Set(ByVal value As String)
            _sXmlTagName = value
        End Set
    End Property

    Public Property ProcessDate() As Date
        Get
            Return _dProcessDate
        End Get
        Set(ByVal value As Date)
            _dProcessDate = value
        End Set
    End Property
    Public Property OutputDataSet() As DataSet
        Get
            Return _dstOutputDataSet
        End Get
        Set(ByVal value As DataSet)
            _dstOutputDataSet = value
        End Set
    End Property
    Public Property OutputString() As String
        Get
            Return _sOutputString
        End Get
        Set(ByVal value As String)
            _sOutputString = value
        End Set
    End Property

    'Public Property InputXml() As Xml.XmlDocument
    '    Get
    '        Return _xmlInputXml
    '    End Get
    '    Set(ByVal value As Xml.XmlDocument)
    '        _xmlInputXml = value
    '    End Set
    'End Property
    Public Property astrFileName() As String
        Get
            Return _sAstrFileName
        End Get
        Set(ByVal value As String)
            _sAstrFileName = value
        End Set
    End Property
    Public Property WrkDirectory() As String
        Get
            Return _sWrkDirectory
        End Get
        Set(ByVal value As String)
            _sWrkDirectory = value
        End Set
    End Property

    Public Property CsvOutput() As DataTable
        Get
            Return _rdrCsvOutput
        End Get
        Set(ByVal value As DataTable)
            _rdrCsvOutput = value
        End Set
    End Property
    Public Property ActionID() As String
        Get
            Return _sActionID
        End Get
        Set(ByVal value As String)
            _sActionID = value
        End Set
    End Property
    Public Property SSOID() As String
        Get
            Return _sActionID
        End Get
        Set(ByVal value As String)
            _sActionID = value
        End Set
    End Property
    Public Property QuoteReplacement() As Boolean
        Get
            Return _bQuoteReplacement
        End Get
        Set(ByVal value As Boolean)
            _bQuoteReplacement = value
        End Set
    End Property
    Public Property CurrencyCode() As String
        Get
            Return _sCurrencyCode
        End Get
        Set(ByVal value As String)
            _sCurrencyCode = value
        End Set
    End Property
    Public Property YieldCurveType() As String
        Get
            Return _sYieldCurveType
        End Get
        Set(ByVal value As String)
            _sYieldCurveType = value
        End Set
    End Property
    Public Property TermPeriod() As String
        Get
            Return _sTermPeriod
        End Get
        Set(ByVal value As String)
            _sTermPeriod = value
        End Set
    End Property

    'Added for Treasury Assessment
    Public Property CostTypes() As String
        Get
            Return _sCostTypes
        End Get
        Set(ByVal value As String)
            _sCostTypes = value
        End Set
    End Property
#End Region

End Class
