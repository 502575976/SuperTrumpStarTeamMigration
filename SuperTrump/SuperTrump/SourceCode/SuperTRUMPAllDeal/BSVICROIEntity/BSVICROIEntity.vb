Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Namespace BSVICROIEntity
    Public Class cSTForAllDealsEntity
        Private m_prmBinary As String
        Private m_prmInputXML As String
        Private m_currentDate As DateTime
        Private m_dtProcess As DateTime
        Private m_folderPath As String
        Private _sCommonSQL As String
        Private _dsOutput As System.Data.DataSet

        Public Property PrmBinary() As [String]
            Get
                Return m_prmBinary
            End Get
            Set(ByVal value As [String])
                m_prmBinary = value
            End Set
        End Property
        Public Property PrmInputXML() As [String]
            Get
                Return m_prmInputXML
            End Get
            Set(ByVal value As [String])
                m_prmInputXML = value
            End Set
        End Property


        Public Property CurrentDate() As DateTime
            Get
                Return m_currentDate
            End Get
            Set(ByVal value As DateTime)
                m_currentDate = value
            End Set
        End Property
        Public Property DtProcess() As DateTime
            Get
                Return m_dtProcess
            End Get
            Set(ByVal value As DateTime)
                m_dtProcess = value
            End Set
        End Property
        Public Property FolderPath() As String
            Get
                Return m_folderPath
            End Get
            Set(ByVal value As String)
                m_folderPath = value
            End Set
        End Property
        Public Property CommonSQL() As String
            Get
                Return _sCommonSQL
            End Get
            Set(ByVal value As String)
                _sCommonSQL = value
            End Set
        End Property
        Public Property OutputDataset() As System.Data.DataSet
            Get
                Return _dsOutput
            End Get
            Set(ByVal value As System.Data.DataSet)
                _dsOutput = value
            End Set
        End Property
    End Class
End Namespace

