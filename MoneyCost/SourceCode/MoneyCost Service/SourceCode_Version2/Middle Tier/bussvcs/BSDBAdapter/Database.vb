Imports System.Text
Imports System.Data.SqlClient
Imports Oracle.DataAccess.Client
Imports System.Data.OleDb
Imports System.Configuration.ConfigurationSettings


Public Enum eCommandType
    ecStoredProcedure
    ecSQLText
End Enum

Public Enum eDBType
    cSQLServer
    cOledb
    cOracleServer
End Enum

Public Class Database    
    Private mstrDBNameKey As String
    Private meDBType As eDBType
    Private mstrConnectionString As String
    Private cFULL_MODULE_NAME As String = "CEF Data"


#Region "Contructors"

    Public Sub New(ByVal astrDBNameKey As String)
        Try
            mstrDBNameKey = astrDBNameKey
            meDBType = eDBType.cSQLServer
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub New(ByVal astrDBNameKey As String, ByVal aeDBType As eDBType)
        Try
            mstrDBNameKey = astrDBNameKey
            meDBType = aeDBType
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Public Sub New(ByVal astrDBNameKey As String, ByVal aeDBType As eDBType, ByVal astrConnectionString As String)
        Try
            mstrDBNameKey = astrDBNameKey
            meDBType = aeDBType
            mstrConnectionString = astrConnectionString
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region


#Region "ExecuteDS"
    '================================================================
    'METHOD  :  ExecuteDS
    'PURPOSE :  execute a SQL statement against the database, 
    '           returning a dataset 
    'PARMS   :  ByVal astrSQL As String
    '           byval aeDBCommandType As eCommandType
    '           ByVal ParamArray aobjSQLParams As Object()
    'RETURN  :  DataSet
    '================================================================
    Public Function ExecuteDS(ByVal astrSQL As String, ByVal aeDBCommandType As eCommandType, ByVal ParamArray aobjSQLParams() As Object) As System.Data.DataSet

        Dim lobjSQLData As SQLData = Nothing       
        Dim lobjOracleData As OracleData = Nothing
        Dim lobjDS As Data.DataSet = Nothing
        Dim lobjOleData As OleDBData = Nothing

        Try
            Select Case meDBType
                Case eDBType.cSQLServer
                    lobjSQLData = New SQLData(mstrDBNameKey, mstrConnectionString)
                    lobjDS = lobjSQLData.ExecuteDS(astrSQL, aeDBCommandType, aobjSQLParams)
                Case eDBType.cOracleServer
                    lobjOracleData = New OracleData(mstrDBNameKey)
                    lobjDS = lobjOracleData.ExecuteDS(astrSQL, aeDBCommandType, aobjSQLParams)
                Case eDBType.cOledb
                    lobjOleData = New OleDBData(mstrDBNameKey, mstrConnectionString)
                    Return lobjOleData.ExecuteDataset(astrSQL)
            End Select

            Return lobjDS

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteDS." & lobjSysEx.Source
            Throw
        Finally
            If Not IsNothing(lobjDS) Then
                lobjDS.Dispose()
            End If
            If Not IsNothing(lobjSQLData) Then
                lobjSQLData = Nothing
            End If
            If Not IsNothing(lobjOleData) Then
                lobjOleData = Nothing
            End If
            If Not IsNothing(lobjOracleData) Then
                lobjOracleData = Nothing
            End If
        End Try
    End Function
#End Region


#Region "ExecuteNonQuery"
    '================================================================
    'METHOD  :  ExecuteNonQuery
    'PURPOSE :  execute a SQL statement against the database, 
    'PARMS   :  ByVal astrSQL As String
    'RETURN  :  N/A
    '================================================================
    Public Function ExecuteNonQuery(ByVal astrSQL As String) As Long
        Dim lobjSQLData As SQLData = Nothing
        Dim lobjOleDBData As OleDBData = Nothing
        Dim lobjOracleData As OracleData = Nothing
        Dim llngAffectedRows As Long
        Try
            Select Case meDBType
                Case eDBType.cSQLServer
                    lobjSQLData = New SQLData(mstrDBNameKey, mstrConnectionString)
                    llngAffectedRows = lobjSQLData.ExecuteNonQuery(astrSQL)
                Case eDBType.cOracleServer
                    lobjOracleData = New OracleData(mstrDBNameKey)
                    llngAffectedRows = lobjOracleData.ExecuteNonQuery(astrSQL)
                Case Else
                    lobjOleDBData = New OleDBData(mstrDBNameKey, mstrConnectionString)
                    llngAffectedRows = lobjOleDBData.ExecuteNonQuery(astrSQL)
            End Select

            Return llngAffectedRows

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteNonQuery." & lobjSysEx.Source
            Throw
        Finally
            If Not IsNothing(lobjSQLData) Then
                lobjSQLData = Nothing
            End If
            If Not IsNothing(lobjOleDBData) Then
                lobjOleDBData = Nothing
            End If
            If Not IsNothing(lobjOracleData) Then
                lobjOracleData = Nothing
            End If
        End Try
    End Function
#End Region


#Region "ExecuteScalar"
    '================================================================
    'METHOD  :  ExecuteScalar
    'PURPOSE :  execute a SQL statement against the database, 
    'PARMS   :  ByVal astrSQL As String
    'RETURN  :  single value returned from DB as Object
    '================================================================
    Public Function ExecuteScalar(ByVal astrSQL As String) As Object

        Dim lobjSQLData As SQLData = Nothing
        Dim lobjOleDBData As OleDBData = Nothing
        Dim lobjOracleData As OracleData = Nothing

        Try
            Select Case meDBType
                Case eDBType.cSQLServer
                    lobjSQLData = New SQLData(mstrDBNameKey, mstrConnectionString)
                    Return lobjSQLData.ExecuteScalar(astrSQL)

                Case eDBType.cOracleServer
                    lobjOracleData = New OracleData(mstrDBNameKey)
                    Return lobjOracleData.ExecuteScalar(astrSQL)

                Case Else
                    lobjOleDBData = New OleDBData(mstrDBNameKey)
                    Return lobjOleDBData.ExecuteScalar(astrSQL)

            End Select

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteScalar." & lobjSysEx.Source
            Throw
        Finally
            If Not IsNothing(lobjSQLData) Then
                lobjSQLData = Nothing
            End If
            If Not IsNothing(lobjOleDBData) Then
                lobjOleDBData = Nothing
            End If
            If Not IsNothing(lobjOracleData) Then
                lobjOracleData = Nothing
            End If
        End Try
    End Function
#End Region

    '================================================================
    'METHOD  :  ExecuteReader
    'PURPOSE :  execute a SQL statement against the database, 
    '           returning a DataReader
    'PARMS   :  ByVal astrSQL As String
    '           byval aeDBCommandType As eCommandType
    '           ByVal ParamArray aobjSQLParams As Object()
    'RETURN  :  DataReader
    '================================================================
    Public Function ExecuteDataset(ByVal astrSQL As String, ByVal aeDBCommandType As eCommandType, ByVal ParamArray aobjSQLParams() As Object) As DataSet

        Dim lobjSQLData As SQLData = Nothing
        Dim lobjOracleData As OracleData = Nothing
        Dim lobjOleData As OleDBData = Nothing

        Try
            Select Case meDBType
                Case eDBType.cSQLServer
                    lobjSQLData = New SQLData(mstrDBNameKey, mstrConnectionString)
                    Return lobjSQLData.ExecuteReader(astrSQL)
                Case eDBType.cOracleServer
                    lobjOracleData = New OracleData(mstrDBNameKey)
                    Return lobjOracleData.ExecuteReader(astrSQL)
                Case eDBType.cOledb
                    lobjOleData = New OleDBData(mstrDBNameKey, mstrConnectionString)
                    Return lobjOleData.ExecuteDataset(astrSQL)
            End Select
        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteScalar." & lobjSysEx.Source
            Throw
        Finally
            If Not IsNothing(lobjSQLData) Then
                lobjSQLData = Nothing
            End If
            If Not IsNothing(lobjOracleData) Then
                lobjOracleData = Nothing
            End If
            If Not IsNothing(lobjOleData) Then
                lobjOleData = Nothing
            End If
        End Try
    End Function

    '================================================================
    'METHOD  :  ExecuteXML
    'PURPOSE :  Executes an SQL statement against the database, 
    '           returning an XML string.
    'PARMS   :  ByVal astrSQL As String
    '           byval aeDBCommandType As eCommandType
    '           ByVal ParamArray aobjSQLParams As Object()
    'RETURN  :  XML.
    'COMMENTS : This method is not supported by OracleData class since 
    '           the Oracle Command class does not support the ExecuteXMLReader() method.
    '================================================================
    Public Function ExecuteXML(ByVal astrSQL As String, ByVal aeDBCommandType As eCommandType, ByVal ParamArray aobjSQLParams As Object()) As String
        Dim lobjSQLData As SQLData = Nothing
        Try
            Select Case meDBType
                Case eDBType.cSQLServer
                    lobjSQLData = New SQLData(mstrDBNameKey, mstrConnectionString)
                    Return lobjSQLData.ExecuteXML(astrSQL, aeDBCommandType, aobjSQLParams)
            End Select
        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteScalar." & lobjSysEx.Source
            Throw
        Finally
            If Not IsNothing(lobjSQLData) Then
                lobjSQLData = Nothing
            End If
        End Try

    End Function
End Class