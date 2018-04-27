Imports System.Data.SqlClient
Imports System.Configuration.ConfigurationSettings
Public Class SQLData
    Private Const cFULL_MODULE_NAME As String = "DBAdaptor.SQLData"

    Private mobjSQLDBConn As SqlConnection
    Private mstrConnString As String
    Private mstrDBNameKey As String

#Region "Constructors"
    'Sub New(ByVal astrDBNameKey As String)
    '    Try

    '        If IsNothing(astrDBNameKey) Then
    '            Throw New Exception("Connection string for this DBNameKey cannot be found")
    '        Else
    '            mstrConnString = AppSettings(astrDBNameKey.ToString)
    '            mstrDBNameKey = astrDBNameKey
    '        End If

    '    Catch lobjSysEx As System.Exception
    '        Throw lobjSysEx
    '    End Try

    'End Sub
    Sub New(ByVal astrDBNameKey As String, ByVal astrConnectionString As String)
        Try

            If IsNothing(astrDBNameKey) Then
                Throw New Exception("Connection string for this DBNameKey cannot be found")
            Else
                mstrConnString = astrConnectionString
                mstrDBNameKey = astrDBNameKey
            End If

        Catch lobjSysEx As System.Exception
            Throw
        End Try

    End Sub
#End Region
#Region "Open/Close Connections"

    Private Function DBConnect() As Object
        Try
            SQLDBConnect()
        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":DBConnect." & lobjSysEx.Source
            Throw
        End Try
    End Function

    Private Function DBDisconnect() As Object
        Try
            If Not IsNothing(mobjSQLDBConn) Then
                SQLDBDisConnect()
            End If

        Catch lobjSysEx As System.Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":SQLDBConnect." & lobjSysEx.Source
            Throw
        End Try
    End Function
    '================================================================
    'METHOD  :  SQLDBConnect
    'PURPOSE :  uses the .Net SQLClient provider to connect to 
    '           the database
    'PARMS   :  N/A
    'RETURN  :  NONE
    '================================================================

    Private Sub SQLDBConnect()
        Try
            'open the connection to SQL svr db
            mobjSQLDBConn = New SqlConnection
            mobjSQLDBConn.ConnectionString = mstrConnString
            mobjSQLDBConn.Open()

        Catch lobjSysEx As System.Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":SQLDBConnect." & lobjSysEx.Source
            Throw
        End Try

    End Sub

    '================================================================
    'METHOD  :  SQLDBDisConnect
    'PURPOSE :  uses the .Net SQLClient provider to disconnect from 
    '           the database if the connection is open
    'PARMS   :  N/A
    'RETURN  :  NONE
    '================================================================

    Private Sub SQLDBDisConnect()
        Try
            'close the connection to the db if it is not closed
            If mobjSQLDBConn.State <> ConnectionState.Closed Then
                mobjSQLDBConn.Close()
                mobjSQLDBConn.Dispose()
            End If

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":SQLDBDisConnect." & lobjSysEx.Source
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

    Public Function ExecuteDS(ByVal astrSQL As String, ByVal aeDBCommandType As eCommandType, _
                                ByVal ParamArray aobjSQLParams As Object()) As Data.DataSet

        Dim lobjSQLDA As Data.SqlClient.SqlDataAdapter
        Dim lobjSQLCmd As SqlCommand
        Dim lobjDS As Data.DataSet

        Try

            DBConnect()
            lobjDS = New DataSet

            'check command type
            Select Case aeDBCommandType

                'stored proc
                Case eCommandType.ecStoredProcedure
                    For i As Integer = 0 To aobjSQLParams.Length - 1
                        If i > 0 Then
                            astrSQL = astrSQL & ","
                        End If

                        astrSQL = astrSQL & " '" & aobjSQLParams(i) & "'"

                    Next

                    'sql text
                Case eCommandType.ecSQLText

            End Select

            lobjSQLDA = New SqlClient.SqlDataAdapter(astrSQL, mobjSQLDBConn)
            lobjSQLDA.SelectCommand.CommandTimeout = 0
            lobjSQLDA.Fill(lobjDS)

            Return lobjDS

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteDS." & lobjSysEx.Source
            Throw
        Finally

            DBDisconnect()

            If Not IsNothing(lobjSQLDA) Then
                lobjSQLDA.Dispose()
            End If

            If Not IsNothing(lobjDS) Then
                lobjDS.Dispose()
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

        Dim lobjSQLCmd As SqlClient.SqlCommand
        Dim llngAffectedRows As Long = Nothing
        'Dim lobjTransCon As SqlTransaction

        Try
            DBConnect()
            lobjSQLCmd = New SqlCommand(astrSQL)
            lobjSQLCmd.Connection = mobjSQLDBConn
            lobjSQLCmd.CommandType = CommandType.Text

            'execute sql passed in
            lobjSQLCmd.CommandTimeout = 0


            'lobjTransCon = mobjSQLDBConn.BeginTransaction(IsolationLevel.ReadCommitted)
            'lobjSQLCmd.Transaction = lobjTransCon
            llngAffectedRows = lobjSQLCmd.ExecuteNonQuery()
            'lobjTransCon.Commit()
            Return llngAffectedRows

        Catch lobjSysEx As Exception
            'lobjTransCon.Rollback()
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteNonQuery." & lobjSysEx.Source
            Throw
        Finally

            DBDisconnect()

            If Not IsNothing(lobjSQLCmd) Then
                lobjSQLCmd.Dispose()
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
        Dim lobjSQLCmd As SqlClient.SqlCommand

        Try
            DBConnect()
            lobjSQLCmd = New SqlCommand(astrSQL)
            lobjSQLCmd.Connection = mobjSQLDBConn
            lobjSQLCmd.CommandType = CommandType.Text
            lobjSQLCmd.CommandTimeout = 0
            'execute sql passed in
            Return lobjSQLCmd.ExecuteScalar()

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteScalar." & lobjSysEx.Source
            Throw

        Finally
            DBDisconnect()

            If Not IsNothing(lobjSQLCmd) Then
                lobjSQLCmd.Dispose()
            End If

        End Try
    End Function

    '================================================================
    'METHOD  :  ExecuteReader
    'PURPOSE :  execute a SQL statement against the database, 
    'PARMS   :  ByVal astrSQL As String
    'RETURN  :  single value returned from DB as Object
    '================================================================

    Public Function ExecuteReader(ByVal astrSQL As String) As IDataReader
        Dim lobjSQLCmd As SqlClient.SqlCommand

        Try
            DBConnect()
            lobjSQLCmd = New SqlCommand(astrSQL)
            lobjSQLCmd.Connection = mobjSQLDBConn
            lobjSQLCmd.CommandType = CommandType.Text
            'execute sql passed in
            lobjSQLCmd.CommandTimeout = 0
            Return lobjSQLCmd.ExecuteReader()

        Catch lobjSysEx As Exception
            DBDisconnect()
            If Not IsNothing(lobjSQLCmd) Then
                lobjSQLCmd.Dispose()
            End If
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteScalar." & lobjSysEx.Source
            Throw
        Finally

        End Try
    End Function

    Public Function ExecuteXML(ByVal astrSQL As String, ByVal aeDBCommandType As eCommandType, ByVal ParamArray aobjSQLParams As Object()) As String
        Dim lobjSQLCmd As SqlCommand
        Dim lobjXMLReader As System.Xml.XmlReader
        Dim lobjSB As System.Text.StringBuilder

        Try
            Select Case aeDBCommandType
                Case eCommandType.ecStoredProcedure
                    For i As Integer = 0 To aobjSQLParams.Length - 1
                        If i > 0 Then
                            astrSQL = astrSQL & ","
                        End If

                        astrSQL = astrSQL & " '" & aobjSQLParams(i) & "'"
                    Next
            End Select

            DBConnect()
            lobjSQLCmd = New SqlCommand(astrSQL, mobjSQLDBConn)
            lobjSQLCmd.CommandTimeout = 0
            lobjXMLReader = lobjSQLCmd.ExecuteXmlReader
            lobjXMLReader.MoveToContent()

            lobjSB = New System.Text.StringBuilder

            Do While lobjXMLReader.ReadState <> lobjXMLReader.ReadState.EndOfFile
                lobjSB.Append(lobjXMLReader.ReadOuterXml())
            Loop

            'Return lobjXMLReader.ReadOuterXml
            Return lobjSB.ToString

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteXML." & lobjSysEx.Source
            Throw

        Finally
            DBDisconnect()

            If Not IsNothing(lobjXMLReader) Then

                If Not lobjXMLReader.ReadState = Xml.ReadState.Closed Then
                    lobjXMLReader.Close()
                End If

                lobjXMLReader = Nothing
            End If

            If Not IsNothing(lobjSB) Then
                lobjSB = Nothing
            End If
            If Not IsNothing(lobjSQLCmd) Then
                lobjSQLCmd.Dispose()
            End If

        End Try
    End Function

#End Region
End Class