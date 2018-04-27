Imports Oracle.DataAccess.Client
Imports System.Configuration.ConfigurationSettings
Public Class OracleData
    Private Const cFULL_MODULE_NAME As String = "DBAdaptor.OracleData"

    Private mobjOracleDBConn As OracleConnection
    Private mstrConnString As String
    Private mstrDBNameKey As String


#Region "Constructors"
    Sub New(ByVal astrDBNameKey As String)
        Try

            If IsNothing(AppSettings(astrDBNameKey.ToString)) Then
                Throw New Exception("Connection string for this DBNameKey cannot be found")
            Else
                mstrConnString = AppSettings(astrDBNameKey.ToString)
                mstrDBNameKey = astrDBNameKey
            End If

        Catch lobjSysEx As System.Exception
            Throw lobjSysEx
        End Try

    End Sub
#End Region


#Region "Open/Close Connections"

    Private Function DBConnect() As Object
        Try
            OracleDBConnect()
        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":DBConnect." & lobjSysEx.Source
            Throw lobjSysEx
        End Try
    End Function

    Private Function DBDisconnect() As Object
        Try
            If Not IsNothing(mobjOracleDBConn) Then
                OracleDBDisConnect()
            End If

        Catch lobjSysEx As System.Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":OracleDBConnect." & lobjSysEx.Source
            Throw lobjSysEx
        End Try
    End Function
    '================================================================
    'METHOD  :  OracleDBConnect
    'PURPOSE :  uses the .Net OracleClient provider to connect to 
    '           the database
    'PARMS   :  N/A
    'RETURN  :  NONE
    '================================================================

    Private Sub OracleDBConnect()
        Try
            'open the connection to Oracle svr db
            mobjOracleDBConn = New OracleConnection
            mobjOracleDBConn.ConnectionString = mstrConnString
            mobjOracleDBConn.Open()

        Catch lobjSysEx As System.Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":OracleDBConnect." & lobjSysEx.Source
            Throw lobjSysEx
        End Try

    End Sub

    '================================================================
    'METHOD  :  OracleDBDisConnect
    'PURPOSE :  uses the .Net OracleClient provider to disconnect from 
    '           the database if the connection is open
    'PARMS   :  N/A
    'RETURN  :  NONE
    '================================================================

    Private Sub OracleDBDisConnect()
        Try
            'close the connection to the db if it is not closed
            If mobjOracleDBConn.State <> ConnectionState.Closed Then
                mobjOracleDBConn.Close()
                mobjOracleDBConn.Dispose()
            End If

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":OracleDBDisConnect." & lobjSysEx.Source
            Throw lobjSysEx
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

        'Dim lobjOracleDA As Data.OracleClient.OracleDataAdapter
        Dim lobjOracleDA As Oracle.DataAccess.Client.OracleDataAdapter
        Dim lobjDS As Data.DataSet = Nothing

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

            'lobjOracleDA = New OracleClient.OracleDataAdapter(astrSQL, mobjOracleDBConn)
            lobjOracleDA = New Oracle.DataAccess.Client.OracleDataAdapter(astrSQL, mobjOracleDBConn)
            lobjOracleDA.SelectCommand.CommandTimeout = 0
            lobjOracleDA.Fill(lobjDS)

            Return lobjDS

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteDS." & lobjSysEx.Source
            Throw
        Finally

            DBDisconnect()

            If Not IsNothing(lobjOracleDA) Then
                lobjOracleDA.Dispose()
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
    'PURPOSE :  Execute a SQL statement against the database, 
    'PARMS   :  ByVal astrSQL As String
    'RETURN  :  N/A
    '================================================================

    Public Function ExecuteNonQuery(ByVal astrSQL As String) As Long

        'Dim lobjOracleCmd As OracleClient.OracleCommand
        Dim lobjOracleCmd As Oracle.DataAccess.Client.OracleCommand
        Dim llngAffectedRows As Long

        Try
            DBConnect()
            lobjOracleCmd = New OracleCommand(astrSQL)
            lobjOracleCmd.Connection = mobjOracleDBConn
            lobjOracleCmd.CommandType = CommandType.Text
            lobjOracleCmd.CommandTimeout = 0
            'execute sql passed in
            llngAffectedRows = lobjOracleCmd.ExecuteNonQuery()
            Return llngAffectedRows

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteNonQuery." & lobjSysEx.Source
            Throw
        Finally

            DBDisconnect()

            If Not IsNothing(lobjOracleCmd) Then
                lobjOracleCmd.Dispose()
            End If

        End Try
    End Function
#End Region


#Region "ExecuteScalar"
    '================================================================
    'METHOD  :  ExecuteScalar
    'PURPOSE :  execute a SQL statement against the database, 
    'PARMS   :  ByVal astrSQL As String
    'RETURN  :  The first column of the first row.
    '================================================================

    Public Function ExecuteScalar(ByVal astrSQL As String) As Object
        ' Dim lobjOracleCmd As OracleClient.OracleCommand
        Dim lobjOracleCmd As Oracle.DataAccess.Client.OracleCommand

        Try
            DBConnect()
            lobjOracleCmd = New OracleCommand(astrSQL)
            lobjOracleCmd.Connection = mobjOracleDBConn
            lobjOracleCmd.CommandType = CommandType.Text
            lobjOracleCmd.CommandTimeout = 0
            ' Execute the SQL.
            Return lobjOracleCmd.ExecuteScalar()

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteScalar." & lobjSysEx.Source
            Throw

        Finally
            DBDisconnect()

            If Not IsNothing(lobjOracleCmd) Then
                lobjOracleCmd.Dispose()
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
        'Dim lobjOracleCmd As OracleClient.OracleCommand
        Dim lobjOracleCmd As Oracle.DataAccess.Client.OracleCommand

        Try
            DBConnect()
            lobjOracleCmd = New OracleCommand(astrSQL)
            lobjOracleCmd.Connection = mobjOracleDBConn
            lobjOracleCmd.CommandType = CommandType.Text
            lobjOracleCmd.CommandTimeout = 0
            ' Execute SQL passed in
            Return lobjOracleCmd.ExecuteReader

        Catch lobjSysEx As Exception
            DBDisconnect()
            If Not IsNothing(lobjOracleCmd) Then
                lobjOracleCmd.Dispose()
            End If
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteScalar." & lobjSysEx.Source
            Throw
        Finally

        End Try
    End Function

    'Public Function ExecuteXML(ByVal astrSQL As String, ByVal aeDBCommandType As eCommandType, ByVal ParamArray aobjSQLParams As Object()) As String
    '    Dim lobjOracleCmd As OracleCommand
    '    Dim lobjXMLReader As System.Xml.XmlReader
    '    Dim lobjSB As System.Text.StringBuilder

    '    Try
    '        Select Case aeDBCommandType
    '            Case eCommandType.ecStoredProcedure
    '                For i As Integer = 0 To aobjSQLParams.Length - 1
    '                    If i > 0 Then
    '                        astrSQL = astrSQL & ","
    '                    End If

    '                    astrSQL = astrSQL & " '" & aobjSQLParams(i) & "'"
    '                Next
    '        End Select

    '        DBConnect()
    '        lobjOracleCmd = New OracleCommand(astrSQL, mobjOracleDBConn)
    '        lobjXMLReader = lobjOracleCmd.ExecuteReader
    '        lobjXMLReader.MoveToContent()

    '        lobjSB = New System.Text.StringBuilder

    '        Do While lobjXMLReader.ReadState <> lobjXMLReader.ReadState.EndOfFile
    '            lobjSB.Append(lobjXMLReader.ReadOuterXml())
    '        Loop

    '        'Return lobjXMLReader.ReadOuterXml
    '        Return lobjSB.ToString

    '    Catch lobjSysEx As Exception
    '        lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteXML." & lobjSysEx.Source
    '        Throw lobjSysEx

    '    Finally
    '        DBDisconnect()

    '        If Not IsNothing(lobjXMLReader) Then

    '            If Not lobjXMLReader.ReadState = Xml.ReadState.Closed Then
    '                lobjXMLReader.Close()
    '            End If

    '            lobjXMLReader = Nothing
    '        End If

    '        If Not IsNothing(lobjSB) Then
    '            lobjSB = Nothing
    '        End If
    '        If Not IsNothing(lobjSQLCmd) Then
    '            lobjSQLCmd.Dispose()
    '        End If

    '    End Try
    'End Function

#End Region


End Class
