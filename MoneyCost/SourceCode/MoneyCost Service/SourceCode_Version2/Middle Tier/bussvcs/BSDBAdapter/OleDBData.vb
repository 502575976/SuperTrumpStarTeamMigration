Imports System.Data.OleDb
Imports System.Configuration.ConfigurationSettings

Public Class OleDBData
    Private Const cFULL_MODULE_NAME As String = "DBAdaptor.OleDBData"
    Private mobjOleDBConn As OleDbConnection
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
            Throw
        End Try

    End Sub
    Sub New(ByVal astrDBNameKey As String, ByVal astrCon As String)
        Try
            If IsNothing(astrDBNameKey.ToString) Then
                Throw New Exception("Connection string for this DBNameKey cannot be found")
            Else
                mstrConnString = astrCon
                mstrDBNameKey = astrDBNameKey
            End If

        Catch lobjSysEx As System.Exception
            Throw
        End Try

    End Sub
#End Region

#Region "Open/Close Connections"

    '================================================================
    'METHOD  :  DBConnect
    'PURPOSE :  connects to the database
    'PARMS   :  N/A
    'RETURN  :  NONE
    '================================================================

    Public Sub DBConnect()
        Try
            OleDBConnect()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    '================================================================
    'METHOD  :  DisConnect
    'PURPOSE :  disconnect from the database if the connection is open
    'PARMS   :  N/A
    'RETURN  :  NONE
    '================================================================

    Private Sub DBDisConnect()
        Try
            If Not IsNothing(mobjOleDBConn) Then
                OleDBDisConnect()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    '================================================================
    'METHOD  :  OleDBConnect
    'PURPOSE :  uses the .Net OleDb provider to connect to 
    '           the database
    'PARMS   :  N/A
    'RETURN  :  NONE
    '================================================================

    Private Sub OleDBConnect()
        Try
            'open the connection to oracle db
            mobjOleDBConn = New OleDbConnection
            mobjOleDBConn.ConnectionString = mstrConnString
            mobjOleDBConn.Open()

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":OleDbConnect." & lobjSysEx.Source
            Throw lobjSysEx
        End Try
    End Sub
    '================================================================
    'METHOD  :  OleDBDisConnect
    'PURPOSE :  uses the .Net OleDB provider to disconnect from 
    '           the database if the connection is open
    'PARMS   :  N/A
    'RETURN  :  NONE
    '================================================================

    Private Sub OleDBDisConnect()
        Try
            'close the connection to the db if it is not closed
            If mobjOleDBConn.State <> ConnectionState.Closed Then
                mobjOleDBConn.Close()
                mobjOleDBConn.Dispose()
            End If

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":OleDBDisConnect." & lobjSysEx.Source
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

        Dim lobjOleDBDA As Data.OleDb.OleDbDataAdapter
        Dim lobjOleDBCmd As OleDbCommand = Nothing
        Dim lobjDS As Data.DataSet

        Try

            DBConnect()
            lobjDS = New DataSet

            'check command type
            Select Case aeDBCommandType

                'stored proc
                Case eCommandType.ecStoredProcedure
                    'Select Case mstrDBNameKey 'sql server/sqlclient
                    '    Case eDBNameKey.BTO_DB, eDBNameKey.ICE_DB

                    '        For i As Integer = 0 To aobjSQLParams.Length - 1
                    '            If i > 0 Then
                    '                astrSQL = astrSQL & ","
                    '            End If

                    '            astrSQL = astrSQL & " '" & aobjSQLParams(i) & "'"

                    '        Next
                    '    Case Else

                    'End Select

                    'sql text
                Case eCommandType.ecSQLText


            End Select

            lobjOleDBDA = New OleDb.OleDbDataAdapter(astrSQL, mobjOleDBConn)
            lobjOleDBDA.SelectCommand.CommandTimeout = 0
            lobjOleDBDA.Fill(lobjDS)

            Return lobjDS

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteDS." & lobjSysEx.Source
            Throw
        Finally

            DBDisConnect()

            If Not IsNothing(lobjOleDBDA) Then
                lobjOleDBDA.Dispose()
            End If

            If Not IsNothing(lobjDS) Then
                lobjDS.Dispose()
            End If

        End Try
    End Function


    Public Function ExecuteDataset(ByVal astrSQL As String) As Data.DataSet


        Dim lobjOleDBDA As Data.OleDb.OleDbDataAdapter = Nothing
        Dim lobjOleDBCmd As OleDbCommand = Nothing
        Dim lobjDS As Data.DataSet = Nothing
        Try

            DBConnect()
            lobjOleDBDA = New Data.OleDb.OleDbDataAdapter(astrSQL, mobjOleDBConn)
            lobjDS = New Data.DataSet
            'check command type

            lobjOleDBDA.Fill(lobjDS)
            Return lobjDS
        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteDS." & lobjSysEx.Source
            Throw
        Finally

            DBDisConnect()

            If Not IsNothing(lobjOleDBDA) Then
                lobjOleDBDA.Dispose()
            End If

            If Not IsNothing(lobjOleDBCmd) Then
                lobjOleDBCmd.Dispose()
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

        Dim lobjOleCmd As OleDb.OleDbCommand
        Dim llngAffectedRows As Long
        Try
            DBConnect()

            lobjOleCmd = New OleDbCommand(astrSQL)
            lobjOleCmd.Connection = mobjOleDBConn
            lobjOleCmd.CommandType = CommandType.Text
            lobjOleCmd.CommandTimeout = 0
            'execute sql passed in
            llngAffectedRows = lobjOleCmd.ExecuteNonQuery()

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteNonQuery." & lobjSysEx.Source
            Throw lobjSysEx
        Finally

            DBDisConnect()

            If Not IsNothing(lobjOleCmd) Then
                lobjOleCmd.Dispose()
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

        Dim lobjOleCmd As OleDb.OleDbCommand = Nothing

        Try
            DBConnect()

            lobjOleCmd = New OleDbCommand(astrSQL)
            lobjOleCmd.Connection = mobjOleDBConn
            lobjOleCmd.CommandType = CommandType.Text
            lobjOleCmd.CommandTimeout = 0
            'execute sql passed in
            Return lobjOleCmd.ExecuteScalar

        Catch lobjSysEx As Exception
            lobjSysEx.Source = cFULL_MODULE_NAME & ":ExecuteScalar." & lobjSysEx.Source
            Throw

        Finally
            DBDisConnect()
            If Not IsNothing(lobjOleCmd) Then
                lobjOleCmd.Dispose()
            End If
        End Try
    End Function
#End Region
End Class
