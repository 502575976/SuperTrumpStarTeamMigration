Imports BSVICROIDBADAPTOR.OracleHelper
Imports BSVICROIEntity.BSVICROIEntity
Imports System.Data.OracleClient
Imports System.Data
Imports System.Text
Imports Microsoft.Win32
Imports System.Xml
Imports System.Reflection
Imports System.Transactions
Imports SuperTRUMPCommon
Imports System.IO

Public Interface IDataClass
    Function GenerateInputXML_DAL(ByVal cSTForAllDealsEntity As cSTForAllDealsEntity) As BSVICROIEntity.BSVICROIEntity.cSTForAllDealsEntity
End Interface
Public Class cDataClass

    Dim STLogger As log4net.ILog

    

    ''' <summary>
    ''' Export the data in DB by passing query with Connection String
    ''' </summary>
    ''' <param name="sbQuery"></param>
    ''' <param name="strCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExportDATFileByQuery(ByVal sbQuery As String, ByVal strCon As String) As String
        Dim objHelper As New OracleHelper(strCon)
        Dim objCommon As BSVICROICommon = New BSVICROICommon()
        STLogger = objCommon.SetLog4Net()
        Try
            STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method in DAL: " + MethodInfo.GetCurrentMethod.Name())
            Return objHelper.ExecuteNonQueryWithMsg(sbQuery.ToString())
        Catch ex As Exception
            Return "Error"
        Finally
            STLogger.Debug("END:  " + DateTime.Now + " Tracing process within method in DAL: " + MethodInfo.GetCurrentMethod.Name())
            objHelper = Nothing
            STLogger = Nothing
            objCommon = Nothing
        End Try

    End Function

    'Public Function GetDataInDataSet(ByVal strCon As String) As DataSet
    '    Dim objHelper As New OracleHelper(strCon)
    '    Return objHelper.GetDatainDataset()
    'End Function

    Public Function TestServiceinDAL() As String
        Return "DAL Successfully - " + OracleHelper.TestServiceinADP()
    End Function


    Public Function GetAllRecordsWithAccountScheduleNo(ByVal strCon As String) As DataSet
        Dim objCommon As BSVICROICommon = New BSVICROICommon()
        Dim STLogger As log4net.ILog
        STLogger = objCommon.SetLog4Net()
        Dim command As OracleCommand
        Dim objAdap As OracleDataAdapter
        Dim dsData As New DataSet

        Try
            STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method in DAL: " + MethodInfo.GetCurrentMethod.Name())
            Dim objConnection As New OracleConnection(strCon)
            If objConnection.State = ConnectionState.Open Then
                objConnection.Close()
                objConnection.Open()
            Else
                objConnection.Open()
            End If

            command = New OracleCommand("GetAllRecordsWithCursor", objConnection)
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add("AccountScheduleFeed", OracleType.Cursor).Direction = ParameterDirection.Output
            STLogger.Debug("Add AccountScheduleFeed Parameter in Command")
            command.Parameters.Add("StreamFeed", OracleType.Cursor).Direction = ParameterDirection.Output
            STLogger.Debug("Add StreamFeed Parameter in Command")
            command.Parameters.Add("AssetLevelFeed", OracleType.Cursor).Direction = ParameterDirection.Output
            STLogger.Debug("Add AssetLevelFeed Parameter in Command")
            command.Parameters.Add("ProductMapping", OracleType.Cursor).Direction = ParameterDirection.Output
            STLogger.Debug("Add ProductMapping Parameter in Command")
            command.Parameters.Add("TemplateMapping", OracleType.Cursor).Direction = ParameterDirection.Output
            STLogger.Debug("Add TemplateMapping Parameter in Command")
            command.Parameters.Add("Depriciation", OracleType.Cursor).Direction = ParameterDirection.Output
            STLogger.Debug("Add Depriciation Parameter in Command")
            objAdap = New OracleDataAdapter(command)
            objAdap.Fill(dsData)
            STLogger.Debug("END:  " + DateTime.Now + " Tracing process within method in DAL: " + MethodInfo.GetCurrentMethod.Name())
            Return dsData
        Catch ex As Exception
            STLogger.Error(" Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + ex.Message)
            Return Nothing
        Finally
            objCommon = Nothing
            STLogger = Nothing
            command = Nothing
            objAdap = Nothing
            dsData = Nothing
        End Try
    End Function


    Public Function InsertDataInMainStreamDetailTable(ByVal strCon As String) As Integer
        Dim objConnection As New OleDb.OleDbConnection(strCon)
        If objConnection.State = ConnectionState.Open Then
            objConnection.Close()
            objConnection.Open()
        Else
            objConnection.Open()
        End If
        Dim command As OleDb.OleDbCommand
        '        Dim objAdap As OleDb.OleDbDataAdapter
        command = New OleDb.OleDbCommand("UpdateStreamTableData", objConnection)
        command.CommandType = CommandType.StoredProcedure

        Dim icount As Int16 = command.ExecuteNonQuery()

        objConnection.Close()
        command.Dispose()
        'objAdap.Dispose()
        Return icount
    End Function

    Public Function GetCapMarketAdder(ByVal strACNumber As String, ByVal nbrTerm As Integer, ByVal strProductName As String, ByVal strProgrameName As String, ByVal strCon As String) As DataSet
        Dim command As OleDb.OleDbCommand
        Dim strMileStone As String = "5"
        Dim dblCapAddValue As Double = 0.0
        Dim objCommon As BSVICROICommon = New BSVICROICommon()
        Dim STLogger As log4net.ILog
        STLogger = objCommon.SetLog4Net()
        Dim dsResidual As New DataSet
        Dim dtResidual As New DataTable("RESIDUALMAPPING")
        Try
            strMileStone = "5.1"
            Dim objConnection As New OleDb.OleDbConnection(strCon)
            If objConnection.State = ConnectionState.Open Then
                objConnection.Close()
                objConnection.Open()
            Else
                objConnection.Open()
            End If
            strMileStone = "5.2"
            command = New OleDb.OleDbCommand("CapMarketAdder", objConnection)
            strMileStone = "5.3"
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add("AccountScheduleNumber", OleDb.OleDbType.VarChar).Value = strACNumber
            command.Parameters.Add("TermValue", OleDb.OleDbType.Numeric).Value = nbrTerm
            command.Parameters.Add("Product", OleDb.OleDbType.VarChar).Value = strProductName
            command.Parameters.Add("ProgramName", OleDb.OleDbType.VarChar).Value = strProgrameName
            command.Parameters.Add("CapADDER", OleDb.OleDbType.Numeric).Direction = ParameterDirection.Output
            strMileStone = "5.4"

            'Dim icount As Int16 = 
            command.ExecuteNonQuery()
            strMileStone = "5.5"
            dblCapAddValue = command.Parameters("CapADDER").Value
            strMileStone = "5.6"
            objConnection.Close()
            
            dtResidual.Columns.Add("CAP_MARKET_ADDER")
            strMileStone = "5.7"
            Dim dr As DataRow
            dr = dtResidual.NewRow()
            dr("CAP_MARKET_ADDER") = dblCapAddValue
            dtResidual.Rows.Add(dr)
            dsResidual.Tables.Add(dtResidual)
            strMileStone = "5.8"
            Return dsResidual
        Catch ex As Exception
            STLogger.Error("MileStone:- " & strMileStone & " Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + ex.Message)
            Return Nothing
        Finally
            command = Nothing
            objCommon = Nothing
            STLogger = Nothing
        End Try
        
    End Function

    '''''''UNUSED FUNCTIONS'''''''''''
    'Public Function GenerateQueryToExportDATFile(ByVal alstColumnValue As ArrayList, ByVal strDBColumnName As String, ByVal strTableName As String) As String
    '    Dim sbQuery As New StringBuilder()
    '    Dim strMileStone As String = "1"
    '    Dim strMiddleQuery As String = String.Empty
    '    Dim Count As Int64 = 0
    '    Dim objCommon As BSVICROICommon = New BSVICROICommon()
    '    STLogger = objCommon.SetLog4Net()
    '    Try
    '        STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method in DAL: " + MethodInfo.GetCurrentMethod.Name())
    '        Dim strColName As String()
    '        Dim strColValue As String()
    '        strColName = strDBColumnName.Split(",")
    '        strMileStone = "1.1"
    '        For iRow As Integer = 0 To alstColumnValue.Count - 1
    '            strMileStone = "1.2"
    '            sbQuery.Append(" Insert Into " + strTableName + "(" + strDBColumnName.Replace("#Date", "").Replace("#Int", "") + ") values(")
    '            'strColValue = alstColumnValue(iRow).Split(Chr(9))
    '            strColValue = alstColumnValue(iRow).Split(Chr(44))
    '            For iCol As Integer = 0 To strColValue.Count - 1
    '                strMileStone = "1.3"
    '                If strColName(iCol).Contains("#Date") = True Then
    '                    If String.IsNullOrEmpty(strColValue(iCol).Trim()) Then
    '                        strMiddleQuery += ",''"
    '                    Else
    '                        'strMiddleQuery += "," + "to_date('" + strColValue(iCol) + "','MM/DD/YYYY HH:MM:SS')"
    '                        strMiddleQuery += ",'" + strColValue(iCol) + "'"
    '                    End If
    '                ElseIf strColName(iCol).Contains("#Int") = True Then

    '                    If String.IsNullOrEmpty(strColValue(iCol).Trim()) Then
    '                        strMiddleQuery += ",''"
    '                    Else
    '                        strMiddleQuery += "," + strColValue(iCol)
    '                    End If
    '                Else
    '                    strMiddleQuery += ",'" + strColValue(iCol).Replace("&", "' ||chr(38)|| '") + "'"
    '                End If
    '            Next

    '            If String.Compare(strTableName, "TBL_ACCOUNTSCHEDULE_DETAIL", True) = 0 And strDBColumnName.Contains("CREATION_DATE") Then
    '                strMiddleQuery += ",'" + DateTime.Now() + "'"
    '            End If
    '            strMileStone = "1.4"
    '            If strMiddleQuery.Length > 0 Then
    '                strMiddleQuery = strMiddleQuery.Substring(1)
    '            End If
    '            strMileStone = "1.5"
    '            sbQuery.Append(strMiddleQuery)
    '            strMiddleQuery = String.Empty
    '            sbQuery.Append("); ")

    '        Next
    '        strMileStone = "1.6"
    '        STLogger.Debug("END:  " + DateTime.Now + " Tracing process within method in DAL: " + MethodInfo.GetCurrentMethod.Name())
    '        Return sbQuery.ToString()
    '    Catch ex As Exception
    '        STLogger.Error("MileStone:- " & strMileStone & " Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
    '        Return "ERROR"
    '        Throw ex
    '    Finally
    '        sbQuery = Nothing
    '        objCommon = Nothing
    '        STLogger = Nothing
    '    End Try
    'End Function

    'Public Function GetRecordsForAccountScheduleNumber(ByVal strAcNum As String, ByVal strCon As String) As DataSet
    '    Dim dblCapAddValue As Double = 0.0
    '    Dim ProductName As String = String.Empty
    '    Dim ProgramName As String = String.Empty
    '    Dim TermValue As String = String.Empty
    '    Dim strQuery As String = String.Empty
    '    Dim strType As String = "Lease"
    '    Dim command As OleDb.OleDbCommand
    '    Dim objAdap As OleDb.OleDbDataAdapter
    '    Dim ds As New System.Data.DataSet
    '    Dim objCommon As BSVICROICommon = New BSVICROICommon()
    '    STLogger = objCommon.SetLog4Net()
    '    Dim objConnection As New OleDb.OleDbConnection(strCon)
    '    If objConnection.State = ConnectionState.Open Then
    '        objConnection.Close()
    '        objConnection.Open()
    '    Else
    '        objConnection.Open()
    '    End If
    '    Try
    '        STLogger.Debug("Account Schedule no for Process is " + strAcNum)
    '        command = New OleDb.OleDbCommand("SELECT * FROM TBL_ACCOUNTSCHEDULE_DETAIL where Account_Schedule_Nbr='" + strAcNum + "'", objConnection)
    '        objAdap = New System.Data.OleDb.OleDbDataAdapter(command)
    '        objAdap.Fill(ds, "AccountScheduleFeed")
    '        STLogger.Debug("Fill AccountScheduleFeed Successfully")

    '        If ds.Tables(0).Rows.Count > 0 Then
    '            TermValue = ds.Tables("AccountScheduleFeed").Rows(0)("Term")
    '            ProductName = ds.Tables("AccountScheduleFeed").Rows(0)("Product")
    '        End If

    '        command.Dispose()
    '        objAdap.Dispose()
    '        command = New OleDb.OleDbCommand("SELECT * FROM TBL_STREAM_DETAIL where ACCOUNT_SCHEDULE_NBR='" + strAcNum + "'", objConnection)
    '        objAdap = New System.Data.OleDb.OleDbDataAdapter(command)
    '        objAdap.Fill(ds, "StreamFeed")
    '        STLogger.Debug("Fill StreamFeed Successfully")


    '        command.Dispose()
    '        objAdap.Dispose()
    '        command = New OleDb.OleDbCommand("SELECT * FROM TBL_ASSET_DETAIL where ACCOUNT_SCHEDULE_NBR='" + strAcNum + "'", objConnection)
    '        objAdap = New System.Data.OleDb.OleDbDataAdapter(command)
    '        objAdap.Fill(ds, "AssetLevelFeed")
    '        STLogger.Debug("Fill AssetLevelFeed Successfully")
    '        command.Dispose()
    '        objAdap.Dispose()

    '        command = New OleDb.OleDbCommand("SELECT * FROM TBL_PMSDATA where PMS_Location=(SELECT Location FROM TBL_ACCOUNTSCHEDULE_DETAIL where Account_Schedule_Nbr='" + strAcNum + "')", objConnection)
    '        objAdap = New System.Data.OleDb.OleDbDataAdapter(command)
    '        objAdap.Fill(ds, "ProductMapping")
    '        STLogger.Debug("Fill ProductMapping Successfully")
    '        If ds.Tables("ProductMapping").Rows.Count > 0 Then
    '            ProgramName = ds.Tables("ProductMapping").Rows(0)("ST_PRODUCT_NAME")
    '        End If

    '        command.Dispose()
    '        objAdap.Dispose()

    '        command = New OleDb.OleDbCommand("select TplMap.* from TBL_TEMPLATEMAPPING TplMap, TBL_ACCOUNTSCHEDULE_DETAIL AcSch where TplMap.Product = AcSch.Product and AcSch.Account_Schedule_Nbr='" + strAcNum + "' and TplMap.TERM_MIN <= AcSch.TERM and TplMap.TERM_MAX >= AcSch.TERM ", objConnection)
    '        objAdap = New System.Data.OleDb.OleDbDataAdapter(command)
    '        objAdap.Fill(ds, "TemplateMapping")
    '        STLogger.Debug("Fill TemplateMapping Successfully")

    '        command.Dispose()
    '        objAdap.Dispose()
    '        command = New OleDb.OleDbCommand("SELECT Method FROM TBL_DEPRICIATION where Depreciation_Type in(SELECT distinct Depreciation_Type FROM TBL_ASSET_DETAIL where ACCOUNT_SCHEDULE_NBR='" + strAcNum + "')", objConnection)
    '        objAdap = New System.Data.OleDb.OleDbDataAdapter(command)
    '        objAdap.Fill(ds, "Depriciation")
    '        STLogger.Debug("Fill Depriciation Successfully")


    '        command.Dispose()
    '        objAdap.Dispose()

    '        command = New OleDb.OleDbCommand("CapMarketAdder", objConnection)
    '        command.CommandType = CommandType.StoredProcedure

    '        command.Parameters.Add("AccountScheduleNumber", OleDb.OleDbType.VarChar).Value = strAcNum
    '        command.Parameters.Add("TermValue", OleDb.OleDbType.Numeric).Value = TermValue
    '        command.Parameters.Add("Product", OleDb.OleDbType.VarChar).Value = ProductName
    '        command.Parameters.Add("ProgramName", OleDb.OleDbType.VarChar).Value = ProgramName
    '        command.Parameters.Add("CapADDER", OleDb.OleDbType.Numeric).Direction = ParameterDirection.Output

    '        If objConnection.State = ConnectionState.Open Then
    '            objConnection.Close()
    '            objConnection.Open()
    '        Else
    '            objConnection.Open()
    '        End If

    '        Dim icount As Int16 = command.ExecuteNonQuery()
    '        dblCapAddValue = command.Parameters("CapADDER").Value
    '        objConnection.Close()
    '        command.Dispose()
    '        objAdap.Dispose()
    '        Dim dt As New DataTable("RESIDUALMAPPING")
    '        dt.Columns.Add("CAP_MARKET_ADDER")
    '        Dim dr As DataRow
    '        dr = dt.NewRow()
    '        dr("CAP_MARKET_ADDER") = dblCapAddValue
    '        dt.Rows.Add(dr)
    '        ds.Tables.Add(dt)

    '    Catch ex As Exception
    '        STLogger.Error(" Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
    '        Throw
    '    Finally
    '        If Not IsNothing(objAdap) Then
    '            objAdap.Dispose()
    '            objAdap = Nothing
    '            objCommon = Nothing
    '        End If
    '        objConnection.Close()
    '    End Try
    '    Return ds
    'End Function


    'Public Function GetAllAccountScheduleNumbers(ByVal strCon As String) As DataSet
    '    Dim strQuery As String = String.Empty
    '    Dim command As OleDb.OleDbCommand
    '    Dim objAdap As OleDb.OleDbDataAdapter
    '    Dim dsASNum As New System.Data.DataSet
    '    Dim objCommon As BSVICROICommon = New BSVICROICommon()
    '    STLogger = objCommon.SetLog4Net()
    '    STLogger.Debug("START:  " + DateTime.Now + " Tracing process within method in DAL: " + MethodInfo.GetCurrentMethod.Name())
    '    Dim objConnection
    '    Try
    '        objConnection = New OleDb.OleDbConnection(strCon)
    '        If objConnection.State = ConnectionState.Open Then
    '            objConnection.Close()
    '            objConnection.Open()
    '        Else
    '            objConnection.Open()
    '        End If

    '        command = New OleDb.OleDbCommand("SELECT Account_Schedule_Nbr FROM TBL_ACCOUNTSCHEDULE_DETAIL", objConnection)
    '        objAdap = New System.Data.OleDb.OleDbDataAdapter(command)
    '        objAdap.Fill(dsASNum, "AccountScheduleFeed")
    '        command.Dispose()
    '        objAdap.Dispose()
    '        STLogger.Debug("END:  " + DateTime.Now + " Tracing process within method in DAL: " + MethodInfo.GetCurrentMethod.Name())
    '    Catch ex As Exception
    '        STLogger.Error(" Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
    '        Throw
    '    Finally
    '        If Not IsNothing(objAdap) Then
    '            objAdap.Dispose()
    '            objAdap = Nothing
    '        End If
    '        objCommon = Nothing
    '        STLogger = Nothing
    '        objConnection.Close()
    '    End Try
    '    Return dsASNum
    'End Function
End Class
