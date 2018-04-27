Imports BSDBAdapter
Imports BSMoneyCostEntity
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.EnterpriseServices
Imports System.Runtime.InteropServices
Public Enum eDCActions
    [_eLowerBound] = -1
    ecRSUpdateBatch
    ecRSUpdateBatchCurrentOnly

    'to use Test SQL
    ecTestSQL

    'to fetch index rates from DWH
    ecGetIndexData

    'To fetch data from DWH without the last_updated_ind flag
    ecGetIndexDataWithoutCondition

    '---All Enums go here--

    [_eUpperBound]
End Enum
Public Interface IMoneyCostAutoDataClass
    Function Test() As DataTable
    Function GetCsvRecords(ByVal objcdataEntity As cDataEntity) As cDataEntity
    Function GetIndexData(ByVal objcdataEntity As cDataEntity) As cDataEntity
    Function UpdateCsvRecords(ByVal objcdataEntity As cDataEntity) As Long
    Function GetSQLStatement(ByVal alActionID As eDCActions, ByVal avParms As Object(), Optional ByVal abQuoteReplacement As Boolean = False) As String
End Interface

'<JustInTimeActivation(), _
' EventTrackingEnabled(), _
' ClassInterface(ClassInterfaceType.None), _
' Transaction(TransactionOption.Supported, Isolation:=TransactionIsolationLevel.Serializable, Timeout:=120), _
' ComponentAccessControl(True)> _
Public Class MoneyCostAutoDataClass
    'Inherits ServicedComponent
    Implements IMoneyCostAutoDataClass

    Private Const cDELIMITER As String = "|~^"
    Private _DBNAME As String = "MoneyCost"
    Private _DEFAULT_CONNECTION As String = "MoneyCost"
    'to fetch index rates from DWH for given Index Rate configuration, given currency, and specific processing date
    Private Const cGetIndexData As String = "SELECT YIELD_CURVE_TYPE, TERM_PERIOD, INTEREST_RATE FROM " & _
                                            "DWC_OWNER.DWC_INTEREST_RATE, DWC_DATE, DWC_TERM_PERIOD, " & _
                                            "DWC_OWNER.DWC_YIELD_CURVE_TYPE, DWC_OWNER.DWC_CURRENCY WHERE " & _
                                            "((DWC_INTEREST_RATE.DATE_KEY = DWC_DATE.DATE_KEY) AND " & _
                                            "(DWC_INTEREST_RATE.CURRENCY_KEY = DWC_CURRENCY.CURRENCY_KEY) AND " & _
                                            "(DWC_INTEREST_RATE.YIELD_CURVE_KEY = DWC_YIELD_CURVE_TYPE.YIELD_CURVE_KEY) " & _
                                            "AND (DWC_INTEREST_RATE.TERM_PERIOD_KEY = DWC_TERM_PERIOD.TERM_PERIOD_KEY)) " & _
                                            "AND CALENDAR_DATE = TO_DATE('|~^1', 'MM/DD/YYYY') AND CURRENCY_CODE = '|~^2' " & _
                                            "AND NVL(DWC_INTEREST_RATE.LAST_UPDATED_IND,'N') <> 'Y' " & _
                                            "AND UPPER(YIELD_CURVE_TYPE) IN (|~^3) AND TERM_PERIOD IN (|~^4) ORDER BY " & _
                                            "YIELD_CURVE_TYPE, TERM_PERIOD"


    'to fetch index rates from DWH for given Index Rate configuration, given currency, and specific processing date
    'where we dont want to fetch the data with the last_updated condition.
    Private Const cGetIndexDataWithoutCondition As String = "SELECT YIELD_CURVE_TYPE, TERM_PERIOD, INTEREST_RATE FROM " & _
                                            "DWC_OWNER.DWC_INTEREST_RATE, DWC_DATE, DWC_TERM_PERIOD, " & _
                                            "DWC_OWNER.DWC_YIELD_CURVE_TYPE, DWC_OWNER.DWC_CURRENCY WHERE " & _
                                            "((DWC_INTEREST_RATE.DATE_KEY = DWC_DATE.DATE_KEY) AND " & _
                                            "(DWC_INTEREST_RATE.CURRENCY_KEY = DWC_CURRENCY.CURRENCY_KEY) AND " & _
                                            "(DWC_INTEREST_RATE.YIELD_CURVE_KEY = DWC_YIELD_CURVE_TYPE.YIELD_CURVE_KEY) " & _
                                            "AND (DWC_INTEREST_RATE.TERM_PERIOD_KEY = DWC_TERM_PERIOD.TERM_PERIOD_KEY)) " & _
                                            "AND CALENDAR_DATE = TO_DATE('|~^1', 'MM/DD/YYYY') AND CURRENCY_CODE = '|~^2' " & _
                                            "AND UPPER(YIELD_CURVE_TYPE) IN (|~^3) AND TERM_PERIOD IN (|~^4) ORDER BY " & _
                                            "YIELD_CURVE_TYPE, TERM_PERIOD"


    '================================================================
    'METHOD  : GetConnectString
    'PURPOSE : This method will get the DB connection string from the
    '          registry.
    'PARMS   : astrKey [String] = Connection string registry Key name
    'RETURN  : String = The connection string to connect to the
    '          database with
    '================================================================
    Private Function GetConnectString(ByVal astrKey As String) As String
        GetConnectString = delGetConfigurationKey(astrKey)
    End Function

    '================================================================
    'METHOD  : Disconnect
    'PURPOSE : Disconnects the current connection from the datasource
    'PARMS   : NONE
    'RETURN  : NONE
    '================================================================
    <AutoComplete()> _
    Public Function Test() As DataTable Implements IMoneyCostAutoDataClass.Test
        Dim lobjDs As DataSet
        lobjDs = New DataSet
        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "UspMoneyCostAutoTest"
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecStoredProcedure
            lobjDs = lobjDataAdaptor.ExecuteDS()                ' Call DataAdaptor Class's ExecuteDS method()
            Return lobjDs.Tables(0)
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(lobjDs) Then
                lobjDs.Dispose()
                lobjDs = Nothing
            End If
        End Try
    End Function

    '5555555555555555555555555 -- Treasury Assessment
    <AutoComplete()> _
    Public Function GetIndexDataForTreasuryAssessment(ByVal objcdataEntity As cDataEntity) As DataSet

        Dim lobjDs As DataSet
        Dim lobjDataAdaptor As DataAdaptor
        lobjDataAdaptor = New DataAdaptor("", delGetConfigurationKey("MCDataWarehouse"), eDBType.cOledb)
        lobjDataAdaptor.replaceQuote = True
        lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecSQLText
        lobjDataAdaptor.commandSQL = delGetConfigurationKey("TreasuryAssessmentQuery").ToString().Replace("CDATE", objcdataEntity.ProcessDate).Replace("$COSTYPE$", objcdataEntity.CostTypes)
        lobjDs = lobjDataAdaptor.ExecuteDS()
        Return lobjDs
    End Function

    <AutoComplete()> _
    Public Function GetCsvRecords(ByVal objcdataEntity As cDataEntity) As cDataEntity Implements IMoneyCostAutoDataClass.GetCsvRecords
        Dim lobjDs As DataSet
        Dim sConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Text;Data Source=" & objcdataEntity.WrkDirectory
        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor("MoneyCost", sConnectionString, eDBType.cOledb)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text            
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecSQLText
            lobjDs = lobjDataAdaptor.ExecuteCsvDS(objcdataEntity.CommonSQL)               ' Call DataAdaptor Class's ExecuteDS method()
            objcdataEntity.CsvOutput = lobjDs.Tables(0)
            Return objcdataEntity
        Catch ex As Exception
            objcdataEntity.CsvOutput = Nothing
            Return objcdataEntity
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(lobjDs) Then
                lobjDs = Nothing
            End If
            If Not IsNothing(objcdataEntity) Then
                objcdataEntity = Nothing
            End If
        End Try
    End Function
    <AutoComplete()> _
    Public Function GetIndexData(ByVal objcdataEntity As cDataEntity) As cDataEntity Implements IMoneyCostAutoDataClass.GetIndexData
        Dim lobjDs As DataSet
        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor("", delGetConfigurationKey("MCDataWarehouse"), eDBType.cOledb)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text            
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecSQLText
            Dim ojbjParame(4) As Object
            ojbjParame(0) = objcdataEntity.ProcessDate
            ojbjParame(1) = objcdataEntity.CurrencyCode
            ojbjParame(2) = objcdataEntity.YieldCurveType
            ojbjParame(3) = objcdataEntity.TermPeriod
            objcdataEntity.CommonSQL = GetSQLStatement(objcdataEntity.ActionID, ojbjParame, objcdataEntity.QuoteReplacement)
            lobjDataAdaptor.commandSQL = objcdataEntity.CommonSQL
            lobjDs = lobjDataAdaptor.ExecuteDS()               ' Call DataAdaptor Class's ExecuteDS method()
            lobjDs.DataSetName = "INDEX_DATASet"
            lobjDs.Tables(0).TableName = "INDEX_DATA"
            objcdataEntity.OutputString = DsToXML(lobjDs)
            Return objcdataEntity
        Catch ex As Exception
            objcdataEntity.OutputString = Nothing
            Return objcdataEntity
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(lobjDs) Then
                lobjDs = Nothing
            End If
            If Not IsNothing(objcdataEntity) Then
                objcdataEntity = Nothing
            End If
        End Try
    End Function
    <AutoComplete()> _
    Public Function UpdateCsvRecords(ByVal objcdataEntity As cDataEntity) As Long Implements IMoneyCostAutoDataClass.UpdateCsvRecords
        Dim lobjDs As DataSet
        Dim sConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Text;Data Source=" & objcdataEntity.WrkDirectory
        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor("", sConnectionString, eDBType.cOledb)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text            
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecSQLText
            'Dim ojbjParame(4) As Object
            'ojbjParame(0) = objcdataEntity.ProcessDate
            'ojbjParame(1) = objcdataEntity.CurrencyCode
            'ojbjParame(2) = objcdataEntity.YieldCurveType
            'ojbjParame(3) = objcdataEntity.TermPeriod

            'objcdataEntity.CommonSQL = GetSQLStatement(objcdataEntity.ActionID, ojbjParame, objcdataEntity.QuoteReplacement)
            Return lobjDataAdaptor.ExecuteNonQuery(objcdataEntity.CommonSQL)               ' Call DataAdaptor Class's ExecuteDS method()           
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(lobjDs) Then
                lobjDs = Nothing
            End If
            If Not IsNothing(objcdataEntity) Then
                objcdataEntity = Nothing
            End If
        End Try
    End Function

    '================================================================
    'METHOD  : GetSQLStatement
    'PURPOSE : Based on an Action ID and an array of parms
    '          GetSQLStatement returns the SQL Statement in
    '          "executeable" form i.e. all parms are replaced in the
    '          string with the parms provided in Parms()
    'PARMS   : alActionID [eDCActions] = SQL Statement to Execute
    '          avParms [Variant] = Parms to Populate SQL Statement
    '          with
    'RETURN  : String = Completed SQL statement
    '================================================================
    <AutoComplete()> _
    Public Function GetSQLStatement(ByVal alActionID As eDCActions, _
                        ByVal avParms As Object(), Optional ByVal abQuoteReplacement As Boolean = False) As String Implements IMoneyCostAutoDataClass.GetSQLStatement
        Dim lstrSQL As String = ""
        Dim liPos As Integer
        Dim liLoopCtr As Integer
        Dim liUBound As Integer

        Try
            Select Case alActionID.ToString()

                Case "ecGetIndexData"
                    lstrSQL = cGetIndexData
                Case "ecGetIndexDataWithoutCondition"
                    lstrSQL = cGetIndexDataWithoutCondition
                    '---All Enum to SQL constant mapping goes here---
                Case Else
                    Throw New Exception("MoneyCostAutoDtaClass :-" & Err.Description)
            End Select

            'If there are no substitution parms, we're done
            liPos = InStr(1, lstrSQL, cDELIMITER)

            If liPos = 0 Then
                GetSQLStatement = lstrSQL
                Exit Function
            End If

            If IsArray(avParms) Then
                If UBound(avParms) = -1 Then
                    Throw New Exception("MoneyCostAutoDtaClass :-" & Err.Description)
                End If
            End If

            'Loop through each of the param
            liUBound = UBound(avParms)
            For liLoopCtr = liUBound + 1 To 1 Step -1

                'First, replace any apostrophes (') in the parm with two ('') for SQL Server.
                If abQuoteReplacement Then avParms(liLoopCtr - 1) = Replace(avParms(liLoopCtr - 1), "'", "''")

                'Then Insert the parm into the SQL statement.
                lstrSQL = Replace(lstrSQL, cDELIMITER & CStr(liLoopCtr), avParms(liLoopCtr - 1))
            Next liLoopCtr

            'REPLACE ANY STRINGS THAT CAME UP 'NULL' WITH NULL
            lstrSQL = Replace(lstrSQL, "'NULL'", "NULL")

            'Make sure all substitution parms have been swapped out
            liPos = InStr(1, lstrSQL, cDELIMITER)
            If liPos > 0 Then Throw New Exception("MoneyCostAutoDtaClass :-" & Err.Description)

            'Return Completed SQL statement
            Return lstrSQL


        Catch ex As Exception
            Throw
        End Try
    End Function
    '=== Function Convert DS To XML ================================
    <AutoComplete()> _
    Public Function DsToXML(ByVal astrDS As DataSet) As String
        Dim iCount, iRow, iCol As Integer
        Dim vXML As String

        vXML = ""
        Try
            If astrDS.Tables.Count > 0 Then
                vXML = "<" & astrDS.DataSetName & ">"
                For iCount = 0 To astrDS.Tables.Count - 1  '' Loop for Table
                    For iRow = 0 To astrDS.Tables(iCount).Rows.Count - 1 'Loop for Record
                        vXML = vXML & "<" & astrDS.Tables(iCount).TableName & ">"
                        For iCol = 0 To astrDS.Tables(iCount).Columns.Count - 1  ''Loop for each column
                            vXML = vXML & "<" & astrDS.Tables(iCount).Columns(iCol).ColumnName & ">"
                            vXML = vXML & astrDS.Tables(iCount).Rows(iRow).Item(iCol)
                            vXML = vXML & "</" & astrDS.Tables(iCount).Columns(iCol).ColumnName & ">"
                        Next
                        vXML = vXML & "</" & astrDS.Tables(iCount).TableName & ">"
                    Next
                Next
                vXML = vXML & "</" & astrDS.DataSetName & ">"
            End If

            vXML = Replace(vXML, "&", "&amp;")
            Return vXML
        Catch ex As Exception
            Throw
        End Try
    End Function
End Class
