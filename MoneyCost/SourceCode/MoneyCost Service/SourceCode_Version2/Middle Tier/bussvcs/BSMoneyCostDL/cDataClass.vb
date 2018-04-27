Imports BSDBAdapter
Imports BSMoneyCostEntity
Imports System.EnterpriseServices
Imports System.Runtime.InteropServices
Public Interface IDataClass
    Function Test() As DataTable
    Function cDataExecute(ByVal cdataEntity As cDataEntity) As DataTable
    Function GetMCFileDetails(ByVal cdataEntity As cDataEntity) As cDataEntity
    Function GetAllMCFiles(ByVal cdataEntity As cDataEntity) As cDataEntity
    Function GetMCFiles(ByVal cdataEntity As cDataEntity) As cDataEntity
    Function GetIndexRates(ByVal cdataEntity As cDataEntity) As cDataEntity
    Sub UpdateMCDetails(ByVal cdataEntity As cDataEntity)
    Sub UpdateMCLogs(ByVal cdataEntity As cDataEntity)
    Sub UpdateMCFile(ByVal cdataEntity As cDataEntity)
    Function GetMCSecurity(ByVal cdataEntity As cDataEntity) As cDataEntity
    Function GetAllMCFilesForMCRUN(ByVal cdataEntity As cDataEntity) As cDataEntity
    '55555555555
    Function GetTreasuryAssessmentData(ByVal cdataEntity As cDataEntity) As DataSet
    Function GetTreasuryDetails(ByVal cdataEntity As cDataEntity) As cDataEntity
    Function GetCostTypes(ByVal cdataEntity As cDataEntity) As cDataEntity

End Interface
'<JustInTimeActivation(), _
' EventTrackingEnabled(), _
' ClassInterface(ClassInterfaceType.None), _
' Transaction(TransactionOption.Supported, Isolation:=TransactionIsolationLevel.Serializable, Timeout:=120), _
' ComponentAccessControl(True)> _
Public Class cDataClass
    'Inherits ServicedComponent
    Implements IDataClass

    Private _DBNAME As String = "MoneyCost"
    Private _DEFAULT_CONNECTION As String = "MoneyCost"

    '================================================================
    'METHOD  : GetConnectString
    'PURPOSE : This method will get the DB connection string from the
    '          registry.
    'PARMS   : astrKey [String] = Connection string registry Key name
    'RETURN  : String = The connection string to connect to the
    '          database with
    '================================================================
    <AutoComplete()> _
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
    Public Function Test() As DataTable Implements IDataClass.Test
        Dim lobjDs As DataSet
        lobjDs = New DataSet
        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "UspMoneyCostTest"
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
    <AutoComplete()> _
    Public Function cDataExecute(ByVal cdataEntity As cDataEntity) As DataTable Implements IDataClass.cDataExecute
        Dim lobjDs As DataSet
        lobjDs = New DataSet
        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = cdataEntity.CommonSQL.ToString()
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecSQLText
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
    <AutoComplete()> _
    Public Sub UpdateMCDetails(ByVal cdataEntity As cDataEntity) Implements IDataClass.UpdateMCDetails

        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = False
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = cdataEntity.CommonSQL.ToString()
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecSQLText
            lobjDataAdaptor.ExecuteNonQuery()                ' Call DataAdaptor Class's ExecuteDS method()            
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
        End Try
    End Sub
    <AutoComplete()> _
    Public Sub UpdateMCLogs(ByVal cdataEntity As cDataEntity) Implements IDataClass.UpdateMCLogs
        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = cdataEntity.CommonSQL.ToString()
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecSQLText
            lobjDataAdaptor.ExecuteNonQuery()                ' Call DataAdaptor Class's ExecuteDS method()            
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
        End Try
    End Sub
    <AutoComplete()> _
    Public Sub UpdateMCFile(ByVal cdataEntity As cDataEntity) Implements IDataClass.UpdateMCFile
        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = cdataEntity.CommonSQL.ToString()
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecSQLText
            lobjDataAdaptor.ExecuteNonQuery()                ' Call DataAdaptor Class's ExecuteDS method()            
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
        End Try
    End Sub
    <AutoComplete()> _
       Public Function GetMCSecurity(ByVal cdataEntity As cDataEntity) As cDataEntity Implements IDataClass.GetMCSecurity

        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "UspGetMCSecurity"
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecStoredProcedure
            cdataEntity.OutputDataSet = lobjDataAdaptor.ExecuteDS(cdataEntity.SSOID)                ' Call DataAdaptor Class's ExecuteDS method()
            cdataEntity.OutputDataSet.DataSetName = "MC_SECURITY_DETAIL"
            Return cdataEntity
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try
    End Function

    '5555555555555555555 treasury assessment
    ' Treasury Assessment method will be used to fetch Adder from Moneycost DB
    <AutoComplete()> _
      Public Function GetTreasuryAssessmentData(ByVal cdataEntity As cDataEntity) As DataSet Implements IDataClass.GetTreasuryAssessmentData

        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "UspGetTreasuryAssessmentData"
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecStoredProcedure
            cdataEntity.OutputDataSet = lobjDataAdaptor.ExecuteDS(cdataEntity.ProcessDate)                ' Call DataAdaptor Class's ExecuteDS method()
            cdataEntity.OutputDataSet.DataSetName = "TREASURY_ASSESSMENT_DATA"
            Return cdataEntity.OutputDataSet
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try
    End Function



    <AutoComplete()> _
    Public Function GetMCFileDetails(ByVal cdataEntity As cDataEntity) As cDataEntity Implements IDataClass.GetMCFileDetails

        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "UspGetMCFileDetails"
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecStoredProcedure
            cdataEntity.OutputDataSet = lobjDataAdaptor.ExecuteDS(cdataEntity.MoneyCostID)                ' Call DataAdaptor Class's ExecuteDS method()
            cdataEntity.OutputDataSet.DataSetName = "MC_FILE_DETAIL"
            Return cdataEntity
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try
    End Function
    <AutoComplete()> _
    Public Function GetAllMCFiles(ByVal cdataEntity As cDataEntity) As cDataEntity Implements IDataClass.GetAllMCFiles

        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "UspGetAllMCFiles"
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecStoredProcedure
            cdataEntity.OutputDataSet = lobjDataAdaptor.ExecuteDS()                ' Call DataAdaptor Class's ExecuteDS method()
            cdataEntity.OutputDataSet.DataSetName = "MC_FILESet"
            cdataEntity.OutputDataSet.Tables(0).TableName = "MC_FILE"
            Return cdataEntity
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try
    End Function
    <AutoComplete()> _
        Public Function GetMCFiles(ByVal cdataEntity As cDataEntity) As cDataEntity Implements IDataClass.GetMCFiles

        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "UspGetMCFiles"
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecStoredProcedure
            cdataEntity.OutputDataSet = lobjDataAdaptor.ExecuteDS(cdataEntity.UserSSOID)                ' Call DataAdaptor Class's ExecuteDS method()
            cdataEntity.OutputDataSet.DataSetName = "MC_FILESet"
            cdataEntity.OutputDataSet.Tables(0).TableName = "MC_FILE"

            Return cdataEntity
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try
    End Function
    <AutoComplete()> _
    Public Function GetIndexRates(ByVal cdataEntity As cDataEntity) As cDataEntity Implements IDataClass.GetIndexRates

        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "UspGetIndexRates"
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecStoredProcedure
            cdataEntity.OutputDataSet = lobjDataAdaptor.ExecuteDS(cdataEntity.MoneyCostID, cdataEntity.ProcessDate)                ' Call DataAdaptor Class's ExecuteDS method()
            cdataEntity.OutputDataSet.DataSetName = "INDEX_RATESet"
            cdataEntity.OutputDataSet.Tables(0).TableName = "INDEX_RATE"
            Return cdataEntity
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try
    End Function

    Public Function GetAllMCFilesForMCRUN(ByVal cdataEntity As cDataEntity) As cDataEntity Implements IDataClass.GetAllMCFilesForMCRUN

        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "UspGetAllMCFilesForMCRUN"
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecStoredProcedure
            cdataEntity.OutputDataSet = lobjDataAdaptor.ExecuteDS()                ' Call DataAdaptor Class's ExecuteDS method()
            cdataEntity.OutputDataSet.DataSetName = "MC_FILESet"
            cdataEntity.OutputDataSet.Tables(0).TableName = "MC_FILE"
            Return cdataEntity
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try
    End Function

    ' Add for Treasury Assessment
    <AutoComplete()> _
    Public Function GetTreasuryDetails(ByVal cdataEntity As cDataEntity) As cDataEntity Implements IDataClass.GetTreasuryDetails

        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "USPGetTreasuryDetails"
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecStoredProcedure
            cdataEntity.OutputDataSet = lobjDataAdaptor.ExecuteDS()                ' Call DataAdaptor Class's ExecuteDS method()
            cdataEntity.OutputDataSet.DataSetName = "TREASURY_DETAIL"
            cdataEntity.OutputDataSet.Tables(0).TableName = "TREASURY_DETAIL"
            Return cdataEntity
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try
    End Function

    ' Add for Treasury Assessment
    <AutoComplete()> _
    Public Function GetCostTypes(ByVal cdataEntity As cDataEntity) As cDataEntity Implements IDataClass.GetCostTypes

        Dim lobjDataAdaptor As DataAdaptor         ' DataAdaptor Class Object
        Try
            lobjDataAdaptor = New DataAdaptor(_DBNAME, delGetConfigurationKey(_DEFAULT_CONNECTION), eDBType.cSQLServer)
            lobjDataAdaptor.replaceQuote = True
            '-- Specify the Command Text 
            lobjDataAdaptor.commandSQL = "UspGetCostTypes"
            lobjDataAdaptor.commandType = DataAdaptor.eCommandType.ecStoredProcedure
            cdataEntity.OutputDataSet = lobjDataAdaptor.ExecuteDS()                ' Call DataAdaptor Class's ExecuteDS method()
            cdataEntity.OutputDataSet.DataSetName = "COST_TYPES"
            cdataEntity.OutputDataSet.Tables(0).TableName = "COST_TYPES"
            Return cdataEntity
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(lobjDataAdaptor) Then
                lobjDataAdaptor = Nothing
            End If
            If Not IsNothing(cdataEntity) Then
                cdataEntity = Nothing
            End If
        End Try
    End Function
End Class