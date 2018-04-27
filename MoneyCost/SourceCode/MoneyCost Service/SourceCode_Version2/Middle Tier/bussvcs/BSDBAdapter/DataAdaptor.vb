Imports System.Xml
Imports System.Data
Imports System.Data.OleDb
Public Class DataAdaptor
    Public Enum eCommandType
        ecStoredProcedure
        ecSQLText
    End Enum
    Public Enum eDBType
        cSQLServer
        cOledb
        cOracleServer
    End Enum
    Public Enum eReturnXMLType As Integer
        eElemementBased
        eAttributeBased
    End Enum
    Private _DBNAME As String = ""
    Private _DEFAULT_CONNECTION As String = ""
    Private _DBType As eDBType = eDBType.cSQLServer
    Private _CommandType As eCommandType = eCommandType.ecSQLText
    Private _ReturnXMLType As eReturnXMLType = eReturnXMLType.eElemementBased
    Private _ReplaceQuote As Boolean = True
    Private _CommandSQL As String

    'Default test SQL statement. All Other Sql statements will defined outside of this class and passed to this class
    'using the commandSQL property
    Private Const cTestSQL As String = "SELECT GETDATE()"
    Private Const cDELIMITER As String = "|~^"
    Private Const cINVALID_PARMS As String = "Invalid SQL command"
    Public WriteOnly Property DBNAME() As String
        Set(ByVal Value As String)
            _DBNAME = Value
        End Set
    End Property
    Public WriteOnly Property DEFAULT_CONNECTION() As String
        Set(ByVal Value As String)
            _DEFAULT_CONNECTION = Value
        End Set
    End Property
    Public WriteOnly Property replaceQuote() As Boolean
        Set(ByVal abQuoteReplacement As Boolean)
            _ReplaceQuote = abQuoteReplacement
        End Set
    End Property
    Public WriteOnly Property commandType() As eCommandType
        Set(ByVal aiCommandType As eCommandType)
            _CommandType = aiCommandType
        End Set
    End Property
    Public WriteOnly Property returnXMLType() As eReturnXMLType
        Set(ByVal aiReturnXMLType As eReturnXMLType)
            _ReturnXMLType = aiReturnXMLType
        End Set
    End Property
    Public WriteOnly Property commandSQL() As String
        Set(ByVal astrCommandSQL As String)
            _CommandSQL = astrCommandSQL
        End Set
    End Property
    Public WriteOnly Property DBType() As eDBType
        Set(ByVal aecDBType As eDBType)
            _DBType = aecDBType
        End Set
    End Property
    Public Sub New(ByVal astrDBName As String, ByVal aeDBType As eDBType)
        _DBNAME = astrDBName
        _DBType = aeDBType
    End Sub
    Public Sub New(ByVal astrDBName As String, ByVal ConStr As String, ByVal aeDBType As eDBType)
        _DBNAME = astrDBName
        _DBType = aeDBType
        _DEFAULT_CONNECTION = ConStr
    End Sub
    Public Sub New(ByVal astrDBName As String, ByVal ConStr As String)
        _DBNAME = astrDBName
        _DEFAULT_CONNECTION = ConStr
    End Sub
    Public Sub New(ByVal aeDBType As eDBType)
        _DBType = aeDBType
    End Sub
    Public Sub New(ByVal astrDBName As String)
        _DBNAME = astrDBName
    End Sub
    Public Sub New()
        '-- Do nothing
    End Sub
    Private Function GetConnectionString() As String
        Return _DEFAULT_CONNECTION
    End Function


    '================================================================
    'METHOD  : GetSQLStatement
    'PURPOSE : Based on an Action ID and an array of parms
    '          GetSQLStatement returns the SQL Statement in
    '          "executeable" form i.e. all parms are replaced in the
    '          string with the parms provided in Parms()
    'PARMS   : alActionID [eDCActions] = SQL Statement to Execute
    '          aobjSQLParams [Object] = Parms to Populate SQL Statement
    '  '        with
    'RETURN  : String = Completed SQL statement
    '================================================================

    Public Function GetSQLStatement(ByVal ParamArray aobjSQLParams As Object()) As String

        Dim lstrSQL As String
        Dim liPos As Integer
        Dim liLoopCtr As Integer
        Dim liUBound As Integer
        Try

            lstrSQL = _CommandSQL
            'If there are no substitution parms, we're done
            liPos = InStr(1, lstrSQL, cDELIMITER)

            If liPos = 0 Then
                GetSQLStatement = lstrSQL
                Exit Function
            End If

            If IsArray(aobjSQLParams) Then
                If UBound(aobjSQLParams) = -1 Then
                    Err.Raise(cINVALID_PARMS)
                End If
            End If

            'Loop through each of the param
            liUBound = UBound(aobjSQLParams)
            For liLoopCtr = liUBound + 1 To 1 Step -1

                'First, replace any apostrophes (') in the parm with two ('') for SQL Server.
                If _ReplaceQuote Then aobjSQLParams(liLoopCtr - 1) = Replace(aobjSQLParams(liLoopCtr - 1), "'", "''")

                'Then Insert the parm into the SQL statement.
                lstrSQL = Replace(lstrSQL, cDELIMITER & CStr(liLoopCtr), aobjSQLParams(liLoopCtr - 1))
            Next liLoopCtr

            'REPLACE ANY STRINGS THAT CAME UP 'NULL' WITH NULL
            lstrSQL = Replace(lstrSQL, "'NULL'", "NULL")

            'Make sure all substitution parms have been swapped out
            liPos = InStr(1, lstrSQL, cDELIMITER)
            If liPos > 0 Then Err.Raise(cINVALID_PARMS)

        Catch ex As Exception
            Throw
        End Try
        'Return Completed SQL statement
        GetSQLStatement = lstrSQL

        Exit Function
    End Function
    '================================================================
    'METHOD  : ExecuteDS
    'PURPOSE : Based on an Action ID and an array of parms
    '          ExecuteDS returns the Dataset .This function can be used to get the Dataset
    '          by passing the Stored Procedure name and parameters or by pssing 
    '          SQL Statement only.
    'PARMS   : alActionID [eDCActions] = SQL Statement to Execute
    '          aobjSQLParams [Object] = Parms to Populate SQL Statement
    '          with
    'RETURN  : Dataset = Dataset Based on SQL Statement and Params to Populate the 
    '          SQL Statement         
    '================================================================

    Public Function ExecuteDS(ByVal ParamArray aobjSQLParams As Object()) As DataSet
        Dim lstrSQLCommand As String
        Dim liLoopCtr As Integer
        Dim liUBound As Integer
        Dim lobjDS As DataSet
        Dim lobjDB As Database

        Try

            lstrSQLCommand = GetSQLStatement(aobjSQLParams)
            ''Replace Single Quotes and NULL in Param
            If _CommandType = eCommandType.ecStoredProcedure Then
                liUBound = UBound(aobjSQLParams)
                For liLoopCtr = liUBound + 1 To 1 Step -1
                    'First, replace any apostrophes (') in the parm with two ('') for SQL Server.
                    If _ReplaceQuote Then aobjSQLParams(liLoopCtr - 1) = Replace(aobjSQLParams(liLoopCtr - 1), "'", "''")
                    aobjSQLParams(liLoopCtr - 1) = Replace(aobjSQLParams(liLoopCtr - 1), "'NULL'", "NULL")
                Next
            End If
            lobjDB = New Database(_DBNAME, _DBType, GetConnectionString())

            lobjDS = lobjDB.ExecuteDS(lstrSQLCommand, _CommandType, aobjSQLParams)
            Return lobjDS
        Catch ex As Exception
            Throw
        Finally
            lobjDB = Nothing
        End Try
    End Function
    '================================================================
    'METHOD  : ExecuteDR
    'PURPOSE : Based on an Action ID , an array of parms
    '          and Node Name Gets the Data Reader and read the Data Reader value
    '          in the XML format and returns it as string
    'PARMS   : alActionID [eDCActions] = SQL Statement to Execute
    '          aobjSQLParams [Object] = Parms to Populate SQL Statement
    '          astrNodeName = Name of the Node for the XML File.
    '          with
    'RETURN  : String =  Data Reader vales in the XML formated string.
    '================================================================

    'Public Function ExecuteDR(ByVal astrNodeName As String, ByVal ParamArray aobjSQLParams As Object()) As String
    '    Dim lstrSQLCommand As String
    '    Dim lobjDataReader As IDataReader = Nothing
    '    Dim lstrOutXML As String
    '    Dim lobjDB As Database

    '    Try
    '        lstrSQLCommand = GetSQLStatement(aobjSQLParams)
    '        lobjDB = New Database(_DBNAME, _DBType, GetConnectionString)

    '        lobjDataReader = lobjDB.ExecuteDataset(lstrSQLCommand, System.Data.CommandType.Text, aobjSQLParams)

    '        lstrOutXML = ReaderToXML(lobjDataReader, astrNodeName)
    '        Return lstrOutXML
    '    Catch Ex As Exception
    '        Throw
    '    Finally
    '        lobjDB = Nothing
    '        If Not lobjDataReader Is Nothing Then
    '            lobjDataReader.Close()
    '            lobjDataReader = Nothing
    '        End If
    '    End Try
    'End Function

    Public Function ExecuteCsvDS(ByVal sqlString As String) As DataSet

        Dim lobjDataset As DataSet = Nothing
        Dim lobjDB As Database

        Try
            lobjDB = New Database(_DBNAME, _DBType, GetConnectionString)
            lobjDataset = lobjDB.ExecuteDS(sqlString, System.Data.CommandType.Text)
            Return lobjDataset
        Catch Ex As Exception
            Throw
        Finally
            lobjDB = Nothing
            If Not lobjDataset Is Nothing Then
                lobjDataset.Dispose()
                lobjDataset = Nothing
            End If
        End Try
    End Function
    '================================================================
    'METHOD  : ExecuteXML
    'PURPOSE : Based on an Action ID and an array of parms
    '          ExecuteXML returns the XML formated string.
    'PARMS   : alActionID [eDCActions] = SQL Statement to Execute
    '          aobjSQLParams [Object] = Parms to Populate SQL Statement
    '          with
    'RETURN  : String = XML formated string
    '================================================================

    Public Function ExecuteXML(ByVal ParamArray aobjSQLParams As Object()) As String
        Dim strSQLCommand As String
        Dim lstrOutXML As String
        Dim lobjDB As Database
        Try
            lstrOutXML = ""

            lobjDB = New Database(_DBNAME, _DBType, GetConnectionString())
            strSQLCommand = GetSQLStatement(aobjSQLParams)
            If _CommandType = eCommandType.ecStoredProcedure Then
                lstrOutXML = lobjDB.ExecuteXML(strSQLCommand, eCommandType.ecStoredProcedure, aobjSQLParams)
            ElseIf _CommandType = eCommandType.ecSQLText Then
                If _ReturnXMLType = eReturnXMLType.eElemementBased Then
                    strSQLCommand += " FOR XML AUTO, ELEMENTS "
                Else
                    strSQLCommand += " FOR XML AUTO "
                End If
                lstrOutXML = lobjDB.ExecuteXML(strSQLCommand, eCommandType.ecSQLText, aobjSQLParams)
            End If

            Return lstrOutXML

        Catch Ex As Exception
            Throw
        Finally
            lobjDB = Nothing
        End Try
    End Function
    '================================================================
    'METHOD  : ExecuteNonQuery
    'PURPOSE : Based on an Action ID and an array of parms
    '          It Executes the SQL Statement or Strored Procedure
    '          and returns the no of record affected.
    'PARMS   : alActionID [eDCActions] = SQL Statement to Execute
    '          aobjSQLParams [Object] = Parms to Populate SQL Statement
    '          with
    'RETURN  : Integer = No of rows affected.
    '================================================================

    Public Function ExecuteNonQuery(ByVal ParamArray aobjSQLParams As Object()) As Long
        Dim strSQLCommand As String
        Dim llResult As Long
        Dim lobjDB As Database

        Try
            strSQLCommand = GetSQLStatement(aobjSQLParams)
            lobjDB = New Database(_DBNAME, _DBType, GetConnectionString())
            llResult = lobjDB.ExecuteNonQuery(strSQLCommand)
            Return llResult
        Catch ex As Exception
            Throw
        Finally
            lobjDB = Nothing
        End Try
    End Function

    Public Function ExecuteNonQuery(ByVal sqlStr As String) As Long
        Dim llResult As Long
        Dim lobjDB As Database

        Try
            lobjDB = New Database(_DBNAME, _DBType, GetConnectionString())
            llResult = lobjDB.ExecuteNonQuery(sqlStr)
            Return llResult
        Catch ex As Exception
            Throw
        Finally
            lobjDB = Nothing
        End Try
    End Function
    '================================================================
    'METHOD  : ExecuteScalar
    'PURPOSE : Based on an Action ID and an array of parms
    '          ExecuteScalar returns the Object.We can use this function
    '          To execute a query where we need to return the @@RowID kind
    '          of value from the query.
    'PARMS   : alActionID [eDCActions] = SQL Statement to Execute
    '          aobjSQLParams [Object] = Parms to Populate SQL Statement
    '          with
    'RETURN  : Object = Having only one value.The Object type should be typecast.
    '================================================================

    Public Function ExecuteScalar(ByVal ParamArray aobjSQLParams As Object()) As Object
        Dim strSQLCommand As String
        Dim lobjResult As Object
        Dim lobjDB As Database
        Try
            strSQLCommand = GetSQLStatement(aobjSQLParams)
            lobjDB = New Database(_DBNAME, _DBType, GetConnectionString())
            lobjResult = lobjDB.ExecuteScalar(strSQLCommand)
            Return lobjResult
        Catch ex As Exception
            Throw
        Finally
            lobjDB = Nothing
        End Try
    End Function
    '================================================================
    'METHOD  : ReaderToXML
    'PURPOSE : Based on the SQL Data Reader and optinal Parent Node Name 
    'PARMS   : objReader  = A Data Reader object to read the data from
    '          Optional ParentNodeName = This is an optional Parameter.If this is supplied it will
    '          read the data in the data reader format it in the XML string and concatenate it with the 
    '          parent node.
    '          with
    'RETURN  : String = XML formated  string
    '================================================================

    Public Function ReaderToXML(ByVal objReader As IDataReader, _
         Optional ByVal ParentNodeName As String = "") As String

        Dim sXML As String
        Dim intColumnCount As Integer
        Dim intCtr As Integer
        Dim sValue As String
        Try
            sXML = ""
            ParentNodeName = Trim(ParentNodeName)

            intColumnCount = objReader.FieldCount
            If ParentNodeName <> "" Then sXML += "<" & ParentNodeName & "Set>"

            Do While objReader.Read

                sXML += "<" & ParentNodeName & ">"
                'Loop through each row
                For intCtr = 0 To intColumnCount - 1
                    'Get the Value of each column
                    'Does not include nodes for null/blank values

                    If Not objReader.IsDBNull(intCtr) Then
                        sValue = objReader.Item(intCtr).ToString
                        If Trim(sValue) <> "" Then
                            sXML += "<" & objReader.GetName(intCtr) & ">" & XMLizeString(sValue) & "</" & objReader.GetName(intCtr) & ">"
                        End If
                    End If
                Next
                sXML += "</" & ParentNodeName & ">"
            Loop
            If ParentNodeName <> "" Then sXML += "</" & ParentNodeName & "Set>"

        Catch ex As Exception
            sXML = ""
            Throw
        End Try
        Return sXML
    End Function
    '================================================================
    'METHOD  : XMLizeString
    'PURPOSE : Based on sInput as a  string 
    '            
    'PARMS   : sInput  = A  string
    '          
    'RETURN  : String = if input string is a alphanumeric string then
    '          itv return it the Return " <![CDATA[" & sInput & "]]>" format.
    '================================================================

    Private Function XMLizeString(ByVal sInput As String) As String
        'SHORTENED VERSION TO REDUCE EXECUTION TIME
        'Return " <![CDATA[" & sInput & "]]>"
        'THIS WILL INCREASE THE SIZE OF YOUR XML String
        If Not (IsAlphaNumeric(sInput)) Then
            Return " <![CDATA[" & sInput & "]]>"
        Else
            Return sInput
        End If
    End Function

    '================================================================
    'METHOD  : IsAlphaNumeric
    'PURPOSE : Based on the string it checks that  the string is alpha numeric or not.
    'PARMS   : TestString = A string
    '          
    'RETURN  : Boolean = TRUE if the parameter string is an Alphanumeric else FALSE
    '================================================================

    Private Function IsAlphaNumeric(ByVal astrTestString As String) As Boolean

        Dim sTemp As String
        Dim iLen As Integer
        Dim iCtr As Integer
        Dim sChar As String

        'returns true if all characters in a string are alphabetical
        '   or numeric
        'returns false otherwise or for empty string
        sTemp = astrTestString
        iLen = Len(sTemp)
        If iLen > 0 Then
            For iCtr = 1 To iLen
                sChar = Mid(sTemp, iCtr, 1)
                If Not sChar Like "[0-9A-Za-z.:, ]" Then _
                     Exit Function
            Next
            IsAlphaNumeric = True
        End If
    End Function

End Class
