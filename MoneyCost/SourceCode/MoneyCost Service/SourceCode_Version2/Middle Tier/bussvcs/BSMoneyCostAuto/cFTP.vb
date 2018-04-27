Imports System.Reflection
Imports BSMoneyCostEntity
Public Class cFTP

    Private Structure FILETIME
        Dim dwLowDateTime As Long
        Dim dwHighDateTime As Long
    End Structure

    Private Structure WIN32_FIND_DATA
        Dim dwFileAttributes As Long
        Dim ftCreationTime As FILETIME
        Dim ftLastAccessTime As FILETIME
        Dim ftLastWriteTime As FILETIME
        Dim nFileSizeHigh As Long
        Dim nFileSizeLow As Long
        Dim dwReserved0 As Long
        Dim dwReserved1 As Long
        Dim cFileName As String ''* MAX_PATH
        Dim cAlternate As String ''''* 14
    End Structure
    Private Const ERROR_NO_MORE_FILES = 18
    Private Const MAX_PATH = 260
    Public Enum FileTransferType
        ftAscii = FTP_TRANSFER_TYPE_ASCII
        ftBinary = FTP_TRANSFER_TYPE_BINARY
    End Enum

#Region "Private Variables"
    Private _sHostName As String
    Private sUserName As String
    Private _sPassword As String
    Private _sDirectory As String
#End Region

    Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
        (ByVal hFind As Long, ByVal lpvFindData As WIN32_FIND_DATA) As Long

    Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
        (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
    ByVal lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

    Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
        (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
          ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
          ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

    Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
    (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
          ByVal lpszRemoteFile As String, _
          ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

    Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
        (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

    ' Initializes an application's use of the Win32 Internet functions
    Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
        (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
        ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
    Private Const INTERNET_OPEN_TYPE_DIRECT = 1
    Private Const INTERNET_OPEN_TYPE_PROXY = 3
    Private Const INTERNET_INVALID_PORT_NUMBER = 0

    Private Const FTP_TRANSFER_TYPE_ASCII = &H1
    Private Const FTP_TRANSFER_TYPE_BINARY = &H1
    Private Const INTERNET_FLAG_PASSIVE = &H8000000

    Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
        (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
        ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long

    Private Const ERROR_INTERNET_EXTENDED_ERROR = 12003

    Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    ByVal lpdwError As Long, _
        ByVal lpszBuffer As String, _
    ByVal lpdwBufferLength As Long) As Boolean

    ' Type of service to access.
    Private Const INTERNET_SERVICE_FTP = 1

    Private Const INTERNET_FLAG_RELOAD = &H80000000
    Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
    Private Const INTERNET_FLAG_MULTIPART = &H200000

    Private Declare Function FtpOpenFile Lib "wininet.dll" Alias _
            "FtpOpenFileA" (ByVal hFtpSession As Long, _
            ByVal sFileName As String, ByVal lAccess As Long, _
            ByVal lFlags As Long, ByVal lContext As Long) As Long
    Private Declare Function FtpDeleteFile Lib "wininet.dll" _
        Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
        ByVal lpszFileName As String) As Boolean

    Private Declare Function FtpRenameFile Lib "wininet.dll" Alias _
        "FtpRenameFileA" (ByVal hFtpSession As Long, _
        ByVal sExistingName As String, _
        ByVal sNewName As String) As Boolean

    ' Closes a single Internet handle or a subtree of Internet handles.
    Private Declare Function InternetCloseHandle Lib "wininet.dll" _
    (ByVal hInet As Long) As Integer
    '
    ' Our Defined Errors
    '
    Public Enum errFtpErrors
        errCannotConnect = vbObjectError + 2001
        errNoDirChange = vbObjectError + 2002
        errCannotRename = vbObjectError + 2003
        errCannotDelete = vbObjectError + 2004
        errNotConnectedToSite = vbObjectError + 2005
        errGetFileError = vbObjectError + 2006
        errInvalidProperty = vbObjectError + 2007
        errFatal = vbObjectError + 2008
    End Enum


    '
    ' Error messages
    '
    Private Const ERRCHANGEDIRSTR As String = "Cannot Change Directory to %s. It either doesn't exist, or is protected"
    Private Const ERRCONNECTERROR As String = "Cannot Connect to %s using User and Password Parameters"
    Private Const ERRNOCONNECTION As String = "Not Connected to FTP Site"
    Private Const ERRNODOWNLOAD As String = "Couldn't Get File %s from Server"
    Private Const ERRNORENAME As String = "Couldn't Rename File %s"
    Private Const ERRNODELETE As String = "Couldn't Delete File %s from Server"
    Private Const ERRALREADYCONNECTED As String = "You cannot change this property while connected to an FTP server"
    Private Const ERRFATALERROR As String = "Cannot get Connection to WinInet.dll !"

    '
    ' Session Identifier to Windows
    '
    Private Const SESSION As String = "CGFtp Instance"
    '
    ' Our INET handle
    '
    Private mlINetHandle As Long
    '
    ' Our FTP Connection Handle
    '
    Private mlConnection As Long
    '
    ' Standard FTP properties for this class
    '
    Private msHostAddress As String
    Private msUser As String
    Private msPassword As String
    Private msDirectory As String

    Private Sub New()
        '
        ' Create Internet session handle
        '

        mlINetHandle = InternetOpen(SESSION, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)

        If mlINetHandle = 0 Then
            mlConnection = 0
            Throw New Exception("CGFTP::Class_Initialise " & ERRFATALERROR)
        End If

        mlConnection = 0
    End Sub

    'Private Sub Class_Terminate()
    '    '
    '    ' Kill off any connection
    '    '
    '    If mlConnection <> 0 Then
    '        InternetCloseHandle(mlConnection)
    '    End If
    '    '
    '    ' Kill off API Handle
    '    '
    '    If mlINetHandle <> 0 Then
    '        InternetCloseHandle(mlINetHandle)
    '    End If
    '    mlConnection = 0
    '    mlINetHandle = 0
    'End Sub
    Public Property Host() As String
        Get
            Return _sHostName
        End Get
        Set(ByVal value As String)
            _sHostName = value
        End Set
    End Property
    Public Property User() As String
        Get
            Return sUserName
        End Get
        Set(ByVal value As String)
            sUserName = value
        End Set
    End Property
    Public Property Password() As String
        Get
            Return _sPassword
        End Get
        Set(ByVal value As String)
            _sPassword = value
        End Set
    End Property
    Public Property Directory() As String
        Get
            Return _sDirectory
        End Get
        Set(ByVal value As String)
            _sDirectory = value
        End Set
    End Property

    Public Function Connect(Optional ByVal objFtpEntity As FTPEntity = Nothing) As cDataEntity
        Dim lobjEntity As New cDataEntity

        '
        ' Connect to the FTP server
        '
        Dim lstrError As String
        '
        ' If we already have a connection then raise an error
        '

        Try
            If mlConnection <> 0 Then
                Throw New Exception("CGFTP::Connect :-You are already connected to FTP Server " & msHostAddress)
                lobjEntity.OutputString = "False"
                Connect = lobjEntity
                Exit Function
            End If
            '
            ' Overwrite any existing properties if they were supplied in the
            ' arguments to this method
            '
            If Len(objFtpEntity.HOST) > 0 Then
                msHostAddress = objFtpEntity.HOST
            End If

            If Len(objFtpEntity.FTPUser) > 0 Then
                msUser = objFtpEntity.FTPUser
            End If

            If Len(objFtpEntity.FTPPassword) > 0 Then
                msPassword = objFtpEntity.FTPPassword
            End If

            '
            ' Connect !
            '

            If Len(msHostAddress) = 0 Then
                Throw New Exception("CGFTP::Connect:- No Host Address Specified!")
            End If

            mlConnection = InternetConnect(mlINetHandle, msHostAddress, INTERNET_INVALID_PORT_NUMBER, _
                msUser, msPassword, INTERNET_SERVICE_FTP, 0, 0)

            '
            ' Check for connection errors
            '
            If mlConnection = 0 Then
                lstrError = Replace(ERRCONNECTERROR, "%s", msHostAddress)
                lstrError = lstrError & vbCrLf & GetINETErrorMsg(Err.LastDllError)
                Throw New Exception("CGFTP::Connect:-" & lstrError)
            End If

            lobjEntity.OutputString = "True"
            Connect = lobjEntity

        Catch ex As Exception
            lobjEntity.OutputString = "False"
            Connect = lobjEntity
            Throw New Exception(ex.Message, ex)
            'Throw
        Finally
            lobjEntity = Nothing
            objFtpEntity = Nothing
        End Try

    End Function

    Public Function Disconnect() As cDataEntity
        Dim lobjEntity As New cDataEntity
        '
        ' Disconnect, only if connected !
        '
        Try
            If mlConnection <> 0 Then
                InternetCloseHandle(mlConnection)
                mlConnection = 0
                lobjEntity.OutputString = "True"
                Disconnect = lobjEntity
            Else
                Throw New Exception("CGFTP::Disconnect:-" & ERRNOCONNECTION)
                lobjEntity.OutputString = "False"
                Disconnect = lobjEntity
            End If
        Catch ex As Exception
            lobjEntity.OutputString = "False"
            Disconnect = lobjEntity
            Throw New Exception(ex.Message, ex)
            'Throw
        Finally
            msHostAddress = ""
            msUser = ""
            msPassword = ""
            msDirectory = ""
            lobjEntity = Nothing
        End Try
    End Function

    Public Function GetDirectoryList(Optional ByVal objFtpEntity As FTPEntity = Nothing) As cDataEntity
        Dim objEntity As New cDataEntity
        Dim pData As WIN32_FIND_DATA
        Dim llFind As Long
        Dim llLastError As Long
        Dim lstrFilter As String
        Dim lstrItemName As String
        Dim lstrFileNames As String = ""
        Dim lbRet As Boolean


        Try
            '
            ' Check if already connected, else raise an error
            '
            If mlConnection = 0 Then
                Throw New Exception("CGFTP::GetDirectoryList:-" & ERRNOCONNECTION)
            End If

            '
            ' Change directory if required
            '
            If Len(objFtpEntity.Directory) > 0 Then
                RemoteChDir(objFtpEntity.Directory)
            End If

            pData.cFileName = vbNullChar  'String$(MAX_PATH, vbNullChar)

            If Len(objFtpEntity.FilterString) > 0 Then
                lstrFilter = objFtpEntity.FilterString
            Else
                lstrFilter = "*.*"
            End If
            '
            ' Get the first file in the directory
            '
            llFind = FtpFindFirstFile(mlConnection, lstrFilter, pData, 0, 0)
            llLastError = Err.LastDllError
            '
            ' If no files, then return an empty recordset.
            '
            If llFind = 0 Then
                If llLastError = ERROR_NO_MORE_FILES Then
                    ' Empty directory
                    objEntity.OutputString = Nothing
                    GetDirectoryList = objEntity
                    Exit Function
                Else
                    Throw New Exception("cFTP::GetDirectoryList:- " & "Error looking at directory " & objFtpEntity.Directory & "\" & objFtpEntity.FilterString)
                End If
                objEntity.OutputString = Nothing
                GetDirectoryList = objEntity
                Exit Function
            End If
            '
            ' Add the first found file into the recordset
            '
            lstrItemName = Left$(pData.cFileName, InStr(1, pData.cFileName, vbNullChar, vbBinaryCompare) - 1)
            lstrFileNames = lstrFileNames & lstrItemName & ","
            '
            ' Get the rest of the files in the list
            '
            Do
                pData.cFileName = vbNullChar ' String$(MAX_PATH, vbNullChar)
                lbRet = InternetFindNextFile(llFind, pData)
                If Not (lbRet) Then
                    llLastError = Err.LastDllError
                    If llLastError = ERROR_NO_MORE_FILES Then
                        Exit Do
                    Else
                        InternetCloseHandle(llFind)
                        Throw New Exception("cFTP::GetDirectoryList" & "Error looking at directory " & objFtpEntity.Directory & "\" & objFtpEntity.FilterString)
                        objEntity.OutputString = Nothing
                        GetDirectoryList = objEntity
                        Exit Function
                    End If
                Else
                    lstrItemName = Left$(pData.cFileName, InStr(1, pData.cFileName, vbNullChar, vbBinaryCompare) - 1)
                    lstrFileNames = lstrFileNames & lstrItemName & ","
                End If
            Loop
            '
            ' Close the 'find' handle
            '
            InternetCloseHandle(llFind)
            objEntity.OutputString = lstrFileNames
            GetDirectoryList = objEntity
        Catch ex As Exception
            If llFind <> 0 Then
                InternetCloseHandle(llFind)
            End If
            objEntity.OutputString = lstrFileNames
            GetDirectoryList = objEntity
            Throw New Exception(ex.Message, ex)
            'Throw
        Finally
            objEntity = Nothing
            objFtpEntity = Nothing
        End Try

    End Function

    Public Function GetFile(ByVal astrServerFileAndPath As String, _
                            ByVal astrDestinationFileAndPath As String, _
    Optional ByVal TransferType As FileTransferType = FileTransferType.ftAscii) As String
        '
        ' Get the specified file to the desired location using the specified
        ' file transfer type
        '
        Dim lbRet As Boolean
        Dim lstrError As String
        Dim lstrReturnString As String


        '
        ' If not connected, raise an error
        '
        lstrReturnString = ""
        Try
            If mlConnection = 0 Then
                'On Error GoTo 0
                'Err.Raise errNotConnectedToSite, "CGFTP::GetFile", ERRNOCONNECTION
                lstrError = ERRNODOWNLOAD
                lstrError = Replace(lstrError, "%s", astrServerFileAndPath)
                lstrReturnString = "<ERROR_DETAILS>" & lstrError & "</ERROR_DETAILS>"

                Throw New Exception("BSSTMoneyCostAuto.cFTP_GetFile(): Error occured - " & Err.Description)

                GetFile = lstrReturnString
                Exit Function
            End If

            '
            ' Get the file
            '
            lbRet = FtpGetFile(mlConnection, astrServerFileAndPath, astrDestinationFileAndPath, False, INTERNET_FLAG_RELOAD, TransferType, 0)

            If lbRet = False Then
                lstrError = ERRNODOWNLOAD
                lstrError = Replace(lstrError, "%s", astrServerFileAndPath)
                'On Error GoTo 0
                lstrReturnString = "<ERROR_DETAILS>" & lstrError & "</ERROR_DETAILS>"

                Throw New Exception("BSSTMoneyCostAuto.cFTP_GetFile(): Error occured - " & Err.Description)

                GetFile = lstrReturnString
                Exit Function
                'Err.Raise errGetFileError, "CGFTP::GetFile", lstrError
            End If

            GetFile = lstrReturnString
        Catch ex As Exception
            lstrError = ERRNODOWNLOAD
            lstrError = Replace(lstrError, "%s", astrServerFileAndPath)
            If lstrError <> "" Then
                lstrReturnString = "<ERROR_DETAILS>" & lstrError & "</ERROR_DETAILS>"
            Else
                lstrReturnString = "<ERROR_DETAILS>" & _
                                        "<ERROR_NUMBER>" & Err.Number & "</ERROR_NUMBER>" & _
                                        "<ERROR_DESCRIPTION>" & Err.Description & "</ERROR_DESCRIPTION>" & _
                                        "<ERROR_SOURCE>" & Err.Source & "::GetFile()</ERROR_SOURCE>" & _
                                    "</ERROR_DETAILS>"
            End If
            Throw New Exception("BSSTMoneyCostAuto.cFTP_GetFile(): Error occured - " & Err.Description)
            GetFile = lstrReturnString
            'Throw New Exception(ex.Message, ex)
            'Throw
        Finally

        End Try


    End Function

    Public Function PutFile(ByVal lstrLocalFileAndPath As String, _
                            ByVal lstrServerFileAndPath As String, _
    Optional ByVal TransferType As FileTransferType = Nothing) As String


        Dim lbRet As Boolean
        Dim lstrError As String
        Dim lstrReturnString As String

        '
        ' If not connected, raise an error!
        '
        lstrReturnString = ""
        Try
            If mlConnection = 0 Then
                'On Error GoTo 0
                'Err.Raise errNotConnectedToSite, "CGFTP::PutFile", ERRNOCONNECTION
                lstrError = ERRNODOWNLOAD
                lstrError = Replace(lstrError, "%s", lstrServerFileAndPath)
                lstrReturnString = "<ERROR_DETAILS>" & lstrError & "</ERROR_DETAILS>"

                Throw New Exception("BSSTMoneyCostAuto.cFTP_PutFile(): Error occured - " & Err.Description)
                PutFile = lstrReturnString
                Exit Function
            End If

            lbRet = FtpPutFile(mlConnection, lstrLocalFileAndPath, lstrServerFileAndPath, TransferType, 0)

            If lbRet = False Then
                'lstrError = ERRNODOWNLOAD
                'lstrError = Replace(lstrError, "%s", lstrServerFileAndPath)
                'On Error GoTo 0
                'PutFile = False
                'lstrError = lstrError & vbCrLf & GetINETErrorMsg(Err.LastDllError)
                'Err.Raise errCannotRename, "CGFTP::PutFile", lstrError
                lstrError = ERRNODOWNLOAD
                lstrError = Replace(lstrError, "%s", lstrServerFileAndPath)
                lstrReturnString = "<ERROR_DETAILS>" & lstrError & "</ERROR_DETAILS>"

                Throw New Exception("BSSTMoneyCostAuto.cFTP_PutFile(): Error occured - " & Err.Description)

                PutFile = lstrReturnString
                Exit Function
            End If
            PutFile = lstrReturnString
        Catch ex As Exception
            lstrError = ERRNODOWNLOAD
            lstrError = Replace(lstrError, "%s", lstrServerFileAndPath)
            If lstrError <> "" Then
                lstrReturnString = "<ERROR_DETAILS>" & lstrError & "</ERROR_DETAILS>"
            Else
                lstrReturnString = "<ERROR_DETAILS>" & _
                                        "<ERROR_NUMBER>" & Err.Number & "</ERROR_NUMBER>" & _
                                        "<ERROR_DESCRIPTION>" & Err.Description & "</ERROR_DESCRIPTION>" & _
                                        "<ERROR_SOURCE>" & Err.Source & "::PutFile()</ERROR_SOURCE>" & _
                                    "</ERROR_DETAILS>"
            End If

            Throw New Exception("BSSTMoneyCostAuto.cFTP_PutFile(): Error occured - " & Err.Description)
            PutFile = lstrReturnString
            'Throw New Exception(ex.Message, ex)
            'Throw
        End Try

    End Function

    Public Function RenameFile(ByVal ExistingName As String, ByVal NewName As String) As Boolean

        Dim bRet As Boolean
        Dim sError As String

        '
        ' If not connected, raise an error
        '
        Try
            If mlConnection = 0 Then
                Throw New Exception("CGFTP::RenameFile:-" & ERRNOCONNECTION)
            End If

            bRet = FtpRenameFile(mlConnection, ExistingName, NewName)
            '
            ' Raise an error if we couldn't rename the file (most likely that
            ' a file with the new name already exists
            '
            If bRet = False Then
                sError = ERRNORENAME
                sError = Replace(sError, "%s", ExistingName)
                RenameFile = False
                sError = sError & vbCrLf & GetINETErrorMsg(Err.LastDllError)
                Throw New Exception("CGFTP::RenameFile" & sError)
            End If

            RenameFile = True
        Catch ex As Exception
            Throw New Exception("cFTP::RenameFile" & Err.Description)
            'Throw New Exception(ex.Message, ex)
            'Throw
        End Try

    End Function

    Public Function DeleteFile(ByVal ExistingName As String) As Boolean

        Dim bRet As Boolean
        Dim sError As String

        '
        ' Check for a connection
        '
        Try
            If mlConnection = 0 Then
                Throw New Exception("CGFTP::DeleteFile:-" & ERRNOCONNECTION)
            End If

            bRet = FtpDeleteFile(mlConnection, ExistingName)
            '
            ' Raise an error if the file couldn't be deleted
            '
            If bRet = False Then
                sError = ERRNODELETE
                sError = Replace(sError, "%s", ExistingName)
                Throw New Exception("CGFTP::DeleteFile:-" & sError)
            End If

            DeleteFile = True
        Catch ex As Exception
            'Throw New Exception(ex.Message, ex)
            Throw
        End Try


    End Function

    Private Sub RemoteChDir(ByVal sDir As String)

        '
        ' Remote Change Directory Command through WININET
        '
        Dim sPathFromRoot As String
        Dim bRet As Boolean
        Dim sError As String
        '
        ' Needs standard Unix Convention
        '
        sDir = Replace(sDir, "\", "/")
        '
        ' Check for a connection
        '
        Try
            If mlConnection = 0 Then
                Throw New Exception("CGFTP::RemoteChDir" & ERRNOCONNECTION)
                Exit Sub
            End If

            If Len(sDir) = 0 Then
                Exit Sub
            Else
                sPathFromRoot = sDir
                If Len(sPathFromRoot) = 0 Then
                    sPathFromRoot = "/"
                End If
                bRet = FtpSetCurrentDirectory(mlConnection, sPathFromRoot)
                '
                ' If we couldn't change directory - raise an error
                '
                If bRet = False Then
                    sError = ERRCHANGEDIRSTR
                    sError = Replace(sError, "%s", sDir)
                    Throw New Exception("CGFTP::ChangeDirectory" & sError)
                End If
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
            'Throw
        End Try

        Exit Sub

    End Sub

    Private Function GetINETErrorMsg(ByVal ErrNum As Long) As String
        Dim lError As Long
        Dim lLen As Long
        Dim sBuffer As String
        '
        ' Get Extra Info from the WinInet.DLL
        '
        Try
            If ErrNum = ERROR_INTERNET_EXTENDED_ERROR Then
                '
                ' Get Message Size and Number
                '
                InternetGetLastResponseInfo(lError, vbNullString, lLen)
                sBuffer = vbNullChar '  String$(lLen + 1, vbNullChar)
                '
                ' Get Message
                '
                InternetGetLastResponseInfo(lError, sBuffer, lLen)
                GetINETErrorMsg = vbCrLf & sBuffer
            Else
                GetINETErrorMsg = Nothing
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
            GetINETErrorMsg = Nothing
            'Throw

        End Try

    End Function

End Class
