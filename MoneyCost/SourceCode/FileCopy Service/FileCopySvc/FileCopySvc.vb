Imports System.ServiceProcess
Imports System.IO

Public Class FileCopySvc
    Inherits System.ServiceProcess.ServiceBase

    Private Const cFULL_MODULE_NAME As String = "FileCopySvc"
    Private Enum eFileMgrAction
        MOVE
        COPY
    End Enum
    Dim mobjTimer As Timers.Timer
    Dim mobjXMLDoc As Xml.XmlDocument


#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        ' This call is required by the Component Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call

    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New FileCopySvc}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
        Me.ServiceName = "FileCopySvc"
    End Sub

#End Region

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Try
            gobjLogger.Debug("enter OnStart method for svc")

            mobjTimer = New Timers.Timer
            AddHandler mobjTimer.Elapsed, AddressOf Controller
            mobjTimer.Interval = IConfig.ServiceSleepInterval
            mobjTimer.Enabled = True

            gobjLogger.Info("FileCopySvc started")

        Catch ex As Exception
            gobjLogger.Error(ex.Message)
        Finally
            gobjLogger.Debug("exit OnStart method for svc")
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Try
            gobjLogger.Debug("enter OnStop method for svc")

            mobjTimer.Enabled = False
            gobjLogger.Info("FileCopySvc stopped")

        Catch ex As Exception
            gobjLogger.Error(ex.Message)
        Finally
            If Not (mobjTimer Is Nothing) Then
                mobjTimer.Dispose()
            End If
            gobjLogger.Debug("exit OnStop method for svc")
            If Not IsNothing(gobjLogger) Then
                gobjLogger = Nothing
            End If
        End Try
    End Sub
    Private Sub Controller(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)

        Dim lstrSrcFile As String
        Dim lstrDestFile As String
        Dim leFileMgrAction As eFileMgrAction
        Dim lstrErrMsg As String

        Try
            gobjLogger.Debug("Enter Controller method")

            'set the interval on the timer
            If mobjTimer.Interval <> IConfig.ServiceSleepInterval Then
                mobjTimer.Interval = IConfig.ServiceSleepInterval
            End If

            gobjLogger.Info("Start copying files")

            mobjXMLDoc = New Xml.XmlDocument
            mobjXMLDoc.Load(IConfig.ConfigFile)

            ManageDirectories()

            'code to access SAMBA shares 
            If SourceFolderExists() Then

                For Each lobjXMLNode As Xml.XmlNode In mobjXMLDoc.SelectNodes("//FILECOPYSVC/CONFIG/FILES/FILE[@ACTIVE=1]")

                    Select Case lobjXMLNode.Attributes.GetNamedItem("ACTION").InnerText
                        Case "COPY"
                            leFileMgrAction = eFileMgrAction.COPY
                        Case "MOVE"
                            leFileMgrAction = eFileMgrAction.MOVE
                        Case Else
                            leFileMgrAction = eFileMgrAction.COPY
                    End Select

                    lstrSrcFile = lobjXMLNode.SelectSingleNode("SRC").InnerText
                    lstrDestFile = lobjXMLNode.SelectSingleNode("DEST").InnerText
                    ManageFile(lstrSrcFile, lstrDestFile, leFileMgrAction)

                Next
                gobjLogger.Info("Finished copying files")
            Else
                lstrErrMsg = "Shamba share folder: " & IConfig.InputFileLocation & " not found or access denied."
                gobjLogger.Info(lstrErrMsg)

            End If
        Catch ex As Exception
            gobjLogger.Error(ex.Message)
        Finally
            If Not IsNothing(mobjXMLDoc) Then
                mobjXMLDoc = Nothing
            End If

            gobjLogger.Debug("Exit Controller method")
        End Try
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[213043914]	10/5/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub ManageDirectories()
        Try
            gobjLogger.Debug("Entering ManageDirectories method")

            UNCDirectories()
            FTPDirectories()

        Catch ex As Exception
            gobjLogger.Error(ex.Message)
        Finally
            gobjLogger.Debug("Exiting ManageDirectories method")
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[213043914]	10/5/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub UNCDirectories()
        Dim lobjDirInfo As DirectoryInfo
        Dim lstrExt As String
        Dim lstrDestDir As String
        Dim leFileMgrAction As eFileMgrAction

        Try
            gobjLogger.Debug("Entering UNCDirectories method")

            For Each lobjXMLNode As Xml.XmlNode In mobjXMLDoc.SelectNodes("//FILECOPYSVC/CONFIG/DIRS/DIR[@ACTIVE=1]")

                lstrExt = Replace(lobjXMLNode.SelectSingleNode("FILE_EXT").InnerText, " ", "")
                lstrExt = Replace(lstrExt, "*", "")
                lstrExt = Replace(lstrExt, ".", "")

                lstrDestDir = Trim(lobjXMLNode.SelectSingleNode("DEST").InnerText)
                If Right(lstrDestDir, 1) <> "\" Then
                    lstrDestDir = lstrDestDir & "\"
                End If

                Select Case lobjXMLNode.Attributes.GetNamedItem("ACTION").InnerText
                    Case "COPY"
                        leFileMgrAction = eFileMgrAction.COPY
                    Case "MOVE"
                        leFileMgrAction = eFileMgrAction.MOVE
                    Case Else
                        leFileMgrAction = eFileMgrAction.COPY
                End Select

                lobjDirInfo = New DirectoryInfo(lobjXMLNode.SelectSingleNode("SRC").InnerText)

                gobjlogger.Info("Manage directory: Source directory - " & lobjXMLNode.SelectSingleNode("SRC").InnerText & ". Destination directory - " & lstrdestdir & ". Action - ") 

                If lobjDirInfo.Exists Then
                    For Each lobjFile As FileInfo In lobjDirInfo.GetFiles("*." & lstrExt)
                        ManageFile(lobjFile.FullName, lstrDestDir & lobjFile.Name, leFileMgrAction)
                    Next
                End If
            Next

        Catch ex As Exception
            gobjLogger.Error(ex.Message)
        Finally
            gobjLogger.Debug("Exiting UNCDirectories method")
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[213043914]	10/5/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub FTPDirectories()
        Dim lobjFTP As cFTPClient
        Dim lintPort As String
        Dim lstrHost As String
        Dim lstrUser As String
        Dim lstrPassword As String
        Dim lstrSrcPath As String
        Dim lstrDestPath As String
        Dim lstrFiles() As String

        Try
            gobjLogger.Debug("Enter FTPDirectories method")

            For Each lobjXMLNode As Xml.XmlNode In mobjXMLDoc.SelectNodes("//FILECOPYSVC/CONFIG/DIRS/FTP_DIR[@ACTIVE=1]")

                lstrHost = lobjXMLNode.SelectSingleNode("@HOST").InnerText
                lintPort = CInt(lobjXMLNode.SelectSingleNode("@PORT").InnerText)
                lstrUser = lobjXMLNode.SelectSingleNode("@USER").InnerText
                lstrPassword = lobjXMLNode.SelectSingleNode("@PASSWORD").InnerText
                lstrSrcPath = lobjXMLNode.SelectSingleNode("SRC").InnerText
                lstrDestPath = lobjXMLNode.SelectSingleNode("DEST").InnerText

                If Right(lstrDestPath, 1) <> "\" Then
                    lstrDestPath = lstrDestPath & "\"
                End If

                lobjFTP = New cFTPClient(lstrHost, lstrSrcPath, lstrUser, lstrPassword, lintPort)
                lobjFTP.Login()
                lobjFTP.SetBinaryMode(True)

                'get a list of all files in the directory
                lstrFiles = lobjFTP.GetFileList("*.*")
                For i As Integer = 0 To lstrFiles.Length - 1
                    If Trim(Replace(lstrFiles(i), Chr(13), "")) <> "" Then
                        lobjFTP.DownloadFile(lstrFiles(i), lstrDestPath & GetTimeStamp() & Microsoft.VisualBasic.Right(Trim(Replace(lstrFiles(i), Chr(13), "")), 4))
                        gobjLogger.Info(lstrFiles(i) & " copied to " & lstrDestPath)
                        lobjFTP.DeleteFile(lstrFiles(i))
                        gobjLogger.Info(lstrFiles(i) & " deleted.")
                    End If
                Next

                lobjFTP.CloseConnection()
            Next

        Catch ex As Exception
            ex.Source = ex.Source & "." & cFULL_MODULE_NAME & ".ManageDirectories."
            gobjLogger.Error(ex.Message)
        Finally
            If Not IsNothing(lobjFTP) Then
                lobjFTP = Nothing
            End If

            gobjLogger.Debug("Exit FTPDirectories method")
        End Try
    End Sub
    Private Sub ManageFile(ByVal astrSrcFile As String, ByVal astrDestFile As String, ByVal aFileMgrAction As eFileMgrAction)

        Dim lobjFileInfo As System.IO.FileInfo

        Dim ldtSrcFileLastWrite As DateTime
        Dim ldtDestFileLastWrite As DateTime
        Try
            gobjLogger.Debug("Enter FileMgr method")

            ldtSrcFileLastWrite = GetLastFileWriteDateTime(astrSrcFile)
            ldtDestFileLastWrite = GetLastFileWriteDateTime(astrDestFile)

            'if last write times of the 2 files differs by a minute or more then copy or move
            If DateDiff(DateInterval.Minute, ldtSrcFileLastWrite, ldtDestFileLastWrite) <> 0 Then

                lobjFileInfo = New FileInfo(astrSrcFile)

                If lobjFileInfo.Exists Then
                    'move or copy the file depending on the action required
                    Select Case aFileMgrAction
                        Case eFileMgrAction.COPY
                            lobjFileInfo.CopyTo(astrDestFile, True)
                            gobjLogger.Info("File " & astrSrcFile & " copied to " & astrDestFile)
                        Case eFileMgrAction.MOVE
                            lobjFileInfo.MoveTo(astrDestFile)
                            gobjLogger.Info("File " & astrSrcFile & " moved to " & astrDestFile)
                    End Select
                Else
                    Throw New System.IO.IOException("File does not exist: " & astrSrcFile)
                End If

            Else
                gobjLogger.Info("Files " & astrSrcFile & " and " & astrDestFile & " are the same age. No action will be taken.")
            End If

        Catch lobjIOEx As System.IO.IOException
            gobjLogger.Error("The following error occured trying to manage files " & astrSrcFile & " and " & astrDestFile & ": " & lobjIOEx.Message)
        Catch ex As Exception
            gobjLogger.Error(ex.Message)
        Finally
            If Not IsNothing(lobjFileInfo) Then
                lobjFileInfo = Nothing
            End If
            gobjLogger.Debug("Exit FileMgr method")
        End Try
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="astrFilePath"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[213043914]	10/2/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function GetLastFileWriteDateTime(ByVal astrFilePath As String) As DateTime

        Dim lobjFileInfo As FileInfo

        Try
            gobjLogger.Debug("Enter GetLastFileWriteDateTime method")

            lobjFileInfo = New FileInfo(astrFilePath)
            If lobjFileInfo.Exists Then
                Return lobjFileInfo.LastWriteTime
            Else
                Return CDate("1/1/1900")
            End If
        Catch ex As Exception
            gobjLogger.Error(ex.Message)
            Throw ex
        Finally
            If Not IsNothing(lobjFileInfo) Then
                lobjFileInfo = Nothing
            End If
            gobjLogger.Debug("Exit GetLastFileWriteDateTime method")
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     checks for presence of input xml files
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Function SourceFolderExists() As Boolean
        Dim lobjDirInfo As DirectoryInfo
        Dim lstrUNCPath As String

        Try

            lstrUNCPath = IConfig.InputFileLocation
            'Only create the log directories if there are prm files to be processed

            lobjDirInfo = New DirectoryInfo(lstrUNCPath)
            'Check If Directory exists
            If lobjDirInfo.Exists Then
                Return True
            Else
                'If not exists then map the UNC Path to the server
                If MapDrive(lstrUNCPath) Then
                    Return True
                Else
                    Return False
                End If
            End If

        Catch ex As Exception

            Throw ex
        Finally
            If Not IsNothing(lobjDirInfo) Then
                lobjDirInfo = Nothing
            End If

        End Try
    End Function

    Public Function MapDrive(ByVal UNCPath As String) As Boolean
        Dim nr As NETRESOURCE
        Dim strUsername As String
        Dim strPassword As String

        nr = New NETRESOURCE
        nr.lpRemoteName = UNCPath
        nr.lpLocalName = IConfig.MappedDrive
        strUsername = IConfig.SambhaUID
        strPassword = IConfig.SambhaPWD
        nr.dwType = RESOURCETYPE_DISK

        Dim result As Integer
        result = WNetAddConnection2(nr, strPassword, strUsername, 0)

        If result = 0 Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
