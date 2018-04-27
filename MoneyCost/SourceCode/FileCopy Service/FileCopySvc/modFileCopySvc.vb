Imports System.IO

Module modFileCopySvc

    Private Const cFULL_MODULE_NAME As String = "modFileCopySvc"
    Public Const RESOURCETYPE_DISK As Long = &H1

    Public gobjLogger As log4net.ILog = log4net.LogManager.GetLogger("FileCopySvcLogger")


    Public Function GetTimeStamp() As String
        Dim lstrTimeStamp As String

        Try
            gobjLogger.Debug("Enter GetTimeStamp method")

            lstrTimeStamp = Replace(Now, "/", "")
            lstrTimeStamp = Replace(lstrTimeStamp, ":", "")
            lstrTimeStamp = Replace(lstrTimeStamp, " ", "_")
            lstrTimeStamp = Replace(lstrTimeStamp, "\", "")
            lstrTimeStamp = Replace(lstrTimeStamp, ",", "")
            lstrTimeStamp = Replace(lstrTimeStamp, ";", "")

            Return lstrTimeStamp
        Catch lobjSysEx As System.Exception
            gobjLogger.Error(lobjSysEx.Message)
            Throw lobjSysEx
        Finally
            gobjLogger.Debug("Exit GetTimeStamp method")
        End Try
    End Function

    Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" _
   (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, _
   ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer

    Public Declare Function WNetCancelConnection2 Lib "mpr" Alias "WNetCancelConnection2A" _
  (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer
    Public Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String
    End Structure
End Module
