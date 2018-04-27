Imports System.Reflection
Imports Microsoft.Win32
Imports System.Xml
Imports BSVICROIEntity
Imports System.IO

Public Class BSVICROICommon
    Dim STLogger As log4net.ILog

    ''' <summary>
    ''' Used to create log file
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetLog4Net() As log4net.ILog
        Try
            Dim regKey As RegistryKey
            regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\FacilitySettings\SuperTRUMPForAllDeal", True)
            If log4net.LogManager.GetRepository.Configured = False Then
                log4net.Config.XmlConfigurator.ConfigureAndWatch(New System.IO.FileInfo(regKey.GetValue("SuperTRUMPForAllDealLogFile")))
            End If
            Dim STLogger As log4net.ILog = log4net.LogManager.GetLogger("SuperTRUMPForAllDeal")
            Return STLogger
        Catch ex As Exception
            Throw
            Return Nothing
        End Try
    End Function


    ''' <summary>
    ''' Used to get configuration file entry for corresponding tagname, pass as a parameter. *****
    ''' </summary>
    ''' <param name="strTagName">Tag Name</param>
    ''' <returns>string</returns>
    ''' <remarks></remarks>
    Public Function ReadConfigurationFileValue(ByVal strTagName As String) As String
        Dim strMileStone As String = "1"
        Dim objXmlDoc As New XmlDataDocument()
        Dim objXmlNode As XmlNodeList
        Dim strConfigValue As String = String.Empty
        STLogger = SetLog4Net()
        Try
            Dim strMainConfigFilePath As String = ReadRegistry("ConfigFilePath")
            If String.IsNullOrEmpty(strMainConfigFilePath) Then
                STLogger.Debug("Process Error:- Registry Variable ConfigFilePath Not Found in Registry ")
            End If
            If Not File.Exists(strMainConfigFilePath) Then
                STLogger.Debug("Process Error:- " + strMainConfigFilePath + " File not found")
            End If
            Dim objFStreme As New FileStream(strMainConfigFilePath, FileMode.Open, FileAccess.Read)
            strMileStone = "1.1"
            objXmlDoc.Load(objFStreme)
            strMileStone = "1.2"
            objXmlNode = objXmlDoc.GetElementsByTagName(strTagName)
            If objXmlNode.Count = 0 Then
                STLogger.Debug("Process Error:- Tag Name " + strTagName + " Not Found in Config XML ")
            End If
            strMileStone = "1.3"
            strConfigValue = objXmlNode.Item(0).InnerText
            strMileStone = "1.4"
            Return strConfigValue
        Catch ex As Exception
            STLogger.Error("MileStone:- " & strMileStone & " Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
            Return Nothing
        Finally
            objXmlDoc = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Used to read registry *****
    ''' </summary>
    ''' <param name="astrRegFullKeyName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ReadRegistry(ByVal astrRegFullKeyName As String) As String

        Dim regKey As RegistryKey
        Try
            regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\FacilitySettings\SuperTRUMPForAllDeal", True)
            Return regKey.GetValue(astrRegFullKeyName)
        Catch ex As Exception
            STLogger.Error(" Error No:- " & Err.Number & " Method Name:- " & System.Reflection.MethodInfo.GetCurrentMethod.Name() & " Error Desc:- " + Err.Description)
            Err.Raise(Err.Number, MethodInfo.GetCurrentMethod.Name() & ":ReadRegistry()", Err.Description)
            Return vbNullString
        End Try
    End Function
End Class
