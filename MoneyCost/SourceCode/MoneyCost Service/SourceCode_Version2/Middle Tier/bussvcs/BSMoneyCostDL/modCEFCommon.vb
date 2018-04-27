Imports Microsoft.Win32
Imports System.text
Imports System.Xml
Public Module modCEFCommon
    Public Enum eDBType
        cSQLServer
        cOledb
        cOracleServer
    End Enum
    '=== Registry Constants =========================================
    Public Const cCONN_STRINGS_REG_PATH As String = "\ConnectStrings\"
    Public Const cFACILITY_CONFIG_REG_PATH As String = "HKEY_LOCAL_MACHINE\SOFTWARE\FacilitySettings\"
    Public Const cFACILITY_ID As String = "MoneyCost"
    '================================================================
    '================================================================
    ' GE Capital Proprietary and Confidential
    ' Copyright (c) 2001-2002 by GE Capital - All rights reserved.
    '
    ' This code may not be reproduced in any way without express
    ' permission from GE Capital.
    '================================================================
    Function delGetConfigurationKey(ByVal astrKey As String) As String

        Dim ConstantsFilePath As String
        Dim lDocXmlFile As New XmlDocument
        Dim lobjreg As RegistryKey = Registry.LocalMachine
        Dim val As Object

        Try
            lobjreg = lobjreg.OpenSubKey("SOFTWARE\FacilitySettings\MoneyCost\FilePath", False)
            val = lobjreg.GetValue("ConfigFilePath")
            ConstantsFilePath = val
            ConstantsFilePath = ConstantsFilePath + "\Constant.XML"
            'ConstantsFilePath = "D:\StarTeam\MoneyCost\QA-ConstantFile" + "\Constant.XML"
            lDocXmlFile.Load(ConstantsFilePath)

            Return lDocXmlFile.GetElementsByTagName(astrKey).Item(0).InnerText

        Catch ex As Exception
            Throw ex
        Finally
            lDocXmlFile = Nothing
        End Try
    End Function
End Module
