
Public Module modBSMoneyCost

    '=== Constant for module name ===================================
    Private Const cMODULE_NAME As String = "modBSMoneyCost"

    Public Const cCOMPONENT_NAME As String = "BSMoneyCost"
    '================================================================

    '=== Registry Constants =========================================
    Public Const mcFACILITY_CONFIG_REG_PATH As String = "HKEY_LOCAL_MACHINE\SOFTWARE\FacilitySettings\"
    Public Const mcFACILITY_ID As String = "MoneyCost"
    Public Const mcCONFIG_FILE_SETTING_ID As String = "ConfigFilePath"
    Public Const mcCONFIG_FILE_ID As String = "Constant.xml"
    Public Const mcCONN_STRINGS_REG_PATH As String = "\ConnectStrings\"
    Public Const mcCONN_STRING_KEY As String = "MoneyCost"
    Public Const mcHIERARCH_CONN_STRING_KEY As String = "MoneyCostHierarch"
    '================================================================
    Public Const mcDEBUG_LEVEL_SIZE_KEY As String = "DebugLevel"
    Public Const mcDEBUG_LOG_FILE_PATH_NAME_KEY As String = "DebugFile"
    Public Const mcDEBUG_ERROR_FILE_PATH_NAME_KEY As String = "ErrorLogFile"
    Public Const mcDEBUG_MAX_FILE_SIZE_KEY As String = "MaxDebugFileSize"

    '=== Error Constants ============================================
    Public Const cHIGHEST_ERROR As Object = vbObjectError + 256
    Public Const cINVALID_PARMS As Object = cHIGHEST_ERROR + 1000
    Public Const cINVALID_SQL_ID As Object = cHIGHEST_ERROR + 1010
    '================================================================

End Module
