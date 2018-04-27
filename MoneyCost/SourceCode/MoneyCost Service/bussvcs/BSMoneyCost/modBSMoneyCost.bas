Attribute VB_Name = "modBSMoneyCost"
'================================================================
' GE Capital Proprietary and Confidential
' Copyright (c) 2001-2002 by GE Capital - All rights reserved.
'
' This code may not be reproduced in any way without express
' permission from GE Capital.
'================================================================

'================================================================
'MODULE  : modBSMoneyCost
'PURPOSE : This module will contain all common functions, public
'          variables & constants specific to this component.
'================================================================

Option Explicit

'=== Constant for module name ===================================
Private Const cMODULE_NAME      As String = "modBSMoneyCost"

Public Const cCOMPONENT_NAME    As String = "BSMoneyCost"
'================================================================

'=== Registry Constants =========================================
Public Const cFACILITY_CONFIG_REG_PATH      As String = "HKEY_LOCAL_MACHINE\SOFTWARE\FacilitySettings\"
Public Const cFACILITY_ID                   As String = "MoneyCost"
Public Const cCONN_STRINGS_REG_PATH         As String = "\ConnectStrings\"
Public Const cCONN_STRING_KEY               As String = "MoneyCost"
Public Const cHIERARCH_CONN_STRING_KEY      As String = "MoneyCostHierarch"
'================================================================

Public Const cDEBUG_REG_PATH                    As String = "\Debug"
Public Const cDEBUG_LEVEL_COMPONENT_REG_PATH    As String = ""
Public Const cDEBUG_LEVEL_SIZE_KEY              As String = "DebugLevel"
Public Const cDEBUG_LOG_FILE_PATH_NAME_KEY      As String = "DebugFile"
Public Const cDEBUG_ERROR_FILE_PATH_NAME_KEY    As String = "ErrorLogFile"
Public Const cDEBUG_MAX_FILE_SIZE_KEY           As String = "MaxDebugFileSize"

'=== Error Constants ============================================
Public Const cHIGHEST_ERROR     As Variant = vbObjectError + 256
Public Const cINVALID_PARMS     As Variant = cHIGHEST_ERROR + 1000
Public Const cINVALID_SQL_ID    As Variant = cHIGHEST_ERROR + 1010
'================================================================
