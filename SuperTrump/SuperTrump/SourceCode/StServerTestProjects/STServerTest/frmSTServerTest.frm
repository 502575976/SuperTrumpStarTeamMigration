VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Begin VB.Form frmSTServerTest 
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "General Test"
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   4575
      Begin VB.CommandButton btnVersion 
         Caption         =   "STServer Ver"
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "eFile Cr Submittal Test"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton btnGenReports 
         Caption         =   "Generate Text Reports"
         Height          =   525
         Left            =   2640
         TabIndex        =   6
         Top             =   1410
         Width           =   1155
      End
      Begin VB.TextBox txtLogPath 
         Height          =   315
         Left            =   1380
         TabIndex        =   5
         Top             =   360
         Width           =   2505
      End
      Begin VB.TextBox txtOutPath 
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   810
         Width           =   2505
      End
      Begin VB.CommandButton btnLogPath 
         Caption         =   "..."
         Height          =   315
         Left            =   3990
         TabIndex        =   3
         Top             =   420
         Width           =   345
      End
      Begin VB.CommandButton btnOutPath 
         Caption         =   "..."
         Height          =   315
         Left            =   3990
         TabIndex        =   2
         Top             =   840
         Width           =   345
      End
      Begin VB.CommandButton btnSummary 
         Caption         =   "Summary"
         Height          =   525
         Left            =   1440
         TabIndex        =   1
         Top             =   1410
         Width           =   1005
      End
      Begin MSComDlg.CommonDialog cdFiles 
         Left            =   120
         Top             =   1290
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Log dir Path:"
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   390
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Output dir Path:"
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   840
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmSTServerTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const cFacilityID               As Integer = 1000
Private Const cFacilityConfigPath       As String = "HKEY_LOCAL_MACHINE\Software\FacilitySettings\"
Private Const cReportTemplatePathKey    As String = "FilePath\ReportTemplatePath"

Private Sub btnGenReports_Click()
    Screen.MousePointer = 11
    GetData
    Screen.MousePointer = 0
End Sub

Private Sub btnLogPath_Click()
On Error GoTo ErrHandler
    
    Screen.MousePointer = 11
    
    'Open the dialog box
    cdFiles.ShowOpen
    
    'Get the file path
    If cdFiles.FileName <> "" Then
        txtLogPath.Text = Mid(cdFiles.FileName, 1, InStrRev(cdFiles.FileName, "\"))
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured - " & Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub btnOutPath_Click()
On Error GoTo ErrHandler
    
    Screen.MousePointer = 11
    
    'Open the dialog box
    cdFiles.ShowOpen
    
    'Get the file path
    If cdFiles.FileName <> "" Then
        txtOutPath.Text = Mid(cdFiles.FileName, 1, InStrRev(cdFiles.FileName, "\"))
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured - " & Err.Description
    Screen.MousePointer = 0
End Sub

'================================================================
'METHOD  : SavePRMBinaryFile
'PURPOSE : Saves the .PRM file using the given Path & File Name
'PARMS   :
'          astrFileName [String] = Name of the .PRM file to
'          be saved
'          astrFilePath [String] = Path where .PRM file to be
'          saved
'          aarrFileData [Byte] = Data to be saved
'RETURN  : String = Complete Path to the saved .PRM file
'================================================================
Public Function SavePRMBinaryFile(ByVal astrFileName As String, _
                                    ByVal astrFilePath As String, _
                                    ByRef aarrFileData() As Byte) As String
On Error GoTo ErrHandler

Dim liFileDes As Integer

    
    liFileDes = FreeFile()
    astrFileName = astrFilePath & "\" & UCase(astrFileName)
    Open astrFileName For Binary Access Write As liFileDes
    Put liFileDes, , aarrFileData
    Close liFileDes
    
    SavePRMBinaryFile = astrFileName
    Exit Function
    
ErrHandler:
    SavePRMBinaryFile = ""
    Err.Raise Err.Number, "SavePRMBinaryFile", Err.Description
End Function

'================================================================
'METHOD  : ReadRegistry
'PURPOSE : Retrives the spcified registry Key value.
'PARMS   :
'          astrRegKey [String] = Registry Key Name along with
'          it's complete path.
'RETURN  : String = Value of the Registry Key.
'================================================================
Public Function ReadRegistry(ByVal astrRegKey As String) As String
On Error GoTo ErrHandler

Dim lobjReg As New IWshRuntimeLibrary.IWshShell_Class   '.WshShell
            
    'Read the key value from the registry
    ReadRegistry = lobjReg.RegRead(astrRegKey)
    Set lobjReg = Nothing
    Exit Function
        
ErrHandler:
    Set lobjReg = Nothing
    Err.Raise Err.Number, "ReadRegistry", Err.Description
End Function

'================================================================
'METHOD  : SavePricingReport
'PURPOSE : To save the pricing report with the specified
'          file name and to the specified path.
'PARMS   :
'          astrFileName [String] = File name with the complete
'          path information.
'          astrData [String] = Pricing Report data.
'RETURN  : Boolean = True if saved successful, Otherwise false.
'================================================================
Public Function SavePricingReport(ByVal astrFileName As String, _
                                    ByVal astrData As String) As Boolean

Dim lobjFileSystem As New Scripting.FileSystemObject
Dim lobjFile As Scripting.File
Dim lobjTxtStream As Scripting.TextStream
Dim liIOMode As Integer
On Error GoTo ErrHandler:
    
    'Create and open the file
    Set lobjTxtStream = lobjFileSystem.CreateTextFile(astrFileName, True)
    
    'Write data to the file
    lobjTxtStream.WriteLine astrData
    
    'Close the file
    lobjTxtStream.Close
    
    Set lobjFile = Nothing
    Set lobjTxtStream = Nothing
    Set lobjFileSystem = Nothing
    SavePricingReport = True
    
    Exit Function

ErrHandler:
    SavePricingReport = False
    Set lobjFile = Nothing
    Set lobjTxtStream = Nothing
    Set lobjFileSystem = Nothing
    Err.Raise Err.Number, "SavePricingReport", Err.Description
End Function


Private Sub btnSummary_Click()
Dim lobjFileSysObj          As New FileSystemObject
Dim lobjLogFolder           As Folder
Dim lobjLogFile             As File
Dim lobjTextStream          As TextStream

Dim lstrSummaryData         As String
Dim lstrProposalId          As String
Dim lstreFileXML            As String
Dim lobjeFileXML            As New DOMDocument40
Dim lobjNodelst             As IXMLDOMNodeList
Dim liPricingFileCnt        As Integer
Dim liNonPricingFileCnt     As Integer
Dim liTotalFileCnt          As Integer
Dim lstrStartDateTime       As String
Dim lstrEndDateTime         As String
Dim liStart                 As Integer

    Screen.MousePointer = 11
    If ValidateInputs Then
        
        'Read the log directory contents
        Set lobjLogFolder = lobjFileSysObj.GetFolder(txtLogPath.Text)
        
        'For each file in the log directory
        For Each lobjLogFile In lobjLogFolder.Files
                
            lstrSummaryData = ""
            lstreFileXML = ""
            lstrStartDateTime = ""
            lstrEndDateTime = ""
            liPricingFileCnt = 0
            liNonPricingFileCnt = 0
            liTotalFileCnt = 0
            
            'Check if the file is a Efile XML
            If CheckFileName(lobjLogFile.Name, "^(eFileXML)(.)*(\.txt)$") Then
                lstrProposalId = Replace(Mid(lobjLogFile.Name, Len("eFileXML") + 1), ".txt", "")
                
                'Read the file contents
                Set lobjTextStream = lobjFileSysObj.OpenTextFile(lobjLogFile.Path, 1)
                lstreFileXML = lobjTextStream.ReadAll
                Set lobjTextStream = Nothing
                
                'Remove the date time stamp present before the start of the CrSubmittal XML in the file contents
                liStart = InStr(1, lstreFileXML, "<eFileSubmittal>")
                If liStart > 0 Then
                    lstrStartDateTime = Mid(lstreFileXML, 1, liStart - 1)
                    lstreFileXML = Mid(lstreFileXML, liStart)
                End If
                
                'Get # of documents in the eFileXML
                If lobjeFileXML.loadXML(lstreFileXML) Then
                    
                    'Get # of pricing documents
                    Set lobjNodelst = lobjeFileXML.selectNodes("//eFileSubmittal/eFileDocSet/eFileDoc[DocMetaData/eFileDocId='PRICING']")
                    
                    If Not (lobjNodelst Is Nothing) Then
                        liPricingFileCnt = lobjNodelst.length
                    End If
                    Set lobjNodelst = Nothing
                    
                    'Get # of non pricing documents
                    Set lobjNodelst = lobjeFileXML.selectNodes("//eFileSubmittal/eFileDocSet/eFileDoc[DocMetaData/eFileDocId!='PRICING']")
                    
                    If Not (lobjNodelst Is Nothing) Then
                        liNonPricingFileCnt = lobjNodelst.length
                    End If
                    
                    liTotalFileCnt = liPricingFileCnt + liNonPricingFileCnt
                End If
                
                lstrSummaryData = lstrProposalId & "," _
                                    & lobjLogFile.Size & "," _
                                    & CStr(liPricingFileCnt) & "," _
                                    & CStr(liNonPricingFileCnt) & "," _
                                    & CStr(liTotalFileCnt) & "," _
                                    & lstrStartDateTime
                                    
                WriteFile txtOutPath.Text & "SummaryData", lstrSummaryData
            End If
        Next
    End If
    Screen.MousePointer = 0
End Sub

Public Sub GetData()
On Error GoTo ErrHandler

Dim lobjFileSysObj          As New FileSystemObject
Dim lobjLogFolder           As Folder
Dim lobjLogFile             As File
Dim lobjTextStream          As TextStream
Dim lstrLogFileContents     As String
Dim lobjCrSubmittalXML      As New DOMDocument
Dim lobjPRMNodelst          As IXMLDOMNodeList
Dim liCnt                   As Integer
Dim lstrPRMFileName         As String
Dim lstrReportTemplateLoc   As String
Dim liTotalFiles            As Integer
Dim lstrSummary             As String
Dim liStart                 As Integer
    
    Debug.Print "Start Time: " & Format(Time, "hh:mm:ss")
    
    lstrSummary = ""
    
    'Check if directory paths are supplied
    If ValidateInputs Then
    
        liTotalFiles = 0
        
        'Read the report template path
        lstrReportTemplateLoc = ReadRegistry(cFacilityConfigPath & cFacilityID & "\" & cReportTemplatePathKey)
        
        'Read the log directory contents
        Set lobjLogFolder = lobjFileSysObj.GetFolder(txtLogPath.Text)
        
        'For each file in the log directory
        For Each lobjLogFile In lobjLogFolder.Files
            
            'Check if the file is a CrSubmittal XML
            If CheckFileName(lobjLogFile.Name, "^(CrSubmittalXML)(.)*(\.txt)$") Then
                                
                liTotalFiles = liTotalFiles + 1
                
                'Read the file contents
                Set lobjTextStream = lobjFileSysObj.OpenTextFile(lobjLogFile.Path, 1)
                lstrLogFileContents = lobjTextStream.ReadAll
                Set lobjTextStream = Nothing
                
                'Remove the date time stamp present before the start of the CrSubmittal XML in the file contents
                liStart = InStr(1, lstrLogFileContents, "<eFileSubmittal>")
                If liStart > 0 Then lstrLogFileContents = Mid(lstrLogFileContents, liStart)
                                
                'Load the Cr Submittal XML
                If lobjCrSubmittalXML.loadXML(lstrLogFileContents) Then
                
                    'Get a list of PRM files i.e. <eFileDoc> node having DocType = PRICING
                    Set lobjPRMNodelst = lobjCrSubmittalXML.getElementsByTagName("eFileSubmittal/eFileDocSet/eFileDoc[DocMetaData/eFileDocId='PRICING']")
                                                        
                    
                        
                    'For each PRM file
                    For liCnt = 0 To lobjPRMNodelst.length - 1

                        'Save PRM to disk at the output directory
                        lstrPRMFileName = lobjPRMNodelst.Item(liCnt).childNodes(0).selectSingleNode("DocName").Text
                        lobjPRMNodelst.Item(liCnt).selectSingleNode("Document").dataType = "bin.base64"
                        SavePRMBinaryFile lstrPRMFileName, txtOutPath.Text, lobjPRMNodelst.Item(liCnt).selectSingleNode("Document").nodeTypedValue
                            
                        'Generate text reports
                        GenerateReports txtOutPath.Text, lstrPRMFileName, lstrReportTemplateLoc

                        'Check if PRM file exists
                        If lobjFileSysObj.FileExists(txtOutPath.Text & lstrPRMFileName) Then

                            'Delete the PRM file
                            Call lobjFileSysObj.DeleteFile(txtOutPath.Text & lstrPRMFileName)
                        End If
                    
                    'Process next PRM file
                    Next
                    
                    Set lobjPRMNodelst = Nothing
                Else
                    MsgBox "Invalid CrSubmittal XML - " & lobjLogFile.Name & ". Error - " & lobjCrSubmittalXML.parseError.reason
                End If
                
                Set lobjCrSubmittalXML = Nothing
            
            'Skip file
            End If
        
        'Process next file
        Next
        
        Debug.Print "End Time: " & Format(Time, "hh:mm:ss")
        MsgBox "Total Cr Submittal files processed: " & liTotalFiles
        
        Set lobjLogFile = Nothing
        Set lobjLogFolder = Nothing
        Set lobjFileSysObj = Nothing
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "Error Occured - " & Err.Description
End Sub

Private Sub btnVersion_Click()
Dim lobjST As New STSERVER.STApplication

    MsgBox "Version : " & lobjST.Version & vbCrLf & "Build : " & lobjST.BuildInfo
    Debug.Print "Version : " & lobjST.Version & vbCrLf & "Build : " & lobjST.BuildInfo
End Sub

Public Sub GenerateReports(ByVal astrOutPath As String, _
                            ByVal astrPRMFileName As String, _
                            ByVal astrReportTemplateLoc As String)
                            
Dim lobjSTTrans             As New STTransaction
Dim liRepCnt                As Integer
Dim lstrReportTemplate      As String
Dim lstrReportName          As String
Dim lobjSTResults           As STResults
Dim lstrReportContents$

    'Open the PRM file
    If lobjSTTrans.OpenFile(astrOutPath & astrPRMFileName) Then
    
        'Get Mode value
        lobjSTTrans.Calculate
        If lobjSTTrans.Mode = ST_Mode_Lender Or lobjSTTrans.Mode = ST_Mode_Lessor Then
    
            'Loop 3 times for report generation
            For liRepCnt = 1 To 3
    
                'Determine which text report to generate
                If lobjSTTrans.Mode = ST_Mode_Lender And liRepCnt = 1 Then
                    lstrReportTemplate = "WIRE_H01.$CS"
                    lstrReportName = "Summary.txt"
                ElseIf lobjSTTrans.Mode = ST_Mode_Lender And liRepCnt = 2 Then
                    lstrReportTemplate = "AGGRE001.$CU"
                    lstrReportName = "Aggregation.txt"
                ElseIf lobjSTTrans.Mode = ST_Mode_Lender And liRepCnt = 3 Then
                    lstrReportTemplate = "WIRE_H16.$CU"
                    lstrReportName = "LoanAmort.txt"
                ElseIf lobjSTTrans.Mode = ST_Mode_Lessor And liRepCnt = 1 Then
                    lstrReportTemplate = "WIRE_H01.$BA"
                    lstrReportName = "Summary.txt"
                ElseIf lobjSTTrans.Mode = ST_Mode_Lessor And liRepCnt = 2 Then
                    lstrReportTemplate = "TV_Q_002.$XG"
                    lstrReportName = "Termination.txt"
                ElseIf lobjSTTrans.Mode = ST_Mode_Lessor And liRepCnt = 3 Then
                    lstrReportTemplate = "LEASE001.$BA"
                    lstrReportName = "LeaseAmort.txt"
                End If
                                        
                'Get the text report
                Set lobjSTResults = lobjSTTrans.Results
                lobjSTResults.ReportFileName = astrReportTemplateLoc & "\" & lstrReportTemplate
                lstrReportContents$ = lobjSTResults.PrintBuffer
                Set lobjSTResults = Nothing
                                        
                'Save the text reports to disk at the output directory
                SavePricingReport astrOutPath & astrPRMFileName & "_" & lstrReportName, lstrReportContents$
    
            'Process next report
            Next
        Else
            MsgBox "Invalid mode value for PRM file - " & astrPRMFileName
        End If
    Else
        MsgBox "Cannot open PRM file - " & astrPRMFileName
    End If
    
    Set lobjSTTrans = Nothing
End Sub

Public Function ValidateInputs() As Boolean
    
    'Check if directory paths are supplied
    If txtLogPath.Text = "" Then
        MsgBox "Enter the Log directory Path."
        txtLogPath.SetFocus
        ValidateInputs = False
    ElseIf txtOutPath.Text = "" Then
        MsgBox "Enter the output directory Path"
        txtOutPath.SetFocus
        ValidateInputs = False
    Else
        ValidateInputs = True
    End If
End Function

Public Function CheckFileName(ByVal astrFileName As String, _
                                ByVal astrFilePattern As String) As Boolean

Dim lobjValidRegExp As New RegExp
          
     
    lobjValidRegExp.Pattern = astrFilePattern
     
    CheckFileName = lobjValidRegExp.Test(astrFileName)
End Function

Private Sub WriteFile(Path$, result$)
'    On Error GoTo errTrap
    Dim FileHandle As Integer
    
    FileHandle = FreeFile
    Open Path$ For Append As #FileHandle
        
    Write #FileHandle, result$
    
    ' Close before reopening in another mode.
    Close #FileHandle

errTrap:

End Sub
