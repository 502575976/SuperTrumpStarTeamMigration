VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   1980
   ClientTop       =   1380
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   11565
   Begin VB.Frame Frame8 
      Caption         =   "Using PRM binary XML file"
      Height          =   1035
      Left            =   0
      TabIndex        =   21
      Top             =   7710
      Width           =   11565
      Begin VB.CommandButton btnAmortRep4PRMBINXML 
         Caption         =   "Get Amortization Report From PRM binary XML file"
         Height          =   735
         Left            =   5880
         TabIndex        =   23
         Top             =   240
         Width           =   5535
      End
      Begin VB.CommandButton btnPRMXML2DataXML 
         Caption         =   "Convert PRM binary XML 2 PRM XML file"
         Height          =   735
         Left            =   330
         TabIndex        =   22
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "PRM binary file wrapped in XML"
      Height          =   1035
      Left            =   0
      TabIndex        =   19
      Top             =   6660
      Width           =   11565
      Begin VB.CommandButton btnPRMinXML 
         Caption         =   "Get PRM BIN XMLfile"
         Height          =   735
         Left            =   330
         TabIndex        =   20
         Top             =   210
         Width           =   11055
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11040
         Top             =   330
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Manipulate PRM binary file"
      Height          =   1095
      Left            =   0
      TabIndex        =   17
      Top             =   5550
      Width           =   11565
      Begin VB.CommandButton btnModifyPRMFileXML 
         Caption         =   "Modify PRM binary File"
         Height          =   735
         Left            =   330
         TabIndex        =   18
         Top             =   240
         Width           =   11055
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data contained in PRM binary file as XML"
      Height          =   1155
      Left            =   0
      TabIndex        =   11
      Top             =   4380
      Width           =   11565
      Begin VB.CommandButton btnCashFlowLstXML 
         Caption         =   "Get Cash Flow List"
         Height          =   735
         Left            =   8760
         TabIndex        =   16
         Top             =   240
         Width           =   2595
      End
      Begin VB.CommandButton btnTransactionAmtXML 
         Caption         =   "Get Transaction Amount"
         Height          =   735
         Left            =   5865
         TabIndex        =   15
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton btnAppDataXML 
         Caption         =   "Get Super Trump Application Data"
         Height          =   735
         Left            =   330
         TabIndex        =   14
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton btnDTDXML 
         Caption         =   "Get DTD of PRM XML"
         Height          =   735
         Left            =   3990
         TabIndex        =   13
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton btnSchemaXML 
         Caption         =   "Get Schema of PRM XML"
         Height          =   735
         Left            =   2175
         TabIndex        =   12
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Text Reports from text reports XML wrapper file"
      Height          =   1065
      Left            =   0
      TabIndex        =   9
      Top             =   3300
      Width           =   11565
      Begin VB.CommandButton btnAmortRepTxtFromXML 
         Caption         =   "Get Text Report document"
         Height          =   735
         Left            =   330
         TabIndex        =   10
         Top             =   240
         Width           =   10995
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Data wrapped in XML From PRM binary file"
      Height          =   1125
      Left            =   0
      TabIndex        =   7
      Top             =   2160
      Width           =   11565
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   615
         Left            =   3510
         TabIndex        =   25
         Top             =   330
         Width           =   2295
      End
      Begin VB.CommandButton btnAmortRepDataXML 
         Caption         =   "Get Amortization Report Data"
         Height          =   735
         Left            =   330
         TabIndex        =   8
         Top             =   270
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Text Reports wrapped in XML From PRM binary file"
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   1050
      Width           =   11565
      Begin VB.CommandButton btnRentSchedule 
         Caption         =   "Get Rent Schedule Report"
         Height          =   735
         Left            =   9180
         TabIndex        =   24
         Top             =   240
         Width           =   2085
      End
      Begin VB.CommandButton btnAmortRepXML 
         Caption         =   "Get Amortization Report"
         Height          =   735
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton btnSummaryReportXML 
         Caption         =   "Get Summary Report"
         Height          =   735
         Left            =   2370
         TabIndex        =   5
         Top             =   240
         Width           =   2025
      End
      Begin VB.CommandButton btnTerminationFullRepXML 
         Caption         =   "Get Termination Report"
         Height          =   735
         Left            =   4470
         TabIndex        =   4
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton btnAggrLendingRepXML 
         Caption         =   "Get Aggregate Lending Report"
         Height          =   735
         Left            =   6570
         TabIndex        =   3
         Top             =   240
         Width           =   2505
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conversion"
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11565
      Begin VB.CommandButton btnConvertPRM2XML 
         Caption         =   "Convert PRM binary file 2 PRM XML file "
         Height          =   735
         Left            =   360
         TabIndex        =   1
         Top             =   210
         Width           =   10875
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const cREPORT_TEMPLATE_PATH As String = "C:\Mac\Official\eFile\BSCEFSuperTrump\ReportTemplates"

Private Sub btnAggrLendingRepXML_Click()
Dim lobjReportSTTrans As New STTransaction
Dim lobjReportSTResults As STResults
Dim lstrReport As String
Dim lobjFileSystem As New Scripting.FileSystemObject
Dim lvFile As Variant
Dim lstrPRMFilePath As String
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
        lobjReportSTTrans.OpenFile (lstrPRMFilePath)
        lobjReportSTTrans.Calculate
        Set lobjReportSTResults = lobjReportSTTrans.Results
        lobjReportSTResults.ReportFileName = cREPORT_TEMPLATE_PATH & "\AGGRE001.$CU"
        lstrReport = lobjReportSTResults.PrintBuffer
        
        Set lvFile = lobjFileSystem.CreateTextFile(Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2AggrLendingRep.xml"), True)
        lvFile.WriteLine ("<AggrLendingRep>" & "<![CDATA[" & lstrReport & "]]></AggrLendingRep>")
        lvFile.Close
        Set lobjFileSystem = Nothing
        MsgBox "Aggregate Loan Lending Report successfully generated from PRM file and XML saved at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2AggrLendingRep.xml")
    Else
        MsgBox "Invalid PRM File Selected"
    End If

CleanUp:
    Set lobjReportSTResults = Nothing
    Set lobjReportSTTrans = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnAmortRep4PRMBINXML_Click()
Dim lobjReportSTTrans As New STTransaction
Dim lobjReportSTResults As STResults
Dim lstrReport As String
Dim lobjFileSystem As New Scripting.FileSystemObject
Dim lvFile As Variant
Dim lstrPRMFilePath As String
Dim lstrXML As String
Dim lstrXMLFilePath As String
Dim lobjFile As New ADODB.Stream
Dim lvPRMXML As Variant
Dim lobjPRMXML As New DOMDocument
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "XML (*.xml)|*.xml"
    CommonDialog1.ShowOpen
    lstrXMLFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrXMLFilePath), ".XML") > 0 Then
        If lobjPRMXML.Load(lstrXMLFilePath) Then
            lvPRMXML = lobjPRMXML.documentElement.nodeTypedValue
            
            lobjFile.Type = adTypeBinary
            lobjFile.Open
            lobjFile.SetEOS
            lobjFile.Write lvPRMXML
            lobjFile.SaveToFile Replace(UCase(lstrXMLFilePath), ".XML", ".PRM"), adSaveCreateOverWrite
            lobjFile.Close
            
            lobjReportSTTrans.OpenFile (Replace(UCase(lstrXMLFilePath), ".XML", ".PRM"))
            Set lobjReportSTResults = lobjReportSTTrans.Results
            lobjReportSTResults.ReportFileName = cREPORT_TEMPLATE_PATH & "\Wire_h16.$cu"
            lstrReport = lobjReportSTResults.PrintBuffer
            
            
            Set lvFile = lobjFileSystem.CreateTextFile(Replace(UCase(lstrPRMFilePath), ".XML", "_BINXML2AmortRep.xml"), True, True)
            lvFile.WriteLine ("<AmortRep>" & "<![CDATA[" & lstrReport & "]]></AmortRep>")
            lvFile.Close
            Set lobjFileSystem = Nothing
            MsgBox "Amortization Report successfully generated from PRM file and XML saved at : " & Replace(UCase(lstrPRMFilePath), ".XML", "_BINXML2AmortRep.xml")
        Else
            MsgBox "Invalid XML File Selected"
        End If
        
    Else
        MsgBox "Invalid XML File Selected"
    End If
    
CleanUp:
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnAppDataXML_Click()
Dim lobjSTApp As New STApplication
Dim lstrXML As String
On Error GoTo ERR_HANDLER:
    
    lstrXML = "<SuperTRUMP WriteXMLFile = """ & App.Path & "\STAppData.xml"" >" & _
                "<AppData query=""true"" />" & _
              "</SuperTRUMP>"
    lstrXML = lobjSTApp.XmlInOut(lstrXML)
    MsgBox "App Data converted to XML at " & App.Path & "STAppData.xml"
        
CleanUp:
    Set lobjSTApp = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnCashFlowLstXML_Click()
Dim lobjSTApp As New STApplication
Dim lstrXML As String
Dim lstrPRMFilePath As String
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
        lstrXML = "<SuperTRUMP WriteXMLFile = """ & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2CashFlowLst.xml") & """ >" & _
                    "<Transaction id=""TRAN3"">" & _
                        "<ReadFile filename=""" & lstrPRMFilePath & """ />" & _
                        "<LendingLoans>" & _
                            "<Loan index=""*"">" & _
                                "<CashflowSteps>" & _
                                    "<CashflowStep index=""*"">" & _
                                        "<DaysInPeriod query=""true"" />" & _
                                        "<Type query=""true"" />" & _
                                        "<Rate query=""true"" />" & _
                                        "<Amount query=""true"" />" & _
                                    "</CashflowStep>" & _
                                "</CashflowSteps>" & _
                            "</Loan>" & _
                        "</LendingLoans>" & _
                    "</Transaction>" & _
                "</SuperTRUMP>"
        lstrXML = lobjSTApp.XmlInOut(lstrXML)
        MsgBox "Cash Flow list successfully extracted from PRM file and XML saved at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2CashFlowLst.xml")
    Else
        MsgBox "Invalid PRM File Selected"
    End If
    
CleanUp:
    Set lobjSTApp = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnDTDXML_Click()
Dim lobjSTApp As New STApplication
Dim lstrXML As String
On Error GoTo ERR_HANDLER:

    lstrXML = "<SuperTRUMP queryDTD=""SuperTrumpDTD"" WriteXMLFile = """ & App.Path & "\STDTD.xml"" />"
    lstrXML = lobjSTApp.XmlInOut(lstrXML)
    MsgBox "Super Trump DTD Present at : " & App.Path & "\STDTD.xml"
    
CleanUp:
    Set lobjSTApp = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnModifyPRMFileXML_Click()
Dim lobjSTApp As New STApplication
Dim lstrXML As String
Dim lstrPRMFilePath As String
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
'''        lstrXML = "<SuperTRUMP>" & _
'''                    "<Transaction id=""TRAN4"" query=""true"">" & _
'''                        "<Mode>Lender</Mode>" & _
'''                        "<Initialize/>" & _
'''                        "<ReadFile filename=""" & lstrPRMFilePath & """ />" & _
'''                        "<TransactionStartDate>2003-11-4</TransactionStartDate>" & _
'''                        "<WriteFile filename=""" & Replace(UCase(lstrPRMFilePath), ".PRM", "_ModifiedPRM.prm") & """ />" & _
'''                    "</Transaction>" & _
'''                "</SuperTRUMP>"
        lstrXML = "<SuperTRUMP>" & _
                    "<Transaction query=""abc"">" & _
                        "<Mode>Lender</Mode>" & _
                        "<Initialize/>" & _
                        "<ReadFile filename=""" & lstrPRMFilePath & """ />" & _
                        "<TARGETDATA>" & _
                            "<TARGETVALUE>0.075</TARGETVALUE>" & _
                        "</TARGETDATA>" & _
                        "<WriteFile filename=""" & Replace(UCase(lstrPRMFilePath), ".PRM", "_ModifiedPRM.prm") & """ />" & _
                    "</Transaction>" & _
                "</SuperTRUMP>"

        lstrXML = lobjSTApp.XmlInOut(lstrXML)
        MsgBox "Modified PRM file at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_ModifiedPRM.prm")
    Else
        MsgBox "Invalid PRM File Selected"
    End If
    
CleanUp:
    Set lobjSTApp = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnConvertPRM2XML_Click()
Dim lobjSTApp As New STApplication
Dim lstrXML As String
Dim lstrPRMFilePath As String
On Error GoTo ERR_HANDLER:

    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
        lstrXML = "<SuperTRUMP WriteXMLFile = """ & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2XML.xml") & """ >" & _
                    "<Transaction id=""TRAN1"" query=""true"">" & _
                        "<ReadFile filename=""" & lstrPRMFilePath & """ />" & _
                    "</Transaction>" & _
                "</SuperTRUMP>"
        lstrXML = lobjSTApp.XmlInOut(lstrXML)
        MsgBox "PRM file successfully converted to XML and saved at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2XML.xml")
    Else
        MsgBox "Invalid PRM File Selected"
    End If
    
CleanUp:
    Set lobjSTApp = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnAmortRepDataXML_Click()
Dim lobjReportSTTrans As New STTransaction
Dim lobjPaymentSTStream As STStream
Dim lobjInterestSTStream As STStream
Dim lobjPrincipalSTStream As STStream

Dim liIndex  As Integer
Dim liTotalPeriods  As Integer
Dim ldteAmortizationDate  As Date
Dim lstrXML As String
Dim lobjFileSystem As New Scripting.FileSystemObject
Dim lvFile As Variant
Dim lstrPRMFilePath As String
Dim lobjSTQuick As STQuick
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
    
        lobjReportSTTrans.OpenFile (lstrPRMFilePath)
        
        'Get the Amortization Data
        lobjReportSTTrans.GetFreeStream
        If lobjReportSTTrans.Mode = ST_Mode_Lender Then
            Set lobjPaymentSTStream = lobjReportSTTrans.GetEconStream(ST_EDA_LoanLendDS) 'For Loan
            Set lobjInterestSTStream = lobjReportSTTrans.GetEconStream(ST_EDA_LoanLendInt)
            Set lobjPrincipalSTStream = lobjReportSTTrans.GetEconStream(ST_EDA_LoanLendPrin)
        Else
            Set lobjPaymentSTStream = lobjReportSTTrans.GetEconStream(ST_EDA_Rent)  'For Lease
        End If
                
        If lobjPaymentSTStream.Count > 0 Then
            
            liTotalPeriods = lobjPaymentSTStream.Count
            
            lstrXML = "<AmortReportInfo>"
            For liIndex = 0 To liTotalPeriods - 1
                ldteAmortizationDate = lobjPaymentSTStream.GetDate(liIndex)
                
                lstrXML = lstrXML & "<Amort>" & _
                                        "<PaymentNo>" & (liIndex + 1) & "</PaymentNo>" & _
                                        "<AmortDate>" & ldteAmortizationDate & "</AmortDate>" & _
                                        "<PaymentAmt>" & lobjPaymentSTStream.GetAmount(liIndex) & "</PaymentAmt>"
                                        
                If lobjReportSTTrans.Mode = ST_Mode_Lender Then
                    lstrXML = lstrXML & "<InterestAmt>" & lobjInterestSTStream.GetAmount(liIndex) & "</InterestAmt>" & _
                                        "<PrincipalAmt>" & lobjPrincipalSTStream.GetAmount(liIndex) & "</PrincipalAmt>"
                End If
                
                lstrXML = lstrXML & "</Amort>"
            Next liIndex
            
            lstrXML = lstrXML & "</AmortReportInfo>"
        Else
            lstrXML = ""
        End If
    
        
        Set lvFile = lobjFileSystem.CreateTextFile(Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2AmortRepData.xml"), True)
        lvFile.WriteLine (lstrXML)
        lvFile.Close
        Set lobjFileSystem = Nothing
        MsgBox "Amortization Report Data successfully generated from PRM file and XML saved at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2AmortRepData.xml")
    Else
        MsgBox "Invalid PRM File Selected"
    End If
    
CleanUp:
    Set lobjInterestSTStream = Nothing
    Set lobjPaymentSTStream = Nothing
    Set lobjPrincipalSTStream = Nothing
    Set lobjReportSTTrans = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnAmortRepTxtFromXML_Click()
Dim lobjXML As New DOMDocument
Dim lobjFileSystem As New Scripting.FileSystemObject
Dim lvFile As Variant
Dim lstrXMLFilePath As String
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "XML (*.xml)|*.xml"
    CommonDialog1.ShowOpen
    lstrXMLFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrXMLFilePath), ".XML") > 0 Then
        lobjXML.Load (lstrXMLFilePath)
        'MsgBox lobjXML.documentElement.Text
        
        
        Set lvFile = lobjFileSystem.CreateTextFile(Replace(UCase(lstrXMLFilePath), ".XML", "_TxtRepXML.doc"), True)
        lvFile.WriteLine (lobjXML.documentElement.Text)
        lvFile.Close
        Set lobjFileSystem = Nothing
        MsgBox "Text report extracted from XML file and saved at : " & Replace(UCase(lstrXMLFilePath), ".XML", "_TxtRepXML.doc")
    Else
        MsgBox "Invalid XML File Selected"
    End If
    
CleanUp:
    Set lobjXML = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnPRMinXML_Click()
Dim lvXML As Variant
Dim lstrPRMFilePath As String
Dim lobjPRMFile As New ADODB.Stream
Dim lvFileData As Variant
Dim lobjFileSystem As New Scripting.FileSystemObject
Dim lvFile As Variant
Dim lobjPRMXML As New DOMDocument40
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
        lobjPRMFile.Type = adTypeBinary
        lobjPRMFile.Open
        lobjPRMFile.LoadFromFile lstrPRMFilePath
        
        lobjPRMXML.loadXML "<FILE_DATA></FILE_DATA>"
        lobjPRMXML.documentElement.dataType = "bin.base64"
        lobjPRMXML.documentElement.nodeTypedValue = lobjPRMFile.Read
        
        
        Set lvFile = lobjFileSystem.CreateTextFile(Replace(UCase(lstrPRMFilePath), ".PRM", "_BinXML.xml"), True, True)
        lvFile.WriteLine (lobjPRMXML.xml)
        lvFile.Close
        Set lobjFileSystem = Nothing
        
        MsgBox "PRM file successfully wrapped into an XML and saved at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_BinXML.xml")
    Else
        MsgBox "Invalid PRM File Selected"
    End If

CleanUp:
    Set lobjPRMFile = Nothing
    Set lobjPRMXML = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub


Private Sub btnPRMXML2DataXML_Click()
Dim lobjSTApp As New STApplication
Dim lstrXML As String
Dim lstrXMLFilePath As String
Dim lobjFile As New ADODB.Stream
Dim lvPRMXML As Variant
Dim lobjPRMXML As New DOMDocument
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "XML (*.xml)|*.xml"
    CommonDialog1.ShowOpen
    lstrXMLFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrXMLFilePath), ".XML") > 0 Then
        If lobjPRMXML.Load(lstrXMLFilePath) Then
            lvPRMXML = lobjPRMXML.documentElement.nodeTypedValue
            
            lobjFile.Type = adTypeBinary
            lobjFile.Open
            lobjFile.SetEOS
            lobjFile.Write lvPRMXML
            lobjFile.SaveToFile Replace(UCase(lstrXMLFilePath), ".XML", ".PRM"), adSaveCreateOverWrite
            lobjFile.Close
            
            lstrXML = "<SuperTRUMP WriteXMLFile = """ & Replace(UCase(lstrXMLFilePath), ".XML", "_BINXML2PRMXML.xml") & """ >" & _
                        "<Transaction id=""TRAN1"" query=""true"">" & _
                            "<ReadFile filename=""" & Replace(UCase(lstrXMLFilePath), ".XML", ".PRM") & """ />" & _
                        "</Transaction>" & _
                    "</SuperTRUMP>"
            lstrXML = lobjSTApp.XmlInOut(lstrXML)
            MsgBox "PRM XML wrapper file successfully converted to XML and saved at : " & Replace(UCase(lstrXMLFilePath), ".XML", "_BINXML2PRMXML.xml")
        Else
            MsgBox "Invalid XML File Selected"
        End If
        
    Else
        MsgBox "Invalid XML File Selected"
    End If
    
CleanUp:
    Set lobjSTApp = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnRentSchedule_Click()
Dim lobjReportSTTrans As New STTransaction
Dim lobjReportSTResults As STResults
Dim lstrReport As String
Dim lobjFileSystem As New Scripting.FileSystemObject
Dim lvFile As Variant
Dim lstrPRMFilePath As String
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.FilterIndex = 0
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
        lobjReportSTTrans.OpenFile (lstrPRMFilePath)
        lobjReportSTTrans.Calculate
        Set lobjReportSTResults = lobjReportSTTrans.Results
        lobjReportSTResults.ReportFileName = cREPORT_TEMPLATE_PATH & "\BASIC010.$BA"
        lstrReport = lobjReportSTResults.PrintBuffer
        
        
        Set lvFile = lobjFileSystem.CreateTextFile(Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2RentShedRep.xml"), True)
        lvFile.WriteLine ("<RentShedRep>" & "<![CDATA[" & lstrReport & "]]></RentShedRep>")
        lvFile.Close
        Set lobjFileSystem = Nothing
        MsgBox "Rent Schedule Report successfully generated from PRM file and XML saved at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2RentShed.xml")
    Else
        MsgBox "Invalid PRM File Selected"
    End If
    
CleanUp:
    Set lobjReportSTResults = Nothing
    Set lobjReportSTTrans = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnSchemaXML_Click()
Dim lobjSTApp As New STApplication
Dim lstrXML As String
On Error GoTo ERR_HANDLER:

    lstrXML = "<SuperTRUMP querySchema=""SuperTrumpSchema"" WriteXMLFile = """ & App.Path & "\STSchema.xml"" />"
    lstrXML = lobjSTApp.XmlInOut(lstrXML)
    MsgBox "Super Trump Schema Present at : " & App.Path & "\STSchema.xml"
    
CleanUp:
    Set lobjSTApp = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnSummaryReportXML_Click()
Dim lobjReportSTTrans As New STTransaction
Dim lobjReportSTResults As STResults
Dim lstrReport As String
Dim lobjFileSystem As New Scripting.FileSystemObject
Dim lvFile As Variant
Dim lstrPRMFilePath As String
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
        lobjReportSTTrans.OpenFile (lstrPRMFilePath)
        lobjReportSTTrans.Calculate
        Set lobjReportSTResults = lobjReportSTTrans.Results
        'lobjReportSTResults.ReportFileName = cREPORT_TEMPLATE_PATH & "\WIRE_H01.$CS"
        lobjReportSTResults.ReportFileName = cREPORT_TEMPLATE_PATH & "\WIRE_H01.$BA"
        lstrReport = lobjReportSTResults.PrintBuffer
        
        
        Set lvFile = lobjFileSystem.CreateTextFile(Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2SummRep.xml"), True)
        lvFile.WriteLine ("<SummRep>" & "<![CDATA[" & lstrReport & "]]></SummRep>")
        lvFile.Close
        Set lobjFileSystem = Nothing
        MsgBox "Summary Report successfully generated from PRM file and XML saved at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2SummRep.xml")
    Else
        MsgBox "Invalid PRM File Selected"
    End If
    
CleanUp:
    Set lobjReportSTResults = Nothing
    Set lobjReportSTTrans = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnTerminationFullRepXML_Click()
Dim lobjReportSTTrans As New STTransaction
Dim lobjReportSTResults As STResults
Dim lstrReport As String
Dim lobjFileSystem As New Scripting.FileSystemObject
Dim lvFile As Variant
Dim lstrPRMFilePath As String
Dim lobjSTPrvVal As STPresentValue
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
        lobjReportSTTrans.OpenFile (lstrPRMFilePath)
        lobjReportSTTrans.Calculate
        Set lobjSTPrvVal = lobjReportSTTrans.PresentValue(1)
        
        Set lobjReportSTResults = lobjReportSTTrans.Results
        lobjReportSTResults.ReportFileName = cREPORT_TEMPLATE_PATH & "\WIRE_H52.$BA"
        lstrReport = lobjReportSTResults.PrintBuffer
        
        
        Set lvFile = lobjFileSystem.CreateTextFile(Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2TerminationFullRep.xml"), True)
        lvFile.WriteLine ("<TerminationFullRep>" & "<![CDATA[" & lstrReport & "]]></TerminationFullRep>")
        lvFile.Close
        Set lobjFileSystem = Nothing
        MsgBox "Termination Full Report successfully generated from PRM file and XML saved at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2TerminationFullRep.xml")
    Else
        MsgBox "Invalid PRM File Selected"
    End If
    
CleanUp:
    Set lobjReportSTResults = Nothing
    Set lobjReportSTTrans = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnTransactionAmtXML_Click()
Dim lobjSTApp As New STApplication
Dim lstrXML As String
Dim lstrPRMFilePath As String
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
        lstrXML = "<SuperTRUMP WriteXMLFile = """ & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2TransAmt.xml") & """ >" & _
                    "<Transaction id=""TRAN2"">" & _
                        "<ReadFile filename=""" & lstrPRMFilePath & """ />" & _
                        "<TransactionAmount query=""true"" />" & _
                    "</Transaction>" & _
                "</SuperTRUMP>"
        lstrXML = lobjSTApp.XmlInOut(lstrXML)
        MsgBox "Transaction Amount successfully extracted from PRM file and XML saved at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2TransAmt.xml")
    Else
        MsgBox "Invalid PRM File Selected"
    End If
    
CleanUp:
    Set lobjSTApp = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub btnAmortRepXML_Click()
Dim lobjReportSTTrans As New STTransaction
Dim lobjReportSTResults As STResults
Dim lstrReport As String
Dim lobjFileSystem As New Scripting.FileSystemObject
Dim lvFile As Variant
Dim lstrPRMFilePath As String
On Error GoTo ERR_HANDLER:
    
    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
        lobjReportSTTrans.OpenFile (lstrPRMFilePath)
        lobjReportSTTrans.Calculate
        Set lobjReportSTResults = lobjReportSTTrans.Results
        lobjReportSTResults.ReportFileName = cREPORT_TEMPLATE_PATH & "\Wire_h16.$cu"
        lstrReport = lobjReportSTResults.PrintBuffer
        
        
        Set lvFile = lobjFileSystem.CreateTextFile(Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2AmortRep.xml"), True)
        lvFile.WriteLine ("<AmortRep>" & "<![CDATA[" & lstrReport & "]]></AmortRep>")
        lvFile.Close
        Set lobjFileSystem = Nothing
        MsgBox "Amortization Report successfully generated from PRM file and XML saved at : " & Replace(UCase(lstrPRMFilePath), ".PRM", "_BIN2AmortRep.xml")
    Else
        MsgBox "Invalid PRM File Selected"
    End If
    
CleanUp:
    Set lobjReportSTResults = Nothing
    Set lobjReportSTTrans = Nothing
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error# :" & Err.Number & _
            "Error Source :" & Err.Source & _
            "Error Description :" & Err.Description
    Resume CleanUp
End Sub

Private Sub Command1_Click()
'Declare Super Trump Variables
Dim lobjSTTransaction               As New STTransaction
Dim lobjSTQuick                     As STQuick
Dim lobjSTStream                    As STStream

'Declare XML Dom variables
Dim lobjXMLSchemaSpace              As New XMLSchemaCache40
Dim lobjReturnAmortSchedLstXMLDOM   As New DOMDocument40
Dim lobjPRMFileListXMLDOM           As New DOMDocument40
Dim lobjPRMFileBinDataElem          As IXMLDOMElement

'Other Declarations
Dim llErrNbr                        As Long
Dim lstrErrDesc                     As String
Dim lstrFileLoc                     As String
Dim lstrPRMFileListXML              As String
Dim lstrPRMFilePath                 As String
Dim liPRMFileLstCnt                 As Integer
Dim lstrPRMFileName                 As String
Dim liStreamCnt                     As Integer
Dim ldLeaseFactor                   As Double
Dim lbGetAmort                      As Boolean
Dim lstrOUT                         As String

    CommonDialog1.Filter = "PRM (*.prm)|*.prm"
    CommonDialog1.ShowOpen
    lstrPRMFilePath = CommonDialog1.FileName
    If InStr(1, UCase(lstrPRMFilePath), ".PRM") > 0 Then
    
        'Check if the PRM file can be read
        If lobjSTTransaction.OpenFile(lstrPRMFilePath) Then
            lobjSTTransaction.Calculate
            'Initialization
            Set lobjSTQuick = lobjSTTransaction.Quick
            lobjSTQuick.ReTarget
            
            
            'Get the Amortization Data
            lobjSTTransaction.GetFreeStream
            If lobjSTTransaction.Mode = ST_Mode_Lender Then
                Set lobjSTStream = lobjSTTransaction.GetEconStream(ST_EDA_LoanLendDS) 'For Loan
            Else
                Set lobjSTStream = lobjSTTransaction.GetEconStream(ST_EDA_Rent)  'For Lease
            End If
            
            lstrOUT = ""
                            
            'For each payment build the <PAYMENT_LIST> node
            For liStreamCnt = 0 To lobjSTStream.Count - 1
                    
                'For Loans, the last row contains Amt = $0 for balloon type.
                'which is not required. Hence ignore the last row.
                If lobjSTTransaction.Mode = ST_Mode_Lender _
                    And liStreamCnt = lobjSTStream.Count - 1 _
                    And lobjSTStream.GetAmount(liStreamCnt) = 0 Then
    
                    Exit For
                End If
                
                lstrOUT = lstrOUT & liStreamCnt + 1 & " " & lobjSTStream.GetDate(liStreamCnt) & " " & lobjSTStream.GetAmount(liStreamCnt) & vbCrLf
                
                                                
            Next
        End If
    End If
End Sub
