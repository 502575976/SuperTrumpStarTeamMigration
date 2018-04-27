VERSION 5.00
Begin VB.Form frmSTWebSvcTest 
   Caption         =   "ST Web Service Test"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnProcessMQMessage 
      Caption         =   "ProcessMQMessage"
      Height          =   435
      Left            =   2640
      TabIndex        =   19
      Top             =   5100
      Width           =   2145
   End
   Begin VB.CommandButton btnProcessPricingRequest 
      Caption         =   "ProcessPricingRequest"
      Height          =   435
      Left            =   120
      TabIndex        =   18
      Top             =   5100
      Width           =   2145
   End
   Begin VB.TextBox txtInputPath 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   7275
   End
   Begin VB.TextBox txtOutputPath 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1920
      TabIndex        =   7
      Top             =   3360
      Width           =   5475
   End
   Begin VB.CommandButton btnModifyPRMFiles 
      Caption         =   "ModifyPRMFiles"
      Height          =   435
      Left            =   5220
      TabIndex        =   13
      Top             =   4560
      Width           =   2145
   End
   Begin VB.CheckBox chkSave2File 
      Caption         =   "Save output to path:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Test"
      Height          =   345
      Left            =   5940
      TabIndex        =   15
      Top             =   5640
      Width           =   555
   End
   Begin VB.CommandButton btn_Ping 
      Caption         =   "Ping"
      Height          =   345
      Left            =   4800
      TabIndex        =   14
      Top             =   5640
      Width           =   555
   End
   Begin VB.TextBox txtInput 
      Height          =   1455
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   840
      Width           =   7335
   End
   Begin VB.TextBox txtOutput 
      Height          =   1575
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   6000
      Width           =   7275
   End
   Begin VB.CommandButton btnGetPRMParams 
      Caption         =   "GetPRMParams"
      Height          =   435
      Left            =   2655
      TabIndex        =   12
      Top             =   4530
      Width           =   2145
   End
   Begin VB.CommandButton btnGetPricingRep 
      Caption         =   "GetPricingReports"
      Height          =   435
      Left            =   90
      TabIndex        =   11
      Top             =   4530
      Width           =   2145
   End
   Begin VB.CommandButton btnGetAmortSched 
      Caption         =   "GetAmortizationSchedule"
      Height          =   435
      Left            =   5220
      TabIndex        =   10
      Top             =   3930
      Width           =   2145
   End
   Begin VB.CommandButton btnGenPRM 
      Caption         =   "GeneratePRM"
      Height          =   435
      Left            =   2655
      TabIndex        =   9
      Top             =   3930
      Width           =   2145
   End
   Begin VB.CommandButton btnConvertPRM2XML 
      Caption         =   "ConvertPRM2XML"
      Height          =   435
      Left            =   90
      TabIndex        =   8
      Top             =   3930
      Width           =   2145
   End
   Begin VB.ComboBox cboWSDL 
      Height          =   315
      ItemData        =   "frmSTWebSvcTest.frx":0000
      Left            =   660
      List            =   "frmSTWebSvcTest.frx":0016
      TabIndex        =   1
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label4 
      Caption         =   "Note: If Input XML is not specified then default Input XML would be used from the following location:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   7215
   End
   Begin VB.Label Label3 
      Caption         =   "Input XML:"
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Output XML:"
      Height          =   285
      Left            =   180
      TabIndex        =   16
      Top             =   5670
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "WSDL:"
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   675
   End
End
Attribute VB_Name = "frmSTWebSvcTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum eSTMethodOptions
    ecConvertPRMToXML = 0
    ecGeneratePRMFiles = 1
    ecGetAmortizationSchedule = 2
    ecGetPricingReports = 3
    ecGetPRMParams = 4
    ecModifyPRMFiles = 5
    ecProcessPricingRequest = 6
    ecProcessMQMessage = 7
End Enum

Private Sub btn_Ping_Click()
On Error GoTo ErrHandler

Dim lobjSoapClient  As New SoapClient30

    Screen.MousePointer = 11
    'lobjSoapClient.ClientProperty("ServerHTTPRequest") = True
    lobjSoapClient.MSSoapInit cboWSDL.Text, , "ISuperTrumpServiceSoapPort"

    txtOutput.Text = lobjSoapClient.Ping()
    Screen.MousePointer = 0
    Exit Sub

ErrHandler:
    MsgBox "Error occurred - " & Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub btnConvertPRM2XML_Click()
    Screen.MousePointer = 11
    InvokeSTWebSvc ecConvertPRMToXML
    Screen.MousePointer = 0
End Sub

Public Function InvokeSTWebSvc(ByVal astrMethod As eSTMethodOptions)
On Error GoTo ErrHandler

Dim lobjSoapClient  As New SoapClient30
Dim lobjINXML       As New DOMDocument
Dim lstrINXML       As String
Dim lstrOUTXML      As String
Dim lstrOUTFileName As String

    txtOutput.Text = ""
    'lobjSoapClient.ClientProperty("ServerHTTPRequest") = True
    lobjSoapClient.MSSoapInit cboWSDL.Text, , "ISuperTrumpServiceSoapPort"
    lstrINXML = txtInput.Text

    Select Case astrMethod
        Case ecConvertPRMToXML
            If lstrINXML = "" Then
                lobjINXML.Load (txtInputPath.Text & "\ConvertPRMToXMLtest.xml")
                lstrINXML = lobjINXML.xml
            End If

            lstrOUTXML = lobjSoapClient.ConvertPRMToXML(lstrINXML)
            lstrOUTFileName = "ConvertPRMToXML_OUT.xml"

        Case ecGeneratePRMFiles
            If lstrINXML = "" Then
                lobjINXML.Load (txtInputPath.Text & "\GeneratePRMFilestest.xml")
                lstrINXML = lobjINXML.xml
            End If

            lstrOUTXML = lobjSoapClient.GeneratePRMFiles(lstrINXML)
            lstrOUTFileName = "GeneratePRMFiles_OUT.xml"

        Case ecGetAmortizationSchedule
            If lstrINXML = "" Then
                lobjINXML.Load (txtInputPath.Text & "\GetAmortizationScheduletest.xml")
                lstrINXML = lobjINXML.xml
            End If

            lstrOUTXML = lobjSoapClient.GetAmortizationSchedule(lstrINXML)
            lstrOUTFileName = "GetAmortizationSchedule_OUT.xml"

        Case ecGetPricingReports
            If lstrINXML = "" Then
                lobjINXML.Load (txtInputPath.Text & "\GetPricingReportstest.xml")
                lstrINXML = lobjINXML.xml
            End If

            lstrOUTXML = lobjSoapClient.GetPricingReports(lstrINXML)
            lstrOUTFileName = "GetPricingReports_OUT.xml"

        Case ecGetPRMParams
            If lstrINXML = "" Then
                lobjINXML.Load (txtInputPath.Text & "\GetPRMParamstest.xml")
                lstrINXML = lobjINXML.xml
            End If

            lstrOUTXML = lobjSoapClient.GetPRMParams(lstrINXML)
            lstrOUTFileName = "GetPRMParams_OUT.xml"

        Case ecModifyPRMFiles
            If lstrINXML = "" Then
                lobjINXML.Load (txtInputPath.Text & "\ModifyPRMFilestest.xml")
                lstrINXML = lobjINXML.xml
            End If

            lstrOUTXML = lobjSoapClient.ModifyPRMFiles(lstrINXML)
            lstrOUTFileName = "ModifyPRMFiles_OUT.xml"
            
        Case ecProcessPricingRequest
            If lstrINXML = "" Then
                lobjINXML.Load (txtInputPath.Text & "\ProcessPricingRequestTEST.xml")
                lstrINXML = lobjINXML.xml
            End If
            
            lobjSoapClient.MSSoapInit cboWSDL.Text, , "IClientServiceSoapPort"
            
            lstrOUTXML = lobjSoapClient.ProcessPricingRequest(lstrINXML)
            lstrOUTFileName = "ProcessPricingRequest_OUT.xml"
            
        Case ecProcessMQMessage
            If lstrINXML = "" Then
                lobjINXML.Load (txtInputPath.Text & "\ProcessMQMessagetest.xml")
                lstrINXML = lobjINXML.xml
            End If
            lobjSoapClient.MSSoapInit cboWSDL.Text, , "IClientServiceSoapPort"
            lstrOUTXML = lobjSoapClient.ProcessMQMessage(lstrINXML)
            lstrOUTFileName = "ProcessMQMessage_OUT.xml"
    End Select

    If chkSave2File.Value Then
        lstrOUTFileName = txtOutputPath.Text & "\" & lstrOUTFileName
        SaveOutput lstrOUTFileName, lstrOUTXML
        txtOutput.Text = "Output saved to: " & lstrOUTFileName
    Else
        txtOutput.Text = lstrOUTXML
    End If

    Exit Function

ErrHandler:
    MsgBox "Error occurred - " & Err.Description
End Function

Private Sub btnGenPRM_Click()
    Screen.MousePointer = 11
    InvokeSTWebSvc ecGeneratePRMFiles
    Screen.MousePointer = 0
End Sub

Private Sub btnGetAmortSched_Click()
    Screen.MousePointer = 11
    InvokeSTWebSvc ecGetAmortizationSchedule
    Screen.MousePointer = 0
End Sub

Private Sub btnGetPricingRep_Click()
    Screen.MousePointer = 11
    InvokeSTWebSvc ecGetPricingReports
    Screen.MousePointer = 0
End Sub

Private Sub btnGetPRMParams_Click()
    Screen.MousePointer = 11
    InvokeSTWebSvc ecGetPRMParams
    Screen.MousePointer = 0
End Sub

Private Sub btnModifyPRMFiles_Click()
    Screen.MousePointer = 11
    InvokeSTWebSvc ecModifyPRMFiles
    Screen.MousePointer = 0
End Sub

Private Sub btnProcessMQMessage_Click()
    Screen.MousePointer = 11
    InvokeSTWebSvcNew ecProcessMQMessage
    Screen.MousePointer = 0
End Sub

Private Sub btnProcessPricingRequest_Click()
    Screen.MousePointer = 11
    InvokeSTWebSvcNew ecProcessPricingRequest
    Screen.MousePointer = 0
End Sub

Private Sub btnTest_Click()
On Error GoTo ErrHandler

Dim lobjSoapClient  As New SoapClient30

    Screen.MousePointer = 11
    'lobjSoapClient.ClientProperty("ServerHTTPRequest") = True
    lobjSoapClient.MSSoapInit cboWSDL.Text, , "ISuperTrumpServiceSoapPort"

    txtOutput.Text = lobjSoapClient.Test()
    Screen.MousePointer = 0

    Exit Sub

ErrHandler:
    MsgBox "Error occurred - " & Err.Description
    Screen.MousePointer = 0
End Sub

Public Function SaveOutput(ByVal astrFileName As String, _
                                    ByVal astrData As String) As Boolean

On Error GoTo ErrHandler:

Dim lobjFileSystem  As New Scripting.FileSystemObject
Dim lobjFile        As Scripting.File
Dim lobjTxtStream   As Scripting.TextStream
Dim liIOMode        As Integer

    'Create and open the file
    Set lobjTxtStream = lobjFileSystem.CreateTextFile(astrFileName, True)

    'Write data to the file
    lobjTxtStream.WriteLine astrData

    'Close the file
    lobjTxtStream.Close

    Set lobjFile = Nothing
    Set lobjTxtStream = Nothing
    Set lobjFileSystem = Nothing
    SaveOutput = True

    Exit Function

ErrHandler:
    SaveOutput = False
    Set lobjFile = Nothing
    Set lobjTxtStream = Nothing
    Set lobjFileSystem = Nothing
    Err.Raise Err.Number, "SaveOutput", Err.Description
End Function

Private Sub chkSave2File_Click()
    If chkSave2File.Value = vbChecked Then
        txtOutputPath.Enabled = True
        txtOutputPath.Text = App.Path & "\XML_OUT"
    Else
        txtOutputPath.Text = ""
        txtOutputPath.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    txtInputPath.Text = App.Path & "\XML_IN"
End Sub

Public Function InvokeSTWebSvcNew(ByVal astrMethod As eSTMethodOptions)
On Error GoTo ErrHandler

Dim lobjSoapClient  As New SoapClient30
Dim lobjINXML       As New DOMDocument
Dim lstrINXML       As String
Dim lstrOUTXML      As String
Dim lstrOUTFileName As String

    txtOutput.Text = ""
    lobjSoapClient.MSSoapInit cboWSDL.Text, , "IClientServiceSoapPort"
    lstrINXML = txtInput.Text

    Select Case astrMethod
        Case ecProcessPricingRequest
            If lstrINXML = "" Then
                lobjINXML.Load (txtInputPath.Text & "\ProcessPricingRequesttest.xml")
                lstrINXML = lobjINXML.xml
            End If

            lstrOUTXML = lobjSoapClient.ProcessPricingRequest(lstrINXML)
            lstrOUTFileName = "ProcessPricingRequest_OUT.xml"

        Case ecProcessMQMessage
            If lstrINXML = "" Then
                lobjINXML.Load (txtInputPath.Text & "\ProcessMQMessagetest.xml")
                lstrINXML = lobjINXML.xml
            End If

            lstrOUTXML = lobjSoapClient.ProcessMQMessage(lstrINXML)
            lstrOUTFileName = "ProcessMQMessage_OUT.xml"

    End Select

    If chkSave2File.Value Then
        lstrOUTFileName = txtOutputPath.Text & "\" & lstrOUTFileName
        SaveOutput lstrOUTFileName, lstrOUTXML
        txtOutput.Text = "Output saved to: " & lstrOUTFileName
    Else
        txtOutput.Text = lstrOUTXML
    End If

    Exit Function

ErrHandler:
    MsgBox "Error occurred - " & Err.Description
End Function
