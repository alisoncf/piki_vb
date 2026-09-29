VERSION 5.00
Begin VB.Form frmEscolherImpressora 
   Caption         =   "Escolha uma Impressora Padrão"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   7020
   Icon            =   "frmEscolherImpressora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   7020
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Impressora"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox cmbxTamPap 
         BackColor       =   &H80000014&
         Height          =   315
         ItemData        =   "frmEscolherImpressora.frx":08CA
         Left            =   3000
         List            =   "frmEscolherImpressora.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cmbxAltura 
         BackColor       =   &H80000014&
         Height          =   315
         ItemData        =   "frmEscolherImpressora.frx":08F2
         Left            =   3000
         List            =   "frmEscolherImpressora.frx":08FF
         Style           =   1  'Simple Combo
         TabIndex        =   8
         Text            =   "0"
         Top             =   1020
         Width           =   1335
      End
      Begin VB.ComboBox cmbxQuali 
         BackColor       =   &H80000014&
         Height          =   315
         ItemData        =   "frmEscolherImpressora.frx":0922
         Left            =   930
         List            =   "frmEscolherImpressora.frx":092F
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnConfirmar 
         Cancel          =   -1  'True
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   4680
         Picture         =   "frmEscolherImpressora.frx":0952
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cmbxPrinter 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   320
         Width           =   5415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tamanho do Papel:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1500
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Altura:"
         Height          =   195
         Left            =   2520
         TabIndex        =   9
         Top             =   1080
         Width           =   450
      End
      Begin VB.Line Line1 
         X1              =   6720
         X2              =   6720
         Y1              =   120
         Y2              =   2520
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qualidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblPorta 
         AutoSize        =   -1  'True
         Caption         =   "Porta"
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Porta:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmEscolherImpressora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxAltura_LostFocus()
    PerdeFoco cmbxAltura
End Sub

Private Sub cmbxPrinter_Click()
    If cmbxPrinter.tag = "" Then
    Set Printer = Printers(cmbxPrinter.ListIndex)
    lblPorta.caption = Printer.Port
        
        
    End If
End Sub

Private Sub Command1_Click()
    objP.Print "teste"
    objP.EndDoc
End Sub

Private Sub cmndbtnConfirmar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.Hide
    ElseIf KeyCode = vbKeyEscape Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dim p As Printer
    Dim i As Integer
    cmbxPrinter.Clear
    cmbxPrinter.tag = "printer"
    cmbxQuali.ListIndex = 1
    cmbxQuali.Enabled = False
    For Each p In Printers
        cmbxPrinter.AddItem p.DeviceName
    Next
    
    P_Padrao
    cmbxPrinter.tag = ""
    cmbxPrinter_Click
End Sub
Sub P_Padrao()
    Dim iCount As Integer
    For iCount = 0 To cmbxPrinter.ListCount - 1
        cmbxPrinter.ListIndex = iCount
        If cmbxPrinter.Text = Printer.DeviceName Then
         '   MsgBox "encontrou"
            Exit Sub
        End If
    Next
End Sub

Private Sub optnDE_Click()
    txtbxIni.Enabled = True
    txtbxFin.Enabled = True
End Sub

Private Sub optnSelect_Click()
    txtbxIni.Enabled = False
    txtbxFin.Enabled = False
End Sub

Private Sub optnTUdo_Click()
    txtbxIni.Enabled = False
    txtbxFin.Enabled = False
End Sub
