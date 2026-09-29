VERSION 5.00
Begin VB.Form frmTestPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "frmTestPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbxPrinter 
      BackColor       =   &H80000014&
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1035
      Width           =   6135
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Impressão"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   4215
      Begin VB.OptionButton optnTUdo 
         Caption         =   "&Tudo"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optnSelect 
         Caption         =   "&Seleção"
         Enabled         =   0   'False
         Height          =   195
         Left            =   960
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optnDE 
         Caption         =   "&De"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtbxIni 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   11
         Top             =   290
         Width           =   375
      End
      Begin VB.TextBox txtbxFin 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   10
         Top             =   290
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   360
         Width           =   285
      End
   End
   Begin VB.CommandButton cmndbtnConfirmar 
      Cancel          =   -1  'True
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   5520
      Picture         =   "frmTestPrint.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ComboBox cmbxQuali 
      BackColor       =   &H80000014&
      Height          =   315
      ItemData        =   "frmTestPrint.frx":1034
      Left            =   1050
      List            =   "frmTestPrint.frx":1041
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1740
      Width           =   1455
   End
   Begin VB.ComboBox cmbxAltura 
      BackColor       =   &H80000014&
      Height          =   315
      ItemData        =   "frmTestPrint.frx":1064
      Left            =   3120
      List            =   "frmTestPrint.frx":1071
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Text            =   "0"
      Top             =   1740
      Width           =   1335
   End
   Begin VB.ComboBox cmbxTamPap 
      BackColor       =   &H80000014&
      Height          =   315
      ItemData        =   "frmTestPrint.frx":1094
      Left            =   3120
      List            =   "frmTestPrint.frx":10A1
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "frmTestPrint.frx":10BC
      ScaleHeight     =   825
      ScaleWidth      =   7725
      TabIndex        =   7
      Top             =   0
      Width           =   7755
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione uma Impressora"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   960
         TabIndex        =   19
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Porta:"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label lblPorta 
      AutoSize        =   -1  'True
      Caption         =   "Porta"
      Height          =   195
      Left            =   1080
      TabIndex        =   17
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qualidade:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Altura:"
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tamanho do Papel:"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   2220
      Visible         =   0   'False
      Width           =   1395
   End
End
Attribute VB_Name = "frmTestPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxAltura_LostFocus()
    P_ValorLostFocus cmbxAltura
End Sub

Private Sub cmbxPrinter_Click()
    If cmbxPrinter.tag = "" Then
        Set Printer = Printers(cmbxPrinter.ListIndex)
        lblPorta.caption = Printer.Port
        frmMenu.mudaVariaveis
        
    End If
End Sub

Private Sub Command1_Click()
    objP.Print "teste"
    objP.EndDoc
End Sub

Private Sub cmndbtnConfirmar_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
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
