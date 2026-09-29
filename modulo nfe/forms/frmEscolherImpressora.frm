VERSION 5.00
Begin VB.Form frmEscolherImpressora 
   Caption         =   "Impressão"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   7650
   Icon            =   "frmEscolherImpressora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   7650
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   5760
      Picture         =   "frmEscolherImpressora.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox cmbxPrinter 
      BackColor       =   &H80000014&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1395
      Width           =   7215
   End
   Begin VB.CommandButton cmndbtnConfirmar 
      Caption         =   "&Confirmar"
      Default         =   -1  'True
      Height          =   735
      Left            =   3960
      Picture         =   "frmEscolherImpressora.frx":0E58
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   7620
      TabIndex        =   3
      Top             =   0
      Width           =   7650
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         Picture         =   "frmEscolherImpressora.frx":15C2
         ScaleHeight     =   735
         ScaleWidth      =   975
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
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
         Left            =   1080
         TabIndex        =   6
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
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label lblPorta 
      AutoSize        =   -1  'True
      Caption         =   "Porta"
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "frmEscolherImpressora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdSair_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmndbtnConfirmar_Click()
    Me.Hide
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
    
    For Each p In Printers
        cmbxPrinter.AddItem p.deviceName
    Next
    
    P_Padrao
    cmbxPrinter.tag = ""
    
    
End Sub
Sub P_Padrao()
    Dim iCount As Integer
    For iCount = 0 To cmbxPrinter.ListCount - 1
        cmbxPrinter.ListIndex = iCount
        If cmbxPrinter.Text = Printer.deviceName Then
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
