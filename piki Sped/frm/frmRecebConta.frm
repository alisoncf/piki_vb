VERSION 5.00
Begin VB.Form frmRecebConta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$ Multi-Fácil $- Recebimento de Contas"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "ESC -  &Fechar"
      Height          =   615
      Left            =   3360
      Picture         =   "frmRecebConta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtbxValor_Cai 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "0,00"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "F7 - &Confirmar"
      Height          =   615
      Left            =   1920
      Picture         =   "frmRecebConta.frx":00F5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005E392B&
      Height          =   330
      Left            =   3360
      TabIndex        =   9
      Top             =   480
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   240
      Width           =   390
   End
   Begin VB.Label lblUsu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Usuário:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005E392B&
      Height          =   330
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   2205
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   240
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Valor:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label lblCaixa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005E392B&
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   435
   End
End
Attribute VB_Name = "frmRecebConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If txtbxValor_Cai.Text = "0,00" Then
    
        MsgBox "Informe o valor correto!", vbExclamation
        txtbxValor_Cai.SetFocus
        Exit Sub
    End If
    strSQL = F_GravaCai("C", Date, txtbxValor_Cai.Text, "RECEB. MULTI-FÁCIL", "", "MUL", _
        strCod_Cai, 0, 0, strCod_Cai, strCod_Tur)
    If strSQL <> "" Then
        MsgBox "Lançamento confirmado!" & Chr(13) & "Controle Nº:" & strSQL, vbInformation
        txtbxValor_Cai.Text = "0,00"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    End If
    
End Sub

Private Sub Form_Load()
    lblData.caption = Date
    lblCaixa.caption = strCod_Cai
    lblUsu.caption = strNome_Usu
End Sub

Private Sub txtbxValor_Cai_GotFocus()
    ganhaFoco txtbxValor_Cai
End Sub

Private Sub txtbxValor_Cai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValor_Cai_LostFocus()
    PerdeFoco txtbxValor_Cai
End Sub
