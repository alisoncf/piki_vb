VERSION 5.00
Begin VB.Form frmCadSubSec 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Subseções"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtbxComiss 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtbxDescmax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox cmbxSecao 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   4560
      Picture         =   "frmCadSubSec.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Gravar"
      Height          =   735
      Left            =   3120
      Picture         =   "frmCadSubSec.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.PictureBox pcStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00B36A09&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   6345
      TabIndex        =   12
      Top             =   3510
      Width           =   6375
      Begin VB.Label lblSatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3045
         TabIndex        =   13
         Tag             =   "!"
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subseção"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   1155
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5880
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   120
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Comissão:"
      Height          =   195
      Left            =   750
      TabIndex        =   7
      Top             =   1965
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   195
      Left            =   2520
      TabIndex        =   6
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Desc. Máximo:"
      Height          =   195
      Left            =   420
      TabIndex        =   4
      Top             =   1605
      Width           =   1050
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Seção:"
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   900
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome:"
      Height          =   195
      Left            =   1005
      TabIndex        =   2
      Top             =   1245
      Width           =   465
   End
End
Attribute VB_Name = "frmCadSubSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxSecao_Click()
    cmbxSecao.tag = F_BuscaUmCampo("cod_sec", "tab_secoes", "nome_sec='" & cmbxSecao.Text & "'")
    
End Sub

Private Sub cmbxSecao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If cmbxSecao.tag = "" Then
        MsgBox "Informe a Seção!", vbExclamation
        cmbxSecao.SetFocus
        Exit Sub
    End If
    If txtbxNome.Text = "" Then
        MsgBox "Informe o nome da subseção!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    If Me.tag = "" Then
        If subsecAO.F_GravaSub(cmbxSecao.tag, txtbxNome.Text, txtbxDescmax.Text, txtbxComiss.Text) <> "" Then
            MsgBox "Subseção incluída com sucesso!", vbInformation
        End If
    Else
        If subsecAO.F_AlteraSub(Me.tag, cmbxSecao.tag, txtbxNome.Text, txtbxDescmax.Text, txtbxComiss.Text) Then
            MsgBox "Subseção alterada com sucesso!", vbInformation
        End If
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    secao.P_PreencheSecao cmbxSecao
End Sub



Private Sub txtbxComiss_GotFocus()
    P_SelecTexto txtbxComiss
End Sub

Private Sub txtbxComiss_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxComiss_LostFocus()
    P_ValorLostFocus txtbxComiss
    If CCur(txtbxComiss.Text) >= 100 Then
        MsgBox "muito alta a comissão!", vbExclamation
        txtbxComiss.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtbxDescmax_GotFocus()
    P_SelecTexto txtbxDescmax
End Sub

Private Sub txtbxDescmax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDescmax_LostFocus()
    P_ValorLostFocus txtbxDescmax
    If CCur(txtbxDescmax.Text) > 100 Then
        MsgBox "valor muito alto!", vbCritical
        txtbxDescmax.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
    
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub
