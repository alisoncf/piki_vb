VERSION 5.00
Begin VB.Form frmCadCid 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inclusão de Cidades"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5670
   Icon            =   "frmCadCid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEstado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Salvar"
      Height          =   735
      Left            =   1680
      Picture         =   "frmCadCid.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   3360
      Picture         =   "frmCadCid.frx":059E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtbxDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
   Begin VB.ComboBox cmbxUf 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&UF:"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cidade:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   540
   End
End
Attribute VB_Name = "frmCadCid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.tag = "" Then
        If F_BuscaUmCampo("Nome_Cid", TAB_CIDADES, _
        "nome_cid='" & txtbxDesc.Text & "' AND uf_cid ='" & cmbxUF.Text & "'") <> "" Then
            MsgBox "Esta cidade já está cadastrada!", vbExclamation
            Exit Sub
        End If
        If txtbxDesc.Text = "" Then
            MsgBox "Informe o nome da cidade!", vbExclamation
            Exit Sub
        End If
        If cmbxUF.Text = "" Then
            MsgBox "Informe o estado corretamente!", vbExclamation
            cmbxUF.SetFocus
            Exit Sub
        End If
        grava
    Else
        altera
    End If
End Sub
Sub altera()
    Cidades.F_AlteraCidade Me.tag, txtbxDesc.Text, cmbxUF.Text
    MsgBox "Alterado!", vbInformation
End Sub
Sub grava()
    Cidade.P_GravaCidade txtCod.Text, txtbxDesc.Text, txtEstado.Text, cmbxUF.Text
    MsgBox "Cidade inserida", vbInformation
    txtCod.Text = ""
    txtCod.SetFocus
End Sub
Private Sub Form_Load()
    P_PreencheUF cmbxUF
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtbxDesc_GotFocus()
    ganhaFoco txtbxDesc
End Sub

Private Sub txtbxDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDesc_LostFocus()
    AjustaTexto txtbxDesc
End Sub

