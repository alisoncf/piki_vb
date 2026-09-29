VERSION 5.00
Begin VB.Form frmIncCid 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inclusão de Cidades"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5910
   Icon            =   "frmIncCid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Salvar"
      Height          =   735
      Left            =   1680
      Picture         =   "frmIncCid.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   3360
      Picture         =   "frmIncCid.frx":059E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtbxDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.ComboBox cmbxUf 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&UF:"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cidade:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "frmIncCid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.tag = "" Then
        If F_BuscaUmCampo("Nome_Cid", "tab_cidades", _
        "nome_cid='" & txtbxDesc.Text & "' AND uf_cid ='" & cmbxUf.Text & "'") <> "" Then
            MsgBox "Esta cidade já está cadastrada!", vbExclamation
            Exit Sub
        End If
        If txtbxDesc.Text = "" Then
            MsgBox "Informe o nome da cidade!", vbExclamation
            Exit Sub
        End If
        If cmbxUf.Text = "" Then
            MsgBox "Informe o estado corretamente!", vbExclamation
            cmbxUf.SetFocus
            Exit Sub
        End If
        grava
    Else
        altera
    End If
End Sub
Sub altera()
    Dim strSQL As String
    
    
    strSQL = "update tab_cidades set nome_Cid='" & txtbxDesc.Text & "',"
    strSQL = strSQL & "uf_cid='" & cmbxUf.Text & "' "
    strSQL = strSQL & " where cod_cid='" & Me.tag & "' "
    Conn.Execute strSQL
    MsgBox "Cidade Alterada!", vbInformation
    Unload Me
End Sub
Sub grava()
    Dim strSQL As String
    strSQL = "insert into tab_cidades(cod_cid,nome_cid,uf_cid)Values("
    strSQL = strSQL & "'" & MaxKeySQL("tab_cidades", "cod_Cid", "", "0000#") & "',"
    strSQL = strSQL & "'" & txtbxDesc.Text & "',"
    strSQL = strSQL & "'" & cmbxUf.Text & "') "
    Conn.Execute strSQL
    MsgBox "cidade cadastrada!", vbInformation
    txtbxDesc.Text = ""
    txtbxDesc.SetFocus
End Sub
Private Sub Form_Load()
    P_PreencheUF cmbxUf
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsLoad("frmCadCid") Then
        'frmCadCid.dt.Refresh
    End If
End Sub



Private Sub txtbxDesc_GotFocus()
    P_SelecTexto txtbxDesc
End Sub

Private Sub txtbxDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesc_LostFocus()
    P_AjustaTexto txtbxDesc
End Sub
