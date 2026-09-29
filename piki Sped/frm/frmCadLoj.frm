VERSION 5.00
Begin VB.Form frmCadLoj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Lojas"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frmCadLoj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   4410
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Loja:"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtbxTel 
         Height          =   285
         Left            =   240
         MaxLength       =   16
         TabIndex        =   14
         Top             =   2880
         Width           =   2055
      End
      Begin VB.ComboBox cmbxUF 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtbxCidade 
         Height          =   285
         Left            =   240
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox txtbxIE 
         Height          =   285
         Left            =   2400
         MaxLength       =   16
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtbxEndereco 
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txtbxCNPJ 
         Height          =   285
         Left            =   240
         MaxLength       =   18
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtbxRazao 
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Telefone:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&UF"
         Height          =   195
         Left            =   3480
         TabIndex        =   11
         Top             =   2040
         Width           =   210
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ci&dade:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label lblRg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&IE:"
         Height          =   195
         Left            =   2400
         TabIndex        =   5
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Endereço:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblCpf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CNPJ:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   15
      Top             =   3480
      Width           =   4335
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   120
         Picture         =   "frmCadLoj.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   2280
         Picture         =   "frmCadLoj.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmCadLoj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkProd_Click()

End Sub

Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxCidade.Text = "" Then
        MsgBox "Informe a Cidade!", vbExclamation
        txtbxCidade.SetFocus
        Exit Sub
    End If
    If txtbxCNPJ.Text = "" Then
        MsgBox "Informe o CNPJ!", vbExclamation
        txtbxCNPJ.SetFocus
        Exit Sub
    End If
    If txtbxEndereco.Text = "" Then
        MsgBox "Informe o Endereço!", vbExclamation
        txtbxEndereco.SetFocus
        Exit Sub
    End If
    If txtbxRazao.Text = "" Then
        MsgBox "Informe o Nome!", vbExclamation
        txtbxRazao.SetFocus
        Exit Sub
    End If
    
    
    If Me.tag = "" Then
        If F_BuscaUmCampo("cnpj_loj", "tab_lojas", "cnpj_loj='" & txtbxCNPJ.Text & "'") <> "" Then
            MsgBox "Já existe uma loja com este CNPJ!", vbExclamation
            Exit Sub
        End If
        
        P_Grava False
    Else
        P_Altera
    End If
End Sub
Sub P_Grava(bProd As Boolean)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCod As String
    P_Status "Aguarde..."
    
    strSQL = "insert into tab_lojas(cod_loj,nome_loj,cnpj_loj,ie_loj,cidade_loj,uf_loj,"
    strSQL = strSQL & "endereco_loj,fone_loj)Values("
    sCod = MaxKeySQL("tab_lojas", "cod_loj", "", "00#")
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxRazao.Text & "',"
    strSQL = strSQL & "'" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "'" & txtbxIE.Text & "',"
    strSQL = strSQL & "'" & txtbxCidade.Text & "',"
    strSQL = strSQL & "'" & cmbxUf.Text & "',"
    strSQL = strSQL & "'" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "'" & txtbxTel.Text & "') "
    conn.Execute strSQL
    
    MsgBox "Loja Gravada!", vbInformation
    If IsLoad("frmEdiLoj") Then
        frmEdiLoj.dtServ.Refresh
    End If
    txtbxCidade.Text = ""
    txtbxCNPJ.Text = ""
    txtbxTel.Text = ""
    txtbxTel.Text = ""
    P_Status ""
    
End Sub
Sub P_Altera()
    Dim strSQL As String
    strSQL = "update tab_lojas set "
    
    strSQL = strSQL & "nome_loj='" & txtbxRazao.Text & "',"
    strSQL = strSQL & "cnpj_loj='" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "ie_loj='" & txtbxIE.Text & "',"
    strSQL = strSQL & "cidade_loj='" & txtbxCidade.Text & "',"
    strSQL = strSQL & "uf_loj='" & cmbxUf.Text & "',"
    strSQL = strSQL & "endereco_loj='" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "fone_loj='" & txtbxTel.Text & "'  "
    strSQL = strSQL & " where cod_loj='" & Me.tag & "' "
    conn.Execute strSQL
    MsgBox "Loja Alterada!", vbInformation
    If IsLoad("frmEdiLoj") Then
        frmEdiLoj.dtServ.Refresh
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_Load()
    P_PreencheUF cmbxUf
End Sub

Private Sub txtbxCidade_GotFocus()
    ganhaFoco txtbxCidade
End Sub

Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCidade_LostFocus()
    AjustaTexto txtbxCidade
End Sub

Private Sub txtbxCNPJ_GotFocus()
    ganhaFoco txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCNPJ_LostFocus()
    P_FormataCGC txtbxCNPJ
End Sub

Private Sub txtbxEndereco_GotFocus()
    ganhaFoco txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxEndereco_LostFocus()
    AjustaTexto txtbxEndereco
End Sub

Private Sub txtbxIE_GotFocus()
    ganhaFoco txtbxIE
End Sub

Private Sub txtbxIE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxIE_LostFocus()
    AjustaTexto txtbxIE
End Sub

Private Sub txtbxRazao_GotFocus()
    ganhaFoco txtbxRazao
End Sub

Private Sub txtbxRazao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxRazao_LostFocus()
    AjustaTexto txtbxRazao
End Sub

Private Sub txtbxTel_GotFocus()
    ganhaFoco txtbxTel
End Sub

Private Sub txtbxTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxTel_LostFocus()
     txtbxTel.Text = F_AjustaTelefone(txtbxTel.Text)
End Sub
