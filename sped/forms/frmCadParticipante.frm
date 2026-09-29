VERSION 5.00
Begin VB.Form frmCadParticipante 
   Caption         =   "sped_participante"
   ClientHeight    =   6480
   ClientLeft      =   1170
   ClientTop       =   450
   ClientWidth     =   7245
   Icon            =   "frmCadParticipante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   7245
   Begin VB.TextBox txtUF 
      DataField       =   "NUMERO"
      Height          =   285
      Left            =   3360
      TabIndex        =   34
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtCep 
      DataField       =   "ENDERECO"
      Height          =   285
      Left            =   5520
      TabIndex        =   25
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   7245
      TabIndex        =   30
      Top             =   5385
      Width           =   7245
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   4920
         Picture         =   "frmCadParticipante.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Gravar"
         Height          =   735
         Left            =   2880
         Picture         =   "frmCadParticipante.frx":0E58
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.TextBox txtFornecedor 
      BackColor       =   &H00FFC0C0&
      DataField       =   "cod_for"
      Height          =   285
      Left            =   240
      TabIndex        =   29
      Top             =   4800
      Width           =   6495
   End
   Begin VB.TextBox txtCliente 
      BackColor       =   &H0080FFFF&
      DataField       =   "cod_cli"
      Height          =   285
      Left            =   240
      TabIndex        =   27
      Top             =   4200
      Width           =   6495
   End
   Begin VB.TextBox txtBairro 
      DataField       =   "BAIRRO"
      Height          =   285
      Left            =   240
      TabIndex        =   23
      Top             =   3240
      Width           =   5175
   End
   Begin VB.TextBox txtSuframa 
      DataField       =   "COMPLEMENTO"
      Height          =   285
      Left            =   5520
      TabIndex        =   15
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtCodMun 
      DataField       =   "NUMERO"
      Height          =   285
      Left            =   4320
      TabIndex        =   19
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtCodPais 
      DataField       =   "ENDERECO"
      Height          =   285
      Left            =   5520
      TabIndex        =   21
      Text            =   "1058"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtComplemento 
      DataField       =   "SUFRAMA"
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txtNumero 
      DataField       =   "COD_MUN"
      Height          =   285
      Left            =   4320
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtEndereco 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox txtCpf 
      DataField       =   "CPF"
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtCnpj 
      DataField       =   "CNPJ"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtIe 
      DataField       =   "COD_PAIS"
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtNome 
      DataField       =   "NOME"
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   6495
   End
   Begin VB.TextBox txtCod 
      DataField       =   "codigo"
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   300
      Width           =   975
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "UF"
      Height          =   195
      Index           =   15
      Left            =   3360
      TabIndex        =   35
      Top             =   2400
      Width           =   210
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "CEP:"
      Height          =   195
      Index           =   14
      Left            =   5520
      TabIndex        =   24
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label lblCidade 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   33
      Top             =   3000
      Width           =   600
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   28
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   26
      Top             =   3960
      Width           =   525
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "BAIRRO:"
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   22
      Top             =   3000
      Width           =   660
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "COMPLEMENTO:"
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   16
      Top             =   2400
      Width           =   1290
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "NUMERO:"
      Height          =   195
      Index           =   9
      Left            =   4320
      TabIndex        =   12
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "ENDERECO:"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   10
      Top             =   1785
      Width           =   945
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "SUFRAMA:"
      Height          =   195
      Index           =   7
      Left            =   5520
      TabIndex        =   14
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Cod Mun."
      Height          =   195
      Index           =   6
      Left            =   4320
      TabIndex        =   18
      Top             =   2400
      Width           =   690
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "IE:"
      Height          =   195
      Index           =   5
      Left            =   4320
      TabIndex        =   8
      Top             =   1185
      Width           =   195
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "CPF:"
      Height          =   195
      Index           =   4
      Left            =   2640
      TabIndex        =   6
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ:"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Cod. País:"
      Height          =   195
      Index           =   2
      Left            =   5520
      TabIndex        =   20
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "NOME:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   615
      Width           =   525
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
End
Attribute VB_Name = "frmCadParticipante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Sub grava()
    If F_BuscaUmCampo("codigo", "sped_participante", "cnpj='" & txtCNPJ.Text & "'") <> "" Then
        MsgBox "já existe um participante cadastrado com este cnpj!", vbExclamation
        Exit Sub
    End If
    
    If txtCNPJ.Text = "" And txtCPF.Text = "" Then
        MsgBox "Informe o cnpj ou cpf!", vbExclamation
        Exit Sub
    End If
    If txtNome.Text = "" Then
        MsgBox "Informe o nome!", vbExclamation
        Exit Sub
    End If
    
    If SpedParticipante.GravaParticipante(txtNome.Text, txtCodPais.Text, txtCNPJ.Text, txtCPF.Text, _
        txtIE.Text, txtCodMun.Text, txtsuframa.Text, txtendereco.Text, txtNumero.Text, txtcomplemento.Text, _
        txtbairro.Text, txtCliente.tag, txtFornecedor.tag, txtcep.Text, txtUF.Text) Then
        MsgBox "Participante gravado com sucesso!", vbInformation
        
        txtNome.Text = ""
        txtCNPJ.Text = ""
        txtNome.SetFocus
        
        
    End If
    
    
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select nome_for,nome_cli,p.* from sped_participante as p "
    strSQL = strSQL & " left join " & TAB_CLIENTES & " as tc on tc.cod_cli=p.cod_cli and tc.cod_cli <> 0 "
    strSQL = strSQL & " left join " & TAB_FORNECEDORES & " as tf on tf.cod_for=p.cod_for "
    strSQL = strSQL & " where codigo=" & Me.tag & ""
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
        txtUF.Text = TiraNull(!UF)
            txtbairro.Text = TiraNull(!bairro)
            txtCNPJ.Text = TiraNull(!CNPJ)
            txtIE.Text = TiraNull(!IE)
            txtCodMun.Text = TiraNull(!cod_mun)
            txtCodPais.Text = TiraNull(!COD_PAIS)
            txtcomplemento.Text = TiraNull(!complemento)
            txtendereco.Text = TiraNull(!Endereco)
            txtIE.Text = TiraNull(!IE)
            txtNome.Text = TiraNull(!nome)
            txtNumero.Text = TiraNull(!numero)
            txtsuframa.Text = TiraNull(!SUFRAMA)
            txtFornecedor.Text = TiraNull(!nome_for)
            txtFornecedor.tag = TiraNull(!cod_For)
            txtCliente.Text = TiraNull(!nome_cli)
            txtCliente.tag = TiraNull(!cod_cli)
            txtcep.Text = TiraNull(!cep)
            txtCod.Text = TiraNull(!codigo)
            txtCod.Enabled = False
        Else
        
        End If
        .Close
    End With
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub altera()
    If F_BuscaUmCampo("codigo", "sped_participante", "cnpj='" & txtCNPJ.Text & "' and id<>" & Me.tag) <> "" Then
        MsgBox "já existe um participante cadastrado com este cnpj!", vbExclamation
        Exit Sub
    End If
    
    If txtCNPJ.Text = "" And txtCPF.Text = "" Then
        MsgBox "Informe o cnpj ou cpf!", vbExclamation
        Exit Sub
    End If
    If txtNome.Text = "" Then
        MsgBox "Informe o nome!", vbExclamation
        Exit Sub
    End If
    
    If SpedParticipante.AlteraParticipante(Me.tag, txtNome.Text, txtCodPais.Text, txtCNPJ.Text, txtCPF.Text, _
        txtIE.Text, txtCodMun.Text, txtsuframa.Text, txtendereco.Text, txtNumero.Text, txtcomplemento.Text, _
        txtbairro.Text, txtCliente.tag, txtFornecedor.tag, txtcep.Text, txtUF.Text) Then
        MsgBox "Participante alterado com sucesso!", vbInformation
        
    
        
    End If
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    txtCodMun_LostFocus
    
    If Me.tag <> "" Then
        carrega
        txtCod.Visible = False
        lblLabels(0).Visible = False
    End If
    
    
    bActive = True
    
    
    
End Sub

Private Sub Form_Load()
    bActive = False
End Sub

Private Sub txtbairro_GotFocus()
    GanhaFoco txtbairro
End Sub

Private Sub txtbairro_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbairro_LostFocus()
    AjustaTexto txtbairro
End Sub

Private Sub txtceganhafoco()
    GanhaFoco txtcep
End Sub

Private Sub txtcep_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtcep_LostFocus()
    AjustaTexto txtcep
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCli
    End If
End Sub
Sub buscaCli()
    Dim c As Clientes.cliente
    c = Clientes.F_BuscaCliente(txtCliente.Text)
    If c.cod_cli <> "" Then
        txtCliente.Text = c.nome_cli
        txtCliente.tag = c.cod_cli
        txtFornecedor.SetFocus
    End If
End Sub

Sub buscaFor()
    Dim c As fornecedor.fornecedor
    c = fornecedor.F_BuscaFornecedor(txtFornecedor.Text)
    If c.cod_For <> "" Then
        txtFornecedor.Text = c.nome_for
        txtFornecedor.tag = c.cod_For
        cmdSalvar.SetFocus
    End If
End Sub
Private Sub txtCliente_LostFocus()
    AjustaTexto txtCliente
End Sub

Private Sub txtCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCNPJ_LostFocus()
    AjustaTexto txtCNPJ
End Sub

Private Sub txtCod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCodMun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCid
    End If
End Sub
Sub buscaCid()
    Dim c As Cidade
    c = F_BuscaCidades(txtCodMun.Text)
    If c.cod_cid <> "" Then
        txtCodMun.Text = c.cod_cid
        txtCodPais.SetFocus
    End If
End Sub
Private Sub txtCodMun_LostFocus()
    AjustaTexto txtCodMun
    lblCidade.caption = F_BuscaUmCampo("concat(nome_cid,'/' ,uf_cid)", TAB_CIDADES, "cod_cid='" & txtCodMun.Text & "'")
    
End Sub

Private Sub txtCodPais_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCodPais_LostFocus()
    AjustaTexto txtCodPais
End Sub

Private Sub txtComplemento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtComplemento_LinkOpen(Cancel As Integer)
    AjustaTexto txtcomplemento
End Sub

Private Sub txtCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCPF_LostFocus()
    AjustaTexto txtCPF
End Sub

Private Sub txtendereco_GotFocus()
    GanhaFoco txtendereco
End Sub

Private Sub txtendereco_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtendereco_LostFocus()
    AjustaTexto txtendereco
End Sub

Private Sub txtFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFor
    End If
End Sub

Private Sub txtIE_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtIE_LostFocus()
    AjustaTexto txtIE
End Sub

Private Sub txtNome_GotFocus()
    GanhaFoco txtNome
End Sub
Private Sub txtCNPJ_GotFocus()
    GanhaFoco txtCNPJ
End Sub
Private Sub txtIE_GotFocus()
    GanhaFoco txtIE
End Sub
Private Sub txtCPF_GotFocus()
    GanhaFoco txtCPF
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtNome_LostFocus()
    AjustaTexto txtNome
End Sub

Private Sub txtNumero_GotFocus()
    GanhaFoco txtNumero
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtNumero_LostFocus()
    AjustaTexto txtNumero
End Sub

Private Sub txtsuframa_GotFocus()
    GanhaFoco txtsuframa
End Sub
Private Sub txtComplemento_GotFocus()
    GanhaFoco txtcomplemento
End Sub
Private Sub txtCodMun_GotFocus()
    GanhaFoco txtCodMun
End Sub
Private Sub txtCodPais_GotFocus()
    GanhaFoco txtCodPais
End Sub
Private Sub txtCliente_GotFocus()
    GanhaFoco txtCliente
End Sub
Private Sub txtFornecedor_GotFocus()
    GanhaFoco txtFornecedor
End Sub







Private Sub txtsuframa_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtsuframa_LostFocus()
    AjustaTexto txtsuframa
End Sub
