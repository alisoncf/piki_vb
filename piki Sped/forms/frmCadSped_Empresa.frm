VERSION 5.00
Begin VB.Form frmCadsped_empresa 
   AutoRedraw      =   -1  'True
   Caption         =   "Configuração de Empresa SPED"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   Icon            =   "frmCadSped_Empresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCRT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8280
      TabIndex        =   19
      Text            =   "1"
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1785
      ScaleWidth      =   8865
      TabIndex        =   58
      Top             =   4680
      Width           =   8895
      Begin VB.TextBox txtCaminhoQrCodeDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   960
         Width           =   5415
      End
      Begin VB.PictureBox pcQrCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   64
         Top             =   960
         Width           =   615
      End
      Begin VB.PictureBox pcCor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   63
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtCaminhoLogo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   5415
      End
      Begin VB.CommandButton cmdLogo 
         Caption         =   "..."
         Height          =   285
         Left            =   5520
         TabIndex        =   60
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdQrCodeLogo 
         Caption         =   "..."
         Height          =   285
         Left            =   5520
         TabIndex        =   66
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QR Code para Pagamento no Carnês:"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   720
         Width           =   2685
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QRCode"
         Height          =   195
         Left            =   5880
         TabIndex        =   67
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cor"
         Height          =   195
         Left            =   5880
         TabIndex        =   62
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Logo da Empresa:"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Width           =   1290
      End
   End
   Begin VB.PictureBox pcNfe 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   8895
      TabIndex        =   55
      Top             =   3240
      Width           =   8895
      Begin VB.CommandButton cmdCertificado 
         Caption         =   "Setar Certificado"
         Height          =   375
         Left            =   6000
         TabIndex        =   57
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtToken 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   52
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtSenha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         TabIndex        =   48
         Top             =   235
         Width           =   2775
      End
      Begin VB.TextBox txtIdCSC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtCaminhoCertificado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   5415
      End
      Begin VB.CommandButton cmdCaminhoCertificado 
         Caption         =   "..."
         Height          =   285
         Left            =   5520
         TabIndex        =   56
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Token"
         Height          =   195
         Left            =   960
         TabIndex        =   51
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha:"
         Height          =   195
         Left            =   5880
         TabIndex        =   47
         Top             =   0
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id CSC"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caminho Completo do Certificado"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   0
         Width           =   2340
      End
   End
   Begin VB.PictureBox pcTool 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9285
      TabIndex        =   44
      Top             =   6840
      Width           =   9285
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   735
         Left            =   7200
         Picture         =   "frmCadSped_Empresa.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Gravar"
         Height          =   735
         Left            =   5280
         Picture         =   "frmCadSped_Empresa.frx":1100C
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.TextBox txtcodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   280
      Width           =   735
   End
   Begin VB.TextBox txtcnpj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtapelido 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox txtcpf 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   11
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtUF 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   21
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtIE 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      TabIndex        =   13
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtnome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   285
      Width           =   4815
   End
   Begin VB.TextBox txtcod_mun 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   23
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtim 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   25
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtsuframa 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   27
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtind_perfil 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtind_atv 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7560
      TabIndex        =   17
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtfantasia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   29
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txtcep 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   31
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtendereco 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      TabIndex        =   33
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox txtnumero 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      TabIndex        =   35
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtcomplemento 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   37
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtbairro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   39
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtfone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      TabIndex        =   41
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtfax 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   43
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtemail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   1485
      Width           =   4095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CRT"
      Height          =   195
      Left            =   8160
      TabIndex        =   18
      ToolTipText     =   "Código do Regimento Tributário"
      Top             =   600
      Width           =   330
   End
   Begin VB.Label lblcodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "codigo"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   45
      Width           =   630
   End
   Begin VB.Label lblcnpj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cnpj"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   645
      Width           =   630
   End
   Begin VB.Label lblapelido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "apelido"
      Height          =   195
      Left            =   6000
      TabIndex        =   4
      Top             =   45
      Width           =   630
   End
   Begin VB.Label lblcpf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF"
      Height          =   195
      Left            =   3000
      TabIndex        =   10
      Top             =   645
      Width           =   300
   End
   Begin VB.Label lblUF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lblIE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IE"
      Height          =   195
      Left            =   5040
      TabIndex        =   12
      Top             =   645
      Width           =   630
   End
   Begin VB.Label lblnome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nome"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   45
      Width           =   630
   End
   Begin VB.Label lblcod_mun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cod_mun"
      Height          =   195
      Left            =   840
      TabIndex        =   22
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lblim 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "im"
      Height          =   195
      Left            =   2040
      TabIndex        =   24
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lblsuframa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "suframa"
      Height          =   195
      Left            =   3720
      TabIndex        =   26
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lblind_perfil 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ind_perfil"
      Height          =   195
      Left            =   6840
      TabIndex        =   14
      Top             =   600
      Width           =   630
   End
   Begin VB.Label lblind_atv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ind_atv"
      Height          =   195
      Left            =   7560
      TabIndex        =   16
      Top             =   600
      Width           =   630
   End
   Begin VB.Label lblfantasia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fantasia"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label lblcep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cep"
      Height          =   195
      Left            =   3120
      TabIndex        =   30
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label lblendereco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "endereco"
      Height          =   195
      Left            =   4320
      TabIndex        =   32
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label lblnumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "numero"
      Height          =   195
      Left            =   7920
      TabIndex        =   34
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label lblcomplemento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "complemento"
      Height          =   195
      Left            =   240
      TabIndex        =   36
      Top             =   2400
      Width           =   630
   End
   Begin VB.Label lblbairro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bairro"
      Height          =   195
      Left            =   3120
      TabIndex        =   38
      Top             =   2445
      Width           =   630
   End
   Begin VB.Label lblfone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fone"
      Height          =   195
      Left            =   6000
      TabIndex        =   40
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label lblfax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fax"
      Height          =   195
      Left            =   7320
      TabIndex        =   42
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label lblemail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "email"
      Height          =   195
      Left            =   4680
      TabIndex        =   6
      Top             =   1245
      Width           =   630
   End
End
Attribute VB_Name = "frmCadsped_empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCaminhoCertificado_Click()
    txtCaminhoCertificado.Text = BuscaNomeArquivoExterno
End Sub

Private Sub cmdCertificado_Click()
    On Error Resume Next
    setarCertificado txtcodigo.Text
End Sub

Private Sub cmdLogo_Click()
    txtCaminhoLogo.Text = BuscaNomeArquivoExterno
End Sub

Private Sub cmdQrCodeLogo_Click()
    txtCaminhoQrCodeDoc.Text = BuscaNomeArquivoExterno
    buscaQrCode
End Sub

Private Sub cmdSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub
Sub buscaQrCode()
On Error Resume Next
    pcQrCode.Picture = LoadPicture(txtCaminhoQrCodeDoc.Text)
    pcQrCode.AutoSize = True
    pcQrCode.PaintPicture pcQrCode.Picture, 0, 0, 100, 100
End Sub
    
Private Sub cmdSalvar_Click()
    
    
    If Me.tag = "" Then
        grava
        
    Else
        altera
        
    End If
     
End Sub
Sub grava()
    Dim strSQL As String
    
    sped_empresa.gravasped_empresa txtcodigo.Text, txtcnpj.Text, txtapelido.Text, txtcpf.Text, txtUF.Text, txtIE.Text, _
        txtnome.Text, txtcod_mun.Text, txtim.Text, txtsuframa.Text, txtind_perfil.Text, txtind_atv.Text, _
        txtfantasia.Text, txtcep.Text, txtendereco.Text, txtnumero.Text, _
        txtcomplemento.Text, txtbairro.Text, txtfone.Text, txtfax.Text, txtemail.Text, _
        txtCaminhoCertificado.Text, txtSenha.Text, txtIdCSC.Text, txtToken.Text, txtCaminhoLogo.Text, txtCaminhoQrCodeDoc.Text, txtCRT.Text, pcCor.BackColor
        MsgBox "Gravado", vbInformation
        
        Me.tag = txtcodigo.Text
End Sub
Sub altera()
    Dim strSQL As String
    
    sped_empresa.Alterasped_empresa txtcodigo.Text, txtcnpj.Text, txtapelido.Text, txtcpf.Text, txtUF.Text, txtIE.Text, _
        txtnome.Text, txtcod_mun.Text, txtim.Text, txtsuframa.Text, txtind_perfil.Text, txtind_atv.Text, _
        txtfantasia.Text, txtcep.Text, txtendereco.Text, txtnumero.Text, _
        txtcomplemento.Text, txtbairro.Text, txtfone.Text, txtfax.Text, txtemail.Text, txtCaminhoCertificado.Text, txtSenha.Text, _
        txtIdCSC.Text, txtToken.Text, txtCaminhoLogo.Text, txtCaminhoQrCodeDoc.Text, txtCRT.Text, pcCor.BackColor
        
        MsgBox "Alterado", vbInformation
    
End Sub


Private Sub Form_Activate()
    If Me.tag <> "" Then
        carrega
    End If
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from sped_empresa where codigo='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            txtapelido.Text = !apelido
            txtbairro.Text = TiraNull(!bairro)
            txtcep.Text = TiraNull(!cep)
            txtcnpj.Text = !CNPJ
            txtcod_mun.Text = TiraNull(!COD_MUN)
            txtcodigo.Text = TiraNull(!codigo)
            txtcomplemento.Text = TiraNull(!complemento)
            txtcpf.Text = TiraNull(!cpf)
            txtemail.Text = TiraNull(!email)
            txtendereco.Text = TiraNull(!Endereco)
            txtfantasia.Text = TiraNull(!fantasia)
            txtfax.Text = TiraNull(!fax)
            txtfone.Text = TiraNull(!fone)
            txtIE.Text = TiraNull(!IE)
            txtim.Text = TiraNull(!IM)
            txtind_atv.Text = TiraNull(!ind_atv)
            txtind_perfil = TiraNull(!ind_perfil)
            txtnome.Text = !nome
            txtnumero.Text = TiraNull(!numero)
            txtsuframa.Text = TiraNull(!SUFRAMA)
            txtUF.Text = TiraNull(!UF)
            txtcodigo.Enabled = False
            txtCaminhoCertificado.Text = UnAjustaCaminho(TiraNull(!caminho_certificado))
            txtCaminhoLogo.Text = UnAjustaCaminho(TiraNull(!caminho_logo))
            txtCaminhoQrCodeDoc.Text = UnAjustaCaminho(TiraNull(!caminho_qrcodedoc))
            txtSenha.Text = TiraNull(!senha_certificado)
            txtIdCSC.Text = TiraNull(!idCSC)
            txtToken.Text = TiraNull(!token)
            txtCRT.Text = !crt
            If Not IsNull(!cor) Then
                pcCor.BackColor = TiraNull(!cor)
            End If
            buscaQrCode
        End If
        .Close
    End With
End Sub

Private Sub pcCor_Click()
    mdiFrm.cmndlg.ShowColor
    
    
    cor = mdiFrm.cmndlg.color
    
    pcCor.BackColor = cor
    
End Sub

Private Sub txtCaminhoLogo_DblClick()
    On Error GoTo trata_Erro:
    Shell "explorer " & txtCaminhoLogo.Text, vbNormalFocus
    
    
    Exit Sub
trata_Erro:
    
    
    
End Sub

Private Sub txtcodigo_GotFocus()
 GanhaFoco txtcodigo
End Sub
Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtCodigo_LostFocus()
 AjustaTexto txtcodigo
End Sub
Private Sub txtCNPJ_GotFocus()
 GanhaFoco txtcnpj
End Sub
Private Sub txtCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtCNPJ_LostFocus()
 AjustaTexto txtcnpj
End Sub
Private Sub txtapelido_GotFocus()
 GanhaFoco txtapelido
End Sub
Private Sub txtapelido_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtapelido_LostFocus()
 AjustaTexto txtapelido
End Sub
Private Sub txtCPF_GotFocus()
 GanhaFoco txtcpf
End Sub
Private Sub txtCPF_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtCPF_LostFocus()
 AjustaTexto txtcpf
End Sub

Private Sub txtUF_GotFocus()
 GanhaFoco txtUF
End Sub
Private Sub txtUF_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtUF_LostFocus()
 AjustaTexto txtUF
End Sub
Private Sub txtIE_GotFocus()
 GanhaFoco txtIE
End Sub
Private Sub txtIE_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtIE_LostFocus()
 AjustaTexto txtIE
End Sub
Private Sub txtNome_GotFocus()
 GanhaFoco txtnome
End Sub
Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtNome_LostFocus()
 AjustaTexto txtnome
End Sub
Private Sub txtcod_mun_GotFocus()
 GanhaFoco txtcod_mun
End Sub
Private Sub txtcod_mun_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtcod_mun_LostFocus()
 AjustaTexto txtcod_mun
End Sub
Private Sub txtim_GotFocus()
 GanhaFoco txtim
End Sub
Private Sub txtim_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtim_LostFocus()
 AjustaTexto txtim
End Sub
Private Sub txtsuframa_GotFocus()
 GanhaFoco txtsuframa
End Sub
Private Sub txtsuframa_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtsuframa_LostFocus()
 AjustaTexto txtsuframa
End Sub
Private Sub txtind_perfil_GotFocus()
 GanhaFoco txtind_perfil
End Sub
Private Sub txtind_perfil_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtind_perfil_LostFocus()
 AjustaTexto txtind_perfil
End Sub
Private Sub txtind_atv_GotFocus()
 GanhaFoco txtind_atv
End Sub
Private Sub txtind_atv_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtind_atv_LostFocus()
 AjustaTexto txtind_atv
End Sub
Private Sub txtfantasia_GotFocus()
 GanhaFoco txtfantasia
End Sub
Private Sub txtfantasia_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtfantasia_LostFocus()
 AjustaTexto txtfantasia
End Sub
Private Sub txtceganhafoco()
 GanhaFoco txtcep
End Sub
Private Sub txtcep_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtcep_LostFocus()
 AjustaTexto txtcep
End Sub
Private Sub txtendereco_GotFocus()
 GanhaFoco txtendereco
End Sub
Private Sub txtendereco_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtendereco_LostFocus()
 AjustaTexto txtendereco
End Sub
Private Sub txtNumero_GotFocus()
 GanhaFoco txtnumero
End Sub
Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtNumero_LostFocus()
 AjustaTexto txtnumero
End Sub
Private Sub txtComplemento_GotFocus()
 GanhaFoco txtcomplemento
End Sub
Private Sub txtComplemento_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtComplemento_LostFocus()
 AjustaTexto txtcomplemento
End Sub
Private Sub txtbairro_GotFocus()
 GanhaFoco txtbairro
End Sub
Private Sub txtbairro_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtbairro_LostFocus()
 AjustaTexto txtbairro
End Sub
Private Sub txtfone_GotFocus()
 GanhaFoco txtfone
End Sub
Private Sub txtfone_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtfone_LostFocus()
 AjustaTexto txtfone
End Sub
Private Sub txtfax_GotFocus()
 GanhaFoco txtfax
End Sub
Private Sub txtfax_Keydown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtfax_LostFocus()
 AjustaTexto txtfax
End Sub
Private Sub txtEmail_GotFocus()
 GanhaFoco txtemail
End Sub
Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Sendkeys Chr(9)
End Sub
Private Sub txtEmail_LostFocus()
 AjustaTexto txtemail
End Sub

