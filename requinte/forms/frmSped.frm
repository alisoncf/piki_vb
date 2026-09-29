VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSped 
   Caption         =   "SPED"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFim 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   14
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtIni 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eventos SPED"
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtCaminho 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "C:\"
      Top             =   5160
      Width           =   3855
   End
   Begin VB.PictureBox pcBlocos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3705
      ScaleWidth      =   7785
      TabIndex        =   4
      Top             =   840
      Width           =   7815
      Begin MSComctlLib.TreeView tree 
         Height          =   2895
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5106
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.CheckBox chkZero 
         Caption         =   "Bloco 0"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox txtMes 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fim:"
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   285
   End
   Begin VB.Label lblIni 
      AutoSize        =   -1  'True
      Caption         =   "Início"
      Height          =   195
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salvar em:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   750
   End
   Begin VB.Label Label2 
      Caption         =   "Mês:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Ano:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmSped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGerar_Click()
    geraSped
End Sub
Sub geraSped()
    Dim s As String
    Dim ado As New ADODB.Recordset
    Dim r As Reg_0
    Dim sCodMun As String
    Dim sCodVer As String
    Dim sEmailEmp As String
    Dim strSQL As String
    sCodMun = ""
    sCodVer = "007"
    
    AbreArquivoExterno txtCaminho.Text, 1
            
    r.CNPJ = F_SoNumero(strCNPJ_Emp)
    r.cod_fin = Remessa_Arquivo_Original
    r.cod_mun = sCodMun
    r.cod_ver = sCodVer
    r.CPF = ""
    r.DT_FIN = txtFim.Text
    r.DT_INI = txtIni.Text
    r.IE = strIE_Emp
    r.im = ""
    r.ind_atv = outros
    r.ind_perfil = "A"
    r.nome = strNome_Emp
    r.suframa = ""
    r.UF = strUF_Emp
    
    s = Registro_0000(r)
    ImprimeEmArquivoExterno s, 1
    s = Registro_0001("0")
    ImprimeEmArquivoExterno s, 1
    s = Registro_0005(strFantasia_Emp, strCEP_Emp, strEndereco_Emp, strNumero_Emp, strComplemento_Emp, strBairro_Emp, _
        strFone_Emp, strFax_Emp, sEmailEmp)
    ImprimeEmArquivoExterno s, 1
    's = Registro_0015 ( ) verificar se é substituto tributário
    
    strSQL = "select * from contabilista"
    
    abreObjetoAdo ado, strSQL
    With ado
        If .RecordCount > 0 Then
            s = Registro_0100(!nome, !CPF, !crc, !CNPJ, !cep, !Endereco, !Num, !complemento, !bairro, !fone, !fax, !email)
            ImprimeEmArquivoExterno s, 1
        Else
            MsgBox "Não foram encontrados registros de contabilistas!", vbExclamation
            
        End If
        FechaObjeto ado
    End With
    
    strSQL = "select tf.* from tab_fornecedores as tf, tab_compras as tc where tf.cod_for=tc.cod_for "
    strSQL = strSQL & " and tc.emissao_not>=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and tc.emissao_not<=" & F_GravaData(txtFim.Text)
    strSQL = strSQL & " order by tf.cod_for  "
    
    
    With ado
        If .RecordCount > 0 Then
            s = Registro_0150("F" & !cod_for, !nome_for, "1058", !cnpj_for, "", !ie_for, !cod_mun, !suframa_for, !endereco_for, !numero_for, !complemento_for, !bairro_for)
            ImprimeEmArquivoExterno s, 1
        Else
            MsgBox "Não foram encontrados registros de fornecedores com compras no período!", vbExclamation
            
        End If
        FechaObjeto ado
    End With
    
    
    strSQL = "select tf.* from tab_clientes as tc, tab_saidasnf as ts where tc.cod_cli=ts.cod_cli "
    strSQL = strSQL & " and tc.emissao_sai>=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and tc.emissao_sai<=" & F_GravaData(txtFim.Text)
    strSQL = strSQL & " order by tf.cod_cli  "
    abreObjetoAdo ado, strSQL
    
    With ado
        If .RecordCount > 0 Then
            s = Registro_0150("C" & !cod_cli, !nome_Cli, "1058", !cnpj_Cli, !cpf_cli, !ie_for, !cod_mun, !suframa_cli, !endereco_cli, !numero_cli, !complemento_cli, !bairro_cli)
            ImprimeEmArquivoExterno s, 1
        Else
            MsgBox "Não foram encontrados registros de fornecedores com compras no período!", vbExclamation
            
        End If
        FechaObjeto ado
    End With
      
     'fazer uma tabela de alteração de clientes e fornecedores registro 0175
    
    strSQL = "select  * from tab_unidades order by cod_uni "
    abreObjetoAdo ado, strSQL
    With ado
        If .RecordCount > 0 Then
            s = Registro_0190(!cod_uni, !descricao_uni)
            ImprimeEmArquivoExterno s, 1
        Else
            MsgBox "Não foram encontrados registros de unidades!", vbExclamation
            
        End If
        FechaObjeto ado
    End With
    
    strSQL = "select tp.* from tab_produtos as tf, tab_compras as tc, tab_itcompras as ti where tf.seq_not=tc.seq_not "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tc.emissao_not>=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and tc.emissao_not<=" & F_GravaData(txtFim.Text)
    strSQL = strSQL & " order by tf.cod_for  "
    abreObjetoAdo ado, strSQL
    With ado
        If .RecordCount > 0 Then
            s = Registro_0200(!cod_pro, !nome_pro, !codbarra_pro, "", !unidade_pro, mercadoria_para_revenda, !ncm_pro, _
                "", !cod_gem, !cod_lst, !icms_pro)
            ImprimeEmArquivoExterno s, 1
        Else
            MsgBox "Não foram encontrados registros de unidades!", vbExclamation
            
        End If
        FechaObjeto ado
    End With
    
    
    
    
    FechaArquivoExterno
    
     
    
End Sub
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    txtAno.Text = Year(Date)
    txtMes.Text = Month(Date)
    txtCaminho.Text = App.Path & "\sped.txt"
    
End Sub

Private Sub txtAno_Change()

End Sub

Private Sub txtAno_GotFocus()
    P_SelecTexto txtAno
End Sub

Private Sub txtAno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtMes_LostFocus()
    txtMes.Text = Format(txtMes.Text, "MM")
    buscadata
End Sub
Sub buscadata()
    On Error GoTo trata_Erro
    txtIni.Text = F_PrimDiaMes("01/" & txtMes.Text & "/" & txtAno.Text)
    txtFim.Text = F_UltDiaMes("01/" & txtMes.Text & "/" & txtAno.Text)
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Private Sub txtMes_GotFocus()
    P_SelecTexto txtMes
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtAno_LostFocus()
    txtAno.Text = Format(txtAno.Text, "YYYY")
End Sub
