VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmClienteInternet 
   Caption         =   "Dados do Cliente na Internet"
   ClientHeight    =   6990
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10140
   Icon            =   "frmClienteInternet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSincronizar 
      Caption         =   "Sincronizar"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8070
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtRetorno 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   5520
      Width           =   9855
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Pesquisar"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtCPF 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "frmClienteInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As cliente
Dim cInternet As cliente
Dim ret As String
Private Sub cmdPesquisar_Click()
    Dim ret As String
    grade.Visible = False
    ajustaGrid
    
    ret = APIPesquisarCliente(txtCPF.Text)
    If ret <> "" Then
        If Left(ret, 2) = "<!" Then
            MsgBox "Erro" & Chr(13) & ret, vbExclamation
            Exit Sub
        End If
        txtRetorno.Text = ret
        extrai_resultado ret
        carregaCliInternet ret
    End If
    ZebraGrid grade, 1
    grade.Visible = True
    
End Sub
Sub carregaCliInternet(ret As String)
    cInternet.cod_cli = JSONUtil.GetValorJson("id_cli", ret)
    cInternet.nome_cli = JSONUtil.GetValorJson("id_cli", ret)
    cInternet.datanasc_cli = JSONUtil.GetValorJson("data_nasc", ret)
    cInternet.cpf_cli = JSONUtil.GetValorJson("cpf", ret)
    cInternet.cnpj_Cli = JSONUtil.GetValorJson("cpf", ret)
    cInternet.endereco_cli = JSONUtil.GetValorJson("endereco", ret)
    cInternet.Logradouro = JSONUtil.GetValorJson("logradouro", ret)
    cInternet.bairro_cli = JSONUtil.GetValorJson("bairro", ret)
    cInternet.cep_cli = JSONUtil.GetValorJson("cep", ret)
    cInternet.cidade_cli = JSONUtil.GetValorJson("cidade", ret)
    cInternet.uf_cli = JSONUtil.GetValorJson("uf", ret)
    cInternet.mae_cli = Chr(9) & JSONUtil.GetValorJson("mae", ret)
    cInternet.pai_cli = JSONUtil.GetValorJson("pai", ret)
    cInternet.dataalt_cli = JSONUtil.GetValorJson("dataalt_cli", ret)
    
    
    
End Sub
Sub extrai_resultado(ret As String)
    ret = AspasSimplesParaAspasDuplas(ret)
    Dim s As String
    
    c = F_BuscaClienteCodigo(Me.tag)

    
    
    grade.rows = 1
    grade.AddItem "id" & Chr(9) & JSONUtil.GetValorJson("id_cli", ret) & Chr(9) & c.cod_cli
    grade.AddItem "Nome" & Chr(9) & JSONUtil.GetValorJson("nome", ret) & Chr(9) & c.nome_cli
    grade.AddItem "Nascimento" & Chr(9) & JSONUtil.GetValorJson("data_nasc", ret) & Chr(9) & c.nasc_cli
    grade.AddItem "CPF" & Chr(9) & JSONUtil.GetValorJson("cpf", ret) & Chr(9) & c.cpf_cli & c.cnpj_Cli
    grade.AddItem "Endereço" & Chr(9) & JSONUtil.GetValorJson("endereco", ret) & Chr(9) & c.Logradouro
    grade.AddItem "Bairro" & Chr(9) & JSONUtil.GetValorJson("bairro", ret) & Chr(9) & c.bairro_cli
    grade.AddItem "CEP" & Chr(9) & JSONUtil.GetValorJson("cep", ret) & Chr(9) & c.cep_cli
    grade.AddItem "Cidade" & Chr(9) & JSONUtil.GetValorJson("cidade", ret) & Chr(9) & c.set.Cidade
    grade.AddItem "UF" & Chr(9) & JSONUtil.GetValorJson("uf", ret) & Chr(9) & c.set.UF
    grade.AddItem "Nome da Mãe" & Chr(9) & JSONUtil.GetValorJson("mae", ret) & Chr(9) & c.mae_cli
    grade.AddItem "Nome do Pai" & Chr(9) & JSONUtil.GetValorJson("pai", ret) & Chr(9) & c.pai_cli
    grade.AddItem "Data" & Chr(9) & JSONUtil.GetValorJson("dataalt_cli", ret) & Chr(9) & c.dataalt_cli
    grade.ColAlignment(1) = 0
    grade.ColAlignment(2) = 0
End Sub

Private Sub cmdSincronizar_Click()
    Dim ret As String
    If cInternet.cod_cli = "" And c.cod_cli <> "" Then
        ret = APIInserirCliente(c.cod_cli)
        MsgBox ret
    ElseIf cInternet.cod_cli <> "" And c.cod_cli <> "" Then
        If cInternet.dataalt_cli = "0000-00-00" Then
            
        End If
        If cInternet.dataalt_cli = "0000-00-00" Then
            ret = APIAtualizarCliente(c.cod_cli, cInternet.cod_cli)
            MsgBox ret
        ElseIf (CDate(cInternet.dataalt_cli) <= CDate(c.dataalt_cli)) Then
            ret = APIAtualizarCliente(c.cod_cli, cInternet.cod_cli)
            MsgBox ret
        End If
    
    
    End If
End Sub

Private Sub txtCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdPesquisar_Click
    End If
End Sub
Sub ajustaGrid()
    grade.rows = 0
    grade.Cols = 3
    grade.ColWidth(1) = 3500
    grade.ColWidth(2) = 3500
    grade.ColAlignment(1) = 0
    grade.ColAlignment(2) = 0
    
    grade.AllowUserResizing = flexResizeColumns
    grade.AddItem "Atributo" & Chr(9) & "On-line" & Chr(9) & "Local"
    
    
End Sub
