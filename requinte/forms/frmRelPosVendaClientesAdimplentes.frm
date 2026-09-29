VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelPosVendaClientesAdimplentes 
   Caption         =   "Relatório de Pós-venda - Clientes Adimplentes de Crediário"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pCtop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   13275
      TabIndex        =   2
      Top             =   0
      Width           =   13305
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         Height          =   495
         Left            =   8880
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   495
         Left            =   11520
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtFun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   10200
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         ScaleHeight     =   1065
         ScaleWidth      =   2505
         TabIndex        =   3
         Top             =   0
         Width           =   2535
         Begin VB.TextBox txtAte 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtDe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "e:"
            Height          =   195
            Left            =   1200
            TabIndex        =   8
            Top             =   120
            Width           =   135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entre:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   420
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Quando a venda foi gerada?"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   2040
         End
      End
      Begin VB.CommandButton cmnbuscarFun 
         Height          =   285
         Left            =   4320
         Picture         =   "frmRelPosVendaClientesAdimplentes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdCidade 
         Height          =   285
         Left            =   6480
         Picture         =   "frmRelPosVendaClientesAdimplentes.frx":03AB
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   2640
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   4680
         TabIndex        =   18
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do cliente"
         Height          =   195
         Left            =   6840
         TabIndex        =   17
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   13305
      TabIndex        =   0
      Top             =   6285
      Width           =   13305
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   2143
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnCopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu mnEditar 
         Caption         =   "Editar Cliente"
      End
   End
End
Attribute VB_Name = "frmRelPosVendaClientesAdimplentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdCidade_Click()
    Dim c As Cidades.Cidade
    
    c = Cidades.F_BuscaCidades(txtCidade.Text)
    If c.cod_cid <> "" Then
        txtCidade.Text = c.nome_cid
        txtCidade.tag = c.cod_cid
    End If
    
    
End Sub

Private Sub cmdFiltrar_Click()
    
End Sub

Private Sub cmdImprimir_Click()
    P_MsflxToHtml "Relatório de Pós Venda", grid, "De " & txtDe.Text & " até " & txtAte.Text, "Total: " & grid.rows - 1 & " Clientes"
End Sub

Private Sub cmdPesquisar_Click()
    Dim strSQL As String
    Dim codigo As String
    Dim sNotIn As String
    If txtDe.Text = "" Then
        MsgBox "Informe a data inicial do documento que foi gerado.", vbExclamation
        txtDe.SetFocus
        Exit Sub
    End If
    If txtAte.Text = "" Then
        MsgBox "Informe a data final do documento que foi gerado.", vbExclamation
        txtAte.SetFocus
        Exit Sub
    End If
    P_Status "Verificando todos os débitos... aguarde..."
    'sNotIn = F_NotIn("distinct td.cod_ter", "tab_doc as td, tab_itdoc as ti", "td.seq_doc=ti.seq_doc and reman_itd>0")
    
    'P_Status "econtrados... Buscando clientes. Aguarde..."
    
    strSQL = "select tc.cod_cli,tcid.nome_cid,tc.nome_cli,"
    strSQL = strSQL & "(select group_concat(num_tel) FROM tab_telefones WHERE cod_cli=tc.cod_cli) AS num_tel,"
    strSQL = strSQL & " count(td.seq_doc) as compras, (select min(data_doc) from tab_doc where "
    strSQL = strSQL & " cod_ter=td.cod_ter and tipo_ter='C') as primeira_compra,"
    strSQL = strSQL & "(select max(data_doc) from tab_doc where "
    strSQL = strSQL & " cod_ter=td.cod_ter and tipo_ter='C') as ultima_compra,"
    strSQL = strSQL & " min(venc_itd) as vencimento_minimo, sum(reman_itd) as reman, sum(td.valor_doc) as valor_comprado ,"
    strSQL = strSQL & " (select fun.guerra_fun from tab_funcionarios as fun INNER join tab_vendas as tv on"
    strSQL = strSQL & " tv.abrevenda_usu=fun.cod_fun  where tv.cod_cli=tc.cod_cli ORDER BY num_ven DESC LIMIT 1) as vendedor"
    strSQL = strSQL & " from tab_clientes as tc inner join " & TAB_DOC & " as td on td.cod_ter=tc.cod_cli and td.tipo_ter='C'"
    strSQL = strSQL & " inner join tab_setores as tset on tset.cod_set=tc.cod_set "
    strSQL = strSQL & " inner join tab_cidades as tcid on tcid.cod_cid=tset.cod_cid"
    strSQL = strSQL & " inner join tab_telefones as tt on tt.cod_cli=tc.cod_cli and tt.tipo_ter='C' "
    strSQL = strSQL & " inner join " & TAB_ITDOC & " as ti on ti.seq_doc=td.seq_doc "
  
    
    strSQL = strSQL & " where 1 > 0 "
    
      
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and td.data_doc>=" & F_GravaData(txtDe.Text)
    End If
    
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and td.data_doc<=" & F_GravaData(txtAte.Text)
    End If
    If txtCidade.tag <> "" Then
        strSQL = strSQL & " and tcid.cod_cid='" & txtCidade.tag & "' "
    End If
    If txtFun.tag <> "" Then
       ' strSQL = strSQL & " and tc.cod_cli in ( select cod_cli from tab_vendas as vendas where cod_cli=tc.cod_cli and abrevenda_usu='" & txtFun.tag & "')"
    End If
    
    
    'strSQL = strSQL & " and tc.cod_cli not in " & sNotIn
    strSQL = strSQL & " and esp_doc='V' "
    strSQL = strSQL & " and status_cli=1 "
    strSQL = strSQL & " and tc.cod_cli not in(SELECT d.cod_ter from tab_doc d inner join tab_itdoc i on  "
    strSQL = strSQL & " d.seq_doc=i.seq_doc AND d.tipo_ter='C' AND i.reman_itd>0 AND i.venc_itd<='2024-11-01') "
    strSQL = strSQL & " group by tt.cod_cli "
    'strSQL = strSQL & " having Parcelas<=1 "
    strSQL = strSQL & " order by nome_cli "
    abreObjetoAdo ado, strSQL
    ado.filter = ("vendedor='" & txtFun.Text & "' ")
    If ado.RecordCount = 0 Then
        P_Status ""
        Exit Sub
    End If
    grid.rows = 1
    P_Status "Montando grade"
    
    Dim colunas As String
    Dim campos As String
    grid.Visible = False
    colunas = "item,Codigo,Cliente,Vendedor,Telefones,Cidade, Compras, Total,Resta pagar, Primeira Compra, Última Compra"
    campos = "@id,cod_cli,nome_cli,vendedor,num_tel,nome_cid,compras,valor_comprado,reman,primeira_compra,ultima_compra"
    AjustaGridData grid, 6, 1, 0, colunas, campos, "1000,1000,3500,2500,3500,2400,1200,1200,1200,1200", ado, "", True
    grid.Visible = True
    
    P_Status ""
End Sub
Sub filtrar(vendedor As String)
    Dim i As Integer
    Dim f As Integer
    Dim c As New Collection
    Dim linha As Integer
    Dim flex As MSHFlexGrid
    Set flex = grid
    flex.rows = 1
    With grid
        f = .rows - 1
        .Visible = False
        linha = 1
        For i = 1 To .rows - 1
            
            .row = i
            
            .col = 0
            .Text = i
            
            .col = 3
            If .Text <> vendedor Then
                   
            End If
        Next
        'ado.filter
        .Visible = True
    End With
    
    
    
    
    ZebraGrid grid, 1
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmnbuscarFun_Click()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectVendedor(txtFun.Text)
    If f.cod_fun <> "" Then
        txtFun.Text = f.guerra_fun
        txtFun.tag = f.cod_fun
    End If
    
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    resizeForm Me
    
    
End Sub
Sub ajustaGrid()
    With grid
        .rows = 1
        .Cols = 4
        .col = 0: .Text = "Código"
        .col = 1: .Text = "Cliente"
        .col = 2: .Text = "Telefones"
        .ColWidth(0) = 1000
        .ColWidth(1) = 3500
        .ColWidth(2) = 3500
        .ColWidth(3) = 2500
        .ColAlignment(0) = 0
        .ColAlignment(1) = 0
        .ColAlignment(2) = 0
    End With
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnMenu
    End If
End Sub

Private Sub mnCopiar_Click()
    Clipboard.SetText grid.Text
End Sub

Private Sub mnEditar_Click()
    If grid.rows <= 1 Then Exit Sub
    grid.col = 1
    
    frmCadCli.tag = grid.Text
    frmCadCli.Show 1
    
End Sub

Private Sub txtCidade_Change()
    txtCidade.tag = ""
End Sub

Private Sub txtCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCidade_Click
    End If
End Sub

Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub

Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

