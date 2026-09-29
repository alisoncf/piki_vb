VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEdiOrc 
   Caption         =   "Edição de Orçamentos"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   14655
   Icon            =   "frmEdiOrc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   14655
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7223
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   14655
      TabIndex        =   0
      Top             =   6510
      Width           =   14655
      Begin VB.CommandButton cmdDuplicate 
         Caption         =   "Duplicar"
         Height          =   735
         Left            =   3720
         Picture         =   "frmEdiOrc.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   735
         Left            =   9120
         Picture         =   "frmEdiOrc.frx":1CE8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdLista 
         Caption         =   "Gerar Lista"
         Height          =   735
         Left            =   7320
         Picture         =   "frmEdiOrc.frx":21D9
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Emitir Orçamento"
         Height          =   735
         Left            =   5520
         Picture         =   "frmEdiOrc.frx":2687
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   12840
         Picture         =   "frmEdiOrc.frx":2B35
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdEfetivar 
         Caption         =   "Efetivar"
         Height          =   735
         Left            =   10920
         Picture         =   "frmEdiOrc.frx":30C3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Editar"
         Height          =   735
         Left            =   1920
         Picture         =   "frmEdiOrc.frx":354A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiOrc.frx":3A86
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14625
      TabIndex        =   9
      Top             =   5655
      Width           =   14655
      Begin VB.ComboBox cmbxSituacao 
         Height          =   315
         ItemData        =   "frmEdiOrc.frx":3E51
         Left            =   6000
         List            =   "frmEdiOrc.frx":3E61
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtBusca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         Height          =   615
         Left            =   7680
         Picture         =   "frmEdiOrc.frx":3E9B
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situacao:"
         Height          =   195
         Left            =   6000
         TabIndex        =   16
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entre:"
         Height          =   195
         Left            =   3840
         TabIndex        =   12
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   4920
         TabIndex        =   14
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.Menu mnOpt 
      Caption         =   "Opções"
      Begin VB.Menu novo 
         Caption         =   "&Novo Orçamento"
      End
      Begin VB.Menu mnEditar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnDup 
         Caption         =   "Duplicar"
      End
      Begin VB.Menu mnExcluir 
         Caption         =   "Excluir"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnEfetivar 
         Caption         =   "Efetivar Orçamento"
      End
      Begin VB.Menu mnSair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmEdiOrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
End Sub


Private Sub cmdDuplicate_Click()
    On Error GoTo trata_Erro
    Dim fr As New frmOrcamentoUniformes
    grid.col = 1
    
    
    
    fr.tag = orcamento.duplicarOrcamento(grid.Text)
    fr.Show 0, mdiFrm
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo trata_Erro
    Dim fr As New frmOrcamentoUniformes
    grid.col = 1
    fr.tag = grid.Text
    fr.Show 0, mdiFrm
trata_Erro:
End Sub

Private Sub cmdEfetivar_Click()
    Dim id As String
    Dim valor As String
    Dim num_ven As String
    With grid
        .col = 1
        id = .Text
        If MsgBox("Confirma efetivação do orçamento em uma nova venda?", vbYesNo + vbQuestion) Then
            
            num_ven = orcamento.GeraVenda(id)
            AlteraOrcamentoStatus id, fechado
            If num_ven <> "" Then
                frmVenda.tag = num_ven
                frmVenda.Show 1
            End If
        End If
    End With
End Sub

Private Sub cmdExportar_Click()
    ExportaMSFLXGridToCSV grid
End Sub

Private Sub cmdImprimir_Click()
    grid.col = 1
    'ImprimeOrcamento grid.Text
    ImprimeOrcamento_Arquivo grid.Text
    
    
End Sub

Private Sub cmdNovo_Click()
    Dim fr As New frmOrcamentoUniformes
    fr.Show 1
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub
Sub busca()
        Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim campos As String
    Dim cab As String
    
    strSQL = "select o.*,c.nome_cli,c.cnpj_cli, f.guerra_fun,fp.nome_for,if(status=1, 'Em aberto',if(status=2,'Efetivado','Cancelado')) as status_descricao"
    strSQL = strSQL & " from orcamentos as o inner join clientes as c on o.cod_cli=c.cod_cli "
    strSQL = strSQL & " left join " & TAB_FUNCIONARIOS & " as f on o.cod_usu=f.cod_fun "
    strSQL = strSQL & " inner join " & TAB_FORMASPAG & " as  fp on fp.cod_for=o.cod_for "
    strSQL = strSQL & " where 1>0 "
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and data >=" & F_GravaData(txtDe.Text)
    End If
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and left(data,10) <=" & F_GravaData(txtAte.Text)
    End If
    
    
    If txtBusca.Text <> "" Then
        strSQL = strSQL & " (and nome like '%" & txtBusca.Text & "%' "
        strSQL = strSQL & " (or contato like '%" & txtBusca.Text & "%' "
        strSQL = strSQL & " (or descricao like '%" & txtBusca.Text & "%' "
        
        
        
        strSQL = strql & ")"
    End If
    
    
    abreObjetoAdo rPesq, strSQL
    
    cab = "Item,Número, Código, Cliente, contato, telefone, Valor, Data,  Condições Pgto., Usuário, Status"
    campos = "@item,id, cod_cli,cliente,contato,telefone,valor_final,data,nome_for,guerra_fun,status_descricao"
    AjustaGridData grid, 11, 1, 1, cab, campos, "500,800,0,4000,2500,1000,1000,2000,3500,1500", rPesq, "", True, 9829119, False
    grid.ColAlignment(5) = 1
    grid.ColAlignment(6) = 1
    grid.ColAlignment(7) = 0
    grid.ColAlignment(8) = 0
    
    
    
    

End Sub




Private Sub cmdSair_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    busca
    
End Sub

Private Sub Form_Load()
    resizeForm Me
    txtDe.Text = F_PrimDiaMes(Date)
    txtAte.Text = Date
    With grid
        
        .FixedCols = 0
        
        
        .Cols = 10
        .row = 0
        .col = 0: .Text = "Nº."
        .col = 1: .Text = "Cód."
        .col = 2: .Text = "Cliente:"
        .col = 3: .Text = "Valor:"
        .col = 4: .Text = "Desconto:"
        .col = 5: .Text = "Data:"
        .col = 6: .Text = "Validade:"
        .col = 7: .Text = "Usuário:"
        .col = 8: .Text = "Situação:"
        .col = 9: .Text = "Condições Pgto:"
        .ColWidth(0) = 1000
        .ColWidth(1) = 800
        .ColWidth(2) = 3500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 2500
        .ColWidth(6) = 1200
        .ColWidth(7) = 1500
        .ColWidth(8) = 1200
        .ColWidth(9) = 2500
    End With
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_Resize()
    posicionaADireita cmdSair, pcBot
    PosicionaGrade grid, Me, pcBusca
End Sub

Private Sub grid_DblClick()
    cmdEdit_Click
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
    
        Me.PopupMenu mnOpt
    End If
End Sub

Private Sub mnDup_Click()
    cmdDuplicate_Click
End Sub

Private Sub mnEditar_Click()
    cmdEdit_Click
End Sub

Private Sub mnExcluir_Click()
    Dim id As String
    Dim status As String
    
    id = buscaValorColuna(grid, 1)
    status = buscaValorColuna(grid, 10)
    If LCase(status) = "em aberto" Then
        If MsgBox("Deseja excluir permanentemente o orçamento " & id & "?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        Else
            ExcluiOrcamento id, True
        End If
        busca
    End If
    
End Sub

Private Sub mnImprimir_Click()
    mnImprimir_Click
End Sub

Private Sub novo_Click()
    cmdNovo_Click
End Sub

Private Sub txtAte_GotFocus()
    ganhaFoco txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub


Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim c As cliente
        c = F_BuscaCliente(txtbxCliente.Text)
        If c.cod_cli <> "" Then
            txtbxCliente.Text = c.nome_Cli
            txtbxCliente.tag = c.cod_cli
            TrocaTabKey KeyCode
        Else
            txtbxCliente.Text = ""
            txtbxCliente.tag = ""
            txtbxCliente.SetFocus
        End If
    End If
End Sub



Private Sub txtBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtDe_GotFocus()
    ganhaFoco txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub
