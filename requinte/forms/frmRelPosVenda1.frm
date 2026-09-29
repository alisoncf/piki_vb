VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelPosVenda 
   Caption         =   "Relatório de Pós-Venda"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   13140
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   13140
      TabIndex        =   7
      Top             =   8490
      Width           =   13140
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   2143
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pCtop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   13110
      TabIndex        =   0
      Top             =   0
      Width           =   13140
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         ScaleHeight     =   1065
         ScaleWidth      =   2505
         TabIndex        =   14
         Top             =   0
         Width           =   2535
         Begin VB.TextBox txtDe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtAte 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   15
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Quando a venda foi gerada?"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   2040
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entre:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "e:"
            Height          =   195
            Left            =   1200
            TabIndex        =   17
            Top             =   120
            Width           =   135
         End
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   10200
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtFun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   495
         Left            =   11520
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         Height          =   495
         Left            =   8880
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmnbuscarFun 
         Height          =   285
         Left            =   4320
         Picture         =   "frmRelPosVenda1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdCidade 
         Height          =   285
         Left            =   6480
         Picture         =   "frmRelPosVenda1.frx":03AB
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do cliente"
         Height          =   195
         Left            =   6840
         TabIndex        =   13
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   4680
         TabIndex        =   11
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   2640
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmRelPosVenda"
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
    
    strSQL = "select tc.cod_cli,tcid.nome_cid,tc.nome_cli,tt.num_tel, tt.tipo_tel, count(ti.seq_itd) as Parcelas "
    strSQL = strSQL & " from tab_clientes as tc inner join " & TAB_DOC & " as td on td.cod_ter=tc.cod_cli and td.tipo_ter='C'"
    strSQL = strSQL & " inner join tab_setores as tset on tset.cod_set=tc.cod_set "
    strSQL = strSQL & " inner join tab_cidades as tcid on tcid.cod_cid=tset.cod_cid"
    strSQL = strSQL & " inner join tab_telefones as tt on tt.cod_cli=tc.cod_cli "
    strSQL = strSQL & " inner join " & TAB_ITDOC & " as ti on ti.seq_doc=td.seq_doc "
    
    
    
    
    strSQL = strSQL & " where reman_itd=0 "
    
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and td.data_doc>=" & F_GravaData(txtDe.Text)
    End If
    
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and td.data_doc<=" & F_GravaData(txtAte.Text)
    End If
    If txtCidade.tag <> "" Then
        strSQL = strSQL & " and tcid.cod_cid='" & txtCidade.tag & "' "
    End If
    'strSQL = strSQL & " and tc.cod_cli not in " & sNotIn
    strSQL = strSQL & " and esp_doc='V'"
    strSQL = strSQL & " group by tc.cod_cli,tt.num_tel  "
    
    'strSQL = strSQL & " having Parcelas<=1 "
    strSQL = strSQL & " order by nome_cli "
    abreObjetoAdo ado, strSQL
    grid.rows = 1
    P_Status "Montando grade"
    With ado
        Do While Not .EOF
            If codigo <> CStr(!cod_cli) Then
                grid.AddItem CStr(!cod_cli) & Chr(9) & !nome_cli & Chr(9) & !num_tel & Chr(9) & !nome_cid
                codigo = !cod_cli
            Else
                grid.row = grid.rows - 1
                grid.col = 2
                grid.Text = grid.Text & "/" & !num_tel
            End If
            grid.row = grid.rows - 1
            grid.col = 2
            grid.CellAlignment = 0
            If grid.rows Mod 2 = 0 Then
                ZebraLinha grid, grid.row, &H80000013
            End If
            
            .MoveNext
        Loop
        
    End With
    
    P_Status ""
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmnbuscarFun_Click()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectVendedor(txtFun.Text)
    If f.cod_fun <> "" Then
        txtFun.Text = f.guerra_fun
    End If
    
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    ajustaGrid
    
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
