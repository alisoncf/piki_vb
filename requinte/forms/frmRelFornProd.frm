VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelFornProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtos Relacionados a um Fabricante"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   Icon            =   "frmRelFornProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10800
   Begin VB.PictureBox pcECF 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   10800
      TabIndex        =   18
      Top             =   7035
      Width           =   10800
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelFornProd.frx":014A
      Height          =   4095
      Left            =   0
      OleObjectBlob   =   "frmRelFornProd.frx":015B
      TabIndex        =   4
      Top             =   1200
      Width           =   10575
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   7200
      TabIndex        =   11
      Top             =   5280
      Width           =   3375
      Begin VB.Label lblProdutos 
         AutoSize        =   -1  'True
         Caption         =   "Produtos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblQdade 
         AutoSize        =   -1  'True
         Caption         =   "Quant:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   705
         Width           =   570
      End
      Begin VB.Label lblPreco 
         AutoSize        =   -1  'True
         Caption         =   "Preço:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   450
         Width           =   540
      End
      Begin VB.Label lblCusto 
         AutoSize        =   -1  'True
         Caption         =   "Custo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   195
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton cmndbtnBuscaFab 
         Height          =   255
         Left            =   6120
         Picture         =   "frmRelFornProd.frx":18AC
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtbxFabricante 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   480
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Produtos por:"
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1575
         Begin VB.OptionButton optnFor 
            Caption         =   "Fornecedores"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   500
            Width           =   1335
         End
         Begin VB.OptionButton optnMar 
            Caption         =   "Fabricantes"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   190
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ordernar por &Código"
         Height          =   195
         Left            =   3480
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optnNome 
         Caption         =   "Ordernar por &Nome"
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   840
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   7200
         Picture         =   "frmRelFornProd.frx":19A5
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   8880
         Picture         =   "frmRelFornProd.frx":1DB2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblMar 
         AutoSize        =   -1  'True
         Caption         =   "&Fabricante:"
         Height          =   195
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   795
      End
   End
   Begin MSDBGrid.DBGrid dbgrdEntr 
      Bindings        =   "frmRelFornProd.frx":1F1D
      Height          =   1455
      Left            =   0
      OleObjectBlob   =   "frmRelFornProd.frx":1F31
      TabIndex        =   5
      Top             =   5400
      Width           =   6975
   End
End
Attribute VB_Name = "frmRelFornProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bBusca As Boolean
Dim sTabCmp As String
Dim sTabIt As String
Private Sub Command1_Click()
    
End Sub

Private Sub cmbxForn_Click()
    If optnMar.Value Then
        cmbxForn.tag = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxForn.Text & "'")
    Else
        cmbxForn.tag = F_BuscaUmCampo("cod_for", "tab_fornecedores", "nome_for='" & cmbxForn.Text & "'")
    End If
    busca
End Sub
Sub soma()
    Dim strSQL As String
    Dim cCusto As Currency
    Dim cVen As Currency
    Dim iEst As Currency
    
    With dt.Recordset
        bBusca = False
        Do While Not .EOF
            cCusto = cCusto + (!cr_pro * !Estoque)
            iEst = iEst + !Estoque
            cVen = cVen + (!pv_pro * !Estoque)
            .MoveNext
            P_Status "Somando...:" & .AbsolutePosition + 1 & " / " & .RecordCount
            Me.Refresh
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        
        lblCusto.caption = "Custo: "
        lblPreco.caption = "Preço: "
        lblQdade.caption = "Quantidade: "
        lblCusto.caption = lblCusto.caption & " " & FormatVal(cCusto)
        lblPreco.caption = lblPreco.caption & " " & FormatVal(cVen)
        lblQdade.caption = lblQdade.caption & " " & FormatVal(iEst)
        lblProdutos.caption = "Produtos: " & .RecordCount
    End With
    P_Status ""
    bBusca = True
End Sub
Sub busca()
    Dim strSQL As String
    If pcECF.ForeColor = vbYellow Then
        sTabCmp = "tab_cmp"
        sTabIt = "tab_itcmp"
    Else
        sTabCmp = "tab_compras"
        sTabIt = "tab_itcompras "
    End If
    If optnMar.Value Then
        strSQL = Produtos.F_SelectProduto("cod_fab='" & txtbxFabricante.tag & "'")
    Else
        strSQL = "select  distinct sum(er_pro) as estoque,tp.*,tu.nome_sub,ts.nome_sec  From   " & sTabCmp & " as tc,"
        strSQL = strSQL & "tab_fornecedores as tf," & sTabIt & " as ti,tab_produtos as tp,"
        strSQL = strSQL & "tab_subsec as tu, tab_secoes As ts,tab_estoque as te  "
        strSQL = strSQL & " Where tu.cod_sub=tp.cod_sub and tu.cod_sec=ts.cod_sec "
        strSQL = strSQL & " and ti.seq_not = tc.seq_not"
        strSQL = strSQL & " and"
        strSQL = strSQL & " tc.Cod_for = tf.Cod_for and tf.Cod_for = '" & txtbxFabricante.tag & "' "
        strSQL = strSQL & " and  ti.cod_pro = tp.cod_pro "
        strSQL = strSQL & " and tp.cod_pro=te.cod_pro "
        strSQL = strSQL & " group by tp.cod_pro "
        If optnNome.Value = True Then
            strSQL = strSQL & " order by tp.nome_pro "
        Else
            strSQL = strSQL & " order by tp.cod_pro "
        End If
    End If
    dt.RecordSource = strSQL
    dt.Refresh
    soma
End Sub

Private Sub cmndbtnBuscaFab_Click()
    If optnFor Then
        buscaFor
    Else
        
    End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
    f = F_SelectForn(txtbxFabricante.Text)
    If f.cod_for <> "" Then
        txtbxFabricante.Text = f.nome_for
        txtbxFabricante.tag = f.cod_for
        busca
    End If
    
    
End Sub

Private Sub cmndbtnImprimir_Click()
    bBusca = False
    If dt.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que ser impresso!", vbExclamation
        Exit Sub
    End If
    If cmbxForn.tag = "" Then
        MsgBox "Escolha um fornecedor!", vbExclamation
        cmbxForn.SetFocus
        Exit Sub
    End If
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
    bBusca = True
End Sub
Sub imprime(obj As Object)
    Dim iPage As Integer
    Dim iCount As Integer
    P_CabRelTimbr "Listagem de Produtos por Fornecedor", obj
    obj.FontName = "arial"
    obj.fontsize = 12
    obj.FontBold = True
    P_ImprX obj, cmbxForn.Text, 10, False
    obj.fontsize = 8
    obj.FontBold = False
    P_ImprLinha obj, "_"
    With dt.Recordset
        .MoveFirst
        P_ImprX obj, "Cód", 0
        P_ImprX obj, "Nome", 15
        P_ImprX obj, "UN", 100
        P_ImprX obj, "Estoque", 115, True, True
        P_ImprX obj, "Custo", 130, True, True
        P_ImprX obj, "Preço", 145, True, True
        P_ImprX obj, "Categoria", 147
        P_ImprX obj, "Marca", 180, False
        P_ImprLinha obj, "_"
        iPage = 1
        Do While Not .EOF
            If iCount = 50 Then
                
                iCount = 0
                P_RodapeVis obj, constNome_sis & "-" & constVersao_Sis, "Pág:" & iPage
                iPage = iPage + 1
                P_ImprX obj, "Cód", 0
                P_ImprX obj, "Nome", 15
                P_ImprX obj, "UN", 100
                P_ImprX obj, "Estoque", 115, True, True
                P_ImprX obj, "Custo", 130, True, True
                P_ImprX obj, "Preço", 145, True, True
                P_ImprX obj, "Categoria", 147
                P_ImprX obj, "Marca", 180, False
                P_ImprLinha obj, "_"
            End If
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome_pro, 15
            P_ImprX obj, !unidade_pro, 100
            P_ImprX obj, FormatVal(!er_pro), 115, True, True
            P_ImprX obj, FormatVal(!cr_pro), 130, True, True
            P_ImprX obj, FormatVal(!pv_pro), 145, True, True
            P_ImprX obj, !nome_cat, 147
            P_ImprX obj, " " & !nome_mar, 179, False
            
            iCount = iCount + 1
            .MoveNext
        Loop
    End With
    P_RodapeVis obj, constNome_sis & "-" & constVersao_Sis, "Pág:" & iPage & "(Fim)"
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dt_Reposition()
    BuscaCMP
End Sub
Sub BuscaCMP()
    On Error GoTo trata_erro
    Dim strSQL As String
    If Not bBusca Then Exit Sub
    With dt.Recordset
        strSQL = "select tc.num_not,tc.entrada_not,ti.* from " & sTabCmp & " as tc,"
        strSQL = strSQL & sTabIt & " as ti where tc.seq_not=ti.seq_not "
        
        strSQL = strSQL & " and ti.cod_pro='" & !cod_pro & "' "
        'strSQL = strSQL & " and tc.cod_for='" & F_BuscaUmCampo("cod_for", "tab_produtos", "cod_pro ='" & !cod_pro & "'") & "'"
        'cmbxForn.Tag & "'"
        strSQL = strSQL & " order by tc.entrada_not"
    End With
    dtSai.RecordSource = strSQL
    dtSai.Refresh
    Exit Sub

trata_erro:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        pcECF.BackColor = vbGreen
    ElseIf KeyCode = vbKeyF9 Then
        pcECF.BackColor = vbYellow
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dt.Connect = Banco.Connect
    dtSai.Connect = Banco.Connect
    If optnMar.Value Then
        
        lblMar.caption = "Marcas"
        frmRelFornProd.caption = "Produtos Relacionados a Marcas"
    Else
        
        lblMar.caption = "Fornecedores"
        frmRelFornProd.caption = "Produtos Relacionados a Fornecedores"
    End If
    bBusca = True
End Sub

Private Sub optnFor_Click()
    
    lblMar.caption = "Fornecedores"
    frmRelFornProd.caption = "Produtos Relacionados a Fornecedores"

End Sub

Private Sub optnMar_Click()
    
    
    lblMar.caption = "Fabricante"
    frmRelFornProd.caption = "Produtos Relacionados a Fabricante"
    
End Sub
