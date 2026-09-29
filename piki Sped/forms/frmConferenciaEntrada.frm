VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConferenciaEntrada 
   Caption         =   "Conferência de Entrada de Produtos"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15045
   Icon            =   "frmConferenciaEntrada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   15045
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7435
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   15015
      TabIndex        =   1
      Top             =   5850
      Width           =   15045
      Begin VB.PictureBox pcBotoes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   9840
         ScaleHeight     =   705
         ScaleWidth      =   5025
         TabIndex        =   15
         Top             =   120
         Width           =   5055
         Begin VB.CommandButton cmdExportar 
            Caption         =   "Exportar"
            Height          =   735
            Left            =   15
            Picture         =   "frmConferenciaEntrada.frx":1084A
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "Sair"
            Height          =   735
            Left            =   3375
            Picture         =   "frmConferenciaEntrada.frx":10C79
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnImprGrade 
            Caption         =   "&Imprimir"
            Height          =   735
            Left            =   1695
            Picture         =   "frmConferenciaEntrada.frx":111C7
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.TextBox txtId 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtParticipante 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Id:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Participante"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   15015
      TabIndex        =   0
      Top             =   0
      Width           =   15045
      Begin VB.OptionButton optRef 
         Caption         =   "Ordenar por referência"
         Height          =   375
         Left            =   6720
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optNota 
         Caption         =   "Ordenar por item da NF"
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   480
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.TextBox txtProduto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informe o código do produto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   3660
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "&Menu"
      Visible         =   0   'False
      Begin VB.Menu mnCopiar 
         Caption         =   "&Copiar Texto"
      End
   End
End
Attribute VB_Name = "frmConferenciaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim ado As New ADODB.Recordset
Dim iLinha As Integer 'linha onde inicia a busca

Private Sub cmdExportar_Click()
    
 ExportaMSFLXGridToCSV grade
End Sub

Private Sub cmndbtnImprGrade_Click()
    Dim titulo As String
    titulo = "Conferência de Estoque NF (" & txtId.Text & ") " & txtNumero.Text & " - " & txtParticipante.Text
    P_MsflxToHtml titulo, grade
 
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    busca
    txtProduto.SetFocus
    bActive = True
End Sub
Sub verificaContagem()
    Dim i As Integer
    Dim qtd As String
    Dim qtd_checada As String
    grade.Visible = False
    With grade
        For i = 1 To .rows - 1
            .row = i
            .col = 4
            qtd = .Text
            .col = 5
            qtd_checada = .Text
            If CCur(qtd_checada) <> CCur(qtd) Then
                .CellBackColor = vbRed
            Else
                .CellBackColor = vbGreen
            End If
        Next
    End With
    grade.Visible = True
End Sub

Sub busca()
    Dim strSQL As String
    Dim i As Integer
    Dim campos As String
    Dim cab As String
    strSQL = "select sp.nome as participante,n.num_doc,n.data, ce.qtd as qtd_checada,ce.cod_usu as usuario,ce.id as id_conf,"
    strSQL = strSQL & " tp.cod_pro as codigo,tp.codfab_pro,nome_pro,qtd_convert,endereco_pro, tm.nome_mar,"
    strSQL = strSQL & " tp.er_pro,tp.ef_pro,tp.cr_pro,tp.pv_pro, tp.codfab_pro,tp.nome_pro as produto,"
    strSQL = strSQL & "vl_item, vUnCom as unit, nome_mar as marca, tc.nome_cat as categoria, tp.ncm_pro as ncm, ti.id as id_item, ti.*,"
    
    strSQL = strSQL & " tp.cr_pro as cr, tp.pv_pro as pv,pv_pro,cr_pro,tp.lucro_pro "
    strSQL = strSQL & " from sped_participante sp inner join sped_nf n on sp.codigo=n.cod_part "
    strSQL = strSQL & " inner join sped_nf_itens as ti on n.id=ti.id_nf "
    strSQL = strSQL & " left join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " left join " & TAB_MARCAS & "  as tm on tm.cod_mar=tp.cod_mar"
    strSQL = strSQL & " left join " & TAB_CATEGORIAS & " as tc on tc.cod_cat=tp.cod_cat "
    strSQL = strSQL & " left join " & TAB_SUBCAT & " as ss on tp.cod_sub=ss.cod_sub "
    strSQL = strSQL & " left join conferencia_entrada ce on ce.id_item=ti.id "
    
    
    strSQL = strSQL & " where ti.id_nf ='" & Me.tag & "' "
    
    
    
       
    If optNota.Value Then
        strSQL = strSQL & " order by ti.id "
    Else
        strSQL = strSQL & " order by codfab_pro "
    End If
    
    abreObjetoAdo ado, strSQL
    
    txtNumero.Text = ado!num_doc
    txtId.Text = Me.tag
    txtData.Text = ado!data
    txtParticipante.Text = ado!participante
    
    campos = "@Item, codfab_pro, id_item, nome_pro, qtd_convert, qtd_checada, endereco_pro, marca, usuario, id_conf"
    cab = "Item,Referência,id_item,Produto,Quantidade,Qtd Checada,Endereço,Marca,usuario,id_conf"
    AjustaGridData grade, 6, 1, 1, cab, campos, "700,2000,0,4500,1200,1200,1000,3500,1000,0", ado, "4,5", True
    
    
    
    verificaContagem
    
End Sub

Sub ajustaGrid()
        
End Sub

Private Sub Form_Load()
    bActive = False
    resizeForm Me
    
End Sub

Private Sub Form_Resize()
    PosicionaGrade grade, Me, pcBot
    posicionaADireita pcBotoes, Me, 300
End Sub

Private Sub grade_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnMenu
    End If
End Sub

Private Sub mnCopiar_Click()
    Clipboard.SetText grade.Text
End Sub

Private Sub optNota_Click()
    busca
End Sub

Private Sub optRef_Click()
    busca
End Sub

Private Sub txtId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
        busca
        
        
        txtProduto.SetFocus
        txtProduto_GotFocus
    End If
    
End Sub
Sub buscaProd()
    Dim p As Produto
    Dim qtd As String
    Dim qtd_pedido As String
    Dim iAst As Integer
    Dim codigo As String
    Dim sLinha As String
    Dim i As Integer
    Dim pa As produtoAtacado
    Dim emb As String
    Dim id_item As String
    Dim id_conf As String
    Dim iRepete As Integer
    txtProduto.Text = Replace(txtProduto.Text, " ", "")
    
    qtd = Left(Trim(txtProduto.Text), 4)
    
    iAst = InStr(1, qtd, "*")
    If iAst > 0 Then
        qtd = Left(qtd, iAst - 1)
        codigo = Right(txtProduto, Len(txtProduto.Text) - iAst)
    Else
        codigo = txtProduto.Text
        qtd = "1"
    End If
    
    pa = buscaProduto_atacado(qtd, codigo)
    If pa.codigo <> "" Then
        emb = pa.qtd
    Else
        emb = "1"
    End If
    
    qtd = CCur(qtd) * CCur(emb)
    
    p = Produtos.F_BuscaProdSaida(codigo, qtd, "", True)
    
    
    If p.cod_pro <> "" Then
        
        iAst = GridUtil.getRowByText(grade, p.codfab_pro, 1, iLinha)
        
        
pergunta:
        If iAst = -1 Then

            
                If MsgBox("Este produto não foi encontrado no pedido. Verifique", vbAbortRetryIgnore + vbDefaultButton2) = vbAbort Then
                    Exit Sub
                Else
                    GoTo pergunta
                End If
            
            
        Else
compara:
            grade.row = iAst
            grade.col = 9
            id_conf = grade.Text
            grade.col = 2
            id_item = grade.Text
            grade.col = 4
            qtd_pedido = grade.Text
            grade.col = 5
            If grade.Text = "" Then grade.Text = "0"
            
            i = CCur(grade.Text)
            If (i + CCur(qtd)) > CCur(qtd_pedido) Then
                iRepete = GridUtil.getRowByText(grade, p.codfab_pro, 1, iAst + 1)
pergunta_maior:
                If iRepete > -1 Then
                    iAst = iRepete
                    GoTo compara
                Else
                    If MsgBox("A quantidade é maior que a quantidade pedida. Verifique", vbAbortRetryIgnore + vbDefaultButton2) = vbAbort Then
                        Exit Sub
                    Else
                        GoTo pergunta_maior
                    End If
                End If
            
            End If
            grade.Text = CCur(grade.Text) + CCur(qtd)
            If CCur(grade.Text) = CCur(qtd_pedido) Then
                grade.CellBackColor = vbGreen
            Else
                grade.CellBackColor = vbRed
            End If
            
            GravaConferenciaEntrada id_item, grade.Text, id_conf
            
        End If
    End If
   
    
    
End Sub
