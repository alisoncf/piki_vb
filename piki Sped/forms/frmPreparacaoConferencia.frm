VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPreparacaoConferencia 
   Caption         =   "Adiciona os Produtos em Preparação"
   ClientHeight    =   6840
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   14310
   Icon            =   "frmPreparacaoConferencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
   ScaleMode       =   0  'User
   ScaleWidth      =   17616.48
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar (F7)"
      Height          =   855
      Left            =   12360
      Picture         =   "frmPreparacaoConferencia.frx":1084A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
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
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade 
      Height          =   3855
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   6800
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3660
   End
End
Attribute VB_Name = "frmPreparacaoConferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Sub buscaItens()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iRow As Integer
    Dim qtd As Integer
    Dim iProdutos As Integer
    On Error GoTo trata_Erro:
    strSQL = "select sum(ti.qtd_ite) as qtd, 0 as conf, tp.cod_pro,nome_pro,codfab_pro,endereco_pro,nome_mar from " & TAB_ITVENDAS & "  as ti inner join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " where num_ven='" & Me.tag & "' "
    strSQL = strSQL & " group by tp.cod_pro "
    strSQL = strSQL & " order by ti.seq_ite "
    abreObjetoAdo rPesq, strSQL
    
    'With rPesq
    '    Do While Not .EOF
    '    sLinha = grade.rows & Chr(9) & p.cod_pro & qtd & Chr(9) & "0" & Chr(9) & p.codfab_pro & Chr(9) & p.nome_pro
    'End With
    grade.Clear
    GridUtil.AjustaGridData grade, 7, 1, 1, "#,Código,Qtd.Pedido,Conferência,Cód.Fab.,Produto,Endereço", _
        "@item,cod_pro,qtd,conf,codfab_pro,nome_pro,endereco_pro ", "800,1300,1300,1300,2000,3500", rPesq, "", True
Exit Sub
trata_Erro:
    P_MostraErro
End Sub


Function ConfereItens() As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iRow As Integer
    Dim qtd As Integer
    Dim iProdutos As Integer
    On Error GoTo trata_Erro:
    Dim i As Integer
    Dim j As Integer
    ConfereItens = True
    For i = 1 To grade.rows - 1
        grade.row = i
        grade.col = 2
        qtd = CCur(grade.Text)
        grade.col = 3
        If qtd <> CCur(grade.Text) Then
            MsgBox "Quantidade divergente na linha " & grade.row, vbExclamation
            ConfereItens = False
        End If
        
    Next
    
    
    Exit Function
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Function

Private Sub cmdConfirmar_Click()
    If ConfereItens Then
        Me.tag = "1"
        Me.Hide
   Else
        MsgBox "Não é possível confirmar a preparação!", vbExclamation
        
    End If
End Sub

Private Sub cmdReini_Click()
    grade.rows = 0
    Form_Load
    
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    buscaItens
    bActive = True
End Sub

Private Sub Form_Load()
    bActive = False
    
    
   
    
End Sub

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
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
        
        iAst = GridUtil.getRowByText(grade, p.cod_pro, 1)
        
        If iAst = -1 Then
            MsgBox "Este produto não foi encontrado no pedido", vbExclamation
        Else
            grade.row = iAst
            grade.col = 2
            qtd_pedido = grade.Text
            grade.col = 3
            
            
            i = CCur(grade.Text)
pergunta:
            If (i + CCur(qtd)) > CCur(qtd_pedido) Then
                If MsgBox("A quantidade conferida é maior que a quantidade pedida. Verifique", vbAbortRetryIgnore + vbDefaultButton2) = vbAbort Then
                    Exit Sub
                Else
                    GoTo pergunta
                End If
            End If
            grade.Text = CCur(grade.Text) + CCur(qtd)
            If CCur(grade.Text) = CCur(qtd_pedido) Then
                grade.CellBackColor = vbGreen
            Else
            grade.CellBackColor = vbRed
            End If
            
            
            
        End If
    End If
   
    
    
End Sub
