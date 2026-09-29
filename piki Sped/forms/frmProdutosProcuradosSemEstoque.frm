VERSION 5.00
Begin VB.Form frmProdutosProcuradosSemEstoque 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Registrar Procura de Produtos sem Estoque"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "frmProdutosProcuradosSemEstoque.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   855
      Left            =   9720
      Picture         =   "frmProdutosProcuradosSemEstoque.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Confirmar"
      Height          =   855
      Left            =   8160
      Picture         =   "frmProdutosProcuradosSemEstoque.frx":0E58
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtProduto 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   11175
   End
   Begin VB.Label lblCodFab 
      AutoSize        =   -1  'True
      Caption         =   "Marca"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cód.Fab"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblMarca 
      AutoSize        =   -1  'True
      Caption         =   "Marca"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   5
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Marca"
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "frmProdutosProcuradosSemEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Dim est As String
    est = F_BuscaUmCampo("sum(er)", "estoque", "cod_pro='" & txtProduto.tag & "'")
    If est <> "" Then
        If CCur(est) > 0 Then
            
            If MsgBox("O produto selecionado consta com estoque de " & est & " deseja adicionar mesmo assim?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
            
        End If
    End If
    Produtos.gravaProdutoProcuradoSemEstoque txtProduto.tag
    Unload Me
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        carrega
    End If
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select tp.*,tm.nome_mar from " & TAB_PRODUTOS & " as tp left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " and tp.cod_pro='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            txtProduto.Text = !nome_pro
            txtProduto.tag = !cod_pro
            lblCodFab.caption = !codfab_pro
            lblMarca.caption = Format(!nome_mar)
            cmdOk.SetFocus
        End If
    End With
    
    
End Sub
