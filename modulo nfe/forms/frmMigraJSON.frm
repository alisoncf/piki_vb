VERSION 5.00
Begin VB.Form frmMigraJSON 
   Caption         =   "Migração de Dados"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcPush 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2295
      ScaleWidth      =   8655
      TabIndex        =   6
      Top             =   2760
      Width           =   8655
      Begin VB.CommandButton cmdImpProd 
         Caption         =   "Importar Produtos"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.PictureBox pcGerar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   8775
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.TextBox txtDataProduto 
         Height          =   405
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Gerar Clientes"
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton cmdProd 
         Caption         =   "Gerar Produtos"
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   $"frmMigraJSON.frx":0000
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data alvo:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMigraJSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProd_Click()
    Dim c As New Collection
    Dim p(6) As par
    Dim s As String
    Dim prods As String
    Dim caminhoArquivo As String
    Dim strSQL As String
    Dim empresa As String
    Dim sec As String
    Dim subsec As String
    caminhoArquivo = BuscaNomeArquivoExterno("*.json", "Migra_Produtos" & F_SoNumero(Now))
        
    
    p(1).nome = "empresa"
    p(1).type = Char
    p(1).valor = strNome_Emp
    
    p(2).nome = "CNPJ"
    p(2).type = Char
    p(2).valor = strCNPJ_Emp
    
    p(3).nome = "Data"
    p(3).type = data
    p(3).valor = Now
    
    
    p(4).nome = "Ação"
    p(4).type = Char
    p(4).valor = "MigrarProdutos"
    
    p(5).nome = "AlterarEstoque"
    p(5).type = booleano
    p(5).valor = "false"
    
    p(6).nome = "DataAlvo"
    p(6).type = Char
    p(6).valor = txtDataProduto.Text
    
    
    empresa = XMLBD.ConjuntoJSON(p, 0, "generator")
    
    AbreArquivoExterno caminhoArquivo, 1
    
    ImprimeEmArquivoExterno "{" & CrLf & empresa, 1
    
    strSQL = "select * from " & TAB_PRODUTOS & " where 1>0 "
    If txtDataProduto.Text <> "" Then
        strSQL = strSQL & " and dataalt_pro>=" & F_GravaData(txtDataProduto.Text)
    End If
    sec = RecordSetToJson("select * from " & TAB_SECOES, TAB_SECOES, "", 0, 1, False)
    subsec = RecordSetToJson("select * from " & TAB_SUBSEC, TAB_SUBSEC, "", 0, 1, False)
    prods = RecordSetToJson(strSQL, TAB_PRODUTOS, "", 0, 1, False)
    
    
    ImprimeEmArquivoExterno "}", 1
    FechaArquivoExterno 1
           
    shell "explorer " & caminhoArquivo
    
    
    
    
    
    

End Sub

Private Sub txtDataProduto_GotFocus()
    P_SelecTexto txtDataProduto
End Sub

Private Sub txtDataProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDataProduto_LostFocus()
    P_AjustaData txtDataProduto
End Sub
