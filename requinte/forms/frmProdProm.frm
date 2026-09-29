VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmProdProm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relação dos Produtos em Promoção"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9165
   Icon            =   "frmProdProm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Adicionar de:"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   3855
      Begin VB.OptionButton optnFab 
         Caption         =   "Fabricante"
         Height          =   195
         Left            =   1200
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optnPrinc 
         Caption         =   "Princípio"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optnSec 
         Caption         =   "Seção"
         Height          =   195
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optnSub 
         Caption         =   "Sub-Seção"
         Height          =   195
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optnProd 
         Caption         =   "Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtbxDesc 
      Height          =   285
      Left            =   5280
      TabIndex        =   10
      Text            =   "0,00"
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton cmndbtnExcl 
      Caption         =   "&Excluir"
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnAdd 
      Caption         =   "&Adicionar"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnImpr 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5400
      Picture         =   "frmProdProm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   7320
      Picture         =   "frmProdProm.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtbxFinPro 
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtbxProd 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1440
      Top             =   2880
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmProdProm.frx":06D6
      Height          =   5055
      Left            =   120
      OleObjectBlob   =   "frmProdProm.frx":06EA
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "..."
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   195
      Left            =   6000
      TabIndex        =   12
      Top             =   5520
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Desconto:"
      Height          =   195
      Left            =   5280
      TabIndex        =   11
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Até:"
      Height          =   195
      Left            =   4080
      TabIndex        =   4
      Top             =   5280
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Produto:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   600
   End
End
Attribute VB_Name = "frmProdProm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnIr_Click()

End Sub

Private Sub cmndbtnAdd_Click()
    If txtbxProd.tag = "" Then
        MsgBox "Escolha o produto primeiramente!", vbExclamation
        txtbxProd.SetFocus
        Exit Sub
    End If
    If CDate(txtbxFinPro.Text) < Date Then
        MsgBox "A Data Final deve ser posterior ou igual a hoje!", vbExclamation
        txtbxFinPro.SetFocus
        Exit Sub
    End If
    strSQL = "delete from tab_prodpromo where cod_pro='" & txtbxProd.tag & "'"
    Conn.Execute strSQL
    strSQL = "insert into tab_prodpromo(cod_pro,datafin_pro)values("
    strSQL = strSQL & "'" & txtbxProd.tag & "',"
    strSQL = strSQL & F_GravaData(txtbxFinPro.Text) & ") "
    Conn.Execute strSQL
    Timer1_Timer
End Sub

Private Sub cmndbtnBuscar_Click()
    If optnProd.Value Then
        Dim p As Produtos.Produto
        p = Produtos.F_BuscaProdCabecalho(txtbxProd.Text)
        If p.cod_pro <> "" Then
        
        End If
    ElseIf optnSub.Value Then
        Dim s As subsec
        's = subsecAO.F_SelectSubSec
    ElseIf optnSec.Value Then
        
    ElseIf optnPrinc.Value Then
        'principio.
    ElseIf optnFab.Value Then
        'Dim F As fabricante
        'f = fabricante.F_SelectFab(
    End If
    
    
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "delete from tab_prodpromo where cod_pro='" & dtPro.Recordset!cod_pro & "'"
    Conn.Execute strSQL
    Timer1_Timer
    Exit Sub
trata_erro:
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtbxFinPro.Text = Date
    dtPro.Connect = Conn.ConnectionString
    P_CentraForm Me
    Timer1_Timer
End Sub

Private Sub Timer1_Timer()
    Dim strSQL As String
    Dim i As Integer
    On Error Resume Next
    strSQL = "select tp.nome_pro,tp.cod_Pro,tp.codfab_pro,tp.pv_pro,tpp.datafin_pro "
    strSQL = strSQL & " from tab_produtos as tp,tab_prodpromo as tpp"
    strSQL = strSQL & " where tp.cod_pro=tpp.cod_pro "
    strSQL = strSQL & " and tpp.datafin_pro >=" & F_GravaData(Date)
    
        strSQL = strSQL & " order by tp.nome_pro "
    
    
    i = dtPro.Recordset.AbsolutePosition
    dtPro.RecordSource = strSQL
    
    dtPro.Refresh
    If i > -1 Then
        dtPro.Recordset.Move i
    End If
End Sub

Private Sub txtbxFinPro_GotFocus()
    P_SelecTexto txtbxFinPro
End Sub

Private Sub txtbxFinPro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxFinPro_LostFocus()
    P_AjustaData txtbxFinPro
End Sub

Private Sub txtbxProd_GotFocus()
    P_SelecTexto txtbxProd
End Sub

Private Sub txtbxProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxProd_LostFocus()
    P_AjustaTexto txtbxProd
End Sub
