VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEfetivaVen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Efetivação de Vendas"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9885
   Begin VB.Frame Frame1 
      Caption         =   "&Vendas a Efetivar:"
      Height          =   5655
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9855
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   4440
         Top             =   1200
      End
      Begin VB.Data dtVendas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   1260
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEfetivaVen.frx":0000
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "frmEfetivaVen.frx":0017
         TabIndex        =   6
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   9855
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "Imprimir Venda"
         Enabled         =   0   'False
         Height          =   660
         Left            =   2040
         Picture         =   "frmEfetivaVen.frx":15D8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir Grade"
         Height          =   660
         Left            =   3960
         Picture         =   "frmEfetivaVen.frx":1B62
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   660
         Left            =   7800
         Picture         =   "frmEfetivaVen.frx":1EEC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnEfetivar 
         Caption         =   "&Efetivar"
         Height          =   660
         Left            =   120
         Picture         =   "frmEfetivaVen.frx":2036
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmEfetivaVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub busca()
     Dim strSQL As String
    
    
    
    
    
    strSQL = "select tc.cod_cli,tc.nome_cli,tu.guerra_fun,tv.num_ven,tv.data_ven,tv.situacao_ven,tv.valor_ven, "
    strSQL = strSQL & " tf.*,(tf.desconto_ven*tv.valor_ven)/100 as desconto,valor_fec"
    strSQL = strSQL & " from tab_clientes as tc, "
    strSQL = strSQL & " tab_vendas as tv,tab_venfec as tf, tab_funcionarios as tu "
    strSQL = strSQL & " where tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tv.situacao_ven='2' "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    strSQL = strSQL & " and tv.cod_fun=tu.cod_fun "
    If UCase(App.title) = "PIKIBOX_CX" Then
        strSQL = strSQL & " and tf.cod_cai='" & strCod_Cai & "' "
    End If
    strSQL = strSQL & " order by tf.num_ven"
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
End Sub
    
Private Sub cmndbtnCancelar_Click()
    On Error Resume Next
    ImprVenda dtVendas.Recordset!num_ven
    
End Sub

Private Sub cmndbtnImpr_Click()
    P_DbgrdToHtml "Vendas a Efetivar em " & Now, dbgrd, dtVendas

End Sub

Private Sub cmndbtnReabrir_Click()
    Dim strSQL As String
    If dbgrd.row = -1 Then
        MsgBox "selecione a Venda!", vbInformation
        dbgrd.SetFocus
        Exit Sub
    End If
    With dtVendas.Recordset
        If MsgBox("Deseja Realmente Reabrir a Venda?", vbYesNo + vbQuestion) = vbYes Then
            
            strSQL = strSQL & " update tab_vendas set situacao_ven='1'"
            strSQL = strSQL & " where num_ven='" & !num_ven & "' "
            
            
            Banco.Execute strSQL
            strSQL = "delete from tab_venfec where num_ven='" & !num_ven & "' "
            Banco.Execute strSQL
            busca
        End If
    End With
End Sub

Private Sub cmndbtnSair_Click()
        Unload Me
End Sub

Private Sub dtVendas_Reposition()
    If dtVendas.Recordset.RecordCount = 0 Then
        cmndbtnCancelar.Enabled = False
        cmndbtnEfetivar.Enabled = False
        cmndbtnImpr.Enabled = False
        cmndbtnReabrir.Enabled = False
    Else
        cmndbtnCancelar.Enabled = True
        cmndbtnEfetivar.Enabled = True
        cmndbtnImpr.Enabled = True
        cmndbtnReabrir.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    
    busca
    Me.caption = "Vendas a Efetivar  - Usuário: " & strNome_Usu
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    ''dtVendas.Connect = Banco.Connect
End Sub

Private Sub Timer1_Timer()
    
    Dim i As Integer
    On Error Resume Next
    i = dtVendas.Recordset.AbsolutePosition
    busca
    dtVendas.Recordset.Move i
End Sub
