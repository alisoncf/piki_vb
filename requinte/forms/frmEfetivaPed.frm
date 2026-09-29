VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEfetivaPed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Efetivação de Pedidos"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8235
   Begin VB.Frame frControle 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Frame Frame2 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2535
         Begin VB.CommandButton cmndbtnPrint 
            Caption         =   "Emitir &Pedido"
            Enabled         =   0   'False
            Height          =   660
            Left            =   120
            Picture         =   "frmEfetivaPed.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Reabrir Pedido"
            Height          =   660
            Left            =   120
            Picture         =   "frmEfetivaPed.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3120
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnEfetiva 
            Caption         =   "&Efetivar Pedido"
            Height          =   660
            Left            =   120
            Picture         =   "frmEfetivaPed.frx":04D4
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "C&ancelar Pedido"
            Height          =   660
            Left            =   120
            Picture         =   "frmEfetivaPed.frx":0A5E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2400
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   660
            Left            =   120
            Picture         =   "frmEfetivaPed.frx":0BA8
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3840
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnImpr 
            Caption         =   "Emitir &Cupom Fiscal"
            Enabled         =   0   'False
            Height          =   660
            Left            =   120
            Picture         =   "frmEfetivaPed.frx":0F32
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.Data dtVendas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Visible         =   0   'False
         Width           =   1260
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmEfetivaPed.frx":12BC
         Height          =   4575
         Left            =   2760
         OleObjectBlob   =   "frmEfetivaPed.frx":12D3
         TabIndex        =   8
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmEfetivaPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmbxBusca_Click()
    'P_Busca
End Sub

Private Sub cmbxOrd_Click()
    'P_Busca
End Sub

Private Sub cmbxSituacao_ven_Click()
   P_Busca
End Sub

Private Sub cmbxLoja_Click()
    P_Busca
End Sub

Private Sub cmndbtnEdit_Click()
    
    
    With dtVendas.Recordset
        If !situacao_ped <> "2" Then
            MsgBox "Este pedido não pode ser Cancelado!", vbExclamation
            P_Status ""
            Exit Sub
        End If
        
        
    P_Status ""
    End With
End Sub

Private Sub cmndbtnEfetiva_Click()
    If MsgBox("Deseja realmente efetivar o Pedido?", vbYesNo + vbQuestion) = vbYes Then
        P_BaixaPed dtVendas.Recordset!Num_Ped
    End If
    P_Busca
End Sub

Private Sub cmndbtnImpr_Click()
    P_Status "Gerando impressão... "
    P_ImprimePedido frmVisualiza.Pic, dtVendas.Recordset!Num_Ped, False
    P_Status ""
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
        Unload frmVisualiza
        P_ImprimePedido Printer, dtVendas.Recordset!Num_Ped, True
                
    End If
End Sub

Private Sub cmndbtnNew_Click()
    On Error Resume Next
    Dim strSQL As String
    If MsgBox("Confirma abertura de Pedido? ", vbYesNo + vbQuestion) = vbYes Then
        strSQL = F_GeraNovoPedido(strCod_Loj)
        If strSQL <> "" Then
            frmPedido.lblPedido.caption = strSQL
            frmPedido.Tag = strSQL
            frmPedido.Show 1
        End If
    End If
End Sub

Private Sub cmndbtnPrint_Click()
    If MsgBox("Deseja Imprimir Pedido?", vbYesNo + vbQuestion) = vbYes Then
        P_ImprimePedido Printer, dtVendas.Recordset!Num_Ped, True
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub DBGrid1_DblClick()
    If cmndbtnEdit.Enabled = True Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.MarqueeStyle = 3
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DBGrid1_DblClick
    End If
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.MarqueeStyle = 0
End Sub

Private Sub dtVendas_Reposition()
    If dtVendas.Recordset.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
        Exit Sub
    Else
        If dtVendas.Recordset!situacao_ped = "1" Then
            cmndbtnImpr.Enabled = False
            cmndbtnEdit.Enabled = True
        Else
            cmndbtnImpr.Enabled = True
            cmndbtnEdit.Enabled = False
        End If
    End If
    
    
    If dtVendas.Recordset!situacao_ped > 1 Then
        cmndbtnEdit.Enabled = False
        
    End If
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    'Me.Width = 10725
    'Me.Height = 6135
    P_CentraForm Me
    'P_PreencheLojas cmbxLoja, False
    dtVendas.Connect = banco.Connect
    
    
    'cmbxBusca.ListIndex = 0
    'cmbxOrd.ListIndex = 0
    'cmbxSituacao_ven.ListIndex = 0
    P_Busca
    bActive = True
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    On Error GoTo trata_erro
    
    
    
    strSQL = "select tc.nome_loj,tv.* from tab_lojas as tc, "
    strSQL = strSQL & " tab_pedidos as tv "
    
    strSQL = strSQL & " where tv.situacao_ped='2' "
    strSQL = strSQL & " and tc.cod_loj=tv.cod_loj "
    strSQL = strSQL & " order by num_ped "
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxExpr_DblClick()
    P_Busca
End Sub


Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    P_Busca
End Sub





