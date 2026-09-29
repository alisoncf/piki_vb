VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiTransf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Transferências"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   Icon            =   "frmEdiTransf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9660
   Begin VB.Frame frControle 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.Frame Frame3 
         Height          =   2535
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2055
         Begin VB.ComboBox cmbxBusca 
            Height          =   315
            ItemData        =   "frmEdiTransf.frx":014A
            Left            =   120
            List            =   "frmEdiTransf.frx":0157
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtbxExpr 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   9
            Top             =   2160
            Width           =   1815
         End
         Begin VB.ComboBox cmbxOrd 
            Height          =   315
            ItemData        =   "frmEdiTransf.frx":016F
            Left            =   120
            List            =   "frmEdiTransf.frx":017C
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   1815
         End
         Begin VB.ComboBox cmbxSituacao_ven 
            Height          =   315
            ItemData        =   "frmEdiTransf.frx":0194
            Left            =   120
            List            =   "frmEdiTransf.frx":019E
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Buscar Por:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   825
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "E&xpressão:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1920
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Ordenar:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Situação da Transf."
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   2055
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Nova"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiTransf.frx":01BD
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiTransf.frx":0547
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiTransf.frx":0691
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnImpr 
            Caption         =   "&Imprimir"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiTransf.frx":07DB
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   840
            Width           =   1815
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
         Top             =   2760
         Visible         =   0   'False
         Width           =   1260
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmEdiTransf.frx":0B65
         Height          =   5055
         Left            =   2280
         OleObjectBlob   =   "frmEdiTransf.frx":0B7C
         TabIndex        =   15
         Top             =   240
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmEdiTransf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmbxBusca_Click()
    P_Busca
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOrd_Click()
    P_Busca
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSituacao_ven_Click()
   P_Busca
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmTransferencia
    P_Status "Montando Lançamento..."
    With dtVendas.Recordset
        fr.tag = !seq_tra
        fr.txtbxNome_ter.Text = !nome_loj
        fr.txtbxNome_ter.Enabled = False
        fr.txtbxNome_ter.tag = !cod_loj
        fr.caption = "Transferência " & !seq_tra & "-" & !nome_loj
        fr.Show 0
        
    End With
    P_Status ""
    dtVendas.Refresh
End Sub

Private Sub cmndbtnImpr_Click()
    On Error GoTo trata_erro
    If dtVendas.Recordset.RecordCount > 0 Then
        'imprimeTransf dtVendas.Recordset!seq_tra
    Else
        MsgBox "Não há o que imprimir!", vbExclamation
        
    End If
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnNew_Click()
    On Error Resume Next
    Dim fr As New frmTransferencia
    fr.Show 0
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
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
    On Error Resume Next
    If dtVendas.Recordset.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
        Exit Sub
    Else
        cmndbtnImpr.Enabled = True
        cmndbtnEdit.Enabled = True
    End If
    If dtVendas.Recordset!situacao_Tra > 1 Then
        cmndbtnEdit.Enabled = False
        
    End If
End Sub

Private Sub Form_Activate()
    
    MudaAparencia Me
    
    P_Busca
End Sub

Private Sub Form_Load()
    'Me.Width = 10725
    'Me.Height = 6135
    P_CentraForm Me
    dtVendas.Connect = Banco.Connect
    
    cmbxSituacao_ven.ListIndex = 0
    'cmbxSituacao_ven.Enabled = False
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 1
    
    P_Busca
    bActive = True
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    
    If cmbxBusca.ListIndex = 0 Then
        sBusca = "Nome_loj"
    ElseIf cmbxBusca.ListIndex = 1 Then
        sBusca = "seq_tra"
    ElseIf cmbxBusca.ListIndex = 2 Then
        sBusca = "data_Tra"
    Else
        sBusca = "data_tra"
    End If
    If cmbxOrd.ListIndex = 0 Then
        sOrd = "Nome_loj"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrd = "seq_tra"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrd = "data_tra"
    Else
        sOrd = "data_tra"
    End If
    
    
    strSQL = "select tl.nome_loj,tt.*,tu.guerra_fun "
    strSQL = strSQL & " from tab_lojas as tl, "
    strSQL = strSQL & " tab_Transf as tt,tab_funcionarios as tu"
    
    strSQL = strSQL & " where " & sBusca & " like '" & txtbxExpr.Text & "%' "
    strSQL = strSQL & " and tu.cod_fun=tt.cod_usu "
    strSQL = strSQL & " and tl.cod_loj=tt.cod_loj "
    strSQL = strSQL & " and tt.situacao_tra='" & Left(cmbxSituacao_ven.Text, 1) & "' "
    
    
    strSQL = strSQL & " order by " & sOrd
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
End Sub


Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxExpr_DblClick()
    P_Busca
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        P_Busca
        If dtVendas.Recordset.RecordCount > 0 Then
            DBGrid1.SetFocus
        End If
    End If
    If KeyCode = vbKeyDown Then
        If dtVendas.Recordset.RecordCount > 0 Then
            DBGrid1.SetFocus
        End If
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
End Sub



