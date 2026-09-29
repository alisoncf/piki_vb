VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiPerdas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edição de Perdas"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9570
   Icon            =   "frmEdiPerdas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
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
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   2055
         Begin VB.CommandButton cmndbtnImpr 
            Caption         =   "&Imprimir"
            Height          =   585
            Left            =   120
            Picture         =   "frmEdiPerdas.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   585
            Left            =   120
            Picture         =   "frmEdiPerdas.frx":04D4
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   585
            Left            =   120
            Picture         =   "frmEdiPerdas.frx":061E
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Nova"
            Height          =   585
            Left            =   120
            Picture         =   "frmEdiPerdas.frx":0768
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2055
         Begin VB.ComboBox cmbxSituacao_ven 
            Height          =   315
            ItemData        =   "frmEdiPerdas.frx":0AF2
            Left            =   120
            List            =   "frmEdiPerdas.frx":0AFC
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1560
            Width           =   1815
         End
         Begin VB.ComboBox cmbxOrd 
            Height          =   315
            ItemData        =   "frmEdiPerdas.frx":0B1B
            Left            =   120
            List            =   "frmEdiPerdas.frx":0B2B
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtbxExpr 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   3
            Top             =   2160
            Width           =   1815
         End
         Begin VB.ComboBox cmbxBusca 
            Height          =   315
            ItemData        =   "frmEdiPerdas.frx":0B56
            Left            =   120
            List            =   "frmEdiPerdas.frx":0B66
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Situação da Perda:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Width           =   1365
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Ordenar:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "E&xpressão:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1920
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Buscar Por:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   825
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmEdiPerdas.frx":0B91
         Height          =   5055
         Left            =   2280
         OleObjectBlob   =   "frmEdiPerdas.frx":0BA8
         TabIndex        =   15
         Top             =   240
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmEdiPerdas"
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
    Dim fr As New frmPerdas
    P_Status "Montando Lançamento..."
    With dtVendas.Recordset
        fr.tag = !seq_per
        fr.txtbxNOme_Ter.Text = !nome_for
        fr.txtbxNOme_Ter.Enabled = False
        fr.txtbxNOme_Ter.tag = !cod_for
        fr.caption = "Perda " & !seq_per & "-" & !nome_for
        fr.Show 0
        
    End With
    P_Status ""
    dtVendas.Refresh
End Sub

Private Sub cmndbtnImpr_Click()
    On Error GoTo trata_erro
    If dtVendas.Recordset.RecordCount > 0 Then
        imprimePerda dtVendas.Recordset!seq_per
    Else
        MsgBox "Não há o que imprimir!", vbExclamation
        
    End If
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnNew_Click()
    On Error Resume Next
    Dim fr As New frmPerdas
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
    If dtVendas.Recordset!situacao_per > 1 Then
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
    dtVendas.Connect = Conn.ConnectionString
    
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
        sBusca = "Nome_for"
    ElseIf cmbxBusca.ListIndex = 1 Then
        sBusca = "seq_per"
    ElseIf cmbxBusca.ListIndex = 2 Then
        sBusca = "data_per"
    Else
        sBusca = "data_per"
    End If
    If cmbxOrd.ListIndex = 0 Then
        sOrd = "Nome_for"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrd = "seq_per"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrd = "data_per"
    Else
        sOrd = "data_per"
    End If
    
    
    strSQL = "select tf.nome_for,tp.*,tu.guerra_fun "
    strSQL = strSQL & " from tab_fornecedores as tf, "
    strSQL = strSQL & " tab_perdas as tp,tab_funcionarios as tu"
    
    strSQL = strSQL & " where " & sBusca & " like '" & txtbxExpr.Text & "%' "
    strSQL = strSQL & " and tu.cod_fun=tp.cod_usu "
    strSQL = strSQL & " and tf.cod_for=tp.cod_for "
    strSQL = strSQL & " and tp.situacao_per='" & Left(cmbxSituacao_ven.Text, 1) & "' "
    
    
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


