VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiReposicao 
   Caption         =   "Reposições"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12495
   Icon            =   "frmEdiReposicao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   12495
   Begin VB.Timer tmr 
      Interval        =   5000
      Left            =   2640
      Top             =   4200
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   12465
      TabIndex        =   8
      Top             =   6300
      Width           =   12495
      Begin VB.ComboBox cmbxSituacao_ven 
         Height          =   315
         ItemData        =   "frmEdiReposicao.frx":058A
         Left            =   3360
         List            =   "frmEdiReposicao.frx":059D
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiReposicao.frx":05FE
         Left            =   1800
         List            =   "frmEdiReposicao.frx":060E
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         MaxLength       =   40
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmEdiReposicao.frx":063D
         Left            =   120
         List            =   "frmEdiReposicao.frx":0647
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   285
         Left            =   11400
         Picture         =   "frmEdiReposicao.frx":0660
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8040
         MaxLength       =   40
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9840
         MaxLength       =   40
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxHoraIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9000
         MaxLength       =   40
         TabIndex        =   10
         Text            =   "00:00"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtbxHoraFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10800
         MaxLength       =   40
         TabIndex        =   9
         Text            =   "23:59"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação da Venda:"
         Height          =   195
         Left            =   3360
         TabIndex        =   23
         Top             =   0
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   1800
         TabIndex        =   22
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   5760
         TabIndex        =   21
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   825
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   8040
         TabIndex        =   19
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   9840
         TabIndex        =   18
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   12465
      TabIndex        =   0
      Top             =   6915
      Width           =   12495
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   585
         Left            =   1440
         Picture         =   "frmEdiReposicao.frx":0759
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   585
         Left            =   8640
         Picture         =   "frmEdiReposicao.frx":0AE3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   585
         Left            =   2880
         Picture         =   "frmEdiReposicao.frx":0C2D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Nova"
         Height          =   585
         Left            =   0
         Picture         =   "frmEdiReposicao.frx":0D77
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnFinalizarVenda 
         Caption         =   "&Recepcionar"
         Height          =   585
         Left            =   4320
         Picture         =   "frmEdiReposicao.frx":1101
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar"
         Height          =   585
         Left            =   5760
         Picture         =   "frmEdiReposicao.frx":148B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnReabre 
         Caption         =   "&Reabrir Rep."
         Height          =   585
         Left            =   7200
         Picture         =   "frmEdiReposicao.frx":15D5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1335
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
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmEdiReposicao.frx":171F
      Height          =   5895
      Left            =   0
      OleObjectBlob   =   "frmEdiReposicao.frx":1736
      TabIndex        =   24
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmEdiReposicao"
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

Private Sub cmndbtnBuscar_Click()
    P_Busca
End Sub

Private Sub cmndbtnCancel_Click()
    
    Exit Sub
    If strExclVenda_usu = "S" Then
        If MsgBox("Deseja realmente cancelar todos os lancamentos dessa reposicao?", vbYesNo + vbQuestion) = vbYes Then
            If dtVendas.Recordset!status_rep <> "4" Then
                MsgBox "A venda não pode ser cancelada por aqui!", vbCritical
            End If
            
        End If
    Else
        MsgBox "Você não tem permissão para isso!", vbCritical
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmReposicao
    P_Status "Montando Reposição..."
    With dtVendas.Recordset
        fr.tag = !seq_rep
        
        
        
        
        fr.Show 0
        
    End With
    P_Status ""
    dtVendas.Refresh
End Sub

Private Sub cmndbtnFinalizar_Click()
    On Error GoTo trata_erro
    'frmEfetivaVen.Show 0
    
    'Exit Sub
    If dtVendas.Recordset!situacao_ven = "3" Then
        frmLanCheque.txtbxCPF.Text = dtVendas.Recordset!cpf_cli
        frmLanCheque.txtbxNome_Che.Text = dtVendas.Recordset!nome_Cli
        frmLanCheque.tag = dtVendas.Recordset!num_ven
        
        
        frmLanCheque.Show 1
        
        
        
    End If
    
    Exit Sub
trata_erro:
    
    
End Sub

Private Sub cmndbtnFinalizarVenda_Click()
    Dim fr As New frmReposicao
    P_Status "Montando Reposição..."
    With dtVendas.Recordset
        fr.tag = !seq_rep
        
        
        
        
        fr.Show 0
        
    End With
    P_Status ""
    dtVendas.Refresh
    
End Sub

Private Sub cmndbtnImpr_Click()
    ImprReposicao dtVendas.Recordset!seq_rep
End Sub

Private Sub cmndbtnNew_Click()
    Dim strSQL As String
    strSQL = F_GeraNovaReposicao()
    frmReposicao.tag = strSQL
    frmReposicao.Show 0
End Sub

Private Sub cmndbtnPosologia_Click()
    On Error Resume Next
    Venda.P_ImprPosologia dtVendas.Recordset!num_ven
End Sub

Private Sub cmndbtnReabre_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    If dtVendas.Recordset!status_rep <> "2" Then
        MsgBox "Você não pode reabrir esta reposição. Ela não se encontra fechada!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma Reabertura da Reposição?", vbYesNo + vbQuestion) = vbYes Then
        Reposicao.P_AlteraReposicao dtVendas.Recordset!seq_rep, "1"
        
            MsgBox "reposição reaberta com sucesso!", vbInformation
            dtVendas.Refresh
        
        
    End If
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub cmndbtnReem_Click()
    frmReemissao.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSepd_Click()
    On Error GoTo trata_erro
    P_notaSEPD dtVendas.Recordset!num_ven, ""
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub cmndbtnVendasEfet_Click()
    frmEfetivaVen.Show 0
End Sub

Private Sub DBGrid1_DblClick()
    On Error GoTo trata_erro
    If cmndbtnEdit.Enabled = True Then
        cmndbtnEdit_Click
    End If
    If dtVendas.Recordset!situacao_ven = "2" Then
        cmndbtnFinalizarVenda_Click
    End If
    Exit Sub
trata_erro:
    
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
    If dtVendas.Recordset!status_rep > 1 Then
        cmndbtnEdit.Enabled = False
                
    End If
    If dtVendas.Recordset!status_rep = "3" Then
        cmndbtnReabre.Enabled = True
    Else
        cmndbtnReabre.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    
    
    MudaAparencia Me
    
    If LCase(App.EXEName) = "piki_cx" Then
        cmbxSituacao_ven.ListIndex = 1
    End If
    
    
    P_Busca
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        'If cmbxSituacao_ven.ListCount <= 2 Then
        '    cmbxSituacao_ven.AddItem "3- Canceladas"
            
        '    cmbxSituacao_ven.AddItem "4- Efetivadas"
        'End If
    End If
    If KeyCode = vbKeyF5 Then
        P_Busca
    End If
End Sub

Private Sub Form_Load()
    resizeForm Me
    dtVendas.Connect = Banco.Connect
    
    
'    cmbxSituacao_ven.ListIndex = 0
    'If strAltprod_usu <> "S" Then
    '    cmbxSituacao_ven.Enabled = False
    'End If
 '   cmbxBusca.ListIndex = 0
  '  cmbxOrd.ListIndex = 1
    
    P_Busca
    bActive = True
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    
    If cmbxBusca.ListIndex = 0 Then
        sBusca = "solicitante"
    ElseIf cmbxBusca.ListIndex = 1 Then
        sBusca = "seq_rep"
    
    End If
    If cmbxOrd.ListIndex = 0 Then
        sOrd = "solicitante"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrd = "seq_rep"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrd = "abertura_rep"
    Else
        sOrd = "fechamento_rep"
    End If
    
    
    strSQL = "select tr.*,tf.guerra_fun as solicitante from tab_reposicao as tr, "
    
    strSQL = strSQL & "tab_funcionarios as tf "
    strSQL = strSQL & " where tf.cod_fun=tr.solicitante_rep "
    
    
    
    strSQL = strSQL & " and tr.status_rep='" & Left(cmbxSituacao_ven.Text, 1) & "' "
    
    
    
    
    If txtbxDe.Text <> "" Then
        strSQL = strSQL & " and abertura_rep >=" & F_GravaDataTempo(txtbxDe.Text & " " & txtbxHoraIni.Text)
    End If
    If txtbxAte.Text <> "" Then
        strSQL = strSQL & " and abertura_rep <=" & F_GravaDataTempo(txtbxAte.Text & " " & txtbxHoraFin.Text)
    End If
    
    strSQL = strSQL & " order by " & sOrd & " desc "
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
End Sub


Private Sub Form_Resize()
    'DBGrid1.Width = Me.Width - 200
    'DBGrid1.Height = pcBot.Top - DBGrid1.Top - 100
    
    
    
End Sub

Private Sub tmr_Timer()
    If cmbxSituacao_ven.ListIndex = 0 Then
        RefreshDataPosition dtVendas
    End If
End Sub

Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
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


Private Sub txtbxHoraFin_GotFocus()
    P_SelecTexto txtbxHoraFin
End Sub

Private Sub txtbxHoraFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxHoraFin_LostFocus()
    P_AjustaHora txtbxHoraFin
End Sub

Private Sub txtbxHoraIni_GotFocus()
    P_SelecTexto txtbxHoraIni
End Sub

Private Sub txtbxHoraIni_LostFocus()
    P_AjustaHora txtbxHoraIni
End Sub
