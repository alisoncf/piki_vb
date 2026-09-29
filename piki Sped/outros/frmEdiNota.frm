VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiNf 
   Caption         =   "Edição de Notas Fiscais de Saída"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   Icon            =   "frmEdiNota.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   10230
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7185
      ScaleWidth      =   2265
      TabIndex        =   16
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmndbtnImprGrade 
         Caption         =   "&Imprimir Grade"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNota.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNota.frx":09A7
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtbxCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Cancelar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNota.frx":0D31
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5520
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "Nova Nota Fiscal"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNota.frx":0E7B
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNota.frx":1205
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiNota.frx":1612
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6240
         Width           =   2055
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEdiNota.frx":175C
         Left            =   120
         List            =   "frmEdiNota.frx":176C
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         Picture         =   "frmEdiNota.frx":1794
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.Data dtNot 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiNota.frx":1B1E
      Height          =   7095
      Left            =   2400
      OleObjectBlob   =   "frmEdiNota.frx":1B32
      TabIndex        =   15
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmEdiNf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxBuscar_Click()
    
    End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sOrd As String
    If cmbxOrder.ListIndex = 0 Then
        sOrd = "nota_sai desc "
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrd = "Nome_cli, nota_sai desc"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrd = "Emissao_sai desc, nota_sai desc "
    Else
        sOrd = "emissao_sai desc, nota_sai desc "
    End If
    strSQL = "select ts.*,tc.nome_cli,tc.uf_cli from tab_clientes as tc, tab_saidasnf as ts "
    strSQL = strSQL & " where ts.cod_cli=tc.cod_cli "
    If cmndbtnTerc.Visible = True Then
        strSQL = strSQL & " and tc.nome_cli like '" & txtbxExpr.Text & "%' "
    Else
        strSQL = strSQL & " and ts.nota_sai like '%" & txtbxExpr.Text & "%' "
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and ts.emissao_sai < " & F_PesqData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and ts.emissao_sai > " & F_PesqData(txtbxDataIni.Text)
    End If
    strSQL = strSQL & " and ts.serie_sai not in('CUPM','CANC')"
    strSQL = strSQL & " order by " & sOrd
    dtNot.RecordSource = strSQL
    dtNot.Refresh
End Sub

Private Sub cmndbtnCancelar_Click()
    Dim strSQL As String
    With dtNot.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !situacao_sai = "3" Then
            MsgBox "A nota fiscal já foi cancelada!", vbExclamation
            Exit Sub
        End If
        If Not F_senha Then
            Exit Sub
        End If
        If MsgBox("Deseja Realmente Excluir a Nota Fiscal " & !nota_sai & _
            " ?", vbYesNo + vbCritical) = vbYes Then
            
            
            P_cancelaNotaFiscalSaida !seq_sai, ""
            
            
            dtNot.Refresh
        End If
    End With
End Sub

Private Sub cmndbtnEdit_Click()
    
    
    If dtNot.Recordset.RecordCount <= 0 Or dbgrd.Row = -1 Then
        MsgBox "Não há notas fiscais nesta pesquisa!", vbExclamation
        Exit Sub
    End If
    
    With dtNot.Recordset
        If !situacao_sai <> "1" Then
            MsgBox "A Nota Não está Aberta!", vbExclamation
            Exit Sub
        End If
        
        frmGeraNota.Tag = !seq_sai
        
        frmGeraNota.Show 1
    End With
End Sub

Private Sub cmndbtnImprGrade_Click()
    P_DbgrdToHtml dbgrd.caption & "De: " & txtbxDataIni.Text & " até " & txtbxdataFin.Text, _
        dbgrd, dtNot
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_Erro
    With dtNot.Recordset
        imprimenota !seq_sai
    
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNew_Click()
    frmGeraNota.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    dtNot.Recordset.Close
    Unload Me
End Sub

Private Sub cmndbtnTerc_Click()
    Dim c As tCliente
    c = Clientes.F_BuscaCliente("C", txtbxCliente.Text)
    If c.cod_cli <> "" Then
        txtbxCliente.Text = c.nome_Cli
        txtbxCliente.Tag = c.nome_Cli
        cmndbtnBuscar_Click
        txtbxExpr.SetFocus
    End If
    
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtNot_Reposition()
    With dtNot.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        
        
    End With
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        
        
        
    End If
End Sub

Private Sub Form_Load()
   ' Me.Width = 10560
   ' Me.Height = 6585
    P_CentraForm Me
    dtNot.Connect = banco.Connect
    
    cmbxOrder.ListIndex = 0
    
    cmbxBuscar_Click
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Resize()
    dbgrd.Width = Me.Width - dbgrd.Left - 200
    dbgrd.Height = Me.Height - dbgrd.Top - 400
    
End Sub

Private Sub txtbxCliente_Change()
    txtbxCliente.Tag = ""
End Sub

Private Sub txtbxDataFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataFin_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        SendKeys Chr(9)
        cmndbtnBuscar_Click
    End If
    
End Sub


