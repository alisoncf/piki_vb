VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEdiCai 
   Caption         =   "Edição de Fluxo de Caixa"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   13890
   Icon            =   "frmEdiCai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   13890
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridCai 
      Height          =   1095
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   1931
      _Version        =   393216
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.PictureBox pcTotal 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   13890
      TabIndex        =   25
      Top             =   7320
      Width           =   13890
   End
   Begin VB.PictureBox pcLeft 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   13860
      TabIndex        =   7
      Top             =   0
      Width           =   13890
      Begin VB.ComboBox cmbEsp 
         Height          =   315
         ItemData        =   "frmEdiCai.frx":014A
         Left            =   9240
         List            =   "frmEdiCai.frx":0157
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cmbxTurno 
         Height          =   315
         ItemData        =   "frmEdiCai.frx":0175
         Left            =   1920
         List            =   "frmEdiCai.frx":0182
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cmbxCai 
         Height          =   315
         ItemData        =   "frmEdiCai.frx":01A0
         Left            =   120
         List            =   "frmEdiCai.frx":01AD
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         MaxLength       =   40
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiCai.frx":01CB
         Left            =   11040
         List            =   "frmEdiCai.frx":01DB
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cmbxTIpo 
         Height          =   315
         ItemData        =   "frmEdiCai.frx":0203
         Left            =   7800
         List            =   "frmEdiCai.frx":0210
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Pesquisar"
         Height          =   645
         Left            =   12360
         Picture         =   "frmEdiCai.frx":022E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espécie:"
         Height          =   195
         Left            =   9240
         TabIndex        =   24
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         Height          =   195
         Left            =   1920
         TabIndex        =   22
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   435
      End
      Begin VB.Label lbldataFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   4920
         TabIndex        =   18
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lblDataIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   3840
         TabIndex        =   17
         Top             =   120
         Width           =   840
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar por:"
         Height          =   195
         Left            =   6000
         TabIndex        =   16
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   11040
         TabIndex        =   15
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   7800
         TabIndex        =   14
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   13860
      TabIndex        =   0
      Top             =   7455
      Width           =   13890
      Begin VB.CommandButton cmdVales 
         Caption         =   "&Vales"
         Height          =   615
         Left            =   10200
         Picture         =   "frmEdiCai.frx":05B8
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   615
         Left            =   8520
         Picture         =   "frmEdiCai.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   6840
         Picture         =   "frmEdiCai.frx":0BF3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   12000
         Picture         =   "frmEdiCai.frx":1000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnFechaCAi 
         Caption         =   "Fec&hamento"
         Height          =   615
         Left            =   5160
         Picture         =   "frmEdiCai.frx":116B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3480
         Picture         =   "frmEdiCai.frx":12C4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1800
         Picture         =   "frmEdiCai.frx":140E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiCai.frx":1558
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registros"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1080
      Width           =   13455
   End
   Begin VB.Menu mnOpt 
      Caption         =   "&Opções"
      Begin VB.Menu mnCopy 
         Caption         =   "Copiar Texto"
      End
      Begin VB.Menu mnEdit 
         Caption         =   "&Editar"
      End
      Begin VB.Menu mnNovo 
         Caption         =   "Novo Lançamento"
      End
      Begin VB.Menu mnExcl 
         Caption         =   "Excluir"
      End
      Begin VB.Menu mnExport 
         Caption         =   "Exportar"
      End
      Begin VB.Menu mnSair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmEdiCai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Private Sub Text1_Change()

End Sub

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    If ado!num_ven = "" Then
        cmndbtnEdit.Enabled = True
        cmndbtnExcl.Enabled = True
    Else
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
    End If
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmdExportar_Click()
    ExportaMSFLXGridToCSV gridCai
End Sub

Private Sub cmdVales_Click()
    frmEdiVales.Show 0, mdiFrm
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    
    
  '  On Error GoTo trata_Erro
    
    
    If txtbxdataFin.Visible = True Then
        If F_DataMaior(txtbxDataIni.text, txtbxdataFin.text) Then
            MsgBox "A data final deve ser maior que a inicial!", vbExclamation
            txtbxdataFin.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "select distinct tc.*,tc.entrada_cai-tc.saida_cai as valor,tb.nome_cai,tf.guerra_fun from tab_caixa as tc inner join "
    'strSQL = strSQL & " left join tab_baixafin as tbf on tbf.seq_fin=tc.seq_cai "
    'strSQL = strSQL & " inner join tab_baixaitd2 as tbi on tbi.seq_cab=tbf.seq_cab "
    strSQL = strSQL & " tab_caixas as tb on tc.cod_cai=tb.cod_cai "
    strSQL = strSQL & " inner join tab_funcionarios as tf on tc.cod_usu=tf.cod_fun "
    
    
   
        strSQL = strSQL & " where "
 
    
    
    strSQL = strSQL & " 1>0  and ("
    
    strSQL = strSQL & " historico_cai like '" & txtbxExpr.text & "%'"
    

    strSQL = strSQL & " or num_ven like '" & txtbxExpr.text & "%'"
    If CCur(FormatVal(txtbxExpr.text)) > 0 Then
        strSQL = strSQL & " or valor_cai = " & F_GravaDec(txtbxExpr.text)
    End If
    strSQL = strSQL & ") and data_cai >= " & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and data_cai <= " & F_GravaData(txtbxdataFin.text)
    strSQL = strSQL & " and excluido_cai=false "
    If cmbxTipo.ListIndex <> 0 Then
        strSQL = strSQL & " and tipo_cai='" & Left(cmbxTipo.text, 1) & "' "
    End If
    If cmbxCai.text <> "" Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCai.text, 2) & "' "
    End If
    If cmbxTurno.text <> "" Then
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.text, 1) & "' "
    End If
    
  '  If pcLeft.BackColor <> vbYellow Then
   '     strSQL = strSQL & " and visivel_cai='S'"
    
    
  '  End If
    
    If cmbxOrd.ListIndex = 0 Then
        strSQL = strSQL & " order by data_cai,seq_cai"
    ElseIf cmbxOrd.ListIndex = 1 Then
        strSQL = strSQL & " order by seq_cai"
    ElseIf cmbxOrd.ListIndex = 2 Then
        strSQL = strSQL & " order by historico_cai, seq_cai"
    ElseIf cmbxOrd.ListIndex = 3 Then
        strSQL = strSQL & " order by valor_cai,seq_cai"
    End If
        
    
    abreObjetoAdo ado, strSQL
    lblCaption.caption = ado.RecordCount & " registros "
    'ado.RecordSource = strSQL
    'ado.Refresh
    
 gridCai.Visible = False
    
    carregar_grid
    gridCai.Visible = True
    
    

trata_Erro:
Exit Sub
P_MostraErro
End Sub
Sub buscaF8()
    Dim strSQL As String
    If txtbxdataFin.Visible = True Then
        If F_DataMaior(txtbxDataIni.text, txtbxdataFin.text) Then
            MsgBox "A data final deve ser maior que a inicial!", vbExclamation
            txtbxdataFin.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "select distinct tc.*tc.entrada_cai-tc.saida_cai as valor,tb.nome_cai from tab_caixa as tc, tab_caixas as tb "
    If chkVales.Value = 1 Then
        strSQL = strSQL & " ,tab_valescai as tv "
        strSQL = strSQL & " where tv.seq_cai=tc.seq_cai and "
    Else
        strSQL = strSQL & " where "
    End If
    
    strSQL = strSQL & " tc.excluido_cai=false and "
    strSQL = strSQL & "tc.cod_cai=tb.cod_cai  and "
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " historico_cai like '" & txtbxExpr.text & "%'"
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & " num_ven like '" & txtbxExpr.text & "%'"
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " valor_cai = " & F_GravaDec(txtbxExpr.text)
    End If
    strSQL = strSQL & " and data_cai >= " & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and data_cai <= " & F_GravaData(txtbxdataFin.text)
    If cmbxTipo.ListIndex <> 0 Then
        strSQL = strSQL & " and tipo_cai='" & Left(cmbxTipo.text, 1) & "' "
    End If
    If cmbxCai.text <> "" Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCai.text, 2) & "' "
    End If
    If cmbxTurno.text <> "" Then
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.text, 1) & "' "
    End If
    If cmbxOrd.ListIndex = 0 Then
        strSQL = strSQL & " order by data_cai,seq_cai"
    ElseIf cmbxOrd.ListIndex = 1 Then
        strSQL = strSQL & " order by seq_cai"
    ElseIf cmbxOrd.ListIndex = 2 Then
        strSQL = strSQL & " order by historico_cai, seq_cai"
    ElseIf cmbxOrd.ListIndex = 3 Then
        strSQL = strSQL & " order by valor_cai,seq_cai"
    End If
        
    
    
    If optnAsc.Value Then
        strSQL = strSQL & " asc "
    Else
        strSQL = strSQL & " desc "
    End If
    FechaObjeto ado
    ado.Open strSQL
    
    P_Resumo
End Sub

Sub P_Resumo()
Dim cCheDeb As Currency
    Dim cCheCred As Currency
    Dim cDinDeb As Currency
    Dim cDinCred As Currency
    
    With ado
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            
            If !esp_cai = "DIN" Then
                If !tipo_cai = "C" Then
                    cDinCred = cDinCred + !valor_Cai
                Else
                    cDinDeb = cDinDeb + !valor_Cai
                End If
            Else
                If !tipo_cai = "C" Then
                    cCheCred = cCheCred + !valor_Cai
                Else
                    cCheDeb = cCheDeb + !valor_Cai
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    End Sub
    'lblCred_Che.caption = FormatVal(cCheCred)
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmLanCaixa
    With gridCai
    .col = 0
       fr.tag = .text
'        fr.txtbxData_Cai.Text = !data_cai
'        fr.txtbxHist.Text = !historico_cai
'        fr.txtbxValor_Cai.Text = FormatVal(!valor_Cai)
'        If !tipo_cai = "C" Then
'            fr.optnCerd.Value = True
'        Else
'            fr.optnDeb.Value = True
'        End If
'        fr.optnCerd.Enabled = False
'        fr.optnDeb.Enabled = False
        fr.Show 1
    End With
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    Dim NumVen As String
    Dim seqCai As String
    On Error GoTo trata_Erro
    If ado.RecordCount = 0 Then
        cmndbtnExcl.Enabled = False
        Exit Sub
    End If
    
    seqCai = GridUtil.getText(gridCai, 0)
    
    If seqCai <> "" Then
        If CAIXA.isLancamentoDuplicado(seqCai) Then
            CAIXA.excluir_lancamento seqCai
            cmndbtnBuscar_Click
        End If
        
    End If
    
    gridCai.col = 8
    
    
    
    
    If gridCai.text <> "MOV" Then
        MsgBox "Este lançamento não pode ser cancelado!", vbExclamation
        Exit Sub
    End If
    
    If strExclcai_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    gridCai.col = 0
    If MsgBox("Confirma exclusão do Lançamento? ", vbYesNo + vbCritical) = vbYes Then
        If CAIXA.excluir_lancamento(seqCai) Then
            cmndbtnBuscar_Click
        End If
    End If
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmndbtnFechaCAi_Click()
    'frmFechaCaixa.Show 1
    'frmMovDiarioCx.Show 1
    frmFechaCaixa2.Show 1
End Sub

Private Sub cmndbtnImpr_Click()
    P_MsflxToHtml "Lançamento de Caixa -  " & txtbxDataIni.text & " - " & txtbxdataFin.text, gridCai, , , "2,3,4"
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanCaixa
    fr.Show 0
 End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVale_Click()
  '  On Error Resume Next
    With ado
        P_ImprimeValeCai !seq_cai
    End With
End Sub

Private Sub cmndbtnVendasNaoCaixa_Click()
'    P_VerificaVendaNaoCaixa txtbxDataIni.Text, Left(cmbxCai.Text, 2)
End Sub

Private Sub dbgrd_GotFocus()
    
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub ado_Reposition()
    'On Error Resume Next
    If ado.RecordCount = 0 Or dbgrd.row = -1 Then
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
        Exit Sub
    
    End If
    
    If ado!num_ven _
     <> "" Then
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
    Else
        cmndbtnEdit.Enabled = True
        cmndbtnExcl.Enabled = True
    End If
End Sub


Sub carregar_grid()
    gridCai.FontName = "Arial"
    gridCai.Visible = False
     AjustaGridData gridCai, 11, 1, 0, "Seq,Data,Entrada,Saida,Saldo,Histórico,Tipo,Venda,Esp,Hora,Funcionário", _
        "Seq_cai,Data_cai,entrada_cai,saida_cai,valor,historico_cai,tipo_cai,num_ven,esp_cai,hora_cai,guerra_fun", _
        "800,1000,1000,1000,1000,5000,500,1000,1000,1800", ado, "2,3,4", True
    Dim f As fonte
    
    
    f.cor_face = vbBlue
    
    f.negrito = True
        
    ColoreLinhaGridKey gridCai, "DIN", 8, f
    ColoreLinhaGridKey gridCai, "CRT", 8, f
    ColoreLinhaGridKey gridCai, "CRD", 8, f
        
    f.cor_face = &H40C0&
        ColoreLinhaGridKey gridCai, "MOV", 8, f
    f.cor_face = &HFF8080
    ColoreLinhaGridKey gridCai, "RDI", 8, f
    ColoreLinhaGridKey gridCai, "RCC", 8, f
    ColoreLinhaGridKey gridCai, "RCD", 8, f
    gridCai.Visible = True
End Sub
Private Sub Form_Activate()
    'MudaAparencia Me
    If F_IsCaixa Then
        P_LocalizaOperacao strCod_Cai, cmbxCai
        P_LocalizaOperacao strCod_Tur, cmbxTurno
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        If pcLeft.BackColor = vbYellow Then
            pcLeft.BackColor = vbWhite
            cmndbtnBuscar_Click
        Else
            pcLeft.BackColor = vbYellow
            cmndbtnBuscar_Click
        End If
    ElseIf KeyCode = vbKeyF11 Then
        VerificaVendaDuplicada Date
    End If
    
End Sub
Sub VerificaVendaDuplicada(data As String)
    Dim strSQL As String
    strSQL = "select tv.num_ven, tc.seq_cai,tc.valor_cai  from tab_vendas as tv, tab_caixa as tc "
    strSQL = strSQL & " where tv.num_ven=tc.num_ven "
    
End Sub
Private Sub Form_Load()
  

  mbDataChanged = False
    
    Set ado = New ADODB.Recordset
    P_PreencheCaixas cmbxCai
    P_PreencheTurnos cmbxTurno
    P_PreencheEspCai cmbEsp
    cmbxOrd.ListIndex = 0
    cmbxTipo.ListIndex = 0
    txtbxdataFin.text = Date
    txtbxDataIni.text = Date
    'ado.ConnectionString = Conn.ConnectionString
    cmndbtnBuscar_Click
    resizeForm Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ' PosicionaGrade dtgrd, Me, pcTotal
    PosicionaGrade gridCai, Me, pcBot
    posicionaADireita cmndbtnSair, pcBot
    lblCaption.width = gridCai.width
End Sub

Private Sub gridCai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyC And Shift = vbKeyControl Then
        mnCopy_Click
    End If
    If KeyCode = vbKeyF5 Then
        gridCai.SelectionMode = flexSelectionFree
    End If
    If KeyCode = vbKeyF6 Then
        gridCai.SelectionMode = flexSelectionByRow
    End If
End Sub

Private Sub gridCai_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpt
    End If
End Sub

Private Sub mnCopy_Click()
    Clipboard.SetText gridCai.text
End Sub

Private Sub mnEdit_Click()
    cmndbtnEdit_Click
End Sub

Private Sub mnExcl_Click()
    cmndbtnExcl_Click
End Sub

Private Sub mnExport_Click()
    cmdExportar_Click
End Sub

Private Sub mnNovo_Click()
    cmndbtnSair_Click
End Sub

Private Sub mnSair_Click()
    cmndbtnSair_Click
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
    If F_DataMaior(txtbxDataIni.text, txtbxdataFin.text) Then
        MsgBox "A data final deve ser maior que a inicial!", vbExclamation
        txtbxdataFin.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub

