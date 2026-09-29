VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiFun 
   Caption         =   "Edição de Funcionários"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15285
   Icon            =   "frmEdiFun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   15285
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   15225
      TabIndex        =   13
      Top             =   8985
      Width           =   15285
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   15255
      TabIndex        =   10
      Top             =   0
      Width           =   15285
   End
   Begin VB.PictureBox pcBotoes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   15255
      TabIndex        =   9
      Top             =   7260
      Width           =   15285
      Begin VB.CommandButton cmdSair 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   13560
         Picture         =   "frmEdiFun.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelec 
         Caption         =   "&Selecionar"
         Height          =   705
         Left            =   4200
         Picture         =   "frmEdiFun.frx":0E18
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Apenas Funcionários Ativos"
         Height          =   195
         Left            =   4080
         TabIndex        =   11
         Top             =   480
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   315
         Left            =   120
         MaxLength       =   40
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiFun.frx":115A
         Left            =   2280
         List            =   "frmEdiFun.frx":116D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   705
         Left            =   120
         Picture         =   "frmEdiFun.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   705
         Left            =   2160
         Picture         =   "frmEdiFun.frx":151E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "&Histórico"
         Height          =   705
         Left            =   6240
         Picture         =   "frmEdiFun.frx":1668
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnPermiss 
         Caption         =   "Permissões"
         Height          =   705
         Left            =   8280
         Picture         =   "frmEdiFun.frx":1F32
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnDeb 
         Caption         =   "&Débitos"
         Height          =   705
         Left            =   10320
         Picture         =   "frmEdiFun.frx":1FE7
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Pesquisar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   2280
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiFun.frx":2131
      Height          =   4095
      Left            =   0
      TabIndex        =   12
      Top             =   240
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "cod_fun"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nome_fun"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "guerra_Fun"
         Caption         =   "Chama-se"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "cpf_fun"
         Caption         =   "CPF"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "nome_car"
         Caption         =   "Cargo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4814,929
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents AdoFun As ADODB.Recordset
Attribute AdoFun.VB_VarHelpID = -1
Public apenas_vendedores As Boolean
Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub AdoFun_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    
End Sub

Private Sub chkStatus_Click()
    cmndbtnLocaliza_Click
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnCard_Click()
    'P_Cartao AdoFun!cod_cli
End Sub

Private Sub cmdSair_Click()
    On Error Resume Next
    If Me.tag <> "" Then
        Me.tag = ""
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub cmdSelec_Click()
    Me.Hide
End Sub

Private Sub cmndbtnDeb_Click()
    On Error GoTo trata_Erro
    frmParcelasTer.txtbxNomeCli.Text = AdoFun!nome_Cli
    frmParcelasTer.txtbxNomeCli.tag = AdoFun!cod_cli
    frmParcelasTer.tag = AdoFun!cod_cli
    frmParcelasTer.Show 1
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadFun
    If AdoFun.RecordCount <= 0 Then
        MsgBox "Não há o que ser editado!", vbCritical
        Exit Sub
    End If
    strHora = Time
    fr.tag = AdoFun!cod_fun
    
    fr.Show 1
End Sub

Private Sub cmndbtnHist_Click()
'    On Error GoTo trata_erro
    With AdoFun
        If .RecordCount = 0 Then Exit Sub
        P_histCli !cod_cli
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnLocaliza_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    On Error GoTo trata_Erro
    strSQL = "select tf.*,tc.nome_car,tc.venda_car from tab_funcionarios as tf,tab_cargos as tc  where "
    
     
          strSQL = strSQL & "(Nome_fun like '" & txtbxExpr.Text & "%' "
    
        strSQL = strSQL & " or cod_fun like '" & txtbxExpr.Text & "%' "
        
        strSQL = strSQL & "or cpf_fun like '" & F_FormataCPF(txtbxExpr.Text) & "' "
    
        strSQL = strSQL & "or cpf_fun like '" & F_FormataCGC(txtbxExpr.Text) & "' "
    
        strSQL = strSQL & "or rg_fun like '" & txtbxExpr.Text & "') "
    
    strSQL = strSQL & " and tc.cod_car=tf.cod_car "
    
    If apenas_vendedores Then
        strSQL = strSQL & " and venda_car='S'"
    
    End If
    
    If chkStatus.Value = 1 Then
        strSQL = strSQL & " and status_fun=1"
    End If
    If cmbxOrd.ListIndex = 0 Then
        sOrderBy = "tf.cod_fun"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrderBy = "Nome_fun,cpf_fun"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrderBy = "cpf_fun"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrderBy = "Cidade_fun,uf_fun,NOme_fun"
    ElseIf cmbxOrd.ListIndex = 4 Then
        sOrderBy = "UF_fun,Cidade_fun"
    End If
    
    strSQL = strSQL & " order by " & sOrderBy
    abreObjetoAdo AdoFun, strSQL, dtgrd
    If AdoFun.RecordCount = 1 And Me.tag <> "" Then
        Me.tag = AdoFun!cod_fun
        Me.Hide
        Exit Sub
    End If
    dtgrd.caption = "Cadastrados...: " & F_BuscaUmCampo("count(cod_fun)", "tab_funcionarios", "")
    dtgrd.caption = dtgrd.caption & " - Listados...: " & AdoFun.RecordCount
Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadFun
    
    If strCadFun_Usu <> "S" Then
        MsgBox "Você não tem permissão para cadastrar funcionários", vbInformation
        Exit Sub
    End If
    
    
    fr.Show 0
    
    
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmndbtnPremiacao_Click()
    On Error GoTo trata_Erro
    'P_PremiacoesCli AdoFun!cod_cli
    Exit Sub
trata_Erro:
End Sub

Private Sub cmndbtnPromocoes_Click()
    On Error GoTo trata_Erro
    'P_PontosCli AdoFun!cod_cli
    Exit Sub
trata_Erro:
End Sub

Private Sub cmndbtnRes_Click()
    frmReserva.Show 0
End Sub

Private Sub cmndbtnPermiss_Click()
    
    If strCadusu_usu <> "S" Then
        MsgBox "VocÊ não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    If strCadFun_Usu <> "S" Then
        MsgBox "Não tem permissão para cadastrar funcionários ou alterar permissões!", vbExclamation
        Exit Sub
    End If
    frmDefUsu.tag = AdoFun!cod_fun
    frmDefUsu.Show 1
    
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub Command1_Click()
    Dim c As Clientes.cliente
    c = Clientes.F_BuscaCliente("")
    
    MsgBox c.t(0).num_tel & c.t(1).num_tel
    
    
    
End Sub

Private Sub dtgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        cmdSelec_Click
    End If
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

'Private Sub buscaTel()
'    Dim strSQL As String
'    On Error Resume Next
'    With AdoFun
'
'
'        strSQL = "select * from tab_telefones where cod_cli='" & !cod_fun & "' and tipo_ter='U'"
'        FechaObjeto AdoTel
'        AdoTel.Open strSQL, Conn_Client
'        adctel.Refresh
'    End With
'End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    txtbxExpr.SetFocus
    If Me.tag = "" Then
        cmdSelec.Enabled = False
    End If
    cmndbtnLocaliza_Click
End Sub

Private Sub Form_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    resizeForm Me
    'P_CentraForm Me
    Set AdoFun = New ADODB.Recordset
    Set AdoTel = New ADODB.Recordset
    
    
    cmbxOrd.ListIndex = 0
    cmndbtnLocaliza_Click
    'txtbxExpr.SetFocus
End Sub

Private Sub Form_Resize()
    'On Error Resume Next
    
    PosicionaGrade dtgrd, Me, pcBotoes
    
End Sub

Private Sub pcBotoes_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
    If KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
    
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
    
End Sub
