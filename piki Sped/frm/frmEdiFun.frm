VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiFun 
   Caption         =   "Edição de Funcionários"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14325
   Icon            =   "frmEdiFun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   14325
   Begin VB.PictureBox pcTop 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   14295
      TabIndex        =   0
      Top             =   6750
      Width           =   14325
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "&Selecionar"
         Height          =   690
         Left            =   4200
         Picture         =   "frmEdiFun.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtbxBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         ToolTipText     =   "Digite o nome do funcionário e tecle <Enter>"
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmEdiFun.frx":064C
         Left            =   120
         List            =   "frmEdiFun.frx":0656
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnDefUsu 
         Caption         =   "F8 - &Def. de Usuário"
         Height          =   685
         Left            =   9240
         Picture         =   "frmEdiFun.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "F3 - &Editar"
         Height          =   685
         Left            =   7560
         Picture         =   "frmEdiFun.frx":07C6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "F2 - &Novo"
         Height          =   690
         Left            =   5880
         Picture         =   "frmEdiFun.frx":0910
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   685
         Left            =   12600
         Picture         =   "frmEdiFun.frx":0A74
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Redefinir Senha"
         Height          =   685
         Left            =   10920
         Picture         =   "frmEdiFun.frx":0BDF
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/CPF/Código:"
         Height          =   195
         Left            =   1920
         TabIndex        =   9
         Top             =   0
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiFun.frx":0D29
      Height          =   5415
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmbxStatus_click()
    P_Busca
End Sub

Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxStatus.ListIndex = -1
    ElseIf KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbnDefUsuNew_Click()
    P_DefUsu adc.Recordset!cod_fun
End Sub

Private Sub cmdSelecionar_Click()
    '
    Me.Hide
End Sub

Private Sub cmndbtnDefUsu_Click()
    If strCadusu_usu <> "S" Then
        MsgBox "VocÊ não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    If strCadFun_Usu <> "S" Then
        MsgBox "Não tem permissão para cadastrar funcionários ou alterar permissões!", vbExclamation
        Exit Sub
    End If
   
     
    With ado
        If .RecordCount = 0 Then
            Exit Sub
        End If
        frmDefUsu.tag = !cod_fun
        frmDefUsu.caption = "Definições de: " & !nome_fun
        frmDefUsu.Show 1
    End With
End Sub

Private Sub cmndbtnEdit_Click()
     
    Dim fr As New frmCadFun
    On Error Resume Next
    
    If strCadusu_usu <> "S" Then
        MsgBox "Você não permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    With ado
        If .RecordCount <= 0 Then
            MsgBox "Não há o que ser editado!", vbExclamation
            Exit Sub
        End If
        fr.tag = !cod_fun
        'fr.txtbxCpf.Text = !cpf_fun
        'fr.txtbxRG_Fun.Text = !rg_fun
        'fr.txtbxEndereco_Fun.Text = !endereco_fun
        'fr.txtbxGuerra.Text = !guerra_fun
       ' fr.cmbxFuncionario.Text = !guerra_fun
       ' P_LocalizaOperacao !nivel_usu, fr.cmbxTipo_Usu
        
                
        'fr.cmbxCargo.Text = !nome_car
        'If !status_fun = "1" Then
        '    fr.cmbxStatus.ListIndex = 0
        'Else
        '    fr.cmbxStatus.ListIndex = 1
        'End If
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnNew_Click()
    
    If strCadFun_Usu <> "S" Then
        MsgBox "Você não tem permissão para cadastrar funcionários", vbInformation
        Exit Sub
    End If
    frmCadFun.Show 1
    
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub DBGrid1_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.MarqueeStyle = 3
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.MarqueeStyle = 0
End Sub

Private Sub Command1_Click()
    If strCadusu_usu <> "S" Then
        MsgBox "Não permitido!", vbExclamation
        
    End If
    
    strSQL = "Update tab_usuarios set senha_usu='' where cod_fun='" & ado!cod_fun & "'"
    conn.Execute strSQL
    MsgBox "Senha redefinida", vbExclamation
    
    
End Sub

Private Sub dtgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        cmdSelecionar_Click
    End If
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        P_Busca
        cmndbtnDefUsu.Visible = False
        cmndbtnEdit.Visible = False
        cmndbtnDefUsu.Visible = False
        cmndbtnNew.Visible = False
        cmbxStatus.Enabled = False
        Me.caption = "Escolha um funcionário"
    Else
        cmdSelecionar.Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmndbtnNew_Click
    ElseIf KeyCode = vbKeyF3 Then
        cmndbtnEdit_Click
    ElseIf KeyCode = vbKeyF8 Then
        cmndbtnDefUsu_Click
    ElseIf KeyCode = vbKeyF10 Then
        cmndbtnSair_Click
    ElseIf KeyCode = vbKeyDown Then
        dtgrd.SetFocus
        
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    resizeForm Me
    cmbxStatus.ListIndex = 0
    
End Sub

Sub P_Busca()
    Dim strSQL As String
    strSQL = "select tf.*,tu.ultacesso_usu,tc.nome_car,tu.nivel_usu "
    strSQL = strSQL & " from tab_funcionarios  as tf left join tab_usuarios as tu on tu.cod_fun=tf.cod_fun "
    strSQL = strSQL & " left join tab_cargos as tc on tc.cod_car=tf.cod_car "
    
    strSQL = strSQL & " where 1>0 "
        
    strSQL = strSQL & " and (tf.guerra_fun like '%" & txtbxBuscar.Text & "%' "
    strSQL = strSQL & " or tf.nome_fun like '%" & txtbxBuscar.Text & "%' "
    strSQL = strSQL & " or tf.cod_fun like '%" & txtbxBuscar.Text & "%' )"
    If cmbxStatus.ListIndex <> -1 Then
        strSQL = strSQL & " and status_fun ='" & Left(cmbxStatus.Text, 1) & "' "
    End If
    strSQL = strSQL & " order by tf.nome_fun "
    
    abreObjetoAdo ado, strSQL, dtgrd
    
    If Me.tag <> "" And ado.RecordCount = 1 Then
        Me.Hide
    End If
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    PosicionaGrade dtgrd, Me, pcTop
    posicionaADireita cmndbtnSair, pcTop
    
End Sub

Private Sub txtbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_Busca
    ElseIf KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
End Sub


