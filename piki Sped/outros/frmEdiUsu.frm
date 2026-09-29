VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frmEdiUsu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Usuários"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmEdiUsu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8040
   Begin VB.Data dtProd 
      Caption         =   "Edição de Gêneros"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.TextBox txtbxBuscar 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Digite o nome do funcionário e tecle <Enter>"
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmEdiUsu.frx":014A
         Left            =   120
         List            =   "frmEdiUsu.frx":0154
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Expressão:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmEdiUsu.frx":017A
      Height          =   4215
      Left            =   2400
      OleObjectBlob   =   "frmEdiUsu.frx":018F
      TabIndex        =   9
      Top             =   120
      Width           =   5535
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   100
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
      Begin VB.CommandButton cmndbtnDefUsu 
         Caption         =   "F8 - &Def. de Usuário"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiUsu.frx":0ECE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1500
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "F10 - &Sair"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiUsu.frx":1018
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "F3 - &Editar"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiUsu.frx":1162
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "F2 - &Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiUsu.frx":12AC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmEdiUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxStatus_Click()
    P_Busca
End Sub

Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxStatus.ListIndex = -1
    ElseIf KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnDefUsu_Click()
    If strCadusu_usu <> "S" Then
        MsgBox "VocÊ não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    With dtProd.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        frmDefUsu.Tag = !cod_fun
        frmDefUsu.caption = "Definições de: " & !nome_fun
        frmDefUsu.Show 1
    End With
End Sub

Private Sub cmndbtnEdit_Click()
     
    Dim fr As New frmInFun
    On Error Resume Next
    With dtProd.Recordset
        If .RecordCount <= 0 Then
            MsgBox "Não há o que ser editado!", vbExclamation
            Exit Sub
        End If
        fr.Tag = !cod_fun
        fr.txtbxCPF.Text = !cpf_fun
        fr.txtbxRG_Fun.Text = !rg_fun
        fr.txtbxEndereco_Fun.Text = !endereco_fun
        fr.txtbxGuerra.Text = !guerra_fun
        fr.txtbxNome_Fun.Text = !nome_fun
        fr.txtbxTelefone_Fun.Text = !telefone_fun
        fr.cmbxCargo.Text = !nome_car
        If !status_fun = "1" Then
            fr.cmbxStatus.ListIndex = 0
        Else
            fr.cmbxStatus.ListIndex = 1
        End If
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmIncUsu
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
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

Private Sub Form_Activate()
    MudaAparencia Me
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
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dtProd.Connect = Banco.Connect
    cmbxStatus.ListIndex = 0
End Sub

Sub P_Busca()
    Dim strSQL As String
    strSQL = "select tf.*,tu.ultacesso_usu,tc.nome_car "
    strSQL = strSQL & "from tab_funcionarios  as tf, tab_usuarios as tu,tab_cargos as tc  "
    strSQL = strSQL & " where tf.cod_fun=tu.cod_fun and tc.cod_car=tf.cod_car "
    
    strSQL = strSQL & " and tf.guerra_fun like '" & txtbxBuscar.Text & "%' "
    If cmbxStatus.ListIndex <> -1 Then
        strSQL = strSQL & " and status_fun ='" & Left(cmbxStatus.Text, 1) & "' "
    End If
    strSQL = strSQL & " order by tf.nome_fun "
    dtProd.RecordSource = strSQL
    dtProd.Refresh
End Sub

Private Sub txtbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_Busca
    End If
End Sub

