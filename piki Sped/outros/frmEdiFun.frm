VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frmEdiFun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Funcionários"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frmEdiFun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8010
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.ComboBox cmnxOrder 
         Height          =   315
         ItemData        =   "frmEdiFun.frx":014A
         Left            =   120
         List            =   "frmEdiFun.frx":0157
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmEdiFun.frx":0170
         Left            =   120
         List            =   "frmEdiFun.frx":017A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtbxBuscar 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   615
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Expressão"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   90
      TabIndex        =   7
      Top             =   2040
      Width           =   2175
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "F10 - &Sair"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiFun.frx":01A0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "F3 - &Editar"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiFun.frx":052A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "F2 - &Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiFun.frx":0674
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
   End
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
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmEdiFun.frx":09FE
      Height          =   4335
      Left            =   2400
      OleObjectBlob   =   "frmEdiFun.frx":0A13
      TabIndex        =   11
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmEdiFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxStatus_Click()
    P_Busca
End Sub

Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmbxStatus.ListIndex = -1
    If KeyCode = vbKeyReturn Then SendKeys Chr(9)
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
    Dim fr As New frmInFun
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmnxOrder_Click()
    P_Busca
End Sub

Private Sub cmnxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmndbtnNew_Click
    ElseIf KeyCode = vbKeyF3 Then
        cmndbtnEdit_Click
    ElseIf KeyCode = vbKeyF10 Then
        cmndbtnSair_Click
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dtProd.Connect = Banco.Connect
    cmbxStatus.ListIndex = 0
    cmnxOrder.ListIndex = 0
End Sub

Sub P_Busca()
    Dim strSQL As String
    strSQL = "select tf.*,tc.nome_car from tab_funcionarios  as tf, tab_cargos as tc "
    strSQL = strSQL & " where tf.cod_car=tc.cod_car "
    
    strSQL = strSQL & " and nome_fun like '" & txtbxBuscar.Text & "%' "
    If cmbxStatus.ListIndex > -1 Then
        strSQL = strSQL & "and status_fun='" & Left(cmbxStatus.Text, 1) & "'"
    End If
    If cmnxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by guerra_fun "
    ElseIf cmnxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by cod_fun "
    ElseIf cmnxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by tc.nome_car "
    End If
    
    dtProd.RecordSource = strSQL
    dtProd.Refresh
End Sub

Private Sub txtbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_Busca
    End If
End Sub
