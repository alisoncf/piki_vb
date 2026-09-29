VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmBusca_Ter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca de Terceiros"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "frmBusca_Ter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbxBusca 
      Height          =   315
      ItemData        =   "frmBusca_Ter.frx":0442
      Left            =   120
      List            =   "frmBusca_Ter.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Data dtCli 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtbxCli 
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   3
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton cmndbtnNewCli 
      Caption         =   "&Cadastro"
      Height          =   615
      Left            =   5040
      Picture         =   "frmBusca_Ter.frx":0476
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnCad 
      Caption         =   "&Selecionar"
      Height          =   615
      Left            =   4080
      Picture         =   "frmBusca_Ter.frx":0800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   855
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmBusca_Ter.frx":0B0A
      Height          =   4575
      Left            =   120
      OleObjectBlob   =   "frmBusca_Ter.frx":0B1E
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label lblExp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Expressão:"
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   4680
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Buscar Por:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   825
   End
End
Attribute VB_Name = "frmBusca_Ter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxBusca_Click()
    lblExp.caption = cmbxBusca.Text
End Sub

Private Sub cmndbtnCad_Click()
    Dim sTipo As String
    If dbgrd.Row = -1 Then
        MsgBox "Não há um terceiro selecionado!", vbExclamation
    Else
        Me.Tag = dtCli.Recordset!cod_ter
        Me.Hide
    End If
End Sub
Sub P_Inclui(sTipo As String)
    Dim strSQL As String
    With dtCli.Recordset
        If F_BuscaUmCampo("COd_Cli", "tab_tercli", " cod_cli='" & strCOD_CLI & "' " _
            & " and cod_ter='" & !cod_ter & "' ") = "" Then
            strSQL = "insert into tab_tercli(cod_Cli,cod_Ter,tipo_Ter)Values("
            strSQL = strSQL & "'" & strCOD_CLI & "',"
            strSQL = strSQL & "'" & !cod_ter & "',"
            strSQL = strSQL & "'" & sTipo & "') "
            banco.Execute strSQL
            MsgBox "Terceiro Incluído com sucesso!", vbInformation
        Else
            MsgBox "Este terceiro já está cadastrado para este cliente", vbCritical
        End If
    End With
End Sub
Private Sub cmndbtnNewCli_Click()
    On Error Resume Next
    frmSelectTer.Show 1
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub dbgrd_DblClick()
    If cmndbtnCad.Enabled = False Then
        
    Else
        cmndbtnCad_Click
    End If
    
End Sub

Private Sub dbgrd_GotFocus()
    
    dbgrd.HeadFont.Bold = True
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.HeadFont.Bold = True
End Sub

Private Sub Form_Activate()
    dtCli.Connect = banco.Connect
    cmbxBusca.ListIndex = 0
    txtbxCli_Change
    Me.caption = "Cliente: " & mdiFrm.StatusBar1.Panels(1).Text
    MudaAparencia Me
End Sub

Sub P_Busca()
    Dim strSQL As String
    
        strSQL = "select tt.*,ttc.tipo_ter "
        strSQL = strSQL & " from tab_terceiros as tt, tab_tercli as ttc, tab_clientes as tc "
        strSQL = strSQL & " where tt.cod_ter=ttc.cod_ter "
        strSQL = strSQL & " and ttc.cod_cli = tc.cod_cli "
        strSQL = strSQL & " and tc.cod_Cli = '" & strCOD_CLI & "' "
        strSQL = strSQL & " and nome_Ter like '" & txtbxCli.Text & "%' "
        strSQL = strSQL & " order by nome_ter "
    
    

    dtCli.RecordSource = strSQL
    dtCli.Refresh
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    
    mdiFrm.mnEditCli.Enabled = True
End Sub

Private Sub mnCadastrar_Click()
    
End Sub

Private Sub optnTer_Click(Index As Integer)
    If Index = 1 Then
        cmndbtnCad.Enabled = False
    Else
        cmndbtnCad.Enabled = True
    End If
    P_Busca
End Sub

Private Sub txtbxCli_Change()
    P_Busca
End Sub

Private Sub txtbxCli_GotFocus()
    P_SelecTexto txtbxCli
End Sub

Private Sub txtbxCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then
        dbgrd.SetFocus
    End If
End Sub

Private Sub txtbxCli_LostFocus()
    P_AjustaTexto txtbxCli
End Sub




