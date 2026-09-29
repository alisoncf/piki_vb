VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSelectTer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terceiros Cadastrados"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmSelectTer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   5520
      Picture         =   "frmSelectTer.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   975
   End
   Begin VB.OptionButton optnTer 
      Caption         =   "&Apenas deste cliente"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   2055
   End
   Begin VB.OptionButton optnTer 
      Caption         =   "&Todos os Terceiros"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmndbtnEdit 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4440
      Picture         =   "frmSelectTer.frx":07CC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmndbtnCad 
      Height          =   375
      Left            =   4800
      Picture         =   "frmSelectTer.frx":0B56
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5430
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnNewCli 
      Caption         =   "&Novo"
      Height          =   735
      Left            =   3360
      Picture         =   "frmSelectTer.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox txtbxCli 
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   1
      Top             =   5430
      Width           =   3495
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
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmSelectTer.frx":186A
      Height          =   5295
      Left            =   0
      OleObjectBlob   =   "frmSelectTer.frx":187E
      TabIndex        =   2
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Razão Social:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   990
   End
End
Attribute VB_Name = "frmSelectTer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnCad_Click()
    Dim sTipo As String
    If dbgrd.Row = -1 Then
        MsgBox "Não há um terceiro selecionado!", vbExclamation
    Else
        On Error Resume Next
        If F_BuscaUmCampo("COd_Cli", "tab_tercli", " cod_cli='" & strCOD_CLI & "' " _
            & " and cod_ter='" & dtCli.Recordset!Cod_Ter & "' ") <> "" Then
            
            MsgBox "Este terceiro já está cadastrado para este cliente", vbCritical
            Exit Sub
        End If
        frmTipo_Ter.Show 1
            sTipo = frmTipo_Ter.Tag
            Unload frmTipo_Ter
        If MsgBox("Confirma inclusão do terceiro: " & dtCli.Recordset!Nome_Ter & _
            Chr(13) & "para o cliente: " & _
            mdiFrm.StatusBar1.Panels(1).Text & "?", vbYesNo + vbQuestion) = vbYes Then
            
            P_Inclui sTipo
            
        End If
        
    End If
End Sub
Sub P_Inclui(sTipo As String)
    Dim strSQL As String
    With dtCli.Recordset
        If F_BuscaUmCampo("COd_Cli", "tab_tercli", " cod_cli='" & strCOD_CLI & "' " _
            & " and cod_ter='" & !Cod_Ter & "' ") = "" Then
            strSQL = "insert into tab_tercli(cod_Cli,cod_Ter,tipo_Ter)Values("
            strSQL = strSQL & "'" & strCOD_CLI & "',"
            strSQL = strSQL & "'" & !Cod_Ter & "',"
            strSQL = strSQL & "'" & sTipo & "') "
            banco.Execute strSQL
            MsgBox "Terceiro Incluído com sucesso!", vbInformation
        Else
            MsgBox "Este terceiro já está cadastrado para este cliente", vbCritical
        End If
    End With
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_erro
    With dtCli.Recordset
        frmIncTer.Tag = !Cod_Ter
        frmIncTer.txtbxBairro.Text = !Bairro_Ter
        frmIncTer.txtbxCEP.Text = !cep_Ter
        frmIncTer.txtbxCidade.Text = !Cidade_Ter
        frmIncTer.txtbxCNPJ.Text = !cnpj_Ter
        frmIncTer.txtbxEndereco.Text = !Endereco_Ter
        frmIncTer.txtbxIE.Text = !ie_Ter
        frmIncTer.txtbxRazao.Text = !Nome_Ter
        frmIncTer.txtbxTel.Text = !Telefone_Ter
        On Error Resume Next
        frmIncTer.Show 1
        On Error GoTo trata_erro
    End With
    
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmndbtnNewCli_Click()
    frmIncTer.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_dblClick()
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
        dbgrd_dblClick
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.HeadFont.Bold = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    mdiFrm.mnEditCli.Enabled = False
    P_CentraForm Me
    dtCli.Connect = banco.Connect
    P_Busca
    Me.caption = "Cliente: " & mdiFrm.StatusBar1.Panels(1).Text
    If Me.Tag = "" Then
        cmndbtnEdit.Enabled = True
    Else
        cmndbtnEdit.Enabled = False
    End If
    mdiFrm.Enabled = False
End Sub
Sub P_Busca()
    Dim strSQL As String
    If optnTer(1).Value = True Then
        strSQL = "select tt.*,ttc.tipo_ter "
        strSQL = strSQL & " from tab_terceiros as tt, tab_tercli as ttc, tab_clientes as tc "
        strSQL = strSQL & " where tt.cod_ter=ttc.cod_ter "
        strSQL = strSQL & " and ttc.cod_cli = tc.cod_cli "
        strSQL = strSQL & " and tc.cod_Cli = '" & strCOD_CLI & "' "
        strSQL = strSQL & " order by nome_ter "
    Else
        strSQL = "select * from tab_terceiros "
        strSQL = strSQL & " where nome_ter like '" & UCase(txtbxCli.Text) & "%'"
        strSQL = strSQL & " order by nome_ter "
    End If

    dtCli.RecordSource = strSQL
    dtCli.Refresh
    If dtCli.Recordset.RecordCount > 0 Then
        cmndbtnEdit.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    mdiFrm.mnEditCli.Enabled = True
    mdiFrm.Enabled = True
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


