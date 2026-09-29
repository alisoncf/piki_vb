VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSelectCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecione um Cliente"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "frmSelectCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmSelectCliente.frx":0442
      Height          =   5175
      Left            =   0
      OleObjectBlob   =   "frmSelectCliente.frx":0456
      TabIndex        =   2
      Top             =   0
      Width           =   5535
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
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtbxCli 
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   1
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Razão Social:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   990
   End
End
Attribute VB_Name = "frmSelectCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dbgrd_dblClick()
    If dbgrd.Row = -1 Then
        MsgBox "Não um cliente para ser selecionado!", vbCritical
    Else
        If strCOD_CLI = dtCli.Recordset!cod_cli Then
            P_Status "Cliente já Selecionado! Operação Cancelada", False
            MsgBox "este cliente já está selecionado", vbInformation
            P_Status ""
        Else
            strCOD_CLI = dtCli.Recordset!cod_cli
            strCNPJ_Cli = dtCli.Recordset!cnpj_Cli
            strIE_CLI = dtCli.Recordset!ie_Cli
            strNome_CLI = dtCli.Recordset!nome_cli
            strUF_CLI = dtCli.Recordset!uf_Cli
            strEndereco_CLI = dtCli.Recordset!endereco_cli
            strBairro_CLI = dtCli.Recordset!bairro_Cli
            strCidade_CLI = dtCli.Recordset!cidade_Cli
            strCEP_CLI = dtCli.Recordset!cep_Cli
            strFone_CLI = dtCli.Recordset!telefone_cli
            
            mdiFrm.StatusBar1.Panels(1).Text = dtCli.Recordset!cod_cli & "-" & _
            dtCli.Recordset!nome_cli
            
            Me.Hide
            P_Status "Selecionando Cliente", 2, False
            MsgBox "Cliente selecionado com sucesso!" & Chr(13) & dtCli.Recordset!nome_cli, vbInformation
            P_Status ""
            Unload Me
        End If
    End If
End Sub

Private Sub dbgrd_GotFocus()
    
    dbgrd.HeadFont.Bold = True
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_dblClick
    Else
        If F_CharValido(Chr(KeyCode)) Then
            If txtbxCli.Text = "" Then
                txtbxCli.Text = Chr(KeyCode)
            End If
            txtbxCli.SetFocus
        End If
        
        
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
    P_CentraForm Me, 400
    P_Status "Selecione um cliente e tecle <Enter> "
    dtCli.Connect = banco.Connect
    P_Busca
    
End Sub
Sub P_Busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from tab_clientes "
    strSQL = strSQL & " where nome_cli like '" & UCase(txtbxCli.Text) & "%'"
    strSQL = strSQL & " order by nome_Cli "
    dtCli.RecordSource = strSQL
    dtCli.Refresh
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation, constNome_sis
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    mdiFrm.mnEditCli.Enabled = True
    P_Status ""
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
