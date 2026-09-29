VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSelectCli 
   Caption         =   "Selecionar Cliente"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13905
   Icon            =   "frmSelectCli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   13905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnSelect 
      Caption         =   "&Selecionar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   8880
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   13905
      TabIndex        =   3
      Top             =   9225
      Width           =   13905
      Begin VB.TextBox txtbxNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Tag             =   "C"
         Top             =   195
         Width           =   7935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Informe uma Expressão para a busca:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   195
         Width           =   2865
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmSelectCli.frx":058A
      Height          =   8775
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   15478
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
      Caption         =   "Escolha um Cliente"
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "cod_cli"
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
         DataField       =   "nome_cli"
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
         DataField       =   "fantasia_cli"
         Caption         =   "Fantasia"
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
         DataField       =   "cnpj_cli"
         Caption         =   "CNPJ"
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
         DataField       =   "endereco_cli"
         Caption         =   "Endereço"
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
      BeginProperty Column05 
         DataField       =   "bairro_cli"
         Caption         =   "Bairro"
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
      BeginProperty Column06 
         DataField       =   "nome_cid"
         Caption         =   "Cidade"
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
      BeginProperty Column07 
         DataField       =   "uf_cli"
         Caption         =   "UF"
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
      BeginProperty Column08 
         DataField       =   "cod_cid"
         Caption         =   "Código Cidade"
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
      BeginProperty Column09 
         DataField       =   "situacao"
         Caption         =   "Situação"
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
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3704,882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2729,764
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelectCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents adoR As ADODB.Recordset
Attribute adoR.VB_VarHelpID = -1
Dim bActive As Boolean
Public bConv As Boolean
Private Sub cmndbtnSelect_Click()
    If adoR.RecordCount = 0 Then
        MsgBox "Não há o que ser selecionado!", vbExclamation
        Exit Sub
    Else
        Me.tag = adoR!cod_cli
        Me.Hide
    End If
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnSelect_Click
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSelect_Click
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtgrd_DblClick()
    On Error Resume Next
    Me.tag = adoR!cod_cli
    Me.Hide
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    
    P_Busca
    
    bActive = True
End Sub

Sub P_Busca()
    Dim strSQL As String
    Dim bCad As Boolean
    Dim sAntes As String
    Dim sDepois As String
    On Error GoTo trata_Erro
    bActive = True
    
    P_Status "Localizando cliente..."
    
    
        
    
            Me.caption = "Selecione um Cliente"
            
            
                strSQL = "select  c.*,case when status_cli=1 then 'Liberado'"
                strSQL = strSQL & " when status_cli=2 then 'Baixado' "
                strSQL = strSQL & " else 'Bloqueado' end as situacao,tfp.nome_for,tfp.cod_for,cid.nome_Cid,cid.estado_cid,cid.uf_cid from  " & TAB_CLIENTES & " as c "
                strSQL = strSQL & " left join " & TAB_CIDADES & " as cid "
                strSQL = strSQL & " on c.cod_cid=cid.cod_cid "
                
                strSQL = strSQL & " left join " & TAB_FORMASPAG & " as tfp on c.cod_tip=tfp.cod_for "
                
                strSQL = strSQL & " where (nome_cli like '%" & txtbxNOme.Text & "%' or fantasia_cli like  '%" & txtbxNOme.Text & "%' "
                If txtbxNOme.tag <> "" Then
                   ' strSQL = strSQL & " and cod_cli='" & txtbxNOme.tag & "'"
                End If
                If F_SoNumero(txtbxNOme.Text) <> "" Then
                    strSQL = strSQL & " or cpf_cli like '" & F_FormataCPF(F_SoNumero(txtbxNOme.Text)) & "' "
                    strSQL = strSQL & " or cpf_cli like '" & F_FormataCNPJ(F_SoNumero(txtbxNOme.Text)) & "' "
                    strSQL = strSQL & " or c.cod_cli='" & F_SoNumero(txtbxNOme.Text) & "'"
                End If
            
            
                strSQL = strSQL & ") order by nome_cli"
        
        
    
    
    
        abreObjetoAdo adoR, strSQL, dtgrd
    
    If adoR.RecordCount = 1 Then
            On Error Resume Next
            Me.tag = adoR!cod_cli
            P_Status ""
            Me.Hide
            
                        Exit Sub
    ElseIf adoR.RecordCount = 0 Then
            
            
    Else
        dtgrd.caption = adoR.RecordCount & " Registros encontrados"
    End If
    
    If bCad = True Then P_Busca
    P_Status ""
    Exit Sub
    
trata_Erro:
P_Status ""
P_MostraErro
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyEscape Then
        Me.tag = ""
        Me.Hide
    ElseIf KeyCode = vbKeyN And Shift = vbCtrlMask Then
        If strCadcli_usu = "S" Then
            frmCadCli.Show 1
        End If
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me, 0
    Set adoR = New ADODB.Recordset
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bActive = False
End Sub

Private Sub optnFinal_Click()
    P_Busca
End Sub

Private Sub optnIni_Click()
    P_Busca
End Sub

Private Sub optnTodo_Click()
    P_Busca
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyUp Then
        dtgrd.SetFocus
    End If
    If KeyCode = vbKeyReturn Then
        P_Busca
    End If
End Sub
