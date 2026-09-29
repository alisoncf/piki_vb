VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSelectCli 
   Caption         =   "Selecionar Cliente"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   12150
   Icon            =   "frmSelectCli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   12150
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmSelectCli.frx":000C
      Height          =   8535
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15055
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cod"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nome"
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
         DataField       =   "cnpj"
         Caption         =   "CPF"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###\.###\.XXX\-XX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "endereco"
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
      BeginProperty Column04 
         DataField       =   "bairro"
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
      BeginProperty Column05 
         DataField       =   "cidade"
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
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00B4DFE0&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   12090
      TabIndex        =   3
      Top             =   8955
      Width           =   12150
      Begin VB.OptionButton optnFinal 
         BackColor       =   &H00B4DFE0&
         Caption         =   "No &Final"
         Height          =   195
         Left            =   2640
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton optnTodo 
         BackColor       =   &H00B4DFE0&
         Caption         =   "Em &Todo Nome"
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton optnIni 
         BackColor       =   &H00B4DFE0&
         Caption         =   "No &Início"
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox txtbxNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Tag             =   "C"
      Top             =   8640
      Width           =   8175
   End
   Begin VB.CommandButton cmndbtnSelect 
      Caption         =   "&Selecionar"
      Height          =   375
      Left            =   11160
      TabIndex        =   2
      Top             =   8565
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Informe uma Expressão para a busca:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   8640
      Width           =   2865
   End
End
Attribute VB_Name = "frmSelectCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WithEvents adoR As ADODB.Recordset
Attribute adoR.VB_VarHelpID = -1
Public BuscarPorDoc As Boolean

Dim bActive As Boolean
Public bConv As Boolean
Private Sub cmndbtnSelect_Click()
    If adoR.RecordCount = 0 Then
        MsgBox "Não há o que ser selecionado!", vbExclamation
        Exit Sub
    Else
        Me.tag = adoR!cod
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
    Me.tag = adoR!cod
    Me.Hide
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    
    If Visao.strVisao = "Vendas" Or Visao.strVisao = "Estoque" Then
        dtgrd.Columns(2).Visible = False
    End If
    
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
    If optnIni.Value Then
        sAntes = "'"
        sDepois = "%'"
    ElseIf optnTodo Then
        sAntes = "'%"
        sDepois = "%'"
    Else
        sAntes = "'%"
        sDepois = "'"
    End If
    P_Status "Localizando cliente..."
    
    If txtbxNome.tag = "R" Then
        
        Me.caption = "Selecione um Representante"
            Label1.caption = "Representante"
            
                strSQL = "select tc.*, null as dataalt_cli,cod_rep as cod, limite_rep,nome_rep as nome,cpf_rep as cnpj,endereco_rep as endereco, bairro_rep as bairro,"
                strSQL = strSQL & " cidade_rep as cidade, "
                strSQL = strSQL & " cpf_rep as cpf from tab_representantes as tc where excluido_rep='N' and "

            
            
            
                strSQL = strSQL & " (cod_rep='" & txtbxNome.Text & "' or nome_rep like " & sAntes & txtbxNome.Text & sDepois
                
                If F_SoNumero(txtbxNome.Text) = "" Then
                    strSQL = strSQL & " or cpf_rep ='XXXXXXXX'"
                Else
                    If F_SoNumero(txtbxNome.Text) > 0 Then
                        strSQL = strSQL & " or cpf_rep like '" & F_FormataCPF(F_SoNumero(txtbxNome.Text)) & "'"
                        strSQL = strSQL & " or cpf_rep like '" & (F_SoNumero(txtbxNome.Text)) & "'"
                        strSQL = strSQL & " or cod_rep='" & Format(txtbxNome.Text, "0000#") & "' "
                        
                    End If
                End If
            
            strSQL = strSQL & ") order by nome_rep"
    Else
    
            Me.caption = "Selecione um Cliente"
            Label1.caption = "Cliente"
            
                strSQL = "select distinct dataalt_cli,  datanasc_cli, cod_org,  logradouro_cli,numero_cli,complemento_cli,   status_cli,nome_Cli,credito_cli,endereco_cli,rg_cli,limite_cli,"
                strSQL = strSQL & " cep_cli,endtrab_cli,setortrab_cli,"
                strSQL = strSQL & " tc.cod_cli,cpf_cli,rg_cli,tipo_set,localtrab_cli,setortrab_cli,tc.cod_set,cod_cli as cod,nome_cid as cidade, "
                strSQL = strSQL & " nome_cid,uf_cid,ts.cod_cid,cobrador_set,ts.nome_set,"
                strSQL = strSQL & " ts.nome_set as bairro,limite_cli,nome_cli as nome,cpf_cli as cnpj,endereco_cli as endereco,  "
                strSQL = strSQL & "cpf_cli as cpf from tab_clientes as tc "
                strSQL = strSQL & " inner join tab_setores as ts on tc.cod_set=ts.cod_set "
                strSQL = strSQL & " inner join tab_cidades as ti on  ti.cod_cid=ts.cod_cid "
                If BuscarPorDoc Then
                    strSQL = strSQL & " left join " & TAB_DOC & " as td on td.cod_Ter=tc.cod_cli and td.tipo_ter='C'"
                End If
                strSQL = strSQL & " where  1 > 0 "
            
                strSQL = strSQL & " and (nome_Cli like " & sAntes & txtbxNome.Text & sDepois
            
                If F_SoNumero(txtbxNome.Text) <> "" Then
                    strSQL = strSQL & " or cpf_cli like '" & F_FormataCPF(F_SoNumero(txtbxNome.Text)) & "' "
                    strSQL = strSQL & " or cpf_cli like '" & (F_SoNumero(txtbxNome.Text)) & "' "
                    
                    strSQL = strSQL & " or cpf_cli like '" & F_FormataCGC(F_SoNumero(txtbxNome.Text)) & "' "
                    strSQL = strSQL & " or tc.cod_cli='" & F_SoNumero(txtbxNome.Text) & "'"
                    If BuscarPorDoc Then
                        strSQL = strSQL & " or td.num_doc='" & F_SoNumero(txtbxNome.Text) & "'"
                    End If
                End If
            
                strSQL = strSQL & ") order by nome_cli"
        
        
    
    End If
    
        FechaObjeto adoR
        adoR.Open strSQL, Conn_Client
        Set dtgrd.DataSource = adoR
    
    If adoR.RecordCount = 1 Then
            On Error Resume Next
            Me.tag = adoR!cod
            P_Status ""
            Me.Hide
            Me.Hide
                        Exit Sub
        ElseIf adoR.RecordCount = 0 Then
            
            
        Else
            
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
