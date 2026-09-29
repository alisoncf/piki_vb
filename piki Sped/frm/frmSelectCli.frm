VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSelectCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecionar Cliente"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10830
   Icon            =   "frmSelectCli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Tag             =   "C"
      Top             =   7560
      Width           =   8775
   End
   Begin VB.CommandButton cmndbtnSelect 
      Caption         =   "&Selecionar (Enter)"
      Height          =   615
      Left            =   9120
      Picture         =   "frmSelectCli.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmSelectCli.frx":0C54
      Height          =   7095
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   12515
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
      Caption         =   "Clientes"
      ColumnCount     =   11
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
         DataField       =   "cpf_cli"
         Caption         =   "CPF/CNPJ"
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
      BeginProperty Column04 
         DataField       =   "complemento_cli"
         Caption         =   "Complemento"
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
         DataField       =   "numero_cli"
         Caption         =   "Número"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "uf_cid"
         Caption         =   "uf"
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
         DataField       =   "telefone_cli"
         Caption         =   "Telefone"
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
      BeginProperty Column10 
         DataField       =   "celular_cli"
         Caption         =   "Celular"
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
            ColumnWidth     =   659,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3990,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1725,165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2789,858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1665,071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Buscar Por:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   7320
      Width           =   825
   End
End
Attribute VB_Name = "frmSelectCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmndbtnSelect_Click()
    If adc.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que ser selecionado!", vbExclamation
        Exit Sub
    Else
        Me.tag = adc.Recordset!cod_cli
        Me.Hide
    End If
End Sub

Private Sub dtgrd_DblClick()
    cmndbtnSelect_Click
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSelect_Click
    End If
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    adc.ConnectionString = strConnect
    P_Busca
    'MudaAparencia Me
    bActive = True
End Sub

Sub P_Busca()
    Dim strSQL As String
    Dim bCad As Boolean
    On Error GoTo trata_erro
    
    For iCount = 1 To 1
        If txtbxNome.tag = "C" Then
            Me.caption = "Selecione um Cliente"
            Label1.caption = "Cliente"
            strSQL = "select tc.*,cid.* from tab_clientes as tc left join tab_cidades as cid on cid.cod_cid=tc.cod_cid "
            
            
                strSQL = strSQL & " where ( nome_Cli like '" & txtbxNome.Text & "%' "
                
                strSQL = strSQL & "or cpf_cli like '" & txtbxNome.Text & "%'"
                
                strSQL = strSQL & " or cod_cli='" & txtbxNome.Text & "' )"
            
            strSQL = strSQL & " order by nome_cli"
        
        Else
            Me.caption = "Selecione um Fornecedor"
            Label1.caption = "Fornecedor"
            strSQL = "select cod_for as cod, nome_for as nome,cnpj_for as cnpj,telefone_for as telefone,uf_for as uf, "
            strSQL = strSQL & "cidade_for as cidade, bairro_for as bairro, endereco_for as endereco "
            strSQL = strSQL & "from tab_fornecedores where nome_for like '" & txtbxNome.Text & "%'"
            strSQL = strSQL & "order by nome_For"
        End If
    
    
    
        adc.RecordSource = strSQL
        adc.Refresh
    If adc.Recordset.RecordCount = 1 Then
            On Error Resume Next
            Me.tag = adc.Recordset!cod_cli
            Me.Hide
            Exit For
            Exit Sub
        ElseIf adc.Recordset.RecordCount = 0 Then
            
            If iCount = 2 Then
                If F_SoNumero(txtbxNome.Text) = "" Then
                    frmCadCliRapido.txtbxRazao.Text = txtbxNome
                Else
                    frmCadCliRapido.txtbxCNPJ.Text = txtbxNome
                End If
                If txtbxNome.tag = "C" Then
                    On Error GoTo trata_erro
                    
                    frmCadCliRapido.Show 1
                    'If frmCadCliRapido.Tag <> "" Then
                    'P_Busca
                    'End If
                    'Unload frmCadCliRapido
                End If
            
                If frmCadCliRapido.tag <> "" Then
                    bCad = True
                End If
                Unload frmCadCliRapido
                Exit For
            End If
        Else
            Exit For
        End If
    Next
    If bCad = True Then P_Busca
    Exit Sub
trata_erro:
MostraErro
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bActive = False
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyUp Then
        dtgrd.SetFocus
    End If
    If KeyCode = vbKeyReturn Then
        P_Busca
    End If
End Sub
