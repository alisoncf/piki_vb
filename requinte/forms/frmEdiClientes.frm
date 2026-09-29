VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiClientes 
   Caption         =   "Edição de Clientes "
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   16995
   Icon            =   "frmEdiClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   16995
   Begin VB.PictureBox pcBotoes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   16965
      TabIndex        =   13
      Top             =   7815
      Width           =   16995
      Begin VB.CommandButton cmdPrintDev 
         Caption         =   "Nome e Telefone"
         Height          =   810
         Left            =   11760
         Picture         =   "frmEdiClientes.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   810
         Left            =   13200
         Picture         =   "frmEdiClientes.frx":0419
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   1335
      End
      Begin VB.CheckBox chkEspecial 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Especiais"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkSusp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Crédito Suspenso"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkLib 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Liberados"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir Completo"
         Height          =   810
         Left            =   10320
         Picture         =   "frmEdiClientes.frx":0848
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   1335
      End
      Begin VB.CheckBox chkBloq 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bloqueados"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   0
         Width           =   1215
      End
      Begin VB.CheckBox chkSPC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SPC"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnSPC 
         Caption         =   "&SPC"
         Height          =   450
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnCliSPC 
         Caption         =   "&Lista SPC"
         Height          =   450
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnMesc 
         Caption         =   "&Mesclar Clientes"
         Height          =   810
         Left            =   8880
         Picture         =   "frmEdiClientes.frx":0992
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   810
         Left            =   3120
         Picture         =   "frmEdiClientes.frx":165C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   810
         Left            =   4560
         Picture         =   "frmEdiClientes.frx":17C0
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   810
         Left            =   14640
         Picture         =   "frmEdiClientes.frx":190A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "&Histórico"
         Height          =   810
         Left            =   6000
         Picture         =   "frmEdiClientes.frx":1A75
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDeb 
         Caption         =   "&Débitos"
         Height          =   810
         Left            =   7440
         Picture         =   "frmEdiClientes.frx":233F
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiClientes.frx":2489
      Height          =   3255
      Left            =   0
      TabIndex        =   23
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5741
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
      Caption         =   "Registros Encontrados"
      ColumnCount     =   12
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
         DataField       =   "cpf_cli"
         Caption         =   "CPF:"
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
      BeginProperty Column03 
         DataField       =   "logradouro_cli"
         Caption         =   "Logradouro"
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
         DataField       =   "numero_cli"
         Caption         =   "N°"
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
      BeginProperty Column06 
         DataField       =   "nome_set"
         Caption         =   "Bairro/Setor"
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
         DataField       =   "email_cli"
         Caption         =   "email"
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
         DataField       =   "endereco_cli"
         Caption         =   "Endereço Antigo"
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
         DataField       =   "atraso"
         Caption         =   "Em atraso"
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
      BeginProperty Column11 
         DataField       =   "devido"
         Caption         =   "Em Aberto"
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
         BeginProperty Column00 
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2759,811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2055,118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2250,142
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   3075,024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2055,118
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   16965
      TabIndex        =   20
      Top             =   0
      Width           =   16995
      Begin VB.CheckBox chkDebitos 
         Caption         =   "Débitos"
         Height          =   255
         Left            =   13200
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbxCobranca 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmEdiClientes.frx":249B
         Left            =   10560
         List            =   "frmEdiClientes.frx":24AB
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxOrgao 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8160
         MaxLength       =   40
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtbxSetor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5760
         MaxLength       =   40
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtbxCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Caption         =   "&Pesquisar"
         Height          =   435
         Left            =   14280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiClientes.frx":24EB
         Left            =   11760
         List            =   "frmEdiClientes.frx":2501
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         MaxLength       =   40
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cobrança:"
         Height          =   195
         Left            =   10560
         TabIndex        =   8
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Orgão:"
         Height          =   195
         Left            =   8160
         TabIndex        =   6
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Se&tor:"
         Height          =   195
         Left            =   5760
         TabIndex        =   4
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cidade:"
         Height          =   195
         Left            =   3480
         TabIndex        =   2
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   11760
         TabIndex        =   10
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/CPF/RG/Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   1725
      End
   End
   Begin MSDataGridLib.DataGrid dtgrdTel 
      Bindings        =   "frmEdiClientes.frx":2546
      Height          =   3255
      Left            =   9960
      TabIndex        =   24
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5741
      _Version        =   393216
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
      Caption         =   "Telefones"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "num_tel"
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
      BeginProperty Column01 
         DataField       =   "tipo_tel"
         Caption         =   "Tipo"
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
         BeginProperty Column00 
            ColumnWidth     =   1695,118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   569,764
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnGerarArquivoClientes 
         Caption         =   "Gerar Arquivo de Clientes"
      End
      Begin VB.Menu mnGerarBloqueados 
         Caption         =   "Gerar arquivo de Bloqueados"
      End
      Begin VB.Menu mnEnviarInternet 
         Caption         =   "Enviar para Internet"
      End
      Begin VB.Menu mnFilaEnvio 
         Caption         =   "Enviar para fila de envio"
      End
      Begin VB.Menu mnPesqInt 
         Caption         =   "Pesquisar na internet"
      End
   End
End
Attribute VB_Name = "frmEdiClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim WithEvents AdoRClientes As ADODB.Recordset
Attribute AdoRClientes.VB_VarHelpID = -1
Dim WithEvents adoRCTelefones As ADODB.Recordset
Attribute adoRCTelefones.VB_VarHelpID = -1


Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub



Private Sub AdoRClientes_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim strSQL As String
    On Error Resume Next
    With AdoRClientes
        
        
        strSQL = "select * from tab_telefones where cod_cli='" & !cod_cli & "' and tipo_ter='C'"
        FechaObjeto adoRCTelefones
        adoRCTelefones.Open strSQL, Conn_Client
        Set dtgrdTel.DataSource = adoRCTelefones
    End With
End Sub

Private Sub Check1_Click()
    cmndbtnLocaliza_Click
End Sub

Private Sub chkBloq_Click()
    cmndbtnLocaliza_Click
End Sub

Private Sub chkEspecial_Click()
    cmndbtnLocaliza_Click
End Sub

Private Sub chkLib_Click()
    cmndbtnLocaliza_Click
End Sub

'Private Sub adc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    adc_Reposition
'End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnCard_Click()
    'P_Cartao adc.Recordset!cod_cli
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV dtgrd, AdoRClientes
End Sub

Private Sub cmdImprimir_Click()
    Dim bCpf As Boolean
    If strCadcli_usu <> "S" Then
        MsgBox "Você não tem permissão para o acesso!", vbExclamation
        Exit Sub
    End If
    bCpf = MsgBox("Deseja imprimir o CPF?", vbYesNo + vbQuestion) = vbYes
    
    imprime_cli bCpf
    Exit Sub
    
    P_DtgrdToHtml "Relação de Clientes ", dtgrd, AdoRClientes, IIf(chkBloq.Value = 1, "Bloqueados", ""), CStr(Now)
    
End Sub
Sub imprime_cli(bCpf As Boolean, Optional sonome_telefone = False)
    Dim a As String
    a = F_AbreHtml("lista_cli")
    dtgrd.Visible = False
    
    P_HtmlAbreHead
    P_HtmlAbreTitle 1, "Listagem de Clientes em " & Date
    P_HtmlFechaHead
    
    P_HtmlAbreH 1, 4, "Listagem de Clientes em " & Date, True
    
    P_HtmlAbreTable 1, 1
    P_HtmlAbreLinha 1
    P_HtmlAbreColuna 1, "Código"
    If bCpf Then
        P_HtmlAbreColuna 1, "CPF"
    End If
    P_HtmlAbreColuna 1, "Nome"
    P_HtmlAbreColuna 1, "Telefones"
    If sonome_telefone = False Then
        P_HtmlAbreColuna 1, "Endereço"
        P_HtmlAbreColuna 1, "Bairro"
        P_HtmlAbreColuna 1, "Cidade"
        
    End If
    P_HtmlAbreColuna 1, "Em atraso"
    P_HtmlFechaLinha 1
        
    With AdoRClientes
        If .RecordCount > 0 Then
            .MoveFirst
            
            Do While Not .EOF
                
                P_HtmlAbreLinha 1
                P_HtmlAbreColuna 1, !cod_cli
                If bCpf Then
                    P_HtmlAbreColuna 1, !cpf_cli
                End If
                P_HtmlAbreColuna 1, !nome_cli
                P_HtmlAbreColuna 1, F_BuscaTel(!cod_cli, "C")
                If sonome_telefone = False Then
                    P_HtmlAbreColuna 1, !Logradouro_cli & " N°" & !numero_cli & " " & !complemento_cli
                    P_HtmlAbreColuna 1, !nome_set
                    P_HtmlAbreColuna 1, !nome_cid
                    
                End If
                P_HtmlAbreColuna 1, FormatVal(!atraso)
                P_HtmlFechaLinha 1
                .MoveNext
            Loop
            
        End If
        
        P_HtmlFechaTable 1
    End With
    
    
    P_HtmlFechaTable
    
    html.P_FechaHtml 1, True, a
    
    
    
    dtgrd.Visible = True
    
      
End Sub

Private Sub cmdPrintDev_Click()
    Dim bCpf As Boolean
    
    imprime_cli False, True
    Exit Sub
    
    P_DtgrdToHtml "Relação de Clientes ", dtgrd, AdoRClientes, IIf(chkBloq.Value = 1, "Bloqueados", ""), CStr(Now)
    
End Sub

Private Sub cmndbtnCliSPC_Click()
    frmRelClienteSpc.Show 0
End Sub

Private Sub cmndbtnDeb_Click()
    On Error GoTo trata_Erro
    frmParcelasTer.txtbxNomeCli.text = AdoRClientes!nome_cli
    frmParcelasTer.txtbxNomeCli.tag = AdoRClientes!cod_cli
    frmParcelasTer.tag = AdoRClientes!cod_cli
    frmParcelasTer.Show 0
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadCli
    If AdoRClientes.RecordCount <= 0 Then
        MsgBox "Não há o que ser editado!", vbCritical
        Exit Sub
    End If
    strHora = Time
    
    fr.tag = AdoRClientes!cod_cli
    
    fr.Show 1
End Sub

Private Sub cmndbtnHist_Click()
'    On Error GoTo trata_erro
    With AdoRClientes
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
    P_Msg "aguarde..."
    strSQL = "select distinct tc.cpf_cli,logradouro_cli,numero_cli,complemento_cli,tc.endereco_cli,"
    strSQL = strSQL & "email_cli,tc.cod_cli,tc.status_cli,tc.nome_cli,ts.nome_set,tcid.nome_cid,tcid.uf_cid "
    
    If chkDebitos.Value = 1 Then
        strSQL = strSQL & ",(select sum(it.reman_itd) from  " & TAB_DOC & " as d inner join " & TAB_ITDOC & " as it "
        strSQL = strSQL & " on d.seq_doc=it.seq_doc where d.tipo_doc='R'  and tipo_ter='C' and cod_ter=tc.cod_cli) as devido,"
        strSQL = strSQL & "(select sum(it.reman_itd) from  " & TAB_DOC & " as d inner join " & TAB_ITDOC & " as it "
        strSQL = strSQL & " on d.seq_doc=it.seq_doc where d.tipo_doc='R'  and tipo_ter='C' and cod_ter=tc.cod_cli and venc_itd<="
        strSQL = strSQL & F_GravaData(Date) & ") as atraso"
    End If
    strSQL = strSQL & " from "
    strSQL = strSQL & " tab_clientes as tc " 'left join tab_telefones as tel on tc.cod_cli=tel.cod_cli and tipo_ter='C' "
    strSQL = strSQL & "left join tab_setores as ts on ts.cod_set=tc.cod_set "
    strSQL = strSQL & " left join tab_cidades as tcid on tcid.cod_cid=ts.cod_cid "
    If chkSPC.Value Then
        strSQL = strSQL & "left join tab_clientespc as tcs where tcs.cod_cli=tc.cod_cli "
        strSQL = strSQL & " and tcs.baixa_spc='0000-00-00' "
    End If
        strSQL = strSQL & " where 1 > 0 "
    
    
    
        strSQL = strSQL & " and (Nome_cli like '%" & txtbxExpr.text & "%' "
    
        strSQL = strSQL & "or tc.cod_cli like '" & txtbxExpr.text & "%' "
        
        strSQL = strSQL & "or cpf_cli = '" & F_FormataCPF(txtbxExpr.text) & "'  "
        strSQL = strSQL & "or (cpf_cli = '" & F_SoNumero(txtbxExpr.text) & "'   and cpf_cli <> '')"
        
        strSQL = strSQL & "or cpf_cli = '" & F_FormataCGC(txtbxExpr.text) & "'  "
    
    
        strSQL = strSQL & "or rg_cli like '" & txtbxExpr.text & "%') "
    
    strSQL = strSQL & " and tc.cod_cli <> 0 " 'venda a consumidor
    
    If txtbxOrgao.tag <> "" Then
        strSQL = strSQL & " and tc.cod_org='" & txtbxOrgao.tag & "' "
    End If
    
    If txtbxSetor.tag <> "" Then
        strSQL = strSQL & " and tc.cod_set='" & txtbxSetor.tag & "' "
    Else
        strSQL = strSQL & " and nome_set like '%" & txtbxSetor.text & "%' "
    End If
    
    If chkBloq.Value = 1 Then
        strSQL = strSQL & " and status_cli=2 "
    End If
    If chkLib.Value = 1 Then
        strSQL = strSQL & " and status_cli=1 "
    End If
    
    If txtbxCidade.tag <> "" Then
        strSQL = strSQL & " and ts.cod_cid='" & txtbxCidade.tag & "'"
    Else
        strSQL = strSQL & " and nome_cid like '" & txtbxCidade.text & "%'"
    End If
    If chkEspecial.Value = 1 Then
        strSQL = strSQL & " and especial_cli='S' "
    End If
    
    If cmbxCobranca.text <> "" Then
        strSQL = strSQL & " and cobranca_cli='" & Left(cmbxCobranca.text, 1) & "' "
    End If
    If chkSusp.Value = 1 Then
        strSQL = strSQL & " and creditosuspenso_cli='S' "
    End If
    
    If cmbxOrd.ListIndex = 0 Then
        sOrderBy = "tc.cod_cli"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrderBy = "Nome_Cli,cpf_Cli"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrderBy = "CPF_CLI"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrderBy = "Cidade_Cli,uf_Cli,NOme_Cli"
    ElseIf cmbxOrd.ListIndex = 4 Then
        sOrderBy = "UF_CLI,Cidade_CLi"
    ElseIf cmbxOrd.ListIndex = 5 Then
        sOrderBy = "dataalt_cli desc,nome_cli"
    End If
    strSQL = strSQL & " order by " & sOrderBy
    abreObjetoAdo AdoRClientes, strSQL, dtgrd
    
    
    'Set dtgrdTel.DataSource = adoRCTelefones
    
   P_Msg
    dtgrd.caption = "Cadastrados...: " & F_BuscaUmCampo("count(cod_Cli)", "tab_clientes", "") & " - Listados...: " & AdoRClientes.RecordCount
Exit Sub
trata_Erro:
    
    MsgBox "A consulta encontrou um erro ou nenhum resultado foi encontrado! Refaça sua pesquisa", vbExclamation
    P_Msg

End Sub

Private Sub cmndbtnMesc_Click()
    frmMesclarClientes.Show 0
End Sub

Private Sub cmndbtnNew_Click()
    If strCadcli_usu <> "S" Then
        MsgBox "Você não tem permissão pra cadastrar Cliente!", vbExclamation
        Exit Sub
    End If
    
    Dim fr As New frmCadCli
    fr.Show 0
End Sub

Private Sub Command4_Click()

End Sub


Private Sub cmndbtnPromocoes_Click()
    On Error GoTo trata_Erro
    
    Exit Sub
trata_Erro:
End Sub

Private Sub cmndbtnRes_Click()
    frmReserva.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSPC_Click()
    On Error GoTo trata_Erro
    
    P_ClienteSpc adc.Recordset!cod_cli, 0
Exit Sub
trata_Erro:
    If Err.Number = 0 Then
    End If
    P_ClienteSpc ""
End Sub

Private Sub dtgrd_DblClick()
    cmndbtnEdit_Click
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


Private Sub Form_Activate()
    If bActive Then Exit Sub
    
    txtbxExpr.SetFocus
    'cmndbtnLocaliza_Click
    bActive = True
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
    
    Set AdoRClientes = New ADODB.Recordset
    Set adoRCTelefones = New ADODB.Recordset
    
    
    
    
    cmbxOrd.ListIndex = 0
    
    'txtbxExpr.SetFocus
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    PosicionaGrade dtgrd, Me, pcBotoes
    
    dtgrd.width = dtgrd.width - dtgrdTel.width
    
    dtgrdTel.Left = Me.width - dtgrdTel.width - 600
    dtgrdTel.top = dtgrd.top
    dtgrdTel.Left = dtgrd.width + dtgrd.Left + 100
    
    
    
    
    'dtgrd.height = Me.height - dtgrd.Top - 600
    dtgrdTel.height = dtgrd.height
    posicionaADireita cmndbtnSair, pcBotoes
    
End Sub

Private Sub mnEnviarInternet_Click()
    ret = APIInserirCliente(AdoRClientes!cod_cli)
    If ret = "" Then
    End If
    
End Sub

Private Sub mnFilaEnvio_Click()
    On Error GoTo trata_Erro
    dtgrd.Visible = False
    With AdoRClientes
        .MoveFirst
        Do While Not .EOF
            
            insere_fila !cod_cli, 1
            .MoveNext
        Loop
        .MoveFirst
    End With
    MsgBox "Fila adicionada", vbInformation
    P_Status
    dtgrd.Visible = True
trata_Erro:
End Sub

Private Sub mnPesqInt_Click()
    Dim ret As String
    ret = APIPesquisarCliente(AdoRClientes!cpf_cli)
    
    
    If ret = "" Then
        
    End If
    
End Sub

Private Sub pcBotoes_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub txtbxCidade_Change()
    txtbxCidade.tag = ""
End Sub

Private Sub txtbxCidade_GotFocus()
    P_SelecTexto txtbxCidade
End Sub

Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCid
    
    End If
End Sub
Sub buscaCid()
    Dim c As Cidade
    c = F_BuscaCidades(txtbxCidade.text)
    
    If c.cod_cid <> "" Then
        txtbxCidade.text = c.nome_cid
        txtbxCidade.tag = c.cod_cid
        cmndbtnLocaliza_Click
    End If
    
End Sub
Sub buscaSetor()
    Dim s As Setores.setor
    
    s = F_BuscaSetor(txtbxSetor.text)
    If s.cod <> "" Then
        txtbxSetor.text = s.nome
        txtbxSetor.tag = s.cod
        cmndbtnLocaliza_Click
    End If
    
End Sub
Private Sub txtbxCidade_LostFocus()
    P_AjustaTexto txtbxCidade
End Sub

Private Sub txtbxExpr_Change()
    'cmndbtnLocaliza_Click
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

Private Sub txtbxOrgao_Change()
    txtbxOrgao.tag = ""
End Sub

Private Sub txtbxOrgao_DblClick()
    buscaOrgao
End Sub

Private Sub txtbxOrgao_GotFocus()
    P_SelecTexto txtbxOrgao
End Sub

Private Sub txtbxOrgao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaOrgao
    End If
End Sub
Sub buscaOrgao()
    Dim o As orgao.orgao
    o = orgao.F_BuscaOrgao(txtbxOrgao.text)
    If o.cod_org <> "" Then
        txtbxOrgao.text = o.nome_org
        txtbxOrgao.tag = o.cod_org
        txtbxExpr.SetFocus
    End If
End Sub
Private Sub txtbxSetor_Change()
    txtbxSetor.tag = ""
    
End Sub

Private Sub txtbxSetor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaSetor
    End If
End Sub
