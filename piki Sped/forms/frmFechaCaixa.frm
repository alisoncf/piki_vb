VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFechaCaixa 
   Caption         =   "Fechamento de Caixa"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   11955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid gridTurno 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1296
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BorderStyle     =   0
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   -1  'True
         AllowSizing     =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      MaxLength       =   40
      TabIndex        =   20
      Top             =   360
      Width           =   1335
   End
   Begin VB.CheckBox chkBaixa 
      Caption         =   "Baixa"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7320
      TabIndex        =   16
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox pcBot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   11625
      TabIndex        =   9
      Top             =   6960
      Width           =   11655
      Begin VB.Label lblTotalGeral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   15
         Top             =   0
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Geral:"
         Height          =   195
         Left            =   4800
         TabIndex        =   14
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblTotalSimb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Simbólico:"
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efetivo:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lblTotalEfetivo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   0
         Width           =   390
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "&Sair"
      Height          =   735
      Left            =   10200
      Picture         =   "frmFechaCaixa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   735
      Left            =   1920
      Picture         =   "frmFechaCaixa.frx":054E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnFecharCaixa 
      Caption         =   "&Fechar Caixa"
      Height          =   735
      Left            =   120
      Picture         =   "frmFechaCaixa.frx":07AA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   8640
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox cmbxCaixa 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtbxData 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin TabDlg.SSTab sstb 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Resumo"
      TabPicture(0)   =   "frmFechaCaixa.frx":08F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "gridResumo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalhado"
      TabPicture(1)   =   "frmFechaCaixa.frx":0910
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gridDetalhado"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Baixa de Parcelas"
      TabPicture(2)   =   "frmFechaCaixa.frx":092C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "gridBaixas"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid gridResumo 
         Height          =   4455
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7858
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "sigla_cai"
            Caption         =   "Sigla"
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
            DataField       =   "descricao_cai"
            Caption         =   "Descrição"
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
            DataField       =   "tipo_cai"
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
         BeginProperty Column03 
            DataField       =   "total"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,##0.00"
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
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3165,166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridDetalhado 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7646
         _Version        =   393216
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
            DataField       =   "seq_cai"
            Caption         =   "Seq"
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
            DataField       =   "esp_cai"
            Caption         =   "Espécie"
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
            DataField       =   "historico_cai"
            Caption         =   "Histórico"
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
            DataField       =   "valor_cai"
            Caption         =   "Valor"
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
            DataField       =   "tipo_cai"
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
         BeginProperty Column05 
            DataField       =   "sinal"
            Caption         =   "*"
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
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4259,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   285,165
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridBaixas 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   8070
         _Version        =   393216
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "cod"
            Caption         =   "Cód"
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
            DataField       =   "cobranca"
            Caption         =   "Cobrança"
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
            DataField       =   "total"
            Caption         =   "Total"
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
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2910,047
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdBuscarFun 
      Height          =   285
      Left            =   4920
      Picture         =   "frmFechaCaixa.frx":0948
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      Height          =   195
      Left            =   3600
      TabIndex        =   22
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa:"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frmFechaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoResumo As ADODB.Recordset
Attribute adoResumo.VB_VarHelpID = -1
Dim WithEvents adoDetalhado As ADODB.Recordset
Attribute adoDetalhado.VB_VarHelpID = -1
Dim WithEvents adoBaixas As ADODB.Recordset
Attribute adoBaixas.VB_VarHelpID = -1

Private Sub chkBaixa_Click()
    If chkBaixa.Value = 0 Then
        sstb.TabVisible(2) = False
    Else
        sstb.TabVisible(2) = True
    End If
End Sub

Private Sub cmbxCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxCaixa.ListIndex = -1
    End If
End Sub

Private Sub cmbxUsu_Click()
    cmbxUsu.tag = F_BuscaUmCampo("cod_fun", TAB_FUNCIONARIOS, "guerra_fun='" & cmbxUsu.Text & "'")
End Sub

Private Sub cmbxUsu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxUsu.ListIndex = -1
    End If
End Sub

Private Sub cmdBuscarFun_Click()
    buscaUsu
End Sub

Private Sub txtUsuario_Change()
    txtUsuario.tag = ""
End Sub

Private Sub txtUsuario_GotFocus()
    P_SelecTexto txtUsuario
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaUsu
    End If
End Sub
Sub buscaUsu()
    Dim f As funcionario.funcionario
    f = f_BuscaFun(txtUsuario.Text)
    If f.cod_fun <> "" Then
        txtUsuario.Text = f.guerra_fun
        txtUsuario.tag = f.cod_fun
        busca
    End If
End Sub
Private Sub cmndbtnBuscar_Click()
        P_Status "Aguarde..."
    busca
    P_Status ""
End Sub

Private Sub cmndbtnImprimir_Click()
    If sstb.Tab = 0 Then
        P_DtgrdToHtml "Resumo de Movimento de Caixa", gridResumo, adoResumo, _
        "Data: " & txtbxData.Text & " Caixa: " & cmbxCaixa.Text & " Usu:" & txtUsuario.Text, 4
    
    ElseIf sstb.Tab = 1 Then
        P_DtgrdToHtml "Movimento de Caixa Detalhado", gridDetalhado, adoDetalhado, _
        "Data: " & txtbxData.Text & " Caixa: " & cmbxCaixa.Text & " Usu:" & txtUsuario.Text, 3
    Else
        P_DtgrdToHtml "Resumo das Baixas", gridBaixas, adoBaixas, _
        "Data: " & txtbxData.Text & " Caixa: " & cmbxCaixa.Text & " Usu:" & txtUsuario.Text, 2
    End If
        
        
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        If pcBot.BackColor = vbYellow Then
            pcBot.BackColor = vbWhite
        Else
            pcBot.BackColor = vbYellow
        End If
    Else
    
    End If
End Sub

Private Sub Form_Load()
    
    P_PreencheCaixas cmbxCaixa
    
    
    txtbxData.Text = Date
    Set adoBaixas = New ADODB.Recordset
    Set adoDetalhado = New ADODB.Recordset
    Set adoResumo = New ADODB.Recordset
    
    
    
    chkBaixa_Click
    P_LocalizaOperacao strCod_Cai, cmbxCaixa
   ' txtUsuario.Text = strNome_Usu
   ' txtUsuario.tag = strCod_Usu
    
    sstb.Tab = 0
    
    If strFechacai_usu = "S" Then
        txtUsuario.Enabled = True
        cmdBuscarFun.Enabled = True
        cmbxCaixa.Enabled = True
    Else
        txtUsuario.Enabled = False
        cmdBuscarFun.Enabled = False
        cmbxCaixa.Enabled = False
    End If
    
End Sub
Sub ajustaGrid()
    
End Sub
Private Sub cmndbtnFecharCaixa_Click()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim sCod_Con  As String
    
    If Not F_IsCaixa Then
        MsgBox "Este comptuador não é um caixa!", vbExclamation
        Exit Sub
    End If
    
    
    If strFechacai_usu <> "S" Then
        MsgBox "Esse usuário não tem permissão para fechamento de caixa", vbExclamation
        Exit Sub
        
    End If
    If strCod_Tur = "" Then
        strCod_Tur = "1"
    End If
    If txtUsuario.tag <> "" Then
        MsgBox "Não é possível fechar só o movimento de um usuário! Tente novamente", vbExclamation
        txtUsuario.tag = ""
        txtUsuario.Text = ""
        cmndbtnBuscar_Click
        Exit Sub
    End If
    
    If sstb.Tab = 0 Then
        If verificaVendaAberta(Left(cmbxCaixa.Text, 2)) Then
            MsgBox "Não é possível fechar o caixa pois existem vendas em aberto! Procure o responsável!", vbCritical
            Exit Sub
        
        End If
    End If
    'On Error GoTo trata_erro'
    'strSQL = F_BuscaUmCampo("count(*)", TAB_VENDAS, "situacao_ven=2 and cod_cai='" & strCod_Cai & "'")
    'If Int(strSQL) > 0 Then
    '    MsgBox "Não é possível fechar o caixa pois existem vendas a efetivar!", vbExclamation
    '    Exit Sub
    'End If
'    If F_BuscaUmCampo("data_cai", "tab_fechacaixa", "data_cai=" & _
'        F_PesqData(txtbxData.Text) & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "'") <> "" Then
'        MsgBox "Este caixa já foi fechado neste dia!", vbExclamation
'        Exit Sub
'    End If

    'If CCur(lblTotalDin.caption) - CCur(txtbxFundoCaixa.Text) < 0 Then
    '    MsgBox "O fundo de caixa está maior que o valor disponível!", vbExclamation
    '    Exit Sub
    'End If

    If VendasFechadas = True Then
        MsgBox "Existem vendas condicionais a serem efetivadas!", vbExclamation
        On Error Resume Next
        frmEfetivaVen.Show 1
        Exit Sub
    End If
    If MsgBox("Deseja realmente fechar o Turno do dia " & txtbxData.Text _
        & "", vbYesNo + vbCritical) = vbYes Then
        'If CCur(txtbxFundoCaixa.Text) = 0 Then
            
        '    If MsgBox("Este fechamento não contém fundo de caixa!" & Chr(13) & _
        '        "Deseja realmente continuar?", vbYesNo + vbQuestion) = vbYesNo Then
        '        Exit Sub
        '    End If
        'End If
        If F_BuscaUmCampo("status_tur", TAB_TURNCAI, "cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' and cod_tur='" & strCod_Tur & "' and data_tur=" & F_GravaData(txtbxData.Text)) = "F" Then
            MsgBox "Este turno já se encontra Fechado!", vbCritical
            Exit Sub
         End If
'        strSQL = "insert into tab_fechacaixa(data_Cai,valor_FEC,fundo_cai,cod_usu,COD_CAI)values("
'        strSQL = strSQL & F_GravaData(txtbxData.Text) & ","
'        strSQL = strSQL & F_GravaDec(lblTotalAB.caption) & ","
'        strSQL = strSQL & F_GravaDec(txtbxFundoCaixa.Text) & ","
'        strSQL = strSQL & "'" & strCod_Usu & "','" & Left(cmbxCaixa.Text, 2) & "') "
'        Banco.Execute strSQL
        
        
        strSQL = "update " & TAB_TURNCAI & "  set logout_usu='" & strCod_Usu & "',"
        strSQL = strSQL & "Fechamento_tur='" & Format(Time, "hh:mm") & "',"
        strSQL = strSQL & "valorfechamento_tur=" & F_GravaDec(lblTotalEfetivo.caption) & ","
        
        strSQL = strSQL & "status_tur='F' "
        strSQL = strSQL & " where cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' "
        strSQL = strSQL & " and cod_tur='" & strCod_Tur & "' "
        conn.Execute strSQL
        
        'strSQL = F_GravaCai("D", txtbxData.Text, CCur(lblTotalDin.caption) - CCur(txtbxFundoCaixa.Text), _
            "TRANSFERÊNCIA DE CAIXA", "", "DIN", "01", 0, 0, Left(cmbxCaixa.Text, 2), strCod_Tur)
        
        
        
        With adoDetalhado
            If .RecordCount > 0 Then
                .MoveFirst
            End If
            Do While Not .EOF
                If !esp_cai = "DIN" Or !esp_cai = "RDI" Then
                    strSQL = F_GravaFinanceiro(!tipo_cai, !data_cai, !valor_Cai, !historico_cai, "001", _
                    "", strPlaVendasEfetuadas)
                    
                End If
                .MoveNext
            Loop
        End With
        
            
            
        strSQL = "Update tab_Caixas set status_Cai='F', "
        strSQL = strSQL & " ip_cai='" & strIP_Usu & "' "
        strSQL = strSQL & " where cod_Cai='" & Left(cmbxCaixa.Text, 2) & "'"
        conn.Execute strSQL
        MsgBox "Caixa " & cmbxCaixa.Text & "  Fechado com sucesso!", vbInformation
        
        
    End If
        
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Sub soma()
    Dim cTotal As Currency
    Dim cEfetivo As Currency
    Dim cSimb As Currency
    With adoResumo
        If Not .EOF Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            
                
                
            If !sigla_cai = "DIN" Then
                If !tipo_cai = "C" Then
                    cEfetivo = cEfetivo + !total
                Else
                    cEfetivo = cEfetivo - !total
                End If
            Else
                If !tipo_cai = "C" Then
                    cSimb = cSimb + !total
                Else
                    cSimb = cSimb - !total
                End If
            End If
            
            
            cTotal = cTotal + !total
            
            .MoveNext
        Loop
        If Not .EOF Then
            .MoveFirst
        End If
    End With
    lblTotalEfetivo.caption = FormatVal(cEfetivo)
    lblTotalGeral.caption = FormatVal(cTotal)
    lblTotalSimb.caption = FormatVal(cSimb)
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cCred As Currency
    Dim sSigla As String
    Dim cDeb As Currency
    Dim sCab As String
    On Error GoTo trata_Erro
    strSQL = "select tc.esp_cai,tc.tipo_cai, sum(tc.valor_cai) as total,te.* from " & TAB_CAIXA & " as tc left join tab_espcai as te "
    strSQL = strSQL & " on te.sigla_cai=tc.esp_cai "
    strSQL = strSQL & " where data_cai=" & F_GravaData(txtbxData.Text)
    If cmbxCaixa.Text <> "" Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    End If
    If txtUsuario.tag <> "" Then
        strSQL = strSQL & " and tc.cod_usu='" & txtUsuario.tag & "' "
    End If
    
    strSQL = strSQL & " group by tc.esp_cai ,tc.tipo_cai"
    
    abreObjetoAdo adoResumo, strSQL, gridResumo
    

    strSQL = " select tc.*,if(tc.tipo_cai='C','+','-') as sinal,te.sigla_cai,te.descricao_cai from " & TAB_CAIXA & " as tc left join "
    strSQL = strSQL & " tab_espcai as te on te.sigla_cai=tc.esp_cai "
    strSQL = strSQL & " where data_cai=" & F_GravaData(txtbxData.Text)
    'strSQL = strSQL & " and te.sigla_cai=tc.esp_cai "
    
    If cmbxCaixa.Text <> "" Then
        strSQL = strSQL & " and cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' "
    End If
    
    If txtUsuario.tag <> "" Then
        strSQL = strSQL & " and cod_usu='" & txtUsuario.tag & "' "
    End If
    'If pcBot.BackColor <> vbYellow Then
    '    strSQL = strSQL & " and visivel_cai='S' "
    'End If
    strSQL = strSQL & " order by tc.esp_cai,tc.tipo_cai,tc.seq_cai  "
    abreObjetoAdo adoDetalhado, strSQL, gridDetalhado
    
    If chkBaixa.Value = 1 Then
    
        strSQL = F_SelectBaixaCarnesPorCaixa(cmbxCaixa.Text, txtbxData.Text, txtbxData.Text, txtUsuario.tag)
    
        abreObjetoAdo adoBaixas, strSQL, gridBaixas
    End If
    soma
    
    strSQL = "select * from " & TAB_TURNCAI & " where data_tur=" & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " and cod_tur='" & strCod_Tur & "'"
    abreObjetoAdo rPesq, strSQL, gridTurno
    
    Dim colTur As String
    Dim colCampos As String
    'colTur = " Aberto por, fechado por, valor abertura, valor fechamento, status "
    'colCampos = "login_usu, logout_usu, valorabertura_tur, valorfechamento_tur, status_tur "
    'GridUtil.AjustaGridData gridTurno, 4, 1, 1, colTur, colCampos, "1000,1000,1500,1500,1000", rPesq
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub
