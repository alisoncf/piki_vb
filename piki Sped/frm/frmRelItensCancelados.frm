VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelItensCancelados 
   Caption         =   "Relatório de Itens Cancelados"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcLeft 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   7920
         Picture         =   "frmRelItensCancelados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   9720
         Picture         =   "frmRelItensCancelados.frx":040D
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1695
      End
      Begin VB.ComboBox cmbxUsuario 
         Height          =   315
         ItemData        =   "frmRelItensCancelados.frx":0578
         Left            =   2880
         List            =   "frmRelItensCancelados.frx":0588
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Pesquisar"
         Height          =   615
         Left            =   6600
         Picture         =   "frmRelItensCancelados.frx":05B0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox cmbxCai 
         Height          =   315
         ItemData        =   "frmRelItensCancelados.frx":093A
         Left            =   120
         List            =   "frmRelItensCancelados.frx":0947
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbxTurno 
         Height          =   315
         ItemData        =   "frmRelItensCancelados.frx":0965
         Left            =   1440
         List            =   "frmRelItensCancelados.frx":0972
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   2880
         TabIndex        =   10
         Top             =   0
         Width           =   585
      End
      Begin VB.Label lblDataIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   4200
         TabIndex        =   9
         Top             =   0
         Width           =   840
      End
      Begin VB.Label lbldataFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   5280
         TabIndex        =   8
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   0
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   0
         Width           =   465
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   5055
      Left            =   0
      TabIndex        =   14
      Top             =   840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8916
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "num_ven"
         Caption         =   "Venda"
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
         DataField       =   "cod_cai"
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
      BeginProperty Column02 
         DataField       =   "cod_usu"
         Caption         =   "Usuário"
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
         DataField       =   "data_cai"
         Caption         =   "Data"
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
         DataField       =   "seq_ite"
         Caption         =   "seq"
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
      BeginProperty Column06 
         DataField       =   "nome_pro"
         Caption         =   "Produto"
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
         DataField       =   "qtd_ite"
         Caption         =   "qtd"
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
         DataField       =   "unit_ite"
         Caption         =   "unit"
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
         DataField       =   "total"
         Caption         =   "total"
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
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2250,142
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRelItensCancelados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmbxUsuario_Click()
    cmbxUsuario.tag = F_BuscaUmCampo("cod_fun", TAB_FUNCIONARIOS, "guerra_fun='" & cmbxUsuario.Text & "'")
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "select tv.valor_ven,tp.nome_pro,ti.* from " & TAB_VENDAS & " as tv," & TAB_PRODUTOS & " as tp, tab_itvndcncl as ti,"
    strSQL = strSQL & " " & TAB_ITVENDAS & " as tiv "
    strSQL = strSQL & " where tv.num_ven=ti.num_ven and tiv.nuM_ven=ti.num_ven "
    strSQL = strSQL & " and ti.seq_ite=tiv.seq_ite and tiv.cod_pro=tp.cod_pro"
    If cmbxCai.Text <> "" Then
        strSQL = strSQL & " and ti.cod_cai='" & Left(cmbxCai.Text, 2) & "'"
    End If
    If cmbxTurno.Text <> "" Then
        strSQL = strSQL & " and ti.cod_tur='" & cmbxTurno.Text & "'"
    End If
    If cmbxUsuario.tag <> "" Then
        strSQL = strSQL & " and tv.cod_usu='" & cmbxUsuario.tag & "'"
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and ti.data_ven>=" & F_GravaDataTempo(txtbxDataIni.Text)
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and ti.data_ven<=" & F_GravaDataTempo(txtbxdataFin.Text & " 23:59:59")
    End If
    strSQL = strSQL & " order by cod_cai, num_ven desc"
    abreObjetoAdo ado, strSQL, grid
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    txtbxdataFin.Text = Date
    txtbxDataIni.Text = Date
    P_PreencheCaixas cmbxCai
    P_PreencheTurnos cmbxTurno
    P_PreencheFunc cmbxUsuario, 1
    resizeForm Me
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, Me
    grid.Height = Me.Height - grid.Top - 400
    posicionaADireita cmndbtnSair, Me
End Sub

