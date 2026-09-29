VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLocalizaProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localização de Produto"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmLocalizaProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6165
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   5895
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   3360
         Picture         =   "frmLocalizaProd.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Reservar"
         Height          =   735
         Left            =   120
         Picture         =   "frmLocalizaProd.frx":2B2C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2400
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "stt"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dtGrd 
      Bindings        =   "frmLocalizaProd.frx":2C76
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
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
      Caption         =   "Produto Localizado em:"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "cod_loj"
         Caption         =   "cód"
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
         DataField       =   "nome_loj"
         Caption         =   "Loja"
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
         DataField       =   "unidade_pro"
         Caption         =   "UN."
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
         DataField       =   "estoque"
         Caption         =   "Estoque"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00"
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
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2910,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   900,284
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLocalizaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    Me.Tag = "X"
    Me.Hide
    
End Sub

Private Sub cmndbtnSair_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub dtGrd_GotFocus()
    dtGrd.MarqueeStyle = 3
End Sub

Private Sub dtGrd_LostFocus()
    dtGrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    
    MudaAparencia Me
    strSQL = "[ " & F_BuscaUmCampo("nome_pro", "tab_produtos", _
        "cod_pro='" & Me.Tag & "'") & "]"
    Me.caption = strSQL
    dtGrd.caption = strSQL & " localizado em: "
    Adodc1.ConnectionString = "DSN=stt;"
    strSQL = "select tl.cod_loj,tl.nome_loj,"
    strSQL = strSQL & "tp.cod_pro,tp.nome_pro,"
    strSQL = strSQL & "tp.unidade_pro,te.estoque_pro,te.custo_pro,"
    strSQL = strSQL & "te.reservado_pro,"
    strSQL = strSQL & "(te.estoque_pro-te.reservado_pro) as estoque "
    strSQL = strSQL & "from tab_lojas as tl, tab_produtos as tp, "
    strSQL = strSQL & "tab_estoque as te where "
    strSQL = strSQL & "te.cod_pro=tp.cod_pro and tl.cod_loj=te.cod_loj "
    strSQL = strSQL & " and tp.cod_pro='" & Me.Tag & "' "
    strSQL = strSQL & "order by tl.cod_loj"
    Adodc1.RecordSource = strSQL
    Adodc1.Refresh
End Sub

