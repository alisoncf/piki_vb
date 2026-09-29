VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiCargos 
   Caption         =   "Edição de Cargos"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   9090
   Icon            =   "frmEdiCargos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   9090
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9090
      TabIndex        =   1
      Top             =   4455
      Width           =   9090
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   7440
         Picture         =   "frmEdiCargos.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmEdiCargos.frx":02B5
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiCargos.frx":03FF
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   1170
      Left            =   1080
      Top             =   5400
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   2064
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
      DataSourceName  =   ""
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
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiCargos.frx":0789
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadCargos
    With ado
        If .RecordCount <= 0 Then
            MsgBox "Não há o que ser Editado!", vbExclamation
            Exit Sub
        End If
    
        fr.tag = !cod_Car
        P_LocalizaOperacao !venda_car, fr.cmbxVenda
        fr.txtbxNome.Text = !nome_car
        fr.Show 0, mdiFrm
    End With
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadCargos
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        strSQL = "select * from tab_bancos order by cod_bco"
    Else
        strSQL = "select * from tab_bancos order by nome_bco"
    End If
    
    
    adc.RecordSource = strSQL
    adc.Refresh
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    
    
    'Me.Height = 5205
    'Me.Width = 5145
    Set ado = New ADODB.Recordset
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select * from tab_cargos order by nome_car"
    FechaObjeto ado
    ado.Open strSQL, Conn_Client
    Set dtgrd.DataSource = ado
End Sub

