VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiCid 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edição de Cidades"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxNome 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   6015
   End
   Begin VB.CommandButton cmndbtnNovo 
      Caption         =   "&Novo"
      Height          =   615
      Left            =   120
      Picture         =   "frmEdiCid.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnSelect 
      Caption         =   "&Selecionar"
      Height          =   615
      Left            =   2160
      Picture         =   "frmEdiCid.frx":015C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "Sai&r"
      Height          =   615
      Left            =   6600
      Picture         =   "frmEdiCid.frx":022B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnEditar 
      Caption         =   "&Editar"
      Height          =   615
      Left            =   4200
      Picture         =   "frmEdiCid.frx":0320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem:"
      Height          =   615
      Left            =   6360
      TabIndex        =   6
      Top             =   5520
      Width           =   2055
      Begin VB.OptionButton optnCid 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optnUf 
         Caption         =   "UF"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   1800
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   1
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
      Bindings        =   "frmEdiCid.frx":046A
      Height          =   5415
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9551
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "cod_cid"
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
         DataField       =   "nome_cid"
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
         DataField       =   "uf_cid"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5639,812
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   569,764
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   540
   End
End
Attribute VB_Name = "frmEdiCid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEditar_Click()
    On Error GoTo trata_Erro
    With adc.Recordset
        frmCadCid.tag = !cod_cid
        frmCadCid.txtCod.text = !cod_cid
        frmCadCid.cmbxUF.text = !uf_cid
        frmCadCid.txtbxDesc.text = !nome_cid
        frmCadCid.txtEstado.text = !estado_cid
        frmCadCid.Show 1
    End With
Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmndbtnNovo_Click()
    frmCadCid.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub cmndbtnSelect_Click()
    On Error Resume Next
    Me.tag = adc.Recordset!cod_cid
    Me.Hide
End Sub

Private Sub dtgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEditar_Click
    Else
        cmndbtnSelect_Click
    End If
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub Form_Activate()
    busca
    If Me.tag = "" Then
        cmndbtnSelect.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    adc.ConnectionString = strConnect
    
    
    
End Sub
Sub busca()
    On Error GoTo erro
    strSQL = "select * from " & TAB_CIDADES
    strSQL = strSQL & " where nome_cid like '%" & txtbxNOme.text & "%'"
    strSQL = strSQL & " or cod_cid like '" & txtbxNOme.text & "%'"
    
    If optnCid.Value Then
        strSQL = strSQL & " order by nome_cid "
    Else
        strSQL = strSQL & " order by uf_cid,nome_cid "
    End If
    adc.RecordSource = strSQL
    adc.Refresh
    Exit Sub
erro:
    MsgBox "String de conexão " & adc.ConnectionString, vbExclamation
    
End Sub


Private Sub optnCid_Click()
    busca
End Sub

Private Sub optnUf_Click()
    busca
End Sub

Private Sub txtbxNome_Change()
    busca
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
