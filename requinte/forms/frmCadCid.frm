VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCadCid 
   Caption         =   "Cadastro de Cidades"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   8505
   Icon            =   "frmCadCid.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8505
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4470
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7885
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Cidades (IBGE)"
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5369,953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   510,236
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmndbtnSelect 
      Caption         =   "Selecionar"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnEditar 
      Caption         =   "&Editar"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnNovo 
      Caption         =   "&Novo"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   1575
   End
End
Attribute VB_Name = "frmCadCid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEditar_Click()
    On Error GoTo trata_Erro
    With adc.Recordset
        frmIncCid.tag = !cod_cid
        frmIncCid.cmbxUf.Text = !uf_cid
        frmIncCid.txtbxDesc.Text = !nome_cid
        
        frmIncCid.Show 1
    End With
Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmndbtnNovo_Click()
    frmIncCid.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub cmndbtnSelect_Click()
    Me.Hide
End Sub

Private Sub dbgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEditar_Click
    Else
        cmndbtnSelect_Click
    End If
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    
    strSQL = "select * from tab_cidades order by nome_cid "
    
    dt.RecordSource = strSQL
    dt.Refresh
    P_CentraForm Me
End Sub

Private Sub tlbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        With Adodc.Recordset
        frmIncCid.tag = !cod_cid
        frmIncCid.cmbxUf.Text = !uf_cid
        frmIncCid.txtbxDesc.Text = !nome_cid
        
        frmIncCid.Show 1
        End With
    ElseIf Button.Index = 2 Then
        On Error Resume Next
        frmIncCid.Show 1
    ElseIf Button.Index = 3 Then
    ElseIf Button.Index = 6 Then
        Unload Me
    End If
    
End Sub
