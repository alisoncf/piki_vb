VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCidadesIBGE 
   Caption         =   "Cidades IBGE"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9510
   Icon            =   "frmCidadesIBGE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   9480
      TabIndex        =   1
      Top             =   5850
      Width           =   9510
      Begin VB.TextBox txtbxNome 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmCidadesIBGE.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSelect 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   1800
         Picture         =   "frmCidadesIBGE.frx":109A6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sai&r"
         Height          =   615
         Left            =   7320
         Picture         =   "frmCidadesIBGE.frx":10A75
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   540
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmCidadesIBGE.frx":10B6A
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9975
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
End
Attribute VB_Name = "frmCidadesIBGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmndbtnSair_Click()
    If Me.tag <> "" Then
        Me.tag = ""
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub cmndbtnSelect_Click()
    If Me.tag <> "" Then
        Me.tag = ado!cod_cid
        Me.Hide
    End If
End Sub

Private Sub dtgrd_DblClick()
    cmndbtnSelect_Click
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub Form_Activate()
    Me.Visible = True
    bActive = True
End Sub

Private Sub Form_Load()
    Me.Visible = False
    Set ado = New ADODB.Recordset
End Sub

Private Sub txtbxNome_Change()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "select * from cidades where nome_cid like '" & txtbxNome.Text & "%' or cod_cid='" & txtbxNome & "'"
    
    abreObjetoAdo ado, strSQL, dtgrd
    
    If Me.tag <> "" And ado.RecordCount = 1 Then
        Me.Hide
    End If
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    ElseIf KeyCode = vbKeyUp Or keycod = vbKeyDown Then
        dtgrd.SetFocus
    End If
End Sub
