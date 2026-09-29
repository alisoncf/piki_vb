VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiGru 
   Caption         =   "Edição de Grupos de Clientes"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5100
   Icon            =   "frmEdiGru.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid grid 
      Height          =   3495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6165
      _Version        =   393216
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
      Caption         =   "Grupos"
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   4095
      Width           =   4935
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiGru.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmEdiGru.frx":1054
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   3240
         Picture         =   "frmEdiGru.frx":119E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtbxBusca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   3735
      Width           =   4935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   3495
      Width           =   540
   End
End
Attribute VB_Name = "frmEdiGru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnEdit_Click()
    If Me.tag = "" Then
        With ado
            frmCadgru.tag = !cod_bco
            frmCadBan.txtbxCod.Text = !cod_bco
            frmCadBan.txtbxCod.Enabled = False
            frmCadBan.txtbxNome.Text = !nome_bco
            frmCadBan.Show 1
        End With
    Else
        Me.Hide
    End If
End Sub

Private Sub cmndbtnNew_Click()
    frmCadBan.Show 1
    
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    
    If Me.tag <> "" Then
    cmndbtnEdit.caption = "&Selecionar"
    P_Busca
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    
    
End Sub
Sub P_Busca()
    Dim strSQL As String
    strSQL = "select * from tab_grupocli "
    strSQL = strSQL & " where cod_gru='" & F_AjustaTexto(txtbxBusca.Text) & "'"
    strSQL = strSQL & " or nome_gru like '%" & F_AjustaTexto(txtbxBusca.Text) & "%'"
    strSQL = strSQL & " order by nome_gru"
    
    abreObjetoAdo ado, strSQL, grid
    If ado.RecordCount = 1 Then
        If Me.tag <> "" Then
            On Error Resume Next
            Me.Hide
        End If
    Else
    
    End If
End Sub
Private Sub txtbxBusca_Change()
    P_Busca
End Sub

Private Sub txtbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_Busca
    End If
    
End Sub

