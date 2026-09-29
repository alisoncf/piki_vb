VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAgenda 
   Caption         =   "Agenda Telefônica"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   615
      Left            =   6240
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo Contato"
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   615
      Left            =   8640
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9551
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "nome"
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
      BeginProperty Column01 
         DataField       =   "telefone"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   3314,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3374,929
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTel 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisar"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   690
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdEditar_Click()
    frmNovoContato.tag = ado!id
    frmNovoContato.txtNome.text = ado!nome
    frmNovoContato.txtTel.text = ado!telefone
    frmNovoContato.Show 1
    ado.Requery
End Sub

Private Sub cmdExcluir_Click()
    On Error Resume Next
    Dim strSQL As String
    strSQL = "delete from contato where id=" & ado!id
    Conn.Execute strSQL
    busca
    
End Sub

Private Sub cmdNovo_Click()
    frmNovoContato.Show 1
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    
    
    
    
End Sub

Private Sub txtTel_Change()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = BuscaContato(txtTel)
    abreObjetoAdo ado, strSQL, grid
    
End Sub
