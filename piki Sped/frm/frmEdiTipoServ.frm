VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiTipoServ 
   Caption         =   "Edição de Tipos de Serviço"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10845
   Icon            =   "frmEdiTipoServ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   10845
      TabIndex        =   1
      Top             =   4620
      Width           =   10845
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   5400
         Picture         =   "frmEdiTipoServ.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   7200
         Picture         =   "frmEdiTipoServ.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   615
         Left            =   9000
         Picture         =   "frmEdiTipoServ.frx":1568
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtPesquisar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "Sele&cionar"
         Height          =   615
         Left            =   3120
         Picture         =   "frmEdiTipoServ.frx":16D3
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7858
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "Id"
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
         DataField       =   "nome"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5265,071
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiTipoServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdEdit_Click()
    On Error GoTo trata_erro
    frmTipoServico.tag = ado!id
    frmTipoServico.txtId.Text = ado!id
    frmTipoServico.txtNome.Text = ado!nome
    frmTipoServico.Show 1
    busca
    
    Exit Sub
trata_erro:
End Sub

Private Sub cmdNovo_Click()
    frmTipoServico.Show 1
    busca
End Sub

Private Sub cmdSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub cmdSelecionar_Click()
    If Me.tag <> "" Then
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    busca
    If Me.tag = "" Then
        cmdSelecionar.Value = False
    Else
        cmdSelecionar.Value = True
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    
    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select * from tipo_servico "
    strSQL = strSQL & " where nome like '" & txtPesquisar.Text & "%' order by nome"
    abreObjetoAdo ado, strSQL, grid
    
End Sub

Private Sub txtPesquisar_Change()
    busca
End Sub

Private Sub txtPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
