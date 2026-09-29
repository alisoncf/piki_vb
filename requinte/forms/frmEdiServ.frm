VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiServ 
   Caption         =   "Edição de Serviços"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12210
   Icon            =   "frmEdiServ.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   12210
   Begin MSDataGridLib.DataGrid grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9763
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
      ColumnCount     =   5
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
      BeginProperty Column02 
         DataField       =   "tiponome"
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
         DataField       =   "valor"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "comissao"
         Caption         =   "Comissão"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   5
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3809,764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1934,929
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1409,953
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12210
      TabIndex        =   0
      Top             =   5730
      Width           =   12210
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "Sele&cionar"
         Height          =   615
         Left            =   4680
         Picture         =   "frmEdiServ.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtPesquisar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   615
         Left            =   10440
         Picture         =   "frmEdiServ.frx":0DCC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   8520
         Picture         =   "frmEdiServ.frx":0F37
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   6600
         Picture         =   "frmEdiServ.frx":1939
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEdiServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub Form_Activate()
    If Me.tag = "" Then
        cmdSelecionar.Visible = False
    Else
        cmdSelecionar.Visible = True
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
End Sub

Sub busca()
    Dim strSQL As String
    strSQL = Servicos.SelectServico("")
    
    FechaObjeto ado
    ado.Open strSQL, Conn_Client
    
    Set grid.DataSource = ado
    
    
End Sub
