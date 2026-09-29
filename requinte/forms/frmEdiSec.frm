VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiSec 
   Caption         =   "Edição de Seções"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8190
   Icon            =   "frmEdiSec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   8160
      TabIndex        =   1
      Top             =   4965
      Width           =   8190
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   5880
         Picture         =   "frmEdiSec.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1920
         Picture         =   "frmEdiSec.frx":10D98
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiSec.frx":10EE2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3720
         Picture         =   "frmEdiSec.frx":1126C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox txtExpr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   5295
   End
   Begin MSDataGridLib.DataGrid grade 
      Bindings        =   "frmEdiSec.frx":115AE
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7223
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
      Caption         =   "Seções"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cod_sec"
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
         DataField       =   "nome_sec"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6089,953
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisar"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   735
   End
End
Attribute VB_Name = "frmEdiSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ado As ADODB.Recordset
Dim bActive As Boolean

Private Sub cmdSelecionar_Click()
    If Me.tag <> "" Then
        Me.Hide
    End If
End Sub

Private Sub cmndbtnNew_Click()
    frmCadSec.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag <> "" Then
        Me.tag = ""
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        busca
    End If
    bActive = True
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    bActive = False
End Sub

Sub busca()
    Dim strSQL As String
    strSQL = "select * from tab_secoes where 1>0 "
    
    strSQL = strSQL & " and (nome_sec like '" & txtExpr.Text & "%' "
    strSQL = strSQL & " or cod_sec='" & txtExpr.Text & "') "
    abreObjetoAdo ado, strSQL, grade
    
    If Me.tag <> "" And ado.RecordCount = 1 Then
        cmdSelecionar_Click
    End If
    
    
End Sub

Private Sub txtExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
