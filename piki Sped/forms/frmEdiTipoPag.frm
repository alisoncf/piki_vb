VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiTipoPag 
   Caption         =   "Espécie de Tipo de Pagamento"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7260
   Icon            =   "frmEdiTipoPag.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBusca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   7095
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   7230
      TabIndex        =   2
      Top             =   4305
      Width           =   7260
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   7560
         Picture         =   "frmEdiTipoPag.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1920
         Picture         =   "frmEdiTipoPag.frx":10D98
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiTipoPag.frx":10EE2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3720
         Picture         =   "frmEdiTipoPag.frx":1126C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox txtExpr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   6000
      Width           =   4455
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frmEdiTipoPag.frx":116F3
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6800
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
      Caption         =   "Edição de Formas de Pagamento"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cod_tip"
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
         DataField       =   "nome_tip"
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
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4605,166
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   735
   End
End
Attribute VB_Name = "frmEdiTipoPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdSelecionar_Click()
    On Error Resume Next
    If Me.tag <> "" Then
    Me.tag = ado!cod_tip
    Me.Hide
    End If
End Sub

Private Sub cmndbtnNew_Click()
    frmCadTipoPag.Show 1
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        busca
        cmdSelecionar.Enabled = True
    Else
        cmdSelecionar.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
End Sub

Private Sub txtBusca_Change()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select * from tab_tipopag "
    strSQL = strSQL & " where cod_tip like '" & txtBusca.Text & "%' "
    strSQL = strSQL & " or nome_tip like '" & txtBusca.Text & "%' "
    
    strSQL = strSQL & " order by cod_tip "
    
    
    abreObjetoAdo ado, strSQL, grid
    
    
End Sub

Private Sub txtBusca_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
    
End Sub
