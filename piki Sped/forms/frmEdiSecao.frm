VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiSecao 
   Caption         =   "Edição de Seções"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8250
   Icon            =   "frmEdiSecao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   8250
      TabIndex        =   3
      Top             =   4350
      Width           =   8250
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiSecao.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1560
         Picture         =   "frmEdiSecao.frx":6C1D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3000
         Picture         =   "frmEdiSecao.frx":7159
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   4440
         Picture         =   "frmEdiSecao.frx":8E53
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   6240
         Picture         =   "frmEdiSecao.frx":AB4D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   8220
      TabIndex        =   0
      Top             =   3735
      Width           =   8250
      Begin VB.TextBox txtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   540
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   3495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6165
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
      Caption         =   "Lista de Seções"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "codigo"
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
         DataField       =   "Descricao"
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
         BeginProperty Column00 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5595,024
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiSecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Public campo1 As String
Public campo2 As String
Public tabela As String

Private Sub cmdEditar_Click()
    frmCadSec.tabela = tabela
    frmCadSec.campo1 = campo1
    frmCadSec.txtCodigo.Enabled = False
    frmCadSec.campo2 = campo2
    frmCadSec.tag = ado!codigo
    frmCadSec.txtCodigo = ado!codigo
    frmCadSec.txtDescricao = ado!Descricao
    
    frmCadSec.caption = Me.caption
    frmCadSec.Show 1
    ado.Requery
End Sub

Private Sub cmdNovo_Click()
    frmCadSec.tabela = tabela
    frmCadSec.campo1 = campo1
    frmCadSec.campo2 = campo2
    frmCadSec.Show 1
End Sub

Private Sub cmdSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Unload Me
    End If
End Sub
Public Sub Inicia(tabela As String, campo1 As String, campo2 As String, Optional captionForm As String = "")
    Me.campo1 = campo1
    Me.campo2 = campo2
    Me.tabela = tabela
    If captionForm <> "" Then
        Me.caption = captionForm
    Else
        Me.caption = "Edição: (" & tabela & ") " & campo1 & "/" & campo2
    End If
End Sub


Private Sub cmdSelecionar_Click()
    On Error Resume Next
    Me.tag = ado!codigo
    Me.Hide
End Sub

Private Sub Form_Activate()
    If Me.tag = "" Then
        cmdSelecionar.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    tabela = "tab_secoes"
    campo1 = "cod_sec"
    campo2 = "nome_sec"
    resizeForm Me
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBusca
End Sub

Private Sub grid_DblClick()
    cmdSelecionar_Click
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSelecionar_Click
    End If
End Sub

Private Sub txtBuscar_Change()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    
    strSQL = "select " & campo1 & " as codigo, " & campo2 & " as descricao  from " & tabela & " where " & campo2 & " like '" & txtBuscar.Text & "%' "
    strSQL = strSQL & " or " & campo1 & " like '" & txtBuscar.Text & "%'"
    strSQL = strSQL & " order by " & campo1
    abreObjetoAdo ado, strSQL, grid
    
    If ado.RecordCount = 1 And Me.tag <> "" Then
        Me.tag = "x"
        Me.Hide
    End If
    
End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
