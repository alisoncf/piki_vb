VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiPadrao1 
   Caption         =   "Padrão1"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11250
   Icon            =   "frmEdiPadrao1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11220
      TabIndex        =   2
      Top             =   4995
      Width           =   11250
      Begin VB.TextBox txtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4695
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
         TabIndex        =   7
         Top             =   0
         Width           =   540
      End
   End
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
      ScaleWidth      =   11250
      TabIndex        =   1
      Top             =   5610
      Width           =   11250
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   9360
         Picture         =   "frmEdiPadrao1.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   4440
         Picture         =   "frmEdiPadrao1.frx":6DE0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3000
         Picture         =   "frmEdiPadrao1.frx":8ADA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1560
         Picture         =   "frmEdiPadrao1.frx":A7D4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiPadrao1.frx":AD10
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8493
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
      ColumnCount     =   3
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
         DataField       =   "descricao"
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
         DataField       =   "complemento"
         Caption         =   "Complemento"
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
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5114,835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2594,835
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiPadrao1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Public campo1 As String
Public campo2 As String
Public campo3 As String
Public tabela As String

Private Sub cmdEditar_Click()
    frmCadPadrao1.tabela = tabela
    frmCadPadrao1.campo1 = campo1
    frmCadPadrao1.txtCodigo.Enabled = False
    frmCadPadrao1.campo2 = campo2
    frmCadPadrao1.tag = ado!codigo
    frmCadPadrao1.txtCodigo = ado!codigo
    frmCadPadrao1.txtDescricao = ado!descricao
    frmCadPadrao1.campo3 = campo3
    If campo3 <> "" Then
        frmCadPadrao1.txtcomplemento = ado!complemento
    End If
    
    
    frmCadPadrao1.caption = Me.caption
    frmCadPadrao1.Show 1
    ado.Requery
End Sub

Private Sub cmdNovo_Click()
    frmCadPadrao1.tabela = tabela
    frmCadPadrao1.campo1 = campo1
    frmCadPadrao1.campo2 = campo2
    frmCadPadrao1.Show 1
End Sub

Private Sub cmdSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Unload Me
    End If
End Sub
Public Sub Inicia(tabela As String, campo1 As String, campo2 As String, campo3 As String, captionForm As String)
    Me.campo1 = campo1
    Me.campo2 = campo2
    Me.campo3 = campo3
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
    Else
        cmdSelecionar.Enabled = True
    End If
    busca
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBusca
    posicionaADireita cmdSair, pcBusca
    
    On Error Resume Next
    grid.Columns(0).width = 1000
    grid.Columns(1).width = 5000
    grid.Columns(2).width = 3000
    
End Sub

Private Sub grid_DblClick()
    If Me.tag = "" Then
        cmdSelecionar_Click
    Else
        cmdEditar_Click
    End If
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
    On Error Resume Next
    strSQL = "select " & campo1 & " as codigo, " & campo2 & " as descricao  "
    If campo3 <> "" Then
        strSQL = strSQL & "," & campo3 & " as complemento "
    End If
    
    
    strSQL = strSQL & " from " & tabela & " where " & campo2 & " like '" & txtBuscar.Text & "%' "
    strSQL = strSQL & " or " & campo1 & " like '%" & txtBuscar.Text & "%'"
    
    If tabela = "sped_cst_icms" Then
        strSQL = strSQL & " and crt='" & OpcaoEmpresaCRT & "'  "
    End If
    
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
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        grid.SetFocus
    End If
End Sub
