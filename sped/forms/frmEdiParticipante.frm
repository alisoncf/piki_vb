VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiParticipante 
   Caption         =   "Edição de Participantes (SPED)"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12240
   Icon            =   "frmEdiParticipante.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcTop 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   12240
      TabIndex        =   6
      Top             =   5040
      Width           =   12240
      Begin VB.TextBox txtPesquisar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   120
         Width           =   4695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   165
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
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
      Caption         =   "Participantes"
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "codigo"
         Caption         =   "codigo"
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
         DataField       =   "NOME"
         Caption         =   "NOME"
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
         DataField       =   "cod_pais"
         Caption         =   "País"
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
         DataField       =   "CNPJ"
         Caption         =   "CNPJ"
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
      BeginProperty Column04 
         DataField       =   "CPF"
         Caption         =   "CPF"
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
      BeginProperty Column05 
         DataField       =   "IE"
         Caption         =   "IE"
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
      BeginProperty Column06 
         DataField       =   "cod_mun"
         Caption         =   "Município"
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
      BeginProperty Column07 
         DataField       =   "ENDERECO"
         Caption         =   "ENDERECO"
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
      BeginProperty Column08 
         DataField       =   "NUMERO"
         Caption         =   "NUMERO"
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
      BeginProperty Column09 
         DataField       =   "COMPLEMENTO"
         Caption         =   "COMPLEMENTO"
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
      BeginProperty Column10 
         DataField       =   "BAIRRO"
         Caption         =   "BAIRRO"
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
      BeginProperty Column11 
         DataField       =   "cod_cli"
         Caption         =   "cod_cli"
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
      BeginProperty Column12 
         DataField       =   "cod_for"
         Caption         =   "cod_for"
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
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6900,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   659,906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2055,118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1755,213
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1620,284
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   840,189
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   12240
      TabIndex        =   0
      Top             =   5505
      Width           =   12240
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   10200
         Picture         =   "frmEdiParticipante.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSelec 
         Caption         =   "Selecionar"
         Height          =   735
         Left            =   3960
         Picture         =   "frmEdiParticipante.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Editar"
         Height          =   735
         Left            =   2040
         Picture         =   "frmEdiParticipante.frx":0F9B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiParticipante.frx":14D7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmEdiParticipante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdEdit_Click()
    frmCadParticipante.tag = ado!codigo
    frmCadParticipante.Show 1
    
    
    
    
End Sub

Private Sub cmdNovo_Click()
    frmCadParticipante.Show 1
End Sub

Private Sub cmdSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub
Sub Busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from sped_participante where codigo <>0 "
    strSQL = strSQL & " and (cnpj like '" & txtPesquisar.Text & "%' "
    strSQL = strSQL & " or cpf like '" & txtPesquisar.Text & "%' "
    strSQL = strSQL & " or nome like '%" & txtPesquisar.Text & "%' )"
    strSQL = strSQL & " order by nome "
    
    abreObjetoAdo ado, strSQL, grid
    Exit Sub
trata_Erro:
    MostraErro
End Sub

Private Sub cmdSelec_Click()
    If Me.tag <> "" Then
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    If Me.tag = "" Then
        cmdSelec.Enabled = False
    End If
    Busca
    
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    resizeForm Me
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcTop
    posicionaADireita cmdSair, pcBot
    
End Sub

Private Sub grid_DblClick()
    If Me.tag <> "" Then
        cmdSelec_Click
    Else
        cmdEdit_Click
    End If
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grid_DblClick
    End If
    
End Sub

Private Sub txtPesquisar_Change()
    Busca
End Sub

Private Sub txtPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Busca
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        grid.SetFocus
    End If
End Sub
