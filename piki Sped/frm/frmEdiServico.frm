VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiServico 
   Caption         =   "Serviços"
   ClientHeight    =   7245
   ClientLeft      =   1170
   ClientTop       =   450
   ClientWidth     =   11625
   Icon            =   "frmEdiServico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   11625
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11625
      TabIndex        =   0
      Top             =   6390
      Width           =   11625
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "Sele&cionar"
         Height          =   615
         Left            =   4200
         Picture         =   "frmEdiServico.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtPesquisar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3735
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   615
         Left            =   9600
         Picture         =   "frmEdiServico.frx":0DCC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   7800
         Picture         =   "frmEdiServico.frx":0F37
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   6000
         Picture         =   "frmEdiServico.frx":1939
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
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10186
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
         DataField       =   "nome_tipo"
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
         DataField       =   "comissao"
         Caption         =   "Comissão"
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
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5100,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2594,835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   929,764
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ado As ADODB.Recordset

Private Sub cmdEdit_Click()
    On Error GoTo trata_erro
    
    frmServico.tag = ado!id
    frmServico.Show 1
    ado.Requery
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmdNovo_Click()
    frmServico.Show 1
    
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
    If Me.tag = "" Then
        cmdSelecionar.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
        
    
End Sub

Private Sub txtPesquisar_Change()
    busca
End Sub
Sub busca(Optional orderBy = "s.nome")
    Dim strSQL As String
    Dim cod As String
    If orderBy = "" Then
        orderBy = " s.nome"
    End If
    cod = F_SoNumero(txtPesquisar.Text)
    strSQL = "select tt.nome_tip as nome_tip,ts.* from servico as s, tipo_servico as ts "
    strSQL = strSQL & " where ts.id=s.id_tipo"
    strSQL = strSQL & " and (s.nome like '%" & txtPesquisar.Text & "%'"
    If cod <> "" Then
        strSQL = strSQL & " or s.id=" & cod
    End If
    strSQL = strSQL & ") order by "
    strSQL = strSQL & orderBy
    
    
    
    abreObjetoAdo ado, strSQL, grid
    
    
    
End Sub

Private Sub txtPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
