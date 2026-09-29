VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAcertos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acertos"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4815
      Left            =   0
      TabIndex        =   11
      Top             =   840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8493
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
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "Visualizar"
      Height          =   615
      Left            =   3360
      Picture         =   "frmacertos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtbxProduto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   6000
      Width           =   4935
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   615
      Left            =   6360
      Picture         =   "frmacertos.frx":00F9
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnImpr 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   1680
      Picture         =   "frmacertos.frx":0264
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnAbre 
      Caption         =   "Abrir"
      Height          =   615
      Left            =   0
      Picture         =   "frmacertos.frx":0671
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   8010
      TabIndex        =   0
      Top             =   0
      Width           =   8040
      Begin VB.TextBox txtbxRepr 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7575
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   330
         Left            =   7680
         Picture         =   "frmacertos.frx":0AA5
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmndbtnBuscaPro 
      Height          =   285
      Left            =   4920
      Picture         =   "frmacertos.frx":0B9E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   5760
      Width           =   600
   End
End
Attribute VB_Name = "frmAcertos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bActive As Boolean
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdVisualizar_Click()
    
        '  On Error Resume Next
        P_VisualizaAcerto ado!numero
    
End Sub

Private Sub cmndbtnBuscaPro_Click()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxProduto.Text)
    txtbxProduto.Text = p.nome_pro
    txtbxProduto.tag = p.cod_pro
    busca
End Sub

Private Sub cmndbtnCliente_Click()
    Dim r As representante.representante
    
    r = representante.F_BuscaRep(txtbxRepr.Text)
    If r.cod_Rep <> "" Then
        txtbxRepr.Text = r.nome_rep
        txtbxRepr.tag = r.cod_Rep
        busca
        
    End If
    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select distinct data_ace as data,valor_ace as Valor,ta.seq_ace as Numero,"
    strSQL = strSQL & " if(status_ace=1,'Aberto',if(status_ace=2,'Fechado','Cancelado'))  as status "
    strSQL = strSQL & " from tab_acertos as ta, tab_itacertos as ti "
    strSQL = strSQL & " where ta.seq_ace=ti.seq_ace "
    strSQL = strSQL & " and ta.cod_rep='" & txtbxRepr.tag & "' "
    If txtbxProduto.tag <> "" Then
        strSQL = strSQL & " and cod_pro='" & txtbxProduto.tag & "' "
    End If
    
    
    strSQL = strSQL & " order by data_ace desc "
    
    abreObjetoAdo ado, strSQL, grid
    
    
End Sub


Private Sub cmndbtnImpr_Click()
    
    imprAcertoData txtbxRepr.tag, ado!data_ace
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If bActive Then Exit Sub
    cmndbtnCliente_Click
    bActive = True
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    bActive = False
End Sub

Private Sub txtbxRepr_GotFocus()
    P_SelecTexto txtbxRepr
End Sub

Private Sub txtbxRepr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
End Sub

