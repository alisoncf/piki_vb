VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiPerm 
   Caption         =   "Permissões"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   Icon            =   "frmediperm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmediperm.frx":038A
      Height          =   5655
      Left            =   120
      OleObjectBlob   =   "frmediperm.frx":039B
      TabIndex        =   6
      Top             =   120
      Width           =   11655
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   11655
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnNova 
         Caption         =   "&Nova"
         Height          =   615
         Left            =   5040
         Picture         =   "frmediperm.frx":10E4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxTipoPer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   6360
         Picture         =   "frmediperm.frx":1248
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   10320
         Picture         =   "frmediperm.frx":1655
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   9000
         Picture         =   "frmediperm.frx":174A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnEdita 
         Caption         =   "Editar"
         Height          =   615
         Left            =   7680
         Picture         =   "frmediperm.frx":1AD4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   615
         Left            =   3720
         Picture         =   "frmediperm.frx":1E5E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Expressão:"
         Height          =   195
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmEdiPerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    strSQL = permissao.F_SelectPermissoes(txtbxExpr.Text, txtbxTipoPer.tag)
    dt.RecordSource = strSQL
    dt.Refresh
    
End Sub

Private Sub cmndbtnNova_Click()
    frmCadPermissoes.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    adc.ConnectionString = Conn.ConnectionString
End Sub

Private Sub txtbxExpr_Change()
    cmndbtnBuscar_Click
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxTipoPer_Change()
    txtbxTipoPer.tag = ""
End Sub

Private Sub txtbxTipoPer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaTipo
    End If
    
End Sub
Sub buscaTipo()
    Dim t As grupoPer
    t = F_BuscaGrupoPer(txtbxTipoPer.Text)
    
    txtbxTipoPer.Text = t.nome
    txtbxTipoPer.tag = t.codigo
    cmndbtnBuscar_Click
    
    
End Sub
