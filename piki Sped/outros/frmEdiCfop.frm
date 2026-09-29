VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiCfop 
   Caption         =   "Edição de CFOP"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   Icon            =   "frmEdiCfop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbxExpr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   4440
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiCfop.frx":038A
      Height          =   4095
      Left            =   0
      OleObjectBlob   =   "frmEdiCfop.frx":039B
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   6870
      TabIndex        =   0
      Top             =   4845
      Width           =   6900
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   5520
         Picture         =   "frmEdiCfop.frx":0F44
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   4080
         Picture         =   "frmEdiCfop.frx":10AF
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSelect 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   2760
         Picture         =   "frmEdiCfop.frx":1439
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnEditar 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1440
         Picture         =   "frmEdiCfop.frx":1595
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiCfop.frx":1F97
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo:"
      Height          =   615
      Left            =   4080
      TabIndex        =   9
      Top             =   4200
      Width           =   2775
      Begin VB.OptionButton optnSaida 
         Caption         =   "Saída"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optnEntrada 
         Caption         =   "Entrada"
         Height          =   195
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label lblBuscar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscar"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "frmEdiCfop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEditar_Click()
    On Error GoTo trata_Erro
    With dt.Recordset
        frmCadCFOP.Tag = !cfop_cfo
        frmCadCFOP.txtbxCFOP.Text = !cfop_cfo
        frmCadCFOP.txtbxNOme.Text = !nome_cfo
        If !tipo_cfo = "S" Then
            frmCadCFOP.optnSaida.Value = True
        Else
            frmCadCFOP.optnEntrada.Value = True
        End If
        frmCadCFOP.Show 1
        busca
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnExcl_Click()
    On Error GoTo trata_Erro
    If MsgBox("Deseja realmente excluir?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    P_DeleteCFOP dt.Recordset!cfop_cfo
    Exit Sub
trata_Erro:
    
    
End Sub

Private Sub cmndbtnNovo_Click()
    frmCadCFOP.Show 1
    busca
End Sub

Private Sub cmndbtnSair_Click()
    If Me.Tag = "" Then
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub cmndbtnSelect_Click()
    Me.Tag = "x"
    Me.Hide
End Sub

Sub busca()
    Dim strSQL As String
    On Error GoTo trata_Erro:
    strSQL = F_SelectCFOP(txtbxExpr.Text, IIf(optnSaida.Value, "S", "E"))
    
    dt.RecordSource = strSQL
    dt.Refresh
    Exit Sub
trata_Erro:
    P_MostraErro

End Sub

Private Sub dbgrd_DblClick()
    If Me.Tag = "" Then
        cmndbtnEditar_Click
    Else
        cmndbtnSelect_Click
    End If
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    End If
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    dt.Connect = banco.Connect
    
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
    
End Sub
