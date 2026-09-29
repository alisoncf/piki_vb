VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmAddPerm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inclusão de Permissões"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnSair 
      Height          =   285
      Left            =   8040
      Picture         =   "frmAddPerm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5250
      Width           =   375
   End
   Begin VB.CommandButton cmndbtnExcluir 
      Height          =   285
      Left            =   7200
      Picture         =   "frmAddPerm.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5250
      Width           =   375
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Height          =   285
      Left            =   6720
      Picture         =   "frmAddPerm.frx":2532
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5250
      Width           =   375
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtbxPermiss 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   5250
      Width           =   3615
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmAddPerm.frx":28BC
      Height          =   4695
      Left            =   3000
      OleObjectBlob   =   "frmAddPerm.frx":28CD
      TabIndex        =   2
      Top             =   240
      Width           =   5415
   End
   Begin VB.ListBox lstGru 
      Appearance      =   0  'Flat
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incluir Nova Permissão:"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   5040
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo de Permissões"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1500
   End
End
Attribute VB_Name = "frmAddPerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnExcluir_Click()
    Dim strSQL As String
    strSQL = "delete from tab_permissoes where cod_per='" & dt.Recordset!cod_per & "' "
    Banco.Execute strSQL
    buscaPerm
End Sub

Private Sub cmndbtnSalvar_Click()
    If lstGru.tag = "" Then
        MsgBox "Escolha um grupo na lista!", vbAbortRetryIgnore
        lstGru.SetFocus
        Exit Sub
    End If
    grava
    buscaPerm
End Sub

Private Sub Form_Load()
    P_PreenchePermGru lstGru
    dt.Connect = Banco.Connect
End Sub



Private Sub lstGru_Click()
    lstGru.tag = F_BuscaUmCampo("cod_gru", "tab_permgru", "nome_gru='" & lstGru.Text & "'")
    buscaPerm
End Sub
Sub buscaPerm()
    Dim strSQL As String
    strSQL = "select * from tab_permissoes where cod_gru='" & lstGru.tag & "'"
    strSQL = strSQL & " order by nome_per"
    dt.RecordSource = strSQL
    dt.Refresh
    
    
End Sub
Sub grava()
    Dim strSQL As String
    If F_BuscaUmCampo("cod_per", "tab_permissoes", "nome_per='" & txtbxPermiss.Text & "'") <> "" Then
        MsgBox "Já existe uma permissão com este nome!", vbExclamation
        txtbxPermiss.SetFocus
        Exit Sub
    End If
    strSQL = "insert into tab_permissoes(cod_per,cod_gru,nome_per)values('',"
    strSQL = strSQL & "'" & lstGru.tag & "',"
    strSQL = strSQL & "'" & txtbxPermiss.Text & "')"
    Banco.Execute strSQL
    
    
    
End Sub
Private Sub txtbxPermiss_GotFocus()
    P_SelecTexto txtbxPermiss
End Sub

Private Sub txtbxPermiss_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPermiss_LostFocus()
    P_AjustaTexto txtbxPermiss, 1
End Sub
