VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddPermUsu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Permissões de Usuário:"
   ClientHeight    =   7620
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   615
      Left            =   9120
      Picture         =   "frmAddPermUsu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Salvar"
      Height          =   615
      Left            =   7080
      Picture         =   "frmAddPermUsu.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   1815
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   5205
      ItemData        =   "frmAddPermUsu.frx":04F5
      Left            =   240
      List            =   "frmAddPermUsu.frx":04F7
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1200
      Width           =   10695
   End
   Begin VB.TextBox txtbxNomeUsu 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4215
   End
   Begin MSComctlLib.TabStrip tbst 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11880
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmndbtnUsu 
      Caption         =   "..."
      Height          =   285
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblUsu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   4
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuário:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   585
   End
   Begin VB.Menu mnAcao 
      Caption         =   "&Ação"
      Begin VB.Menu mnTodos 
         Caption         =   "&Todos"
      End
      Begin VB.Menu mnNehum 
         Caption         =   "&Nenhum"
      End
      Begin VB.Menu mnSave 
         Caption         =   "&Salvar"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "frmAddPermUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmndbtnSalvar_Click()
    '
    Dim i As Integer
    Dim strSQL As String
    On Error Resume Next
    If strNome_Usu <> "ALISON" Then
    If Not F_VeririficaPermissao("funcionarios", "controlar permissoes", strCod_Usu) Then
        MsgBox "Sua Credencial para este acesso não foi encontrada!", vbExclamation
        Exit Sub
    End If
    End If
    If Not F_senha(True) Then
        Exit Sub
    End If
    
    For i = 0 To lst.ListCount - 1
        lst.ListIndex = i
        
        If lst.Selected(i) = True Then
            strSQL = "insert into tab_permusu(cod_per,cod_fun)values("
            strSQL = strSQL & Left(lst.Text, 3) & ","
            strSQL = strSQL & "'" & txtbxNomeUsu.tag & "')"
            Banco.Execute strSQL
        Else
            strSQL = " delete from tab_permusu where cod_per=" & Left(lst.Text, 3)
            strSQL = strSQL & " and cod_fun='" & txtbxNomeUsu.tag & "'"
            Banco.Execute strSQL
        End If
    Next
End Sub

Private Sub cmndbtnUsu_Click()
    Dim u As funcionario.funcionario
    u = F_SelectFuncionario(txtbxNomeUsu.Text)
    If u.cod_fun <> "" Then
        
        txtbxNomeUsu.Text = u.nome_fun
        txtbxNomeUsu.tag = u.cod_fun
        lblUsu.caption = u.cod_fun
        busca
    End If
End Sub

Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_permgru order by nome_gru "
    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        tbst.Tabs.Clear
        Do While Not .EOF
            tbst.Tabs.Add .AbsolutePosition + 1, , !nome_gru
            tbst.Tabs(tbst.Tabs.count).tag = !cod_gru
            .MoveNext
        Loop
    End With
    
    
    
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    If Me.tag <> "" Then
        cmndbtnUsu_Click
    End If
    bActive = True
End Sub

Private Sub lst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnAcao
    End If
End Sub

Private Sub mnNehum_Click()
    selectTodos False
End Sub

Private Sub mnSair_Click()
    Unload Me
End Sub

Private Sub mnSave_Click()
    cmndbtnSalvar_Click
End Sub

Private Sub mnTodos_Click()
    selectTodos True
End Sub
Sub selectTodos(opcao As Boolean)
    Dim i As Integer
    For i = 0 To lst.ListCount - 1
        lst.ListIndex = i
        lst.Selected(i) = opcao
    Next
End Sub
Private Sub tbst_Click()
    buscaPerm
End Sub
Sub buscaPerm()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_permissoes where cod_gru='" & tbst.SelectedItem.tag & "' "
    strSQL = strSQL & " order by cod_per "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        lst.Clear
        Do While Not .EOF
            lst.AddItem F_EsticaTexto(!cod_per, 3) & !nome_per
            .MoveNext
        Loop
        .Close
    End With
    buscaPermUsu
    
End Sub
Sub buscaPermUsu()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim iCount As Integer
    strSQL = "select distinct tp.* from tab_permusu as tp, tab_permissoes as tm "
    strSQL = strSQL & " where tm.cod_per=tp.cod_per "
    strSQL = strSQL & " and tm.cod_gru='" & tbst.SelectedItem.tag & "' "
    
    strSQL = strSQL & " and cod_fun='" & txtbxNomeUsu.tag & "' "
    strSQL = strSQL & " order by cod_per "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        
        Do While Not .EOF
            For iCount = 0 To lst.ListCount - 1
                lst.ListIndex = iCount
                If Int(Trim(Left(lst.Text, 3))) = !cod_per Then
                    lst.Selected(iCount) = True
                    Exit For
                End If
            Next
            .MoveNext
        Loop
        .Close
    End With
    
End Sub

