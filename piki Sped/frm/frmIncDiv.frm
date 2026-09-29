VERSION 5.00
Begin VB.Form frmIncDiv 
   Caption         =   "Inclusão de Divisões"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3285
   Icon            =   "frmIncDiv.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   3285
   WindowState     =   2  'Maximized
   Begin VB.Frame frControle 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   2280
         Picture         =   "frmIncDiv.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   120
         Picture         =   "frmIncDiv.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtbxNome 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox cmbxGenero 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome da Divisão:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label lblGen 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   2805
         TabIndex        =   3
         Tag             =   "!"
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Gênero:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmIncDiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxGenero_Click()
    lblGen.caption = F_BuscaUmCampo("Cod_Gen", "tab_Generos", "nome_gen='" & cmbxGenero.Text & "'")
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If F_Vazio(txtbxNome) Then
        MsgBox "Informe o nome da Divisão", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    If cmbxGenero.Text = "" Then
        MsgBox "Informe o Gênero!", vbExclamation
        cmbxGenero.SetFocus
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("Cod_Div", "Tab_Divisoes", "Nome_Div='" & txtbxNome.Text & "'")
    If Me.Tag = "" Then
        If strSQL <> "" Then
            MsgBox "Já existe esta Divisão!", vbExclamation
            Exit Sub
        End If
        P_Salva
    Else
        
        If strSQL <> "" Then
            If strSQL <> Me.Tag Then
                MsgBox "Já existe esta Divisão!", vbExclamation
                Exit Sub
            End If
        End If
        P_Altera
    End If
    
End Sub
Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_erro
    sCod = F_BuscaUmCampo(" max(cod_div)", "tab_divisoes", "")
    If sCod = "" Then
        sCod = "001"
    Else
        
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    strSQL = "Insert into tab_divisoes(cod_div,nome_div,cod_gen)values("
    strSQL = strSQL & "'" & sCod & "', "
    strSQL = strSQL & "'" & txtbxNome.Text & "', "
    strSQL = strSQL & "'" & lblGen.caption & "')"
    banco.Execute strSQL
    
    MsgBox "Divisão Incluída com Sucesso.", vbInformation
    txtbxNome.Text = ""
    txtbxNome.SetFocus
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Altera()
    Dim strSQL As String
    
    On Error GoTo trata_erro
    
    strSQL = "update tab_divisoes set "
    
    strSQL = strSQL & " nome_div='" & txtbxNome.Text & "'  "
    
    strSQL = strSQL & " where cod_div='" & Me.Tag & "' "
    banco.Execute strSQL
    
    MsgBox "Divisão Alterada com Sucesso.", vbInformation
    Unload Me
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    Me.Height = 2550
    Me.Width = 3345
    P_CentraForm Me
    P_PreencheGen cmbxGenero
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub

