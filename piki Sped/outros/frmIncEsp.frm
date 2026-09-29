VERSION 5.00
Begin VB.Form frmIncEsp 
   Caption         =   "Inclusão de Espécies:"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3150
   Icon            =   "frmIncEsp.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1905
   ScaleWidth      =   3150
   WindowState     =   2  'Maximized
   Begin VB.Frame frControle 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   2040
         Picture         =   "frmIncEsp.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   1080
         Picture         =   "frmIncEsp.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtbxNome 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome da Espécie:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmIncEsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If F_Vazio(txtbxNome) Then
        MsgBox "Informe o nome da Espécie", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("Cod_Esp", "Tab_Especies", "Nome_Esp='" & txtbxNome.Text & "'")
    If Me.Tag = "" Then
        If strSQL <> "" Then
            MsgBox "Já existe esta Espécie!", vbExclamation
            Exit Sub
        End If
        P_Salva
    Else
        
        If strSQL <> "" Then
            If strSQL <> Me.Tag Then
                MsgBox "Já existe esta Espécie!", vbExclamation
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
    sCod = F_BuscaUmCampo(" max(cod_Esp)", "tab_Especies", "")
    If sCod = "" Then
        sCod = "001"
    Else
        
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    strSQL = "Insert into tab_Especies(cod_Esp,nome_Esp)values("
    strSQL = strSQL & "'" & sCod & "', "
    strSQL = strSQL & "'" & txtbxNome.Text & "')  "
    banco.Execute strSQL
    
    MsgBox "Espécie Incluída com Sucesso.", vbInformation
    txtbxNome.Text = ""
    txtbxNome.SetFocus
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Altera()
    Dim strSQL As String
    
    On Error GoTo trata_erro
    
    strSQL = "update tab_Especies set "
    
    strSQL = strSQL & " nome_Esp='" & txtbxNome.Text & "'  "
    strSQL = strSQL & " where cod_Esp='" & Me.Tag & "' "
    banco.Execute strSQL
    
    MsgBox "Espécie Alterado com Sucesso.", vbInformation
    Unload Me
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    Me.Height = 2100
    Me.Width = 3195
    P_CentraForm Me
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

