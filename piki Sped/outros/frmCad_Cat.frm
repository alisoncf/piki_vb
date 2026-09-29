VERSION 5.00
Begin VB.Form frmCadCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de Categorias"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   Icon            =   "frmCad_Cat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3225
   Begin VB.Frame frControle 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.TextBox txtbxNome 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   240
         Picture         =   "frmCad_Cat.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   1680
         Picture         =   "frmCad_Cat.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome da Categoria:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmCadCat"
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
        MsgBox "Informe o nome da Categoria", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("Cod_Cat", "Tab_categorias", "Nome_Cat='" & txtbxNome.Text & "'")
    If Me.Tag = "" Then
        If strSQL <> "" Then
            MsgBox "Já existe esta Categoria!", vbExclamation
            Exit Sub
        End If
        P_Salva
    Else
        
        If strSQL <> "" Then
            If strSQL <> Me.Tag Then
                MsgBox "Já existe esta Categoria!", vbExclamation
                Exit Sub
            End If
        End If
        P_Altera
    End If
    If IsLoad("frmEdiCat") Then
        frmEdiCat.dtProd.Refresh
    End If
    If IsLoad("frmCadPro") Then
        P_PreencheCateg frmCadPro.cmbxCat
    End If
End Sub
Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_erro
    sCod = F_BuscaUmCampo(" max(cod_cat)", "tab_categorias", "")
    If sCod = "" Then
        sCod = "001"
    Else
        
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    strSQL = "Insert into tab_categorias(cod_cat,nome_cat)values("
    strSQL = strSQL & "'" & sCod & "', "
    strSQL = strSQL & "'" & txtbxNome.Text & "')  "
    Banco.Execute strSQL
    
    MsgBox "Categoria Incluída com Sucesso.", vbInformation
    txtbxNome.Text = ""
    txtbxNome.SetFocus
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Altera()
    Dim strSQL As String
    
    On Error GoTo trata_erro
    
    strSQL = "update tab_categorias set "
    
    strSQL = strSQL & " nome_cat='" & txtbxNome.Text & "'  "
    strSQL = strSQL & " where cod_cat='" & Me.Tag & "' "
    Banco.Execute strSQL
    
    MsgBox "Categoria Alterada com Sucesso.", vbInformation
    Unload Me
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    
End Sub

Private Sub Form_Load()
    Me.Height = 2070
    Me.Width = 3195
    P_CentraForm Me
    MudaAparencia Me
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
