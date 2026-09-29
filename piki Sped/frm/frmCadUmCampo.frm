VERSION 5.00
Begin VB.Form frmCadUmCampo 
   Caption         =   "Cadastro de Departamentos"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   6225
   Begin VB.PictureBox pcStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00B36A09&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   6195
      TabIndex        =   4
      Top             =   3420
      Width           =   6225
      Begin VB.Label lblSatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3045
         TabIndex        =   5
         Tag             =   "!"
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Gravar"
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   465
   End
End
Attribute VB_Name = "frmCadUmCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sTab As String
Dim sCod As String
Dim sCampo As String
Dim sNome As String
Dim sMask As String
Public operacao As operacao_um_campo
Private Sub Option1_Click()

End Sub




Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    
    If F_Vazio(txtbxNome.Text) Then
        MsgBox "Informe o nome que deseja incluir!", vbExclamation
        Exit Sub
    End If
    If Me.tag = "" Then
        If operacao = secao_ Then
            If secao.F_GravaSec(txtbxNome.Text) <> "" Then
                MsgBox "Dados Gravados com sucesso!", vbInformation
                txtbxNome.Text = ""
                txtbxNome.SetFocus
            End If
        End If
        
        If operacao = exp_mov_int Then
            If MovInt.F_GravaEspMovInt(txtbxNome.Text) <> "" Then
                MsgBox "Dados Gravados com sucesso!", vbInformation
                txtbxNome.Text = ""
                txtbxNome.SetFocus
            End If
        End If
        If operacao = grupo_cliente Then
            If Clientes.F_GravaGrupoCliente(txtbxNome.Text) Then
                MsgBox "Dados Gravados com sucesso!", vbInformation
                txtbxNome.Text = ""
                txtbxNome.SetFocus
            End If
        End If
        
        
        
    Else
        If operacao = "1" Then
            If secao.F_AlteraSec(Me.tag, txtbxNome.Text) Then
                MsgBox "Dados Alterados com sucesso!", vbInformation
                Unload Me
            End If
        End If
        If operacao = "5" Then
            If MovInt.F_AlteraEspMovInt(Me.tag, txtbxNome.Text) = True Then
                MsgBox "Dados Gravados com sucesso!", vbInformation
                txtbxNome.Text = ""
                txtbxNome.SetFocus
            End If
        End If
    End If
    
    
End Sub
Sub Grava()
    Dim strSQL As String
    Dim sNovoCod As String
    sNovoCod = MaxKeySQL(sTab, sCod, "", sMask)
    strSQL = "insert into " & sTab & "(" & sCod & "," & sCampo & ")Values("
    
    strSQL = strSQL & "'" & sNovoCod & "',"
    strSQL = strSQL & "'" & txtbxNome.Text & "')"
    
    conn.Execute strSQL
    MsgBox "Registro incluído com sucesso!", vbInformation
    
End Sub
Sub Altera()
    Dim strSQL As String
    strSQL = "Update " & sTab & " set "
    
    strSQL = strSQL & sCampo & "='" & txtbxNome.Text & "' "
    strSQL = strSQL & " where " & sCod & "='" & Me.tag & "'     "
    conn.Execute strSQL
    MsgBox "Registro alterado com sucesso!", vbInformation
    
    Unload Me
End Sub

Private Sub Form_Activate()
    
    If operacao = secao_ Then
        sNome = "Seções"
        sTab = "tab_secoes"
        sCod = "cod_sec"
        sCampo = "Nome_sec"
        sMask = "00#"
    ElseIf operacao = grupo_cliente Then
        sNome = "Princípios"
        sTab = "tab_Principios"
        sCod = "cod_pri"
        sCampo = "Nome_pri"
        sMask = "00#"
    ElseIf operacao = "3" Then
        sNome = "Categorias"
        sTab = "tab_categorias"
        sCod = "cod_cat"
        sCampo = "Nome_cat"
        sMask = "00#"
    ElseIf operacao = "4" Then
        sNome = "Tipos de Serviço"
        sTab = "tab_tipos"
        sCod = "cod_tip"
        sCampo = "Nome_tip"
        sMask = "00#"
    ElseIf operacao = exp_mov_int Then
        sNome = "Destinos de Movimentação Interna"
        sTab = "tab_destmovint"
        sCod = "cod_des"
        sCampo = "Nome_tip"
        sMask = "#"
    End If
    
    
    If Me.tag = "" Then
        Me.caption = "Cadastro de " & sNome
    Else
        Me.caption = Me.tag & "-" & txtbxNome
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
End Sub

Private Sub optnCred_Click()
    lblTipo.caption = "Crédito"
    lblTipo.ForeColor = vbBlue
End Sub

Private Sub optnDeb_Click()
    lblTipo.caption = "Débito"
    lblTipo.ForeColor = vbRed
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
