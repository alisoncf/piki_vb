VERSION 5.00
Begin VB.Form frmCadPlano 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Receita/Despesa"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmPlanoContas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6690
   Begin VB.PictureBox pcTipoPLa 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
      Begin VB.OptionButton optnCred 
         Caption         =   "&Crédito"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optnDeb 
         Caption         =   "&Débito"
         Height          =   195
         Left            =   1170
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame frPri 
      BackColor       =   &H80000004&
      Height          =   975
      Left            =   3240
      TabIndex        =   7
      Top             =   2880
      Width           =   3015
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   1560
         Picture         =   "frmPlanoContas.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   160
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Gravar"
         Height          =   735
         Left            =   120
         Picture         =   "frmPlanoContas.frx":06D8
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   160
         Width           =   1335
      End
   End
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Código da Conta:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   4080
      TabIndex        =   10
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome da Conta:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1155
   End
End
Attribute VB_Name = "frmCadPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()

End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If F_Vazio(txtbxNOme.Text) Then
        MsgBox "Informe o nome da Conta que deseja incluir!", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "" Then
        If F_Vazio(txtCodigo.Text) Then
            MsgBox "Informe o código da Conta que deseja incluir!", vbExclamation
            Exit Sub
        End If
        If F_BuscaUmCampo("Cod_Pla", TAB_PLANOCONTAS, "Nome_Pla='" & _
            txtbxNOme.Text & "'") <> "" Then
            MsgBox "Já existe uma conta com este Nome!", vbExclamation
            Exit Sub
        End If
        
        grava
        txtbxNOme.Text = ""
        txtbxNOme.SetFocus
    Else
        If F_BuscaUmCampo("Cod_Pla", TAB_PLANOCONTAS, "Nome_Pla='" & _
            txtbxNOme.Text & "' and cod_pla<>'" & Me.tag & "'") <> "" Then
            MsgBox "Você está tentanto mudar o nome desta conta para um que já existe!", vbExclamation
            Exit Sub
        End If
        altera
    End If
End Sub
Sub grava()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into " & TAB_PLANOCONTAS & "(cod_pla,nome_pla,tipo_pla)Values("
    strSQL = strSQL & "'" & txtCodigo.Text & "',"
    strSQL = strSQL & "'" & txtbxNOme.Text & "',"
    strSQL = strSQL & "'" & Left(lblTipo.caption, 1) & "')"
    conn.Execute strSQL
    
    MsgBox "Conta incluída com sucesso!", vbInformation
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub altera()
    Dim strSQL As String
    strSQL = "Update " & TAB_PLANOCONTAS & " set "
    
    strSQL = strSQL & "Nome_Pla='" & txtbxNOme.Text & "' "
    strSQL = strSQL & " where cod_pla='" & Me.tag & "'     "
    conn.Execute strSQL
    MsgBox "Conta alterada!", vbInformation
    
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        pcTipoPLa.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
End Sub

Private Sub optnCred_Click()
    lblTipo.caption = "Crédito"
    lblTipo.ForeColor = vbBlue
End Sub
Sub carrega(codigo As String)
    Dim strSQL As String
    Dim rPesq As ADODB.Recordset
    strSQL = "select * from " & TAB_PLANOCONTAS & " where cod_pla='" & codigo & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            txtCodigo.Text = !cod_pla
            txtbxNOme.Text = !nome_pla
            If !tipo_pla = "C" Then
                optnCred.Value = True
            Else
                optnCred.Value = False
            End If
            
        End If
        
    End With
    
End Sub
Private Sub optnDeb_Click()
    lblTipo.caption = "Débito"
    lblTipo.ForeColor = vbRed
End Sub

Private Sub txtbxNome_GotFocus()
    ganhaFoco txtbxNOme
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    AjustaTexto txtbxNOme
End Sub

Private Sub txtcodigo_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        carrega txtCodigo.Text
    End If
End Sub
