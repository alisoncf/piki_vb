VERSION 5.00
Begin VB.Form frmLanCaixa 
   Caption         =   "Lançamento em Caixa"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   8910
   Icon            =   "frmLanCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtbxValor_Cai 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "0,00"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtbxHist 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Caixa"
      Height          =   735
      Left            =   4080
      TabIndex        =   12
      Top             =   120
      Width           =   4215
      Begin VB.Label lblCaixa 
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome da Conta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Tag             =   "!"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Turno:"
         Height          =   195
         Left            =   2880
         TabIndex        =   16
         Top             =   0
         Width           =   465
      End
      Begin VB.Label lblTur 
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   2880
         TabIndex        =   15
         Tag             =   "!"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtbxClass 
      Appearance      =   0  'Flat
      BackColor       =   &H0005C9EB&
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Tag             =   "000"
      Top             =   3240
      Width           =   5895
   End
   Begin VB.TextBox txtbxData_Cai 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo do Movimento:"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton optnDeb 
         Height          =   615
         Left            =   1320
         Picture         =   "frmLanCaixa.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optnCerd 
         Height          =   615
         Left            =   120
         Picture         =   "frmLanCaixa.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblEsp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MOV"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2640
         TabIndex        =   19
         Tag             =   "mov"
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lblTipo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Débito"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2640
         TabIndex        =   18
         Tag             =   "D"
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   8880
      TabIndex        =   20
      Top             =   4230
      Width           =   8910
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   7200
         Picture         =   "frmLanCaixa.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   5640
         Picture         =   "frmLanCaixa.frx":0CF9
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmndbtnPlanoContas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6120
      Picture         =   "frmLanCaixa.frx":11A1
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Valor:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Histórico:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   660
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Classificação:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   390
   End
End
Attribute VB_Name = "frmLanCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxConta_Click()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select nome_cai,cod_tur from tab_caixas where cod_cai='" & Left(cmbxConta.text, 2) & "'"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblNome_Con.caption = !nome_cai
            lblTur.caption = !cod_tur
            
        Else
            lblNome_Con.caption = ""
            lblTur.caption = ""
        End If
        .Close
    End With
    
    
    
    
End Sub

Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

'Private Sub cmndbtnFunc_Click()
'    Dim f As funcionario.funcionario
'    f = F_SelectFuncionario(txtbxFunc.Text)
'    If f.cod_fun <> "" Then
'        txtbxFunc.Text = f.nome_fun
'        txtbxFunc.tag = f.cod_fun
'    End If
'End Sub

Private Sub cmndbtnPlanoContas_Click()
    Dim p As PlanoContas
    
    p = F_BuscaPlanoContas(Left(lblTipo.caption, 1), txtbxClass.text)
    If p.cod_pla <> "" Then
        txtbxClass.tag = p.cod_pla
        txtbxClass.text = p.nome_pla
        If p.cod_pla = "002" Then
            txtbxFunc.Enabled = True
            cmndbtnFunc.Enabled = True
            txtbxFunc.SetFocus
        Else
           ' txtbxFunc.Enabled = False
           ' cmndbtnFunc.Enabled = False
          '  txtbxFunc.Text = ""
          '  txtbxFunc.tag = ""
            cmndbtnSalvar.SetFocus
        End If
        
    End If

End Sub


Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strCAI As String
    Dim strSQL As String
    'strCAI = F_BuscaUmCampo("COD_CAI", "tab_caixas", "ip_cai='" & strIP_Usu & "'")
    If Not F_IsCaixa Then
        MsgBox "Esta computador não é um caixa, não pode fazer lançamento por tanto!", vbExclamation
        Exit Sub
    End If
    If Not F_senha(True) Then
        Exit Sub
    End If
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para essa operação!", vbExclamation
        Exit Sub
    End If
    
    
    
    
'    If F_BuscaUmCampo("status_tur", "tab_turncai", " cod_cai='" & lblCaixa.caption & "' and cod_tur='" & _
'        lblTur.caption & "' order by data_Tur desc ") = "F" Then
'        MsgBox "Não é possível fazer lançamento neste caixa, ele está fechado!", vbCritical
'        Exit Sub
'    End If
    
    If txtbxValor_Cai.text = "0,00" Then
            MsgBox "Informe um Valor válido!", vbExclamation
            txtbxValor_Cai.SetFocus
            Exit Sub
        End If
        If F_Vazio(txtbxHist.text) Then
            MsgBox "Informe um histórico!", vbExclamation
            txtbxHist.SetFocus
            Exit Sub
        End If
    If Me.tag = "" Then
        
       
        P_Grava
        
    Else
         If UCase(lblEsp.caption) <> "MOV" Then
            MsgBox "Este lançamento é bloqueado para alteração devido à sua espécie! " & lblEsp.caption, vbExclamation
            Exit Sub
        End If
        P_Altera
    End If
End Sub
Sub P_Grava()
    Dim strSQL As String
    strSQL = F_GravaCai(Left(lblTipo.caption, 1), txtbxData_Cai.text, _
                txtbxValor_Cai.text, txtbxHist.text, "", "MOV", _
                strCod_Cai, "0", 0, strCod_Cai, strCod_Tur, txtbxClass.tag, "S")
    If strSQL = "" Then
        MsgBox "Erro!", vbCritical
        Exit Sub
    End If
    
'    If txtbxFunc.tag <> "" Then
'        If Not F_GravaValesCai(strSQL, txtbxFunc.tag) Then
'            MsgBox "Erro na gravação do vale!", vbExclamation
'            Exit Sub
'        End If
'        P_ImprimeValeCai strSQL
'    End If
    
    'P_Class CCur(txtbxValor_Cai.Text), Left(lblTipo.caption, 1), strSQL
    MsgBox "Lançamento Gravado!", vbInformation
    
    
    
    txtbxData_Cai.text = Date
    txtbxHist.text = ""
    txtbxValor_Cai.text = "0,00"
    txtbxHist.SetFocus
End Sub

Sub P_Altera()
    Dim strSQL As String
    strSQL = "Update tab_caixa set historico_Cai='" & txtbxHist.text & "',"
    strSQL = strSQL & "cod_pla='" & txtbxClass.tag & "', "
    strSQL = strSQL & "valor_Cai=" & F_GravaDec(txtbxValor_Cai.text)
    strSQL = strSQL & ", COD_cON='" & cmbxConta.text & "' "
    strSQL = strSQL & " where seq_cai='" & Me.tag & "' "
    Conn.Execute strSQL
    MsgBox "Dados Alterados!", vbInformation
    
    Unload Me
    
End Sub


Private Sub Form_Activate()
    If bActive Then Exit Sub
    bActive = True
    
    lblTur.caption = strCod_Tur
    lblCaixa.caption = strCod_Cai
    
    carrega
End Sub

Private Sub Form_Load()
    On Error Resume Next
    

    
    
    txtbxData_Cai.text = Date
    
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If Me.tag = "" Then Exit Sub
    strSQL = F_SelectCai(Me.tag)
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then Exit Sub
        
        txtbxClass.text = !nome_pla
        txtbxData_Cai.text = !data_cai
        txtbxHist.text = !historico_cai
        txtbxValor_Cai.text = FormatVal(!valor_Cai)
        lblEsp.caption = !esp_cai
        
        .Close
    End With
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Unload Me
End Sub

Private Sub optnCerd_Click()
    lblTipo.caption = "Crédito"
End Sub

Private Sub optnCerd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnDeb_Click()
        lblTipo.caption = "Débito"
End Sub

Private Sub optnDeb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxClass_GotFocus()
    P_SelecTexto txtbxClass
End Sub

Private Sub txtbxClass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnPlanoContas_Click
    End If
End Sub

Private Sub txtbxFunc_Change()
    txtbxFunc.tag = ""
End Sub

Private Sub txtbxFunc_GotFocus()
    End Sub

Private Sub txtbxFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
    End If
End Sub

Private Sub txtbxHist_GotFocus()
    P_SelecTexto txtbxHist
End Sub

Private Sub txtbxHist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxHist_LostFocus()
    P_AjustaTexto txtbxHist
End Sub

Private Sub txtbxValor_Cai_GotFocus()
    P_SelecTexto txtbxValor_Cai
End Sub

Private Sub txtbxValor_Cai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Cai_LostFocus()
    P_ValorLostFocus txtbxValor_Cai
End Sub
