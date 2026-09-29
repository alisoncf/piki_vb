VERSION 5.00
Begin VB.Form frmLanFin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lançamento Financeiro"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6855
   Icon            =   "frmLanFin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCco2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CommandButton cmdCco2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6120
      Picture         =   "frmLanFin.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   2895
      TabIndex        =   24
      Top             =   360
      Width           =   2895
      Begin VB.OptionButton optnLan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lançamento"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optTransf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Transferência"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.TextBox txtbxClass 
      Appearance      =   0  'Flat
      BackColor       =   &H0005C9EB&
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   3480
      Width           =   5895
   End
   Begin VB.CommandButton cmndbtnPlanoContas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6120
      Picture         =   "frmLanFin.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox pcTipoOperacao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   3135
      TabIndex        =   1
      Top             =   360
      Width           =   3135
      Begin VB.OptionButton optnDeb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Débito"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optnCred 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Crédito"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox txtCCo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   5895
   End
   Begin VB.CommandButton cmdCco1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6120
      Picture         =   "frmLanFin.frx":0A1E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtbxData_Cai 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   3360
      Picture         =   "frmLanFin.frx":0DA8
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   4920
      Picture         =   "frmLanFin.frx":1250
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtbxHist 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   4320
      Width           =   6135
   End
   Begin VB.TextBox txtbxValor_Cai 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "0,00"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblCco2 
      AutoSize        =   -1  'True
      Caption         =   "Conta Corrente Creditada:"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Movimento"
      Height          =   195
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Natureza do movimento"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label lblCCo1 
      AutoSize        =   -1  'True
      Caption         =   "Conta Corrente Debitada:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data:"
      Height          =   195
      Left            =   5160
      TabIndex        =   9
      Top             =   1200
      Width           =   390
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Classificação:"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Histórico:"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   4080
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Valor:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   405
   End
End
Attribute VB_Name = "frmLanFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmbxConta_Click()
    
'    lblNome_Con.caption = F_BuscaUmCampo("NOme_con", "tab_contas", "cod_con='" & cmbxConta.Text & "'")
    If cmbxConta.Text <> constContaCaixa And optnDeb.Value = True Then
        Dim ch As chePag
        'ch = F_DadosChePag(txtbxData_Cai.Text, lblNum_Che.caption, _
                cmbxConta.Text, txtbxValor_Cai.Text, True)
        'If ch.num_che <> "" Then
        '    frch.Visible = True
        '    lblNum_Che.caption = ch.num_che
        '    lblCod_Con.caption = ch.cod_con
        'Else
        '    frch.Visible = False
        'End If
    End If
End Sub

Private Sub cmdCco1_Click()
    Dim c As ContaCorrente
    c = financeiro.F_BuscaContaCorrente(txtCCo1.Text)
    If c.cod_con <> "" Then
        txtCCo1.Text = c.nome_con
        txtCCo1.tag = c.cod_con
        
    End If
End Sub

Private Sub cmdCco2_Click()
    Dim c As ContaCorrente
    c = financeiro.F_BuscaContaCorrente(txtCco2.Text)
    If c.cod_con <> "" Then
        txtCco2.Text = c.nome_con
        txtCco2.tag = c.cod_con
        
    End If
End Sub

Private Sub cmndbtnPlanoContas_Click()
    Dim p As PlanoContas
    p = F_BuscaPlanoContas(IIf(optnCred.Value = True, "C", "D"), "")
    If p.cod_pla <> "" Then
        txtbxClass.Text = p.nome_pla
        txtbxClass.tag = p.cod_pla
        
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If strLancarFin_Usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    If txtCCo1.Text = "" Then
            MsgBox "Informe a conta!", vbExclamation
            txtCCo1.SetFocus
            Exit Sub
        End If
        If txtbxClass.tag = "" Then
            MsgBox "Escolha uma conta do plano de contas!", vbExclamation
            txtbxClass.SetFocus
            Exit Sub
        End If
        
    If optTransf.Value And txtCco2.Text = "" Then
            MsgBox "Informe a conta!", vbExclamation
            txtCco2.SetFocus
            Exit Sub
        End If
       
        
        If txtbxValor_Cai.Text = "0,00" Then
            MsgBox "Informe um Valor válido!", vbExclamation
            txtbxValor_Cai.SetFocus
            Exit Sub
        End If
        If F_Vazio(txtbxHist.Text) Then
            MsgBox "Informe um histórico!", vbExclamation
            txtbxHist.SetFocus
            Exit Sub
        End If
    If Me.tag = "" Then
    
        
        
        
        
        P_Grava
        
    Else
        P_Altera
    End If
    
End Sub
Sub P_Grava()
    Dim strSQL As String
    Dim tipo As String
    Dim tipoR As String
    Dim s1 As String
    Dim s2 As String
    If optnCred.Value Then
        tipo = "C"
        tipoR = "D"
    Else
        tipo = "D"
        tipoR = "C"
    End If

    
    
    
    
    If optTransf.Value Then
        s1 = F_GravaFinanceiro(tipo, txtbxData_Cai.Text, _
                txtbxValor_Cai.Text, "Transf: " & txtbxHist.Text, txtCCo1.tag, _
                 "N", txtbxClass.tag)
    
        s2 = F_GravaFinanceiro(tipoR, txtbxData_Cai.Text, _
                txtbxValor_Cai.Text, "Transf: " & txtbxHist.Text, txtCco2.tag, _
                 "N", txtbxClass.tag, , s1)
                 
        strSQL = "update " & TAB_FINANCEIRO & " set seq_tra='" & s2 & "' "
        strSQL = strSQL & " where seq_fin='" & s1 & "'"
        conn.Execute strSQL
        strSQL = s2
    Else
        strSQL = F_GravaFinanceiro(tipo, txtbxData_Cai.Text, _
                txtbxValor_Cai.Text, txtbxHist.Text, txtCCo1.tag, _
                 "N", txtbxClass.tag)
    
    
    End If
    If strSQL = "" Then
        MsgBox "Erro!", vbCritical
        Exit Sub
    End If
    
    
    
    MsgBox "Lançamento Gravado!", vbInformation
    txtbxData_Cai.Text = Date
    txtbxHist.Text = ""
    txtbxValor_Cai.Text = "0,00"
    txtbxHist.SetFocus
End Sub

Sub P_Altera()
    Dim strSQL As String
    strSQL = "Update " & TAB_FINANCEIRO & "  set historico_FIN='" & txtbxHist.Text & "',"
    strSQL = strSQL & "valor_FIN=" & F_GravaDec(txtbxValor_Cai.Text)
    strSQL = strSQL & ", COD_cON='" & txtCCo1.tag & "' "
    strSQL = strSQL & " where seq_FIN='" & Me.tag & "' "
    conn.Execute strSQL
    
   
    
    
    MsgBox "Dados Alterados!", vbInformation
     'P_Class CCur(txtbxValor_Cai.Text), Left(lblTipo.caption, 1), Me.Tag
    
    
    
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
    
    End If
    bActive = True
End Sub


Private Sub Form_Load()
    bActive = False
    P_CentraForm Me

    'cmbxConta.ListIndex = 0
    txtbxData_Cai.Text = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsLoad("frmEdiCai") Then
        
    End If
    Unload Me
End Sub



Private Sub optnCerd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnCred_Click()
    lblCCo1.caption = "Conta Corrente Creditada"
    lblCco2.caption = "Conta Corrente Debitada"
    txtbxClass.Text = ""
End Sub

Private Sub optnDeb_Click()
        lblCco2.caption = "Conta Corrente Creditada"
    lblCCo1.caption = "Conta Corrente Debitada"
    txtbxClass.Text = ""
    'txtbxClass.tag = "000"
End Sub

Private Sub optnDeb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnLan_Click()
    lblCco2.Visible = False
    txtCco2.Visible = False
    cmdCco2.Visible = False
End Sub

Private Sub optTransf_Click()
    lblCco2.Visible = True
    txtCco2.Visible = True
    cmdCco2.Visible = True
End Sub

Private Sub txtbxClass_Change()
    txtbxClass.tag = ""
End Sub

Private Sub txtbxClass_GotFocus()
    P_SelecTexto txtbxClass
End Sub

Private Sub txtbxClass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnPlanoContas_Click
    End If
End Sub

Private Sub txtbxData_Cai_GotFocus()
    P_SelecTexto txtbxData_Cai
End Sub

Private Sub txtbxData_Cai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxData_Cai_LostFocus()
    P_AjustaData txtbxData_Cai
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

Private Sub txtCCo1_GotFocus()
    P_SelecTexto txtCCo1
End Sub
Private Sub txtCCo2_GotFocus()
    P_SelecTexto txtCco2
End Sub
Private Sub txtCco1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCco1_Click
    End If
End Sub

Private Sub txtCco2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCco2_Click
    End If
End Sub
