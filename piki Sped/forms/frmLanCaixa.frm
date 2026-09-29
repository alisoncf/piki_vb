VERSION 5.00
Begin VB.Form frmLanCaixa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lançamento em Caixa"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5805
   Icon            =   "frmLanCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtbxFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H00CF614E&
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   3480
         Width           =   4815
      End
      Begin VB.TextBox txtbxClass 
         Appearance      =   0  'Flat
         BackColor       =   &H0005C9EB&
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Tag             =   "000"
         Top             =   2880
         Width           =   4815
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   2520
         Picture         =   "frmLanCaixa.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   4080
         Picture         =   "frmLanCaixa.frx":07B2
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   5295
         Begin VB.TextBox txtbxHist 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   5055
         End
         Begin VB.TextBox txtbxValor_Cai 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Text            =   "0,00"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Histórico:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Valor:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   405
         End
      End
      Begin VB.Frame frch 
         Caption         =   "Cheque"
         Height          =   855
         Left            =   120
         TabIndex        =   19
         Top             =   4920
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conta:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   600
         End
         Begin VB.Label lblCod_Con 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "001"
            ForeColor       =   &H00800080&
            Height          =   195
            Left            =   720
            TabIndex        =   21
            Tag             =   "!"
            Top             =   240
            Width           =   270
         End
         Begin VB.Label lblNum_Che 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00800080&
            Height          =   195
            Left            =   720
            TabIndex        =   20
            Tag             =   "!"
            Top             =   480
            Width           =   45
         End
      End
      Begin VB.TextBox txtbxData_Cai 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo do Movimento:"
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2655
         Begin VB.OptionButton optnCerd 
            Height          =   615
            Left            =   120
            Picture         =   "frmLanCaixa.frx":091D
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optnDeb 
            Height          =   615
            Left            =   1320
            Picture         =   "frmLanCaixa.frx":0D5F
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "&Caixa"
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   4080
         Width           =   5175
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
            Left            =   4080
            TabIndex        =   26
            Tag             =   "!"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Turno:"
            Height          =   195
            Left            =   4080
            TabIndex        =   25
            Top             =   0
            Width           =   465
         End
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
            TabIndex        =   18
            Tag             =   "!"
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.CommandButton cmndbtnPlanoContas 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         Picture         =   "frmLanCaixa.frx":11A1
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnFunc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         Picture         =   "frmLanCaixa.frx":152B
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Funcionário:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   3240
         Width           =   870
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Classificação:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data:"
         Height          =   195
         Left            =   3840
         TabIndex        =   4
         Top             =   120
         Width           =   390
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   1800
         TabIndex        =   24
         Tag             =   "!"
         Top             =   840
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmLanCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxConta_Click()
    Dim rPesq As New ADODB.Recordset
    strSQL = "select nome_cai,cod_tur from tab_caixas where cod_cai='" & Left(cmbxConta.Text, 2) & "'"
    rPesq.Open strSQL, conn
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnFunc_Click()
    Dim f As funcionario.funcionario
    f = F_SelectFuncionario(txtbxFunc.Text)
    If f.cod_fun <> "" Then
        txtbxFunc.Text = f.nome_fun
        txtbxFunc.tag = f.cod_fun
    End If
End Sub

Private Sub cmndbtnPlanoContas_Click()
    Dim p As PlanoContas
    
    p = F_BuscaPlanoContas(Left(lblTipo.caption, 1), txtbxClass.Text)
    If p.cod_pla <> "" Then
        txtbxClass.tag = p.cod_pla
        txtbxClass.Text = p.nome_pla
        If p.cod_pla = "002" Then
            txtbxFunc.Enabled = True
            cmndbtnFunc.Enabled = True
            txtbxFunc.SetFocus
        Else
            txtbxFunc.Enabled = False
            cmndbtnFunc.Enabled = False
            txtbxFunc.Text = ""
            txtbxFunc.tag = ""
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
    
    
    
    
'    If F_BuscaUmCampo("status_tur", "tab_turncai", " cod_cai='" & lblCaixa.caption & "' and cod_tur='" & _
'        lblTur.caption & "' order by data_Tur desc ") = "F" Then
'        MsgBox "Não é possível fazer lançamento neste caixa, ele está fechado!", vbCritical
'        Exit Sub
'    End If
    
    
    If Me.tag = "" Then
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
        
        P_Grava
        
    Else
        P_Altera
    End If
End Sub
Sub P_Grava()
    Dim strSQL As String
    strSQL = F_GravaCai(Left(lblTipo.caption, 1), txtbxData_Cai.Text, _
                txtbxValor_Cai.Text, txtbxHist.Text, "", "MOV", _
                strCod_Cai, CCur(txtbxValor_Cai.Text), 0, strCod_Cai, strCod_Tur, txtbxClass.tag, "S")
    If strSQL = "" Then
        MsgBox "Erro!", vbCritical
        Exit Sub
    End If
    
    If txtbxFunc.tag <> "" Then
        If Not F_GravaValesCai(strSQL, txtbxFunc.tag) Then
            MsgBox "Erro na gravação do vale!", vbExclamation
            Exit Sub
        End If
        P_ImprimeValeCai strSQL
    End If
    
    'P_Class CCur(txtbxValor_Cai.Text), Left(lblTipo.caption, 1), strSQL
    MsgBox "Lançamento Gravado!", vbInformation
    
    
    
    txtbxData_Cai.Text = Date
    txtbxHist.Text = ""
    txtbxValor_Cai.Text = "0,00"
    txtbxHist.SetFocus
End Sub

Sub P_Altera()
    Dim strSQL As String
    strSQL = "Update tab_caixa set historico_Cai='" & txtbxHist.Text & "',"
    strSQL = strSQL & "cod_pla='" & txtbxClass.tag & "', "
    strSQL = strSQL & "valor_Cai=" & F_GravaDec(txtbxValor_Cai.Text)
    STRSQAL = strSQL & ", COD_cON='" & cmbxConta.Text & "' "
    strSQL = strSQL & " where seq_cai='" & Me.tag & "' "
    conn.Execute strSQL
    MsgBox "Dados Alterados!", vbInformation
    
    Unload Me
    
End Sub


Private Sub Form_Activate()
    MudaAparencia Me
    
    lblTur.caption = strCod_Tur
    lblCaixa.caption = strCod_Cai
        
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    P_CentraForm Me
    
    
    txtbxData_Cai.Text = Date
    
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Unload Me
End Sub

Private Sub optnCerd_Click()
    lblTipo.caption = "Crédito"
End Sub

Private Sub optnCerd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnDeb_Click()
        lblTipo.caption = "Débito"
End Sub

Private Sub optnDeb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
    P_SelecTexto txtbxFunc
End Sub

Private Sub txtbxFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnFunc_Click
    End If
End Sub

Private Sub txtbxHist_GotFocus()
    P_SelecTexto txtbxHist
End Sub

Private Sub txtbxHist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Cai_LostFocus()
    P_ValorLostFocus txtbxValor_Cai
End Sub
