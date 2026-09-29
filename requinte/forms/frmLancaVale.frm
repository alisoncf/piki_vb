VERSION 5.00
Begin VB.Form frmLancaVale 
   Caption         =   "Lançamento de Vale"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtValorTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F79971&
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Text            =   "0,00"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtParcelas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Text            =   "1"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Gravar"
      Height          =   735
      Left            =   1800
      Picture         =   "frmLancaVale.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   735
      Left            =   3600
      Picture         =   "frmLancaVale.frx":015C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txtHistorico 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2520
      Width           =   5055
   End
   Begin VB.TextBox txtDataRef 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtFuncionario 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Text            =   "0,00"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdLocaliza 
      Height          =   315
      Left            =   5040
      Picture         =   "frmLancaVale.frx":02C7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Parcelas"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblMes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mês"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   11
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Histórico:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   660
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Referência"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   780
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Funcionário:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   870
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Parcela:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   990
   End
End
Attribute VB_Name = "frmLancaVale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLocaliza_Click()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectFuncionario(txtFuncionario.Text)
    If f.cod_fun <> "" Then
        txtFuncionario.Text = f.nome_fun
        txtFuncionario.tag = f.cod_fun
        If f.vale_fun = "N" Then
            MsgBox "Vale bloqueado", vbExclamation
            txtFuncionario.tag = ""
        End If
        txtValorTotal.SetFocus
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim i As Integer
    Dim strSQL As String
    Dim isSaidaCaixa As Boolean
    If txtFuncionario.tag = "" Then
        MsgBox "Informe o funcionário!", vbExclamation
        Exit Sub
    End If
        If Not F_senha Then
        Exit Sub
    End If

    If MsgBox("Este vale resulta em Saída do caixa?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        isSaidaCaixa = True
    End If

    For i = 1 To CCur(txtParcelas.Text)
        If gravaVale(txtFuncionario.tag, txtHistorico.Text, txtValor.Text, Date, _
            DateAdd("M", i - 1, txtDataRef.Text), strCod_Usu, i & "/" & txtParcelas.Text) Then
            
        Else
        
            GoTo trata_Erro
        End If
    Next
    
    
    
    strSQL = F_GravaCai("D", Date, txtValorTotal.Text, "VALE - " & txtFuncionario.Text, "", "VAL", "", txtValorTotal.Text, txtValorTotal.Text, strCod_Cai, strCod_Tur, "000", "S")
    
    If isSaidaCaixa Then
        strSQL = F_GravaCai("D", Date, txtValorTotal.Text, "SAÍDA VALE - " & txtFuncionario.Text, "", "MOV", "", "0", txtValorTotal.Text, strCod_Cai, strCod_Tur, "000", "S")
    End If
    
    
    
    MsgBox "Vale gravado com sucesso!", vbInformation
    txtFuncionario.Text = ""
    txtFuncionario.tag = ""
    txtFuncionario.SetFocus
Exit Sub
trata_Erro:
    MsgBox "Erro!", vbExclamation
    
End Sub

Private Sub Form_Load()
    txtDataRef.Text = F_PrimDiaMes(Date)
    
    
    
End Sub

Private Sub txtDataRef_Change()
    On Error Resume Next
    lblMes.caption = MonthName(Month(txtDataRef.Text))
End Sub

Private Sub txtDataRef_GotFocus()
    P_SelecTexto txtDataRef
End Sub

Private Sub txtDataRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDataRef_LostFocus()
    P_AjustaData txtDataRef
    'txtDataRef = F_PrimDiaMes(txtDataRef)
End Sub

Private Sub txtFuncionario_GotFocus()
    P_SelecTexto txtFuncionario
End Sub

Private Sub txtFuncionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdLocaliza_Click
    End If
    
End Sub

Private Sub txtHistorico_GotFocus()
    P_SelecTexto txtHistorico
End Sub

Private Sub txtHistorico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtHistorico_LostFocus()
    P_AjustaTexto txtHistorico
End Sub

Private Sub txtParcelas_GotFocus()
    P_SelecTexto txtParcelas
End Sub

Private Sub txtParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtParcelas_LostFocus()
    P_ValorLostFocus txtParcelas, "###"
    If CCur(txtParcelas.Text) <= 0 Then
        txtParcelas.Text = "1"
    End If
    CalcTotal
    
End Sub

Private Sub txtValor_GotFocus()
    P_SelecTexto txtValor
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtValor_LostFocus()
    P_ValorLostFocus txtValor
    CalcTotal True
End Sub

Private Sub txtValorTotal_GotFocus()
    P_SelecTexto txtValorTotal
End Sub

Private Sub txtValorTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtValorTotal_LostFocus()
    P_ValorLostFocus txtValorTotal
    CalcTotal
End Sub
Sub CalcTotal(Optional parc As Boolean = False)
    On Error Resume Next
    If Not parc Then
        txtValor.Text = FormatVal(CCur(txtValorTotal.Text) / CCur(txtParcelas.Text))
    Else
        txtValorTotal.Text = FormatVal(CCur(txtValor.Text) * CCur(txtParcelas.Text))
    End If
End Sub
