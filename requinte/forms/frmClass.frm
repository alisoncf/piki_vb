VERSION 5.00
Begin VB.Form frmClass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Classificação"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbxDep 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmClass.frx":014A
      Left            =   2160
      List            =   "frmClass.frx":0154
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.ComboBox cmbxConta 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmClass.frx":0172
      Left            =   2160
      List            =   "frmClass.frx":017C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar "
      Height          =   615
      Left            =   2880
      Picture         =   "frmClass.frx":019A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Cancela&r"
      Height          =   615
      Left            =   4320
      Picture         =   "frmClass.frx":02E4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
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
      Left            =   1080
      TabIndex        =   9
      Top             =   1920
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label lblHistorico 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   1080
      TabIndex        =   7
      Top             =   1680
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Histórico:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   660
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Departamento:"
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   540
      Width           =   1050
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C&onta:"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   1020
      Width           =   465
   End
End
Attribute VB_Name = "frmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public operacao As String
Function F_Restante() As Currency
    'F_Restante = CCur(lblTotal.caption) - CCur(lblDistr.caption)
End Function

Private Sub cmbxConta_Click()
    cmbxConta.tag = F_BuscaUmCampo("cod_pla", "tab_planocontas", "nome_pla='" & cmbxConta.text & "' and tipo_pla='" & cmndbtnConfirma.tag & "'")
End Sub

Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxDep_Click()
    cmbxDep.tag = F_BuscaUmCampo("cod_dep", "tab_departamentos", "nome_dep='" & cmbxDep.text & "'")
End Sub

Private Sub cmbxDep_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnConfirma_Click()
    'ReDim class(msflxgrd.Rows - 1)
    Dim strSQL As String
    Dim iCount As Integer
    Dim c As Currency
    c = F_Restante
    If c > 0 Then
        MsgBox "O Valor não foi completamente distribuído!", vbExclamation
        Exit Sub
    ElseIf c < 0 Then
        MsgBox "O Valor distribuído excede o total!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "delete from tab_class where seq_cla ='" & Me.tag & "' "
    Conn.Execute strSQL
    strSQL = ""
    
    strSQL = strSQL & "insert into tab_class(seq_cla,cod_dep,cod_pla,valor_cla,cod_usu,tipo_cla)values("
    strSQL = strSQL & "'" & Me.tag & "', "
    strSQL = strSQL & "'" & cmbxDep.tag & "',"
    strSQL = strSQL & "'" & F_PlanoConta(False, cmbxConta.text, cmndbtnConfirma.tag) & "', "
    
    strSQL = strSQL & F_GravaDec(lblTotal.caption) & ","
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & operacao & "') "
    Conn.Execute strSQL
    Me.Hide
End Sub

Private Sub cmndbtnExcl_Click()
    Dim sVal As Currency
    If msflxgrd.rows = 1 Then
        MsgBox "Não há o que ser excluído!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma exlcusão do Lançamento?", vbYesNo + vbCritical) = vbYes Then
        msflxgrd.col = 1
        sVal = CCur(msflxgrd.text)
        lblDistr.caption = FormatVal(CCur(lblDistr.caption) - sVal)
        If msflxgrd.rows = 2 Then
            msflxgrd.rows = 1
        Else
            msflxgrd.RemoveItem msflxgrd.row
        End If
        
    End If
End Sub

Private Sub cmndbtnLancar_Click()
    Dim strSQL As String
    If cmbxConta.text = "" Then
        MsgBox "Informe a Conta!", vbExclamation
        cmbxConta.SetFocus
        Exit Sub
    End If
    If txtbxValorBaixa.text = "0,00" Then
        MsgBox "Informe um valor válido!", vbExclamation
        txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    'If F_InFlxgrd(msflxgrd, cmbxConta.Text, 0) <> -1 Then
    '    MsgBox "Já existe um lançamento para esta Conta!", vbExclamation
    '    Exit Sub
    'End If
    strSQL = cmbxConta.text & Chr(9)
    strSQL = strSQL & txtbxValorBaixa.text
    msflxgrd.AddItem strSQL
    lblDistr.caption = FormatVal(CCur(lblDistr.caption) + CCur(txtbxValorBaixa.text))
    
End Sub

Private Sub cmndbtnSair_Click()
    If F_Restante < 0 Then
        MsgBox "O Valor distribuído excede o total!", vbExclamation
        Exit Sub
    End If
    Me.tag = ""
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    P_PreenchePlanoContas cmbxConta, IIf(cmndbtnConfirma.tag <> "D" And cmndbtnConfirma.tag <> "P", True, False)
    P_PreencheDepto cmbxDep
    busca
    With msflxgrd
        
    End With
    strSQL = "select * from tab_class where seq_cla='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = F_PlanoConta(True, !cod_pla, cmndbtnConfirma.tag) & _
            Chr(9) & FormatVal(!valor_cla)
            'msflxgrd.AddItem strSQL
            cTotal = cTotal + !valor_cla
            .MoveNext
        Loop
        'lblDistr.caption = FormatVal(cTotal)
        .Close
    End With
    MudaAparencia Me
End Sub
Sub busca()
    
    
End Sub
Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmndbtnExcl_Click
    End If
End Sub

Private Sub txtbxValorBaixa_GotFocus()
    txtbxValorBaixa.text = F_Restante
    'P_SelecTexto txtbxValorBaixa
End Sub

Private Sub txtbxValorBaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValorBaixa_LostFocus()
    'P_ValorLostFocus txtbxValorBaixa
End Sub

