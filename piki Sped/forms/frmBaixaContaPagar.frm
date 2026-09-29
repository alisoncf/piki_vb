VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBaixaContaPagar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa de Parcela"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "frmBaixaContaPagar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxTotalBaixa 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   19
      Text            =   "0,00"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Cancela&r"
      Height          =   735
      Left            =   8040
      Picture         =   "frmBaixaContaPagar.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar Baixa (F7)"
      Height          =   735
      Left            =   6120
      Picture         =   "frmBaixaContaPagar.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txtConta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   2280
         Width           =   4215
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "<< E&xcluir"
         Height          =   405
         Left            =   8280
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Text            =   "0,00"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtbxJuros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Text            =   "0,00"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnLancar 
         Caption         =   ">> &Lançar"
         Height          =   405
         Left            =   7200
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtbxValorBaixa 
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
         Left            =   4920
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cmbxEspecie 
         Height          =   315
         ItemData        =   "frmBaixaContaPagar.frx":061E
         Left            =   2520
         List            =   "frmBaixaContaPagar.frx":0628
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1650
         Width           =   2295
      End
      Begin VB.TextBox txtbxReman 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtbxValor_Par 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   1680
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   1095
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   1931
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         FillStyle       =   1
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.CommandButton cmdConta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         Picture         =   "frmBaixaContaPagar.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&onta:"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto(R$):"
         Height          =   195
         Left            =   1200
         TabIndex        =   24
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Juros(R$):"
         Height          =   195
         Left            =   1200
         TabIndex        =   22
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lblParc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   6120
         TabIndex        =   18
         Tag             =   "!"
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor a Baixar:"
         Height          =   195
         Left            =   4920
         TabIndex        =   3
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&spécie:"
         Height          =   195
         Left            =   2520
         TabIndex        =   1
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remanesc.:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label lblNome_Ter 
         AutoSize        =   -1  'True
         Caption         =   "Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Tag             =   "!"
         Top             =   0
         Width           =   630
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total da Baixa:"
      Height          =   195
      Left            =   2880
      TabIndex        =   20
      Top             =   3480
      Width           =   1065
   End
End
Attribute VB_Name = "frmBaixaContaPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cmbxConta_Click()
'    cmbxConta.tag = F_BuscaUmCampo("cod_con", "tab_contas", "nome_con='" & cmbxConta.Text & "'")
'
'End Sub

'Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
'    TrocaTabKey KeyCode
'End Sub

Private Sub cmbxEspecie_Click()
    If cmbxEspecie.ListIndex = 1 Then
        txtbxValorBaixa.Text = "0,00"
        txtConta.Enabled = False
        txtbxValorBaixa.Enabled = False
    Else
        txtbxValorBaixa.Enabled = True
        txtConta.Enabled = True
    End If
End Sub

Private Sub cmbxEspecie_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub cmdConta_Click()
    buscaConta
End Sub
Sub buscaConta()
    Dim c As ContaCorrente
    c = F_BuscaContaCorrente(txtConta.Text)
    If c.cod_con <> "" Then
        txtConta.Text = c.nome_con
        txtConta.tag = c.cod_con
        cmndbtnLancar.SetFocus
    End If
    
    
    
    
    
End Sub
Private Sub cmndbtnConfirma_Click()
    Dim iCount As Integer
    Dim conta As String
    Dim valor As String
    Dim ESP As String
    Dim num_ch As String
    Dim Banco As String
    Dim titular As String
    Dim Agencia As String
    Dim cpf As String
    Dim venc As String
    Dim strSQL As String
    Dim sSEQ_Bai As String
    Dim sFin As String
    Dim s2 As String
    If msflxgrd.rows = 1 Then
        MsgBox "Não foi informado nenhum valor para ser baixado!", vbExclamation
        Exit Sub
    Else
        'Me.tag = "ok"
        'Me.Hide
    End If
    
    sSEQ_Bai = F_BuscaChaveBaixa
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            .col = 0: venc = .Text
            .col = 1: valor = .Text
            .col = 2: ESP = .Text
            .col = 3: num_ch = .Text
            .col = 4: Banco = .Text
            .col = 5: Agencia = .Text
            .col = 6: conta = .Text
            .col = 7: titular = .Text
            .Row = iCount
            
                    
            strSQL = "update " & TAB_ITDOC & " set reman_itd=reman_itd-" & F_GravaDec(valor)
            strSQL = strSQL & " where seq_itd='" & Me.tag & "' "
            conn.Execute strSQL
            
            sFin = F_GravaFinanceiro("D", Date, valor, "BAIXA DOC.P " & Left(lblNome_Ter.caption, 30), conta, "N", strPlaNaoClassD, "")
            
            strSQL = F_GravaBaixaItd(sSEQ_Bai, Me.tag, valor, ESP, sFin, "")
            s2 = F_GravaBaixaItdFin(sSEQ_Bai, sFin, ESP, valor)
            
            
            
            
            
            
        Next
        
    End With
    MsgBox "Parcela baixada!", vbInformation
    Unload Me
End Sub

Private Sub cmndbtnExcl_Click()
      Dim cValor As Currency
    If msflxgrd.rows = 1 Then
            MsgBox "Não há o que ser deletado!", vbExclamation
        Else
            If MsgBox("Confirma exclusão da linha?", vbYesNo + vbCritical) = vbYes Then
                msflxgrd.col = 1
                cValor = CCur(msflxgrd.Text)
                If msflxgrd.rows = 2 Then
                    msflxgrd.rows = 1
                Else
                    msflxgrd.RemoveItem msflxgrd.Row
                End If
                txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) - cValor)
            End If
            
        End If
End Sub

Private Sub cmndbtnLancar_Click()
    
    Dim strSQL As String
    
    If txtConta.tag = "" And Left(cmbxEspecie.Text, 1) = "1" Then
        MsgBox "Escolha a conta!", vbExclamation
        Exit Sub
    End If
    
    
    If CCur(txtbxValorBaixa.Text) > F_Restante Then   'cmbxEspecie.Text = "1 - Dinheiro" And + CCur(txtbxTotalBaixa.Text)
        MsgBox "O Valor a ser lançado deve ser menor ou igual ao remanescente!", vbExclamation
          
        'txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    If cmbxEspecie.ListIndex = 0 Then
        strSQL = Date & Chr(9) & txtbxValorBaixa.Text & Chr(9) & cmbxEspecie.Text
        strSQL = strSQL & Chr(9) & Chr(9) & Chr(9) & Chr(9) & txtConta.tag
        msflxgrd.AddItem strSQL
        
        txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) + CCur(txtbxValorBaixa.Text))
        
    ElseIf cmbxEspecie.ListIndex = 1 Then
        Dim c As cheque
        If cmndbtnConfirma.tag = "R" Then
'            If cmbxConta.Text = "" And cmbxEspecie.Text = "1 - Dinheiro" Then
'                MsgBox "Escolha a Conta!", vbInformation
'                cmbxConta.SetFocus
'                Exit Sub
'            End If
            c = F_DadosCheque(True, Date, lblNome_Ter.caption, "", "", "", "", F_Restante)
            If c.Valor_che = 0 Then
                Unload frmDadosChe
                Exit Sub
            End If
            Do While c.Valor_che > F_Restante
                MsgBox "O Valor a ser lançado deve ser menor ou igual ao remanescente!", vbExclamation
                c = F_DadosCheque(False, Date, lblNome_Ter.caption, c.num_che, c.cco_che, c.age_che, c.cod_bco, "0,00")
        
            
            Loop
            strSQL = c.Venc_che & Chr(9) & c.Valor_che & Chr(9)
            strSQL = strSQL & cmbxEspecie.Text & Chr(9) & c.num_che & Chr(9)
            strSQL = strSQL & c.cod_bco & Chr(9) & c.age_che & Chr(9)
            strSQL = strSQL & c.cco_che & Chr(9) & c.nome_che & Chr(9)
            strSQL = strSQL & c.cpf_che
            msflxgrd.AddItem strSQL
            txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) + c.Valor_che)
        Else
            Dim cp As chePag
            cp = F_DadosChePag(Date, , , F_Restante)
            If cp.Valor_che = 0 Then
                Unload frmDadosChe
                Exit Sub
            End If
            If cp.Valor_che > F_Restante Then
                MsgBox "O Valor está maior que o restante!", vbExclamation
                Unload frmDadosChe
                Exit Sub
            End If
            strSQL = cp.Venc_che & Chr(9) & cp.Valor_che & Chr(9)
            strSQL = strSQL & cmbxEspecie.Text & Chr(9) & cp.num_che & Chr(9)
            strSQL = strSQL & cp.cod_con & Chr(9) & "" & Chr(9)
            strSQL = strSQL & cp.cod_con & Chr(9) & F_BuscaUmCampo("Nome_con", "tab_Contas", "cod_con='" & cp.num_che & "'")
            msflxgrd.AddItem strSQL
            txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) + cp.Valor_che)
        End If
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnConfirma_Click
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me, 0
    'P_PreencheContas cmbxConta, 3
    msflxgrd.rows = 1
    msflxgrd.Cols = 9
    msflxgrd.col = 0
    msflxgrd.Text = "Vencimento"
    msflxgrd.col = 1
    msflxgrd.Text = "Valor"
    msflxgrd.col = 2
    msflxgrd.Text = "Espécie"
    msflxgrd.col = 3
    msflxgrd.Text = "Cheque Nº"
    msflxgrd.col = 4
    msflxgrd.Text = "Banco"
    msflxgrd.col = 5
    msflxgrd.Text = "Agencia"
    msflxgrd.col = 6
    msflxgrd.Text = "Conta C."
    
    msflxgrd.col = 7
    msflxgrd.Text = "Titular"
    msflxgrd.col = 8
    msflxgrd.Text = "CPF"
    msflxgrd.ColAlignment(2) = 1
    
    msflxgrd.ColWidth(0) = 1000
    msflxgrd.ColWidth(1) = 1000
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyDelete Then
        cmndbtnExcl_Click
    End If
    
End Sub

Private Sub txtbxDesconto_GotFocus()
    ganhaFoco txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    PerdeFoco txtbxDesconto
End Sub

Private Sub txtbxJuros_GotFocus()
    ganhaFoco txtbxJuros
End Sub

Private Sub txtbxJuros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxJuros_LostFocus()
    PerdeFoco txtbxJuros
End Sub

Private Sub txtbxValorBaixa_GotFocus()
    txtbxValorBaixa.Text = F_Restante  'CCur(txtbxReman.Text) - CCur(txtbxTotalBaixa.Text)
    ganhaFoco txtbxValorBaixa
        
End Sub
Function F_Restante() As Currency
    F_Restante = CCur(txtbxReman.Text) - CCur(txtbxTotalBaixa.Text) + CCur(txtbxJuros.Text) - CCur(txtbxDesconto.Text)
End Function

Private Sub txtbxValorBaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValorBaixa_LostFocus()
    PerdeFoco txtbxValorBaixa
    If CCur(txtbxValorBaixa.Text) > F_Restante Then
        MsgBox "O Valor não pode ser maior que o restante!", vbExclamation
        txtbxValorBaixa_GotFocus
        txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtConta_Change()
    txtConta.tag = ""
End Sub

Private Sub txtConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaConta
    End If
End Sub
