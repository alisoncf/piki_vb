VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBaixaPar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa de Parcela"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "frmBaixaPar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxValor_Par 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3120
      TabIndex        =   23
      Text            =   "0,00"
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtbxReman 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2040
      TabIndex        =   22
      Text            =   "0,00"
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtbxTotalBaixa 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   12
      Text            =   "0,00"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Cancela&r"
      Height          =   735
      Left            =   7680
      Picture         =   "frmBaixaPar.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar Baixa"
      Height          =   735
      Left            =   6240
      Picture         =   "frmBaixaPar.frx":04F5
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   525
         Left            =   7680
         Picture         =   "frmBaixaPar.frx":0C5F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0005C9EB&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtbxJuros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0005C9EB&
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnLancar 
         Caption         =   "&Lançar"
         Height          =   525
         Left            =   6360
         Picture         =   "frmBaixaPar.frx":11E9
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtbxValorBaixa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.ComboBox cmbxEspecie 
         Height          =   315
         ItemData        =   "frmBaixaPar.frx":1773
         Left            =   120
         List            =   "frmBaixaPar.frx":1783
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   4080
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   3375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5953
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
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar:"
         Height          =   195
         Left            =   2520
         TabIndex        =   3
         Top             =   3840
         Width           =   540
      End
      Begin VB.Label lblTur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         Height          =   195
         Left            =   4440
         TabIndex        =   19
         Top             =   2280
         Width           =   180
      End
      Begin VB.Label lblCaixa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         Height          =   195
         Left            =   4200
         TabIndex        =   18
         Top             =   2280
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto(R$):"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   4440
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Juros(R$):"
         Height          =   195
         Left            =   1320
         TabIndex        =   15
         Top             =   4440
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor a Baixar:"
         Height          =   195
         Left            =   4800
         TabIndex        =   5
         Top             =   3840
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fluxo de Saída:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   3840
         Width           =   1125
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicial:"
      Height          =   195
      Left            =   3120
      TabIndex        =   25
      Top             =   5520
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remanesc.:"
      Height          =   195
      Left            =   2040
      TabIndex        =   24
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblNome_Ter 
      AutoSize        =   -1  'True
      Caption         =   "Terceiro:"
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
      Left            =   120
      TabIndex        =   21
      Tag             =   "!"
      Top             =   5160
      Width           =   780
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
      Left            =   4200
      TabIndex        =   20
      Tag             =   "!"
      Top             =   5520
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total da Baixa:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   1065
   End
End
Attribute VB_Name = "frmBaixaPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim che As cheque 'caso chama o cheque
Dim cheP As chePag
Private Sub cmbxConta_Click()
    If cmbxConta.ListIndex = 1 Then
        cmbxConta.tag = F_BuscaUmCampo("cod_con", "tab_contas", "nome_con='" & cmbxConta.text & "'")
    Else
        cmbxConta.tag = F_BuscaUmCampo("cod_fin", "tab_financeiras", "nome_fin='" & cmbxConta.text & "'")
    End If
End Sub

Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxEspecie_Click()
    If cmbxEspecie.ListIndex = 2 Then
        txtbxValorBaixa.text = "0,00"
        txtbxValorBaixa.Enabled = False
        
    ElseIf cmbxEspecie.ListIndex = 3 Then
        'P_PreencheFinanceiras cmbxConta
        txtbxValorBaixa.Enabled = True
        'cmbxConta.Enabled = True
    ElseIf cmbxEspecie.ListIndex = 0 Then
        txtbxExpr.text = "Caixa " & strCod_Cai
        txtbxExpr.tag = strCod_Cai
    Else
        'P_PreencheContas cmbxConta, 3
        txtbxValorBaixa.Enabled = True
        txtbxExpr.text = ""
                
        
    End If
    lblExpr.caption = Mid(cmbxEspecie.text, 5)
End Sub

Private Sub cmbxEspecie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnConfirma_Click()
    If msflxgrd.rows = 1 Then
        MsgBox "Não foi informado nenhum valor para ser baixado!", vbExclamation
    Else
        Me.tag = "ok"
        Me.Hide
    End If
End Sub

Private Sub cmndbtnExcl_Click()
      Dim cValor As Currency
    If msflxgrd.rows = 1 Then
            MsgBox "Não há o que ser deletado!", vbExclamation
        Else
            If MsgBox("Confirma exclusão da linha?", vbYesNo + vbCritical) = vbYes Then
                msflxgrd.col = 1
                cValor = CCur(msflxgrd.text)
                If msflxgrd.rows = 2 Then
                    msflxgrd.rows = 1
                Else
                    msflxgrd.RemoveItem msflxgrd.row
                End If
                txtbxTotalBaixa.text = FormatVal(CCur(txtbxTotalBaixa.text) - cValor)
            End If
            
        End If
End Sub

Private Sub cmndbtnLancar_Click()
    
    Dim strSQL As String
    On Error Resume Next
    If CCur(txtbxValorBaixa.text) > F_Restante Then   'cmbxEspecie.Text = "1 - Dinheiro" And + CCur(txtbxTotalBaixa.Text)
        MsgBox "O Valor a ser lançado deve ser menor ou igual ao remanescente!", vbExclamation
          
        'txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    If CCur(txtbxJuros.text) < 0 Then
        MsgBox "Juros inválidos!", vbExclamation
        txtbxJuros.SetFocus
        Exit Sub
    End If
    If CCur(txtbxDesconto.text) < 0 Then
        MsgBox "Juros inválidos!", vbExclamation
        txtbxJuros.SetFocus
        Exit Sub
    End If
    
    
    If cmbxEspecie.ListIndex = 0 Then
        If Not F_IsCaixa Then
            MsgBox "Este computador não é um caixa, por tanto não pode lançar no caixa por aqui!", vbExclamation
            Exit Sub
        End If
        
        strSQL = Date & Chr(9) & txtbxValorBaixa.text & Chr(9) & cmbxEspecie.text
        strSQL = strSQL & Chr(9) & Chr(9) & Chr(9) & Chr(9) & strCod_Cai
        msflxgrd.AddItem strSQL
        txtbxTotalBaixa.text = FormatVal(CCur(txtbxTotalBaixa.text) + CCur(txtbxValorBaixa.text))
        
    
    ElseIf cmbxEspecie.ListIndex = 1 Then
    
        If txtbxExpr.tag = "" Then
            MsgBox "Escolha a conta corrente", vbExclamation
            buscaCco
        End If
        strSQL = Date & Chr(9) & txtbxValorBaixa.text & Chr(9) & cmbxEspecie.text
        strSQL = strSQL & Chr(9) & Chr(9) & Chr(9) & Chr(9) & txtbxExpr.tag
        msflxgrd.AddItem strSQL
        
        txtbxTotalBaixa.text = FormatVal(CCur(txtbxTotalBaixa.text) + CCur(txtbxValorBaixa.text))
        
    ElseIf cmbxEspecie.ListIndex = 2 Then
        strSQL = cheP.Venc_che & Chr(9) & cheP.Valor_che & Chr(9)
            strSQL = strSQL & cmbxEspecie.text & Chr(9) & cheP.num_che & Chr(9)
            strSQL = strSQL & cheP.cod_con & Chr(9) & "" & Chr(9)
            strSQL = strSQL & cheP.cod_con & Chr(9) & F_BuscaUmCampo("Nome_con", "tab_Contas", "cod_con='" & cp.num_che & "'")
            msflxgrd.AddItem strSQL
            txtbxTotalBaixa.text = FormatVal(CCur(txtbxTotalBaixa.text) + cheP.Valor_che)
    
    ElseIf cmbxEspecie.ListIndex = 3 Then
        strSQL = che.Venc_che & Chr(9)
        strSQL = strSQL & che.Valor_che & Chr(9)
        strSQL = strSQL & cmbxEspecie.text & Chr(9)
        strSQL = strSQL & che.num_che & Chr(9)
        strSQL = strSQL & che.cod_bco & Chr(9)
        strSQL = strSQL & che.age_che & Chr(9)
        strSQL = strSQL & che.cco_che & Chr(9)
        strSQL = strSQL & che.nome_che & Chr(9)
        strSQL = strSQL & che.cpf_che & Chr(9)
        strSQL = strSQL & che.seq_che
        msflxgrd.AddItem strSQL
        txtbxTotalBaixa.text = FormatVal(CCur(txtbxTotalBaixa.text) + che.Valor_che)
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub Form_Load()
    P_CentraForm Me, 0
    'P_PreencheContas cmbxConta, 3
    msflxgrd.rows = 1
    msflxgrd.Cols = 10
    msflxgrd.col = 0
    msflxgrd.text = "Vencimento"
    msflxgrd.col = 1
    msflxgrd.text = "Valor"
    msflxgrd.col = 2
    msflxgrd.text = "Espécie"
    msflxgrd.col = 3
    msflxgrd.text = "Cheque Nº"
    msflxgrd.col = 4
    msflxgrd.text = "Banco"
    msflxgrd.col = 5
    msflxgrd.text = "Agencia"
    msflxgrd.col = 6
    msflxgrd.text = "Conta C."
    
    msflxgrd.col = 7
    msflxgrd.text = "Titular"
    msflxgrd.col = 8
    msflxgrd.text = "CPF"
    msflxgrd.ColAlignment(2) = 1
    msflxgrd.col = 9
    msflxgrd.text = "Seq.Che"
    msflxgrd.ColAlignment(2) = 1
    
    msflxgrd.ColWidth(0) = 1000
    msflxgrd.ColWidth(1) = 1000
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyDelete Then
        cmndbtnExcl_Click
    End If
    
End Sub

Private Sub optnCaixa_Click()
    cmbxConta.Enabled = False
    Dim sStatus As String
    If strCod_Cai = "" Or strCod_Tur = "" Then
        MsgBox "Operação impossível!", vbExclamation
        optnFin.Value = True
        Exit Sub
    End If
    sStatus = F_BuscaUmCampo("status_Cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
    If sStatus = "F" Then
        MsgBox "Não é possível efetuar esse recebimento pois o caixa se encontra fechado!", vbExclamation
        optnFin.Value = True
        Exit Sub
    End If
    
    
    
    
    
End Sub

Private Sub optnChequeCli_Click()
    
End Sub

Private Sub optnFin_Click()
    cmbxConta.Enabled = True
    cmbxConta.SetFocus
End Sub

Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    P_ValorLostFocus txtbxDesconto
End Sub

Private Sub txtbxExpr_Change()
    txtbxExpr.tag = ""
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaConta
    End If
End Sub
Sub buscaConta()
    If cmbxEspecie.ListIndex = 0 Then
        txtbxExpr.text = "Caixa " & strCod_Cai
    ElseIf cmbxEspecie.ListIndex = 1 Then
        buscaCco
    ElseIf cmbxEspecie.ListIndex = 2 Then
        buscaChe
    ElseIf cmbxEspecie.ListIndex = 3 Then
        buscaCheTer
    End If
End Sub
Sub buscaChe()
    cheP = F_DadosChePag()
    
End Sub
Sub buscaCheTer()
    che = F_BuscaChequeTer(txtbxExpr.text)
    txtbxExpr.text = che.num_che
    txtbxExpr.tag = che.seq_che
End Sub
Sub buscaCco()
    Dim c As ContaCorrente
    c = financeiro.F_BuscaContaCorrente("")
    If c.cod_con <> "" Then
        txtbxExpr.text = c.nome_con
        txtbxExpr.tag = c.cod_con
        txtbxValorBaixa.SetFocus
    End If
End Sub
Private Sub txtbxJuros_GotFocus()
    P_SelecTexto txtbxJuros
End Sub

Private Sub txtbxJuros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxJuros_LostFocus()
    P_ValorLostFocus txtbxJuros
End Sub

Private Sub txtbxValorBaixa_GotFocus()
    txtbxValorBaixa.text = F_Restante  'CCur(txtbxReman.Text) - CCur(txtbxTotalBaixa.Text)
    P_SelecTexto txtbxValorBaixa
        
End Sub
Function F_Restante() As Currency
    F_Restante = CCur(txtbxReman.text) - CCur(txtbxTotalBaixa.text) + CCur(txtbxJuros.text) - CCur(txtbxDesconto.text)
End Function

Private Sub txtbxValorBaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValorBaixa_LostFocus()
    P_ValorLostFocus txtbxValorBaixa
    If CCur(txtbxValorBaixa.text) > F_Restante Then
        MsgBox "O Valor não pode ser maior que o restante!", vbExclamation
        txtbxValorBaixa_GotFocus
        txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    
End Sub
