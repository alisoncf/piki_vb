VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBaixaPar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa de Parcela"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "frmBaixaPar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxTotalBaixa 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   7200
      TabIndex        =   18
      Text            =   "0,00"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Cancela&r"
      Height          =   735
      Left            =   7680
      Picture         =   "frmBaixaPar.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar Baixa (F7)"
      Height          =   735
      Left            =   7800
      Picture         =   "frmBaixaPar.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.ComboBox cmbxConta 
         BackColor       =   &H00008000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmBaixaPar.frx":061E
         Left            =   2520
         List            =   "frmBaixaPar.frx":0628
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "!"
         Top             =   2250
         Width           =   3135
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "<< E&xcluir"
         Height          =   405
         Left            =   5760
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtbxJuros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Text            =   "0,00"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnLancar 
         Caption         =   ">> &Lançar"
         Height          =   405
         Left            =   5760
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtbxValorBaixa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cmbxEspecie 
         Height          =   315
         ItemData        =   "frmBaixaPar.frx":0646
         Left            =   2520
         List            =   "frmBaixaPar.frx":0650
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1650
         Width           =   1815
      End
      Begin VB.TextBox txtbxReman 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   16
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
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   1680
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
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
         TabIndex        =   23
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Juros(R$):"
         Height          =   195
         Left            =   1200
         TabIndex        =   21
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
         TabIndex        =   17
         Tag             =   "!"
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor a Baixar:"
         Height          =   195
         Left            =   4560
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
         TabIndex        =   15
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label lblNome_Ter 
         AutoSize        =   -1  'True
         Caption         =   "Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
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
      Left            =   7200
      TabIndex        =   19
      Top             =   2040
      Width           =   1065
   End
End
Attribute VB_Name = "frmBaixaPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxConta_Click()
    cmbxConta.tag = F_BuscaUmCampo("cod_con", "tab_contas", "nome_con='" & cmbxConta.Text & "'")
    
End Sub

Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub cmbxEspecie_Click()
    If cmbxEspecie.ListIndex = 1 Then
        txtbxValorBaixa.Text = "0,00"
        cmbxConta.Enabled = False
        txtbxValorBaixa.Enabled = False
    Else
        txtbxValorBaixa.Enabled = True
        cmbxConta.Enabled = True
    End If
End Sub

Private Sub cmbxEspecie_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub cmndbtnConfirma_Click()
    If Msflxgrd.rows = 1 Then
        MsgBox "Não foi informado nenhum valor para ser baixado!", vbExclamation
    Else
        Me.tag = "ok"
        Me.Hide
    End If
End Sub

Private Sub cmndbtnExcl_Click()
      Dim cValor As Currency
    If Msflxgrd.rows = 1 Then
            MsgBox "Não há o que ser deletado!", vbExclamation
        Else
            If MsgBox("Confirma exclusão da linha?", vbYesNo + vbCritical) = vbYes Then
                Msflxgrd.col = 1
                cValor = CCur(Msflxgrd.Text)
                If Msflxgrd.rows = 2 Then
                    Msflxgrd.rows = 1
                Else
                    Msflxgrd.RemoveItem Msflxgrd.Row
                End If
                txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) - cValor)
            End If
            
        End If
End Sub

Private Sub cmndbtnLancar_Click()
    
    Dim strSQL As String
    If CCur(txtbxValorBaixa.Text) > F_Restante Then   'cmbxEspecie.Text = "1 - Dinheiro" And + CCur(txtbxTotalBaixa.Text)
        MsgBox "O Valor a ser lançado deve ser menor ou igual ao remanescente!", vbExclamation
          
        'txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    If cmbxEspecie.ListIndex = 0 Then
        strSQL = Date & Chr(9) & txtbxValorBaixa.Text & Chr(9) & cmbxEspecie.Text
        strSQL = strSQL & Chr(9) & Chr(9) & Chr(9) & Chr(9) & cmbxConta.tag
        Msflxgrd.AddItem strSQL
        txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) + CCur(txtbxValorBaixa.Text))
        If cmbxConta.Text = "" And cmbxEspecie.Text = "1 - Dinheiro" Then
                MsgBox "Escolha a Conta!", vbInformation
                cmbxConta.SetFocus
                Exit Sub
            End If
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
            Msflxgrd.AddItem strSQL
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
            Msflxgrd.AddItem strSQL
            txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) + cp.Valor_che)
        End If
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnConfirma_Click
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me, 0
    P_PreencheContas cmbxConta, 3
    Msflxgrd.rows = 1
    Msflxgrd.Cols = 9
    Msflxgrd.col = 0
    Msflxgrd.Text = "Vencimento"
    Msflxgrd.col = 1
    Msflxgrd.Text = "Valor"
    Msflxgrd.col = 2
    Msflxgrd.Text = "Espécie"
    Msflxgrd.col = 3
    Msflxgrd.Text = "Cheque Nº"
    Msflxgrd.col = 4
    Msflxgrd.Text = "Banco"
    Msflxgrd.col = 5
    Msflxgrd.Text = "Agencia"
    Msflxgrd.col = 6
    Msflxgrd.Text = "Conta C."
    
    Msflxgrd.col = 7
    Msflxgrd.Text = "Titular"
    Msflxgrd.col = 8
    Msflxgrd.Text = "CPF"
    Msflxgrd.ColAlignment(2) = 1
    
    Msflxgrd.ColWidth(0) = 1000
    Msflxgrd.ColWidth(1) = 1000
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
