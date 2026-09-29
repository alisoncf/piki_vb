VERSION 5.00
Begin VB.Form frmInfPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informações sobre o Produto:"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "frmInfPro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbxER 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      TabIndex        =   35
      Text            =   "0"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtbxEF 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      TabIndex        =   34
      Text            =   "0"
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmndbtnAlt 
      Caption         =   "Alterar"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   5400
      TabIndex        =   21
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ER:"
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   4200
      Width           =   270
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EF:"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   4560
      Width           =   240
   End
   Begin VB.Label lblForn 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedores:"
      Height          =   195
      Left            =   1200
      TabIndex        =   32
      Top             =   3360
      Width           =   1020
   End
   Begin VB.Label lblDiscr 
      Alignment       =   1  'Right Justify
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
      Left            =   3405
      TabIndex        =   31
      Top             =   2760
      Width           =   390
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   195
      Left            =   3945
      TabIndex        =   30
      Top             =   2760
      Width           =   120
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Índice de Discrepância:"
      Height          =   195
      Left            =   1200
      TabIndex        =   29
      Top             =   2760
      Width           =   1680
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Itens"
      Height          =   195
      Left            =   3945
      TabIndex        =   28
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Itens"
      Height          =   195
      Left            =   3945
      TabIndex        =   27
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Itens"
      Height          =   195
      Left            =   3945
      TabIndex        =   26
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R$"
      Height          =   195
      Left            =   3945
      TabIndex        =   25
      Top             =   1800
      Width           =   210
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R$"
      Height          =   195
      Left            =   3945
      TabIndex        =   24
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R$"
      Height          =   195
      Left            =   3945
      TabIndex        =   23
      Top             =   2280
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R$"
      Height          =   195
      Left            =   3945
      TabIndex        =   22
      Top             =   2520
      Width           =   210
   End
   Begin VB.Label lblEndereco 
      AutoSize        =   -1  'True
      Caption         =   "Unidade"
      Height          =   195
      Left            =   3720
      TabIndex        =   20
      Top             =   600
      Width           =   600
   End
   Begin VB.Label lblUltComp 
      AutoSize        =   -1  'True
      Caption         =   "Última Compra"
      Height          =   195
      Left            =   1200
      TabIndex        =   19
      Top             =   3000
      Width           =   1020
   End
   Begin VB.Label lblSai 
      Alignment       =   1  'Right Justify
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
      Left            =   3405
      TabIndex        =   18
      Top             =   2520
      Width           =   390
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Preço de Agregado"
      Height          =   195
      Left            =   1200
      TabIndex        =   17
      Top             =   2520
      Width           =   1380
   End
   Begin VB.Label lblSit 
      Alignment       =   1  'Right Justify
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
      Left            =   3645
      TabIndex        =   16
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Situação:"
      Height          =   195
      Left            =   1200
      TabIndex        =   15
      Top             =   840
      Width           =   675
   End
   Begin VB.Label lblUnidade 
      AutoSize        =   -1  'True
      Caption         =   "Unidade"
      Height          =   195
      Left            =   1200
      TabIndex        =   14
      Top             =   600
      Width           =   600
   End
   Begin VB.Label lblPreco 
      Alignment       =   1  'Right Justify
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
      Left            =   3405
      TabIndex        =   13
      Top             =   2280
      Width           =   390
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Preço de Venda"
      Height          =   195
      Left            =   1200
      TabIndex        =   12
      Top             =   2280
      Width           =   1155
   End
   Begin VB.Label lblCustoFiscal 
      Alignment       =   1  'Right Justify
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
      Left            =   3405
      TabIndex        =   11
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Custo F:"
      Height          =   195
      Left            =   1200
      TabIndex        =   10
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label lblCusto 
      Alignment       =   1  'Right Justify
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
      Left            =   3405
      TabIndex        =   9
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Custo Venda:"
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label lblPre 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3675
      TabIndex        =   7
      Top             =   1560
      Width           =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade a Efetivar:"
      Height          =   195
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   1590
   End
   Begin VB.Label lblRes 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3675
      TabIndex        =   5
      Top             =   1320
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Est. Fiscal:"
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label lblEst 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3675
      TabIndex        =   3
      Top             =   1080
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Estoque:"
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   630
   End
   Begin VB.Label lblcodigo 
      AutoSize        =   -1  'True
      Caption         =   "000.000-0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label lblNome_Pro 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Produto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2430
   End
End
Attribute VB_Name = "frmInfPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnAlt_Click()
    Dim strSQL As String
    If strAltpreco_usu <> "S" Then
        MsgBox "Você não tem permissão para este!", vbCritical
    Else
        If MsgBox("Confirma alteração do estoque do produto?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "update tab_produtos set ef_pro=" & F_GravaDec(txtbxEF.Text) & ","
            strSQL = strSQL & "er_pro=" & F_GravaDec(txtbxER.Text)
            strSQL = strSQL & " where cod_Pro='" & Me.Tag & "' "
            Banco.Execute strSQL
            MsgBox "Produto alterado!", vbInformation
        End If
    End If
    
End Sub

Private Sub txtbxEF_GotFocus()
    
    P_SelecTexto txtbxER
End Sub

Private Sub txtbxEF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEF_LostFocus()
    
    If CCur(txtbxEF.Text) > 0 Then
        P_ValorLostFocus txtbxEF
    Else
        P_ValorLostFocus txtbxEF
        
        txtbxEF.Text = "-" & txtbxEF.Text
        
        
        
    End If
End Sub

Private Sub txtbxER_GotFocus()
    P_SelecTexto txtbxER
     
End Sub

Private Sub txtbxER_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxER_LostFocus()
    If CCur(txtbxER.Text) > 0 Then
        P_ValorLostFocus txtbxER
    Else
        P_ValorLostFocus txtbxER
        
        txtbxER.Text = "-" & txtbxER.Text
        
        
        
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    Dim rPesq As Recordset
    'On Error Resume Next
    strSQL = "select tp.* "
    strSQL = strSQL & " from  tab_produtos as tp "
    strSQL = strSQL & "where tp.cod_pro='" & Me.Tag & "'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        lblcodigo.caption = Format(!cod_pro, "00#\.###\-#")
        lblNome_pro.caption = !nome_pro
        lblEndereco.caption = "Endereço: " & !Endereco_pro
        lblUnidade.caption = "UN.: " & !unidade_pro
        lblCusto.caption = FormatVal(!cr_pro)
        lblCustoFiscal.caption = FormatVal(!cf_pro)
        lblEst.caption = FormatVal(!er_pro)
        lblSit.caption = !st_pro
        txtbxER.Text = FormatVal(!er_pro)
        txtbxEF.Text = FormatVal(!ef_pro)
        lblPre.caption = F_BuscaUmCampo("sum(tp.qtd_ite)", _
        "tab_itcmp as tp, tab_cmp as tv", "tp.cod_pro='" & _
            !cod_pro & "' and tp.cod_for=tv.cod_for " & _
            " and tv.serie_not=tp.serie_not " & _
            " and tp.num_not=tv.num_not and tv.status_not <3")
        If lblPre.caption = "" Then
            lblPre.caption = "0,00"
        Else
            lblPre.caption = FormatVal(lblPre.caption)
        End If
        lblPreco.caption = FormatVal(!pV_pro)
        lblRes.caption = FormatVal(!ef_pro)
        lblSai.caption = FormatVal(!pV_pro)
        lblUltComp.caption = ""
        
        lblUltComp.caption = ""
        lblDiscr.caption = FormatVal(!cf_pro / !cr_pro * 100)
        .Close
    End With
     On Error GoTo trata_erro
    strSQL = "select distinct tf.cod_for,tf.nome_for from tab_fornecedores as tf,"
    strSQL = strSQL & "tab_cmp as tc, tab_itcmp as ti "
    strSQL = strSQL & " where ti.num_not=tc.num_not and ti.serie_not=tc.serie_not "
    strSQL = strSQL & " and tc.cod_for=ti.cod_for and tc.cod_for=tf.cod_for "
    strSQL = strSQL & "  and ti.cod_pro='" & Me.Tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        strSQL = ""
        Do While Not .EOF
            
            strSQL = strSQL & !nome_for & Chr(13)
            
            .MoveNext
        Loop
        .Close
    End With
    lblForn.caption = "Fornecedores: " & Chr(13) & strSQL
    strSQL = "select max(tc.entrada_not) as max ,tc.entrada_not,tf.nome_for,tc.num_not from tab_fornecedores as tf,"
    strSQL = strSQL & "tab_cmp as tc, tab_itcmp as ti "
    strSQL = strSQL & " where ti.num_not=tc.num_not and ti.serie_not=tc.serie_not "
    strSQL = strSQL & " and tc.cod_for=ti.cod_for and tc.cod_for=tf.cod_for "
    strSQL = strSQL & "  and ti.cod_pro='" & Me.Tag & "' group by ti.cod_pro   "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        strSQL = ""
        Do While Not .EOF
            
            lblUltComp.caption = "Nota: " & !Num_Not & " Entrada Em: " & !entrada_not & Chr(13) & "Fornecedor: " & !nome_for
            
            .MoveNext
        Loop
        .Close
    End With
    
trata_erro:
    Exit Sub
End Sub

