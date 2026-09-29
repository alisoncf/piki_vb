VERSION 5.00
Begin VB.Form frmTesteNFE 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInserirNF 
      Caption         =   "InserirNF e Item"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveXml 
      Caption         =   "Salvar XML"
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtEndXml 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   6600
      Width           =   6255
   End
   Begin VB.TextBox txtRetorno 
      Height          =   5295
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   9735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Testa nova NFc-e"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Endereço do XML"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   6360
      Width           =   1290
   End
End
Attribute VB_Name = "frmTesteNFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInserirNF_Click()
    Dim strSQL As String
    Dim sSeq As String
    Dim iItem As String
    
    
    sSeq = NFe.NFE_GravaNFE("", "", "GO", "00000000", "Venda", "1", "65", "1", "1", "2017-01-01T03:00:00 -03:00", _
    "2017-01-01T03:00:00 -03:00", "0", operacao_interna, strCodigoMunicipioIBGE, danf_nfc_e, _
    Normal, "0", homologacao, Nfe_normal, consumidor_final, Operacao_presencial, Aplicativo_contribuinte, "2.0", _
    "", "")
    
    
    If Not GravaNFe_Dest(sSeq, "", "86075314172", "", "Alison") Then
        MsgBox "erro"
    End If
    
    If Not GravaNFe_Emit(sSeq, F_SoNumero(strCNPJ_Emp), "", strNome_Emp, strFantasia_Emp) Then
        MsgBox "erro"
    End If
    If Not GravaNfe_enderdest(sSeq, "rua", "2", "s/n", "Jundiai", strCodigoMunicipioIBGE, "ANAPOLIS", "GO", _
        "75115443", strCodigoPaisIBGE, "BRASIL", "6233137671", "102345671", "", "", "", "a@a.com") Then
        MsgBox "erro "
    End If
    If Not GravaNfe_enderEmit(sSeq, "rua", "2", "s/n", "Jundiai", strCodigoMunicipioIBGE, "ANAPOLIS", "GO", _
        "75115443", strCodigoPaisIBGE, "BRASIL", "6233137671", "102345671", "", "", "", SIMPLES_NACIONAL) Then
        MsgBox "erro "
    End If
    
    
    iItem = NFe.GravaNFe_Item(sSeq, "1", "", "meu produto de teste", "01234567", "", "", "5102", _
    "UN", "1", "125", "125", "", "UN", "1", "125", "0", "0", "0", "0", vProd_Sim_copoe_total_nota, "", "")
    If Not GravaNfe_item_imposto(iItem, sSeq, (125 * crcTxFederativa_Par / 100), "ICMSSN102", "PISOutr", "COFINSOutr", "") Then
        MsgBox "erro"
    End If

    


End Sub

Private Sub cmdSaveXml_Click()
    If AbreArquivoExterno(txtEndXml.Text, 1) Then
        
        Print #1, Trim(StrConv(txtRetorno.Text, vbUnicode))
        
        
        Close #1
        If MsgBox("Arquivo salvo com sucesso, deseja abrir", vbYesNo + vbQuestion) = vbYes Then
            Shell "explorer " & txtEndXml.Text
        End If
        
        
    End If
    
End Sub

Private Sub Command1_Click()
    Dim s As String
    Dim c As New Collection
    
    Dim p As prod_nfe.TProd
    
    
    
    p = prod_nfe.setProduto("1", "", "produto text", "000000", "5102", "UN", "1.0000", "246.0000", "1.0000", "246.0000", _
        "", "UN", "1.0000", "246,00", "0.0000", vProd_Sim_copoe_total_nota)
    
    
    
    'ZeraListaProduto
    'AddListaProduto p, ListaProduto
    
    s = GeraNovaXmlNFc_E(Nfe_normal, interna, consumidor_final, Operacao_presencial, avista, "Venda de Mercadorias", _
        saida, modelo_Nfe.NFCe_65)
        
    txtRetorno.Text = s
    
    
    txtEndXml.Text = "c:\NFe" & Mid(s, 99, 43) & ".xml"
    
End Sub
