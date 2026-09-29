VERSION 5.00
Begin VB.Form frmEdiNfeDevolucoes 
   Caption         =   "Manifestos e Devoluções"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18690
   Icon            =   "frmEdiNFeDevolucoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   18690
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   18660
      TabIndex        =   36
      Top             =   7020
      Width           =   18690
      Begin VB.TextBox txtEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEDFD5&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscarNfeExterna 
         Caption         =   "Buscar Externas"
         Height          =   495
         Left            =   16440
         TabIndex        =   38
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdDuplicidade 
         Caption         =   "Duplicidade"
         Height          =   495
         Left            =   17520
         TabIndex        =   37
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtUF 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   12720
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtVenda 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11760
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbxSerie 
         Height          =   315
         ItemData        =   "frmEdiNFeDevolucoes.frx":1CFA
         Left            =   7920
         List            =   "frmEdiNFeDevolucoes.frx":1D07
         TabIndex        =   12
         Text            =   "cmbxSerie"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbxModelo 
         Height          =   315
         ItemData        =   "frmEdiNFeDevolucoes.frx":1D2B
         Left            =   6480
         List            =   "frmEdiNFeDevolucoes.frx":1D38
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         Height          =   495
         Left            =   13440
         Picture         =   "frmEdiNFeDevolucoes.frx":1D5C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         ItemData        =   "frmEdiNFeDevolucoes.frx":20AC
         Left            =   4560
         List            =   "frmEdiNFeDevolucoes.frx":20BC
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtChave 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8880
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10320
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdEmpresa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4200
         Picture         =   "frmEdiNFeDevolucoes.frx":20F7
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   0
         Width           =   660
      End
      Begin VB.Label lblTotal 
         Caption         =   "Total: 0"
         Height          =   495
         Left            =   14880
         TabIndex        =   39
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         Height          =   195
         Left            =   12720
         TabIndex        =   19
         Top             =   0
         Width           =   210
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda:"
         Height          =   195
         Left            =   11760
         TabIndex        =   17
         Top             =   0
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         Height          =   195
         Left            =   7920
         TabIndex        =   11
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo:"
         Height          =   195
         Left            =   6480
         TabIndex        =   9
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   4560
         TabIndex        =   7
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   195
         Left            =   8880
         TabIndex        =   13
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destinatário"
         Height          =   195
         Left            =   10320
         TabIndex        =   15
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   18690
      TabIndex        =   35
      Top             =   7755
      Width           =   18690
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   735
         Left            =   15960
         Picture         =   "frmEdiNFeDevolucoes.frx":24A2
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "E-mail"
         Height          =   735
         Left            =   5880
         Picture         =   "frmEdiNFeDevolucoes.frx":28D1
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Editar"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiNFeDevolucoes.frx":300C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdInutilizar 
         Caption         =   "Inutilizar"
         Height          =   735
         Left            =   14520
         Picture         =   "frmEdiNFeDevolucoes.frx":3548
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   735
         Left            =   11640
         Picture         =   "frmEdiNFeDevolucoes.frx":4567
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   17400
         Picture         =   "frmEdiNFeDevolucoes.frx":4A3B
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdPDF 
         Caption         =   "DANFE PDF"
         Height          =   735
         Left            =   7320
         Picture         =   "frmEdiNFeDevolucoes.frx":4C85
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdAtlzRetorno 
         Caption         =   "Atualizar Retorno"
         Height          =   735
         Left            =   13080
         Picture         =   "frmEdiNFeDevolucoes.frx":5133
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar NFe"
         Height          =   735
         Left            =   3000
         Picture         =   "frmEdiNFeDevolucoes.frx":563B
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "DANFE Ext"
         Height          =   735
         Left            =   8760
         Picture         =   "frmEdiNFeDevolucoes.frx":5ADD
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar NFe"
         Height          =   735
         Left            =   10200
         Picture         =   "frmEdiNFeDevolucoes.frx":5F8B
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdCupom 
         Caption         =   "Cupom"
         Height          =   735
         Left            =   4440
         Picture         =   "frmEdiNFeDevolucoes.frx":6519
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdRetransm 
         Caption         =   "Retransmitir"
         Height          =   735
         Left            =   1560
         Picture         =   "frmEdiNFeDevolucoes.frx":69C7
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnCopiarCaminho 
         Caption         =   "Copiar Caminho XML"
      End
      Begin VB.Menu mnEmail 
         Caption         =   "Enviar Email"
      End
      Begin VB.Menu mnCSV 
         Caption         =   "Exportar Tabela para CSV"
      End
   End
End
Attribute VB_Name = "frmEdiNfeDevolucoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Sub soma()
    On Error GoTo trata_Erro
    Dim c As Currency
    grade.Visible = False
    With ado
        Do While Not .EOF
            c = c + !vNF
            .MoveNext
        Loop
        .MoveFirst
    End With
    grade.Visible = True
    lblTotal.caption = "Total.: " & FormatVal(c)
    Exit Sub
trata_Erro:
grade.Visible = True
End Sub
Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    If ado!modelo = "55" Then
        cmdCupom.caption = "Gerar PDF"
    Else
        cmdCupom.caption = "Gerar Cupom"
    End If
End Sub

Private Sub cmbxModelo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmdAtlzRetorno_Click()
    Dim s As String
    Dim chNFe As String
    Dim DigestValue As String
    Dim autorizacao As String
    Dim sRecibo As String
    Dim Inf As infProt
    Dim venda As String
    Dim strSQL As String
    Dim Chave As String
    Dim caminho As String
    venda = BuscaVendaPeloID(ado!id)
    
    s = XmlRetornoNFe(ado!Chave)
    caminho = RetornoCaminhoNFeXML(ado!Chave)
    If s = "" Then
        If MsgBox("Arquivo não encontrado!, Deseja localizar manualmente?", vbYesNo + vbQuestion) = vbYes Then
            mdiFrm.cmndlg.filter = "*.xml"
            mdiFrm.cmndlg.ShowOpen
            
            s = mdiFrm.cmndlg.FileName
            s = XmlRetornoNFeCaminho(s)
            If s = "" Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    DigestValue = RetornaValorTagXML(s, "<DigestValue>")
    Inf = RetornaInfProt(s)
    If Inf.cStat = "100" Then
        
        
        MsgBox Inf.xMotivo, vbInformation
            
        If Inf.chNFe <> ado!Chave Then
            If MsgBox("As chaves da notas são diferentes. Deseja mesmo alterar para esta chave?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
            strSQL = "update nfe set chave='" & Inf.chNFe & "' where id=" & ado!id
            Conn.Execute strSQL
        End If
            
        
        AlteraDadosFinaisNFe ado!id, "2", mudarCaractere(RetornoCaminhoNFePeloID(ado!id), "\", "/"), Inf.nProt, Inf.digVal, DigestValue
        CopiaXML caminho, CaminhoXMLCopia
        If ado!status = "1" Then
        
            AlteraStatusVen venda, "2"
            P_BaixaSaidaNFe ado!id
            
        End If
        ado.Requery
    Else
        If Inf.cStat = "101" Then
            AlteraDadosFinaisNFe ado!id, "5", mudarCaractere(RetornoCaminhoNFePeloID(ado!id), "\", "/"), Inf.nProt, Inf.digVal, DigestValue
            AlteraStatusVen venda, "5"
            MsgBox Inf.xMotivo, vbInformation
        End If
        
    End If
    
    RefreshAdoPosition ado
    
    
    'MsgBox s
End Sub

Private Sub cmdBuscarNfeExterna_Click()
        '    frmImportrNFe.Show 1
End Sub

Private Sub cmdCancelar_Click()
    Dim just As String
    Dim s As String
    On Error GoTo trata_Erro
    
    If ado!status <> "2" Then
        MsgBox "A NFe não pode ser cancelada!", vbExclamation
        Exit Sub
    End If
    
    
    just = InputBox("Motivo do cancelamento")
    If just = "" Then
        Exit Sub
    End If
    
    If MsgBox("Confirma cancelamento da NFe?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    
        
    
    s = AcBrMonitor.CancelarNFe(ado!Chave, just, "", "")
    
    MsgBox s
    Exit Sub
trata_Erro:
    
    
    
End Sub

Private Sub cmdConsultar_Click()
    Dim s As String
    Dim ret As String
    
    InicializaEmpresa ado!cod_emp
     
     
    s = ConsultarNFe(ado!Chave, Format(Left(ado!dhEmi, 10), "YYYYMM"))
    
    If ado!status = "1" Then
    ret = F_SoNumero(RetornaValorTagArquivo(s, "CStat=", 4))
        If ret = "100" Then
            AlteraStatusNfe ado!id, "2"
            ado.Requery
        End If
    
    End If
    
    MsgBox s
    
    
    
    
    
End Sub


Private Sub cmdCupom_Click()
    Dim s As String
    Dim p As String
    Dim via_cartao As String
    Dim num_ven As String
    Dim modelo As String
    Dim strSQL As String
    Dim i As Integer
    On Error GoTo trata_Erro
        
        
    If ado!modelo = "55" Then
        
        
        
        s = RetornoConteudoNFeXML(ado!Chave, Format(Left(ado!dhEmi, 10), "YYYYMM"))
        p = bioPDF.ApontarPrinterToPDF(ado!Chave & ".pdf")
        i = ImprimirXMLNFeDANFE(s, Printer)
        Printer.EndDoc
        Exit Sub
    End If
        
    If ado!modelo <> "65" Then
        MsgBox "Esta NFe não é um Cupom Fiscal Eletrônico modelo 65", vbExclamation
        Exit Sub
    End If
        
        
    via_cartao = F_BuscaUmCampo("mensagem", "tab_venfin", "num_ven='" & ado!num_ven & "'")
    
    
    
    s = RetornoConteudoNFeXML(ado!Chave, Format(Left(ado!dhEmi, 10), "YYYYMM"))
    
    
     
    
    If s <> "" Then
        ImprimirXMLNFeCUPOM s, frmVisualiza.Pic, True, via_cartao
        frmVisualiza.Show 1
        If frmVisualiza.tag <> "" Then
            ImprimirXMLNFeCUPOM s, Printer, True, via_cartao
        End If
        Unload frmVisualiza
    Else
        MsgBox "XML não encontrado. Essa nota não foi gerada.", vbExclamation
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmdDuplicidade_Click()
    If ado!status <> "1" Then
        Exit Sub
    End If
    If MsgBox("Confirma nota em duplicidade? Se confirmada ela muda o status pra enviada. Lembre-se de buscar a NFe e guardar uma copia junta com as demais", vbYesNo + vbQuestion) = vbYes Then
        NFe.AlteraStatusNfe ado!id, "2"
    End If
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo trata_Erro
    frmNFE_Detalhe.tag = ado!id
     frmNFE_Detalhe.Show 1
    
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmdEmail_Click()
    Dim s As String
    's = InputBox("informe o email", "Email", BuscaEmailClientePeloId(ado!id))
    
    EnviarDanfeXmlPorEmail ado!Chave
    
    
    
    'AcBrMonitor.EnviaEmailNFe s, RetornoCaminhoNFePeloID(ado!id)
End Sub

'Private Sub cmdEnviar_Click()
'    Dim strSQL As String
'
'    strSQL = TransmitirNFcE(ado!id, "")
'
'    If InStr(1, "CStat=100", strSQL) > 0 Then
'        AlteraStatusNfe ado!id, "2"
'        'MsgBox strSQL
'    Else
'       ' MsgBox strSQL
'    End If
'
'
'
'
'
'
'
'End Sub

Private Sub cmdExcluir_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    
    With ado
        
        
        If !status <> "1" Then
            MsgBox "Não pode ser excluída!", vbExclamation
            Exit Sub
        End If
        If !tpEmis <> "1" Then
            MsgBox "Não pode ser excluída! Está lançada em contingência", vbExclamation
            Exit Sub
        End If
        If MsgBox("Confirma exclusão da NFe?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
                
                
        strSQL = "delete from nfe where id=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_item where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_item_comb where id_nfe=" & ado!id
        Conn.Execute strSQL
        
        strSQL = "delete from nfe_dest where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_emit where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_enderdest where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_enderemit where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_dest where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_icmstot where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_infadic where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_dest where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_item_imposto where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_dest where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_pag where id_nfe=" & ado!id
        Conn.Execute strSQL
        strSQL = "delete from nfe_transp where id_nfe=" & ado!id
        Conn.Execute strSQL
        ado.Requery
        cmdPesquisar_Click
    End With
    Exit Sub
trata_Erro:
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV grade, ado
End Sub

Private Sub cmdImprimir_Click()
    Dim s As String
    s = RetornoConteudoNFeXML(ado!Chave, Format(Left(ado!dhEmi, 10), "YYYYMM"))
    
    If s <> "" Then
        ImprimirDanfe s
    Else
        MsgBox "XML não encontrado.", vbExclamation
    End If
    
    
    
    
    
    
End Sub

Private Sub cmdInutilizar_Click()
    frmInutilizarNFe.Show 1
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub

Private Sub cmdXml_Click()
    GeraXML_NFcE_Dados ado!id
End Sub

Private Sub cmdPDF_Click()
    Dim s As String
    Dim b As Boolean
    Dim modelo As String
    Dim caminho As String
    On Error GoTo trata_Erro
        
    BuscaPDFDanfe ado!id
    
    
    Exit Sub
trata_Erro:
    
    
        
    
    
End Sub


Private Sub cmdReimp_Click()
    'ReimportarItens ado!num_ven
End Sub

Private Sub cmdRetransm_Click()
    On Error GoTo trata_Erro
    If ado!status = "1" Then
        TransmitirNFcE ado!id, ""

    Else
        MsgBox "O status dessa NFe não permite retransmissão!", vbExclamation
    End If
    Exit Sub
trata_Erro:
P_MostraErro
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim s As String
    Dim via_cartao As String
    Dim num_ven As String
    Dim modelo As String
    Dim strSQL As String
    On Error GoTo trata_Erro
        
    If ado!modelo <> "55" Then
        MsgBox "Esta NFe não é um DANFE modelo 55", vbExclamation
        Exit Sub
    End If
        
    If ado!status <> "2" Then
        MsgBox "Esta NFe não está fechada.", vbExclamation
        Exit Sub
    End If
        
    'via_cartao = F_BuscaUmCampo("mensagem", "tab_venfin", "num_ven='" & ado!num_ven & "'")
    
    
    
    s = RetornoConteudoNFeXML(ado!Chave, Format(Left(ado!dhEmi, 10), "YYYYMM"))
    
    
     
    Dim iPrinter As Integer
    If s <> "" Then
        Load frmVisualizadorImpressao
        ImprimirXMLNFeDANFE s, frmVisualizadorImpressao.pcImpressao, False, False
        frmVisualizadorImpressao.addNewPage
        P_Status ""
        frmVisualizadorImpressao.Show 1
        If frmVisualizadorImpressao.tag <> "" Then
            iPrinter = BuscaImpressoraPadrao
            Set Printer = Printers(frmVisualizadorImpressao.iPrinter)
            'ImprimirXMLNFeCUPOM s, Printer, True, via_cartao
            
            ImprimirXMLNFeDANFE s, Printer, False, False
            
            Set Printer = Printers(iPrinter)
            
        End If
        Unload frmVisualizadorImpressao
    End If
    Exit Sub
trata_Erro:
    P_MostraErro

End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    'busca
    bActive = True
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    txtDe.text = Date
    txtAte.text = Date
    P_PreencheSerie cmbxSerie
    resizeForm Me
    
    If strCod_Cai <> "" Then
        cmbxSerie.text = Int(strCod_Cai)
    End If
    
End Sub
Private Sub cmdEmpresa_Click()
    buscaEmpresa
End Sub
Sub buscaEmpresa()
    Dim e As sped_empresa.sped_empresa
    
    e = sped_empresa.SelecionarSped_Empresa(txtEmpresa.text)
    If e.codigo <> "" Then
        txtEmpresa.text = e.nome
        txtEmpresa.tag = e.codigo
        busca
    End If
        
    
    
End Sub
Sub busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    P_Status "Aguarde..."
    strSQL = "select distinct n.*, d.xNome, t.vNF from nfe as n inner join nfe_dest as d on n.id=d.id_nfe "
    strSQL = strSQL & " inner join nfe_icmstot as t on n.id=t.id_nfe  "
    
    strSQL = strSQL & " where 1 > 0 "
    If txtDe.text <> "" Then
        strSQL = strSQL & " and left(dhemi,10)>=" & F_GravaData(txtDe.text)
    End If
    If txtAte.text <> "" Then
        strSQL = strSQL & " and left(dhemi,10)<=" & F_GravaData(txtAte.text)
    End If
    If cmbStatus.text <> "" Then
        strSQL = strSQL & " and status=" & Left(cmbStatus.text, 1)
    End If
    If txtChave.text <> "" Then
        strSQL = strSQL & " and (nnf = '" & txtChave.text & "'"
        strSQL = strSQL & " or chave = '" & txtChave.text & "')"
    End If
    If txtVenda.text <> "" Then
        strSQL = strSQL & " and num_ven ='" & txtVenda.text & "'"
    End If
    If txtUF.text <> "" Then
        strSQL = strSQL & " and cUF ='" & txtUF.text & "'"
    End If
    If txtEmpresa.tag <> "" Then
        strSQL = strSQL & " and cod_emp ='" & txtEmpresa.tag & "'"
    End If
    
    If txtCliente.text <> "" Then
        strSQL = strSQL & " and (d.xNome like '%" & txtCliente.text & "%' or d.CPF='" & txtCliente.text & "')"
    End If
    If cmbxModelo.text <> "" Then
        strSQL = strSQL & " and modelo='" & Left(cmbxModelo.text, 2) & "'"
    End If
    If cmbxSerie.text <> "" Then
        strSQL = strSQL & " and serie='" & Int(cmbxSerie.text) & "' "
    End If
    
    
    strSQL = strSQL & " order by nnf desc "
    abreObjetoAdo ado, strSQL, grade
    grade.caption = ado.RecordCount & " registros encontrados..."
    P_Status ""
    
    soma
    Exit Sub
trata_Erro:
    P_Status ""
    P_MostraErro
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade grade, Me, pcBusca
End Sub

Private Sub grade_HeadClick(ByVal ColIndex As Integer)
    OrdenaGrid grade, ado
End Sub

Private Sub grade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnMenu
    End If
End Sub

Private Sub mnCopiarCaminho_Click()
    On Error Resume Next
    Dim s As String
    s = ado!caminhoXML
    s = mudarCaractere(s, "/", "\")
    Clipboard.Clear
    Clipboard.SetText s
    
End Sub

Private Sub mnCSV_Click()
    On Error GoTo trata_Erro

    ExportaGridToCSV grade, ado


Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub mnEmail_Click()
    
    'EnviaEmailNFe ado!id
    cmdEmail_Click
End Sub

Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtChave_GotFocus()
    P_SelecTexto txtChave
End Sub

Private Sub txtChave_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtCliente_GotFocus()
    P_SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub

Private Sub txtEmpresa_GotFocus()
    P_SelecTexto txtEmpresa
End Sub

Private Sub txtEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaEmpresa
    End If
    
End Sub

Private Sub txtUF_GotFocus()
    P_SelecTexto txtUF
End Sub

Private Sub txtUF_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtUF_LostFocus()
    P_AjustaTexto txtUF
End Sub

Private Sub txtVenda_GotFocus()
    P_SelecTexto txtVenda
End Sub

Private Sub txtVenda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtVenda_LostFocus()
    P_AjustaNota txtVenda, 7
End Sub
