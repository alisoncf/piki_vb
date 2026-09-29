VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmImportrNFe 
   Caption         =   "Importar NFe"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13665
   Icon            =   "frmImportrNFe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCsv 
      Caption         =   "Gerar Planilha"
      Height          =   495
      Left            =   12120
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "3 >> Importar"
      Height          =   495
      Left            =   8040
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   6015
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   10610
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdAnalisar 
      Caption         =   "2 >> Analisar"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.FileListBox flXML 
      Height          =   5745
      Left            =   360
      Pattern         =   "*.xml"
      TabIndex        =   1
      Top             =   1080
      Width           =   5055
   End
   Begin VB.CommandButton cmdCaminhoXML 
      Caption         =   "1 >> Localizar XMLs"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "frmImportrNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnalisar_Click()
    Analisa
End Sub

Private Sub cmdCaminhoXML_Click()
    flXML.path = BuscaNomeArquivoExternoAbrirPasta
     
End Sub

Sub Analisa(Optional analisar As Boolean = True)
    Dim i As Integer
    Dim path As String
    Dim ret As String
    AjustaGrid
    For i = 0 To flXML.ListCount - 1
        flXML.ListIndex = i
        path = flXML.path
        ret = ImportaXML(flXML.FileName, analisar)
        
        If ret <> "" Then
            grid.col = 1
            grid.Text = "Não"
            ColoreLinhaGrid grid, grid.row, vbRed
        Else
            grid.col = 1
            grid.Text = "Sim"
            
            ColoreLinhaGrid grid, grid.row, vbGreen
        End If
        If ret = "" Then
            ret = "Autorizado"
        End If
        grid.col = 10: grid.Text = ret
        
    Next
End Sub
Function ImportaXML(caminho As String, Optional analisar As Boolean = True) As String
    Dim X As XMLSped
    Dim critica As String
    Dim idNfe As String
    X = Le_XML(caminho)
    
    With grid
        grid.AddItem grid.rows
    
        grid.row = grid.rows - 1
        
                
        .col = 2: .Text = X.cab.e.CNPJ
        .col = 3: .Text = X.cab.nNF
        .col = 4: .Text = X.cab.serie
        .col = 5: .Text = X.cab.id
        .col = 6: .Text = X.cab.d.xNome
        .col = 7: .Text = FormatValNFe(X.tot.vNF)
        .col = 8: .Text = X.infProt.dhRecbto
        .col = 9: .Text = X.infProt.cStat
    End With
    
    If X.cab.e.CNPJ <> strCNPJ_Emp Then
         critica = "Erro: CNPJ da NFe (" & X.cab.e.CNPJ & ") é diferente do emitente cadastrado (" & strCNPJ_Emp & ")"
         
    End If
    If InStr(1, LCase(X.cab.natOp), "venda") = 0 Then
        critica = critica + "Erro: natureza divergente - " & X.cab.natOp
    End If
    If F_BuscaUmCampo("id", "nfe", "chave='" & F_SoNumero(X.cab.id) & "'") <> "" Then
        critica = critica + "Erro: NFe já existe na base"
    End If
    If X.infProt.cStat <> "100" Then
        critica = critica + "Erro: Status diferente de 100 (" & X.infProt.cStat & " " & X.infProt.xMotivo
    End If
    
    If analisar = True Then
        If critica <> "" Then
            ImportaXML = critica
            Exit Function
        End If
        Exit Function
    Else
        If critica <> "" Then
            ImportaXML = critica
            Exit Function
        End If
    End If
    
    idNfe = NFE_GravaNFE("", F_SoNumero(X.cab.id), X.cab.cUF, X.cab.cNF, X.cab.natOp, X.cab.indPag, X.cab.mod, X.cab.serie, _
        X.cab.nNF, X.cab.dEmi, X.cab.dEmi, Int(X.cab.tpNf), X.cab.idDst, X.cab.cMunFG, Int(X.cab.tpImp), Int(X.cab.tpEmis), _
        X.cab.cDV, Int(X.cab.tpAmb), Int(X.cab.finNFe), X.cab.indFnl, X.cab.indPr, Int(X.cab.procEmi), X.cab.verProc, X.cab.dhCont, X.cab.xJust, "2")
    Dim sId_NFe As String
    sId_NFe = idNfe
    
    
    If idNfe = "" Then
        critica = "Erro: erro na geração da NFe"
        Exit Function
    End If
        
        
    If Not GravaEmitenteComEndereco(idNfe) Then
        critica = "Erro: erro na gravação do emitente"
        Exit Function
    End If
        
        
    If X.cab.d.xNome <> "" Then
        
        If Not NFe.GravaNFe_Dest(sId_NFe, X.cab.d.CNPJ, X.cab.d.cpf, "", X.cab.d.xNome) Then
            'MsgBox "Erro"
            Exit Function
        End If
            
        If Not NFe.GravaNfe_enderdest(sId_NFe, X.cab.d.ende.xLgr, X.cab.d.ende.nro, X.cab.d.ende.xCpl, X.cab.d.ende.xBairro, _
                X.cab.d.ende.cMun, X.cab.d.ende.xMun, X.cab.d.ende.UF, X.cab.d.ende.cep, strCodigoPaisIBGE, strPais, X.cab.d.ende.fone, X.cab.d.ende.IE, _
                  X.cab.d.ende.IndIEDest, "", "", X.cab.d.ende.email) Then
        
                MsgBox "Erro gravando endereço do Destinatário"
        End If
    End If
    
        Dim i As Integer
        Dim tipo_icms As String
        Dim tipo_pis As String
        Dim tipo_ipi As String
        Dim sId_Item As String
        For i = 0 To UBound(X.p) - 1
        
            sId_Item = NFe.GravaNFe_Item(sId_NFe, Format(X.p(i).cProd, "0000000#"), X.p(i).cEAN, X.p(i).xProd, X.p(i).ncm, "", "", X.p(i).CFOP, _
                X.p(i).uCom, FormatValNFe(X.p(i).qCom), FormatValNFe(X.p(i).vUnCom), FormatValNFe(X.p(i).vProd), X.p(i).cEAN, X.p(i).uCom, FormatValNFe(X.p(i).qCom), _
                        FormatValNFe(X.p(i).vUnTrib), "0", "0", FormatValNFe(X.p(i).vOutro), "0", vProd_Sim_copoe_total_nota, "", "")
                    
                    If sId_Item = "" Then
                        MsgBox "Erro na gravação do item " & "", vbExclamation
                        Exit Function
                    End If
                    
                    tipo_icms = X.p(i).icms_.cst
                    tipo_ipi = X.p(i).ipi_.cst
                    tipo_ipi = X.p(i).pis_.cst
                    If NFe.GravaNfe_item_imposto(sId_Item, sId_NFe, "0", _
                          CStr(tipo_icms), CStr(tipo_pis), "", tipo_ipi) Then
                        
                        If tipo_icms = "ICMSSN102" Or tipo_icms = "102" Then
                    
                            If Not NFe.GravaNfe_item_imposto_icmssn102(sId_Item, sId_NFe, "", TipoICMS.ICMSSN102) Then
                                MsgBox "erro na gravação do item_Imposto_icmssn102"
                            
                            End If
                        End If
                        
                        
            
            End If
        Next
        
            
            'If Not GravaNfe_IcmsTot(sID_Nfe, "0", "0", "0", "0", "0", CStr(cTotalProd), "0", _
            '            "0", CStr(cTotalDesc), "0", "0", "0", "0", "0", _
            '            CStr((cTotalProd - cTotalDesc)), CStr((cTotalProd - cTotalDesc) * crcTxFederativa_Par / 100)) Then
            'End If
                        
            
        
        
        
        
        
            
        
    
End Function
Sub AjustaGrid()
    With grid
        .rows = 2
        .Cols = 11
        .FixedRows = 1
        .rows = 1
        .col = 0: .Text = "...": .ColWidth(0) = 500
        .col = 1: .Text = "Apta?": .ColWidth(1) = 600
        .col = 2: .Text = "CNPJ"
        .col = 3: .Text = "nNF"
        .col = 4: .Text = "Serie"
        .col = 5: .Text = "Chave"
        .col = 6: .Text = "Destinatario": .ColWidth(6) = 1000
        .col = 7: .Text = "Valor"
        .col = 8: .Text = "dhEmi"
        .col = 9: .Text = "xStatus"
        .col = 10: .Text = "Crítica": .ColWidth(10) = 2500
        
        
        
        
        
    End With
End Sub

Private Sub cmdCsv_Click()
    ExportaMSFLXGridToCSV grid
End Sub

Private Sub cmdImportar_Click()
    Analisa False
End Sub

Private Sub Form_Load()
    AjustaGrid
End Sub
