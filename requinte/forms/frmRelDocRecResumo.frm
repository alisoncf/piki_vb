VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelDocRecResumo 
   Caption         =   "Resumo de Documentos a Receber"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16185
   Icon            =   "frmRelDocRecResumo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   16185
   Begin VB.PictureBox pcbot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   16155
      TabIndex        =   42
      Top             =   7725
      Width           =   16185
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   14280
         Picture         =   "frmRelDocRecResumo.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   12360
         Picture         =   "frmRelDocRecResumo.frx":06F5
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   6135
      Left            =   0
      TabIndex        =   41
      Top             =   1440
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   10821
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   0
      ScaleHeight     =   1320
      ScaleWidth      =   16185
      TabIndex        =   0
      Top             =   0
      Width           =   16185
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Height          =   600
         Left            =   14520
         Picture         =   "frmRelDocRecResumo.frx":0B02
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   60
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmRelDocRecResumo.frx":1504
         Left            =   14520
         List            =   "frmRelDocRecResumo.frx":1514
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmRelDocRecResumo.frx":1540
         Left            =   2400
         List            =   "frmRelDocRecResumo.frx":154D
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtbxContaCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxTerceiro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   10
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtbxTipoCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxSeqDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   60
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxSeqItd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   60
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cmbxEspDoc 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmRelDocRecResumo.frx":157F
         Left            =   10560
         List            =   "frmRelDocRecResumo.frx":158F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtbxRepr 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   12480
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxHist 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   60
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxGrupo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9360
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtbxOrgao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11520
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton cmndbtnTipoCob 
         Height          =   285
         Left            =   7920
         Picture         =   "frmRelDocRecResumo.frx":15B0
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnCob 
         Height          =   285
         Left            =   10200
         Picture         =   "frmRelDocRecResumo.frx":193A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnTerc 
         Height          =   285
         Left            =   8880
         Picture         =   "frmRelDocRecResumo.frx":1CC4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnGru 
         Height          =   285
         Left            =   11160
         Picture         =   "frmRelDocRecResumo.frx":204E
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnRep 
         Enabled         =   0   'False
         Height          =   285
         Left            =   14040
         Picture         =   "frmRelDocRecResumo.frx":23D8
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnOrg 
         Height          =   285
         Left            =   14040
         Picture         =   "frmRelDocRecResumo.frx":2762
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Histórico:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Tag             =   "!"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   38
         Tag             =   "!"
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Doc."
         Height          =   195
         Left            =   2160
         TabIndex        =   37
         Tag             =   "!"
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar Por:"
         Height          =   195
         Left            =   14520
         TabIndex        =   36
         Tag             =   "!"
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Terceiro:"
         Height          =   195
         Left            =   2400
         TabIndex        =   35
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cobrança:"
         Height          =   195
         Left            =   6000
         TabIndex        =   34
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label lblTipoCob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Cobrança:"
         Height          =   195
         Left            =   8280
         TabIndex        =   33
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label lblTerCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         Height          =   195
         Left            =   4680
         TabIndex        =   32
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. Doc:"
         Height          =   195
         Left            =   4320
         TabIndex        =   31
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. Parc:"
         Height          =   195
         Left            =   5160
         TabIndex        =   30
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblCPF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         Height          =   195
         Left            =   8160
         TabIndex        =   29
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espécie de Doc:"
         Height          =   195
         Left            =   10560
         TabIndex        =   28
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label lblRepr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Representante:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   12480
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   3360
         TabIndex        =   26
         Tag             =   "!"
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lblGru 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Clientes:"
         Height          =   195
         Left            =   9360
         TabIndex        =   25
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label lblOrgao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Órgão:"
         Height          =   195
         Left            =   11520
         TabIndex        =   24
         Top             =   600
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmRelDocRecResumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub chkAberto_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub chkProrr_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_Click()
    If Not bActive Then Exit Sub
    If cmbxBuscar.ListIndex = 3 Then
        dbgrd.Columns(6).caption = "Boletos"
        dbgrd.Columns(6).Visible = True
        dbgrd.Columns(6).DataField = "num_bol"
    ElseIf cmbxBuscar.ListIndex = 2 Then
        txtbxExpr.Text = "0"
    Else
        dbgrd.Columns(6).Visible = False
    End If
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxEspDoc_Click()
    If Left(cmbxEspDoc.Text, 1) = "A" Then
        cmndbtnRep.Enabled = True
        txtbxRepr.Enabled = True
        lblRepr.Enabled = True
        txtbxRepr.SetFocus
    Else
        cmndbtnRep.Enabled = False
        txtbxRepr.Enabled = False
        lblRepr.Enabled = False
    End If
End Sub

Private Sub cmbxOrder_Click()
    If Not bActive Then Exit Sub
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    If cmbxTipo.ListIndex = 0 Then
        cmbxTipoTer.ListIndex = 0
        cmbxTipo.Enabled = True
    Else
        cmbxTipo.ListIndex = -1
        cmbxTipo.Enabled = False
    End If
    If Not bActive Then Exit Sub
    
    
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub


Private Sub cmndbtnBaixar_Click2()
    Dim cJuros As Currency
    Dim iCount As Integer
    Dim jCount As Integer
    On Error GoTo trata_Erro
    With dtNot.Recordset
        cJuros = F_CalculaJuros(!venc_itd, Date, !reman_itd)
        
        frmBaixaDocReq.tag = !cod_ter
        
                
        frmBaixaDocReq.txtbxCliente.Text = !nome
        frmBaixaDocReq.txtbxCliente.tag = !Codigo
        frmBaixaDocReq.txtbxJuros.Text = FormatVal(cJuros)
        frmBaixaDocReq.txtbxReman.Text = frmListaDoc.lblTotal.caption
        frmBaixaDocReq.txtbxValor_Par.Text = frmListaDoc.lblTotal.caption
        'frmBaixaDocReq.lblParc.caption = "Parcelas: " & frmListaDoc.msflxgrd.rows
        frmBaixaDocReq.txtbxTotalBaixa.Text = "0,00"
        frmBaixaDocReq.cmndbtnCOnfirma.tag = "R"
        frmBaixaDocReq.lblCPF.caption = lblCPF.caption
        frmBaixaDocReq.msflxParcelas.Cols = frmListaDoc.msflxgrd.Cols
        frmBaixaDocReq.msflxParcelas.rows = frmListaDoc.msflxgrd.rows + 1
        For iCount = 0 To frmListaDoc.msflxgrd.rows - 1
            frmBaixaDocReq.msflxParcelas.Row = iCount + 1
            frmListaDoc.msflxgrd.Row = iCount
            For jCount = 0 To frmListaDoc.msflxgrd.Cols - 1
                frmBaixaDocReq.msflxParcelas.Col = jCount
                frmListaDoc.msflxgrd.Col = jCount
                frmBaixaDocReq.msflxParcelas.Text = frmListaDoc.msflxgrd.Text
            Next
            
        Next
        frmBaixaDocReq.Show 1
        Unload frmBaixaDocReq
        Unload frmListaDoc
        dtNot.Refresh
    End With
trata_Erro:
    P_MostraErro
    Exit Sub
End Sub

Private Sub cmbxTipoTer_Click()
    lblTerCap.caption = "&" & cmbxTipoTer.Text
    If Left(cmbxTipoTer.Text, 1) = "C" Then
        txtbxGrupo.Enabled = True
        txtbxOrgao.Enabled = True
        lblOrgao.Enabled = True
        lblGru.Enabled = True
        txtbxGrupo.Enabled = True
        txtbxOrgao.Enabled = True
    Else
        txtbxGrupo.Enabled = False
        txtbxOrgao.Enabled = False
        lblOrgao.Enabled = False
        lblGru.Enabled = False
        txtbxGrupo.Enabled = False
        txtbxOrgao.Enabled = False
    
    End If
End Sub

Private Sub cmndbrnPrejuizo_Click()
    On Erro GoTo trata_Erro
    With dtNot.Recordset
        If MsgBox("Confirma Inclusão do Carnê nos Prejuízos?", vbYesNo + vbQuestion) = vbYes Then
            If !prejuizo_itd = "S" Then
                strSQL = "Update tab_itdoc set prejuizo_itd='N' where seq_itd='" & !seq_itd & "'"
                Banco.Execute strSQL
            Else
                strSQL = "Update tab_itdoc set prejuizo_itd='S' where seq_itd='" & !seq_itd & "'"
                Banco.Execute strSQL
            End If
        End If
        cmndbtnBuscar_Click
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnBaixar_Click()
    On Error GoTo trata_Erro
    With dtNot.Recordset
        frmBaixaDocReq.txtbxCliente.Text = !nome
        frmBaixaDocReq.txtbxCliente.tag = !Codigo
        frmBaixaDocReq.lblCPF.caption = !CPF
        frmBaixaDocReq.Show 1
    End With
trata_Erro:
End Sub
Sub Monta()
    Dim iCount As Integer
    Dim jCount As Integer
    
    
    
    
End Sub

Sub monta_grid1(strSQL As String)
    Dim rPesq As Recordset
    Dim sCod As String
    Dim iCount As Integer
    Dim iiTem As Integer
    Dim iPrimLinha  As Integer
    
    
    ajusta_Grid
    iPrimLinha = 10
    msflxgrd.Visible = False
    msflxgrd.Visible = True
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            If !cod_ter <> sCod Then
                sCod = !cod_ter
                iiTem = iiTem + 1
                strSQL = iiTem & Chr(9) & !seq_itd & Chr(9)
                strSQL = strSQL & !seq_doc & Chr(9)
                
                strSQL = strSQL & !num_doc & Chr(9)
                strSQL = strSQL & !historico_doc & Chr(9)
                strSQL = strSQL & !esp_Doc & Chr(9)
                strSQL = strSQL & !tipo_ter & Chr(9)
                strSQL = strSQL & !nome & Chr(9)
                strSQL = strSQL & !parcela_itd & "/" & !parcelas_doc & Chr(9)
                strSQL = strSQL & FormatVal(!reman_itd) & Chr(9)
                strSQL = strSQL & !venc_itd
                
                msflxgrd.AddItem strSQL
                
            Else
                
                For iCount = 10 To msflxgrd.Cols - 1
                    msflxgrd.Col = iCount
                    msflxgrd.Row = msflxgrd.rows - 1
                    If msflxgrd.Text <> "" Then
                        
                        msflxgrd.Row = msflxgrd.rows - 1
                        If msflxgrd.Col >= msflxgrd.Cols - 1 Then
                            msflxgrd.Cols = msflxgrd.Cols + 1
                            msflxgrd.Col = msflxgrd.Cols - 1
                        Else
                            msflxgrd.Col = msflxgrd.Col + 1
                        End If
                        
                        msflxgrd.Text = !parcela_itd & "/" & !parcelas_doc
                        msflxgrd.Cols = msflxgrd.Cols + 1
                        
                        If msflxgrd.Col >= msflxgrd.Cols - 1 Then
                            msflxgrd.Cols = msflxgrd.Cols + 1
                            msflxgrd.Col = msflxgrd.Cols - 1
                        Else
                            msflxgrd.Col = msflxgrd.Col + 1
                        End If
                        msflxgrd.Text = FormatVal(!reman_itd)
                        If msflxgrd.Col >= msflxgrd.Cols - 1 Then
                            msflxgrd.Cols = msflxgrd.Cols + 1
                            msflxgrd.Col = msflxgrd.Cols - 1
                        Else
                            msflxgrd.Col = msflxgrd.Col + 1
                        End If
                        
                        msflxgrd.Text = !venc_itd
                        msflxgrd.Row = 0
'                        msflxgrd.Col = msflxgrd.Cols - 3
'                        msflxgrd.Text = "Parc."
'                        msflxgrd.Col = msflxgrd.Cols - 2
'                        msflxgrd.Text = "Valor."
'                        msflxgrd.Col = msflxgrd.Cols - 1
'                        msflxgrd.Text = "Venc."
                    Else
                        MsgBox ""
                    End If
                Next
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    
    msflxgrd.Visible = True
    
End Sub
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sOrderby As String
    Dim sCampo As String
    On Error GoTo trata_Erro
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sTipoTer As String
    Dim sOrd As String
    Dim rPesq As Recordset
    Dim cTotalReman As Currency
    Dim cTotalDoc As Currency
    Dim sWhere As String
    Dim sIn As String
    Dim iCount As Integer
    Dim jCount As Integer
    Dim iAcha As Integer
    Dim iiTem As Integer
    ajusta_Grid
    sOrd = "td.seq_doc,ti.seq_itd"
    
    
    
    
    sWhere = sWhere & " td.seq_doc=ti.seq_doc "
    If txtbxDataFin.Text <> "" Then
        sWhere = sWhere & "  and venc_itd<=" & F_GravaData(txtbxDataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        sWhere = sWhere & "  and venc_itd>=" & F_GravaData(txtbxDataIni.Text)
    End If
    
    P_Status "Aguarde... buscando dados..."
    
    sIn = F_NotIn("td.seq_doc", "tab_doc as td,tab_itdoc as ti", sWhere, "'")
    
    
    If Left(cmbxEspDoc.Text, 1) = "A" Then
        strSQL = doc.F_SelectDocsReceberAcerto(Left(cmbxTipoTer.Text, 1), "", "", _
         "", txtbxTipoCob.tag, txtbxContaCob.tag, txtbxNumDoc.Text, txtbxHist.Text, txtbxSeqItd.Text, sIn, _
        txtbxGrupo.tag, "", txtbxTerceiro.tag, chkAberto.Value, "", txtbxRepr.tag, sOrd, IIf(chkPrej.Value = 1, "S", ""), Venc_che)
    Else
        If txtbxGrupo.tag <> "" Or txtbxOrgao.tag <> "" Then
            strSQL = doc.F_SelectDocsReceberGruposOrgaos(Left(cmbxTipoTer.Text, 1), "", "", _
            "", txtbxTipoCob.tag, txtbxContaCob.tag, txtbxNumDoc.Text, txtbxHist.Text, txtbxSeqItd.Text, sIn, _
             txtbxGrupo.tag, txtbxOrgao.tag, "", txtbxTerceiro.tag, chkAberto.Value, sOrd, IIf(chkPrej.Value = 1, "S", ""), Venc_che)
        Else
        strSQL = doc.F_SelectDocsReceber(Left(cmbxTipoTer.Text, 1), "", "", _
            "", txtbxTipoCob.tag, txtbxContaCob.tag, txtbxNumDoc.Text, txtbxHist.Text, txtbxSeqItd.Text, sIn, _
            txtbxGrupo.tag, "", txtbxTerceiro.tag, , sOrd, , Venc_che)
        End If
    End If
    
    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            iAcha = acha_linha(!seq_doc)
            If iAcha = 0 Then
                strSQL = iiTem + 1 & Chr(9)
                strSQL = strSQL & !seq_itd & Chr(9)
                strSQL = strSQL & !seq_doc & Chr(9)
                strSQL = strSQL & !num_doc & Chr(9)
                strSQL = strSQL & !historico_doc & Chr(9)
                strSQL = strSQL & !esp_Doc & Chr(9)
                strSQL = strSQL & !tipo_ter & Chr(9)
                strSQL = strSQL & !nome & Chr(9)
                strSQL = strSQL & !parcela_itd & "/" & !parcelas_doc & Chr(9)
                strSQL = strSQL & FormatVal(!reman_itd) & Chr(9)
                strSQL = strSQL & !venc_itd & Chr(9)
                msflxgrd.AddItem strSQL
                cTotalDoc = !reman_itd
                iiTem = iiTem + 1
            Else
                cTotalDoc = cTotalDoc + !reman_itd
                msflxgrd.Row = iAcha
                msflxgrd.Col = msflxgrd.Cols - 1
                
                If msflxgrd.Text <> "" Then
                    msflxgrd.Cols = msflxgrd.Cols + 1
                    msflxgrd.Row = iAcha
                    msflxgrd.Col = msflxgrd.Cols - 1: msflxgrd.Text = !parcela_itd & "/" & !parcelas_doc
                    msflxgrd.Row = 0: msflxgrd.Text = "Parcela"
                    msflxgrd.Cols = msflxgrd.Cols + 1
                    
                    msflxgrd.Row = iAcha
                    msflxgrd.Col = msflxgrd.Cols - 1: msflxgrd.Text = FormatVal(!reman_itd)
                    msflxgrd.Row = 0: msflxgrd.Text = "Reman"
                    msflxgrd.Row = iAcha
                    msflxgrd.Cols = msflxgrd.Cols + 1
                    msflxgrd.Col = msflxgrd.Cols - 1: msflxgrd.Text = !venc_itd
                    msflxgrd.Row = 0: msflxgrd.Text = "Vencimento"
                Else
                    For iCount = 1 To msflxgrd.Cols - 1
                        msflxgrd.Row = iAcha
                        msflxgrd.Col = iCount
                        If msflxgrd.Text = "" Then
                            
                            msflxgrd.Col = iCount: msflxgrd.Text = !parcela_itd & "/" & !parcelas_doc
                            msflxgrd.Col = iCount + 1: msflxgrd.Text = FormatVal(!reman_itd)
                            msflxgrd.Col = iCount + 2: msflxgrd.Text = !venc_itd
                            Exit For
                        End If
                    Next
                
                End If
            End If
            cTotalReman = cTotalReman + !reman_itd
            P_Status .AbsolutePosition + 1 & " & /" & .RecordCount
            .MoveNext
        Loop
    End With
    
    
    
    
    
    
    
    
    
    
    
    ZebraGrid msflxgrd, 1
    
    P_Status
    bActive = True
    Exit Sub
trata_Erro:
    P_Status
    MsgBox Err.Description, vbCritical
End Sub
Function acha_linha(seq_doc As String) As Integer
    Dim iCount As Integer
    Dim jCount As Integer
    With msflxgrd
        For iCount = 1 To msflxgrd.rows - 1
            .Col = 2
            .Row = iCount
            If .Text = seq_doc Then
                
                acha_linha = .Row
                Exit For
                
            End If
        Next
    End With
End Function
Private Sub cmndbtnCarne_Click()
    imprimeCarne dtNot.Recordset!seq_doc
End Sub

Private Sub cmndbtnCob_Click()
    buscaContaCob
End Sub
Sub buscaContaCob()
    Dim c As ContaCob
    c = F_BuscaContaCob(txtbxContaCob.Text)
    If c.cod_cob <> "" Then
        txtbxContaCob.Text = c.nome_cob
        txtbxContaCob.tag = c.cod_cob
    End If
    
    
End Sub

Private Sub cmndbtnContrato_Click()
    'imprimeContratoHTML dtNot.Recordset!seq_doc
    imprimeContrato dtNot.Recordset!seq_doc
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmEdiLan
    On Error GoTo trata_Erro
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    
    With dtNot.Recordset
        If .RecordCount < 0 Then
            Exit Sub
        End If
        fr.tag = !seq_itd
        
        fr.Show 1
        'RefreshDataPosition dtNot
        Exit Sub
        
        fr.lblData.caption = !data_doc
        fr.lblNUM_Doc.caption = !num_doc
        fr.lblParc.caption = !parcela_itd & "/" & !parcelas_doc
        fr.lblSeq_Doc.caption = !seq_doc
        fr.lblSeq_Itd.caption = !seq_itd
        fr.lblValor.caption = FormatVal(!valor_itd)
        fr.txtbxHist.Text = !historico_doc
        fr.txtbxNome.Text = !nome
        fr.txtbxNome.tag = !Codigo
        fr.txtbxVenc.Text = !venc_itd
        fr.txtbxValor.Text = FormatVal(!reman_itd)
        fr.txtbxBoleto.Text = !boleto_itd
        fr.txtbxClass.tag = !cod_pla
        fr.txtbxClass.Text = !nome_pla
        fr.txtbxContaCob.Text = !nome_cob
        fr.txtbxContaCob.tag = !cod_cob
        fr.txtbxTipoCob.Text = !nome_tip
        fr.txtbxTipoCob.tag = !cod_tip
        fr.lblTipo.caption = "Crédito"
        fr.Show 1
        dtNot.Refresh
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnExcl_Click()
    If strExcldoc_usu <> "S" Then
        MsgBox "Você não pode excluir este documento!", vbExclamation
        Exit Sub
    End If
    
    With dtNot.Recordset
        If !num_ven = "" Then
            If F_BuscaUmCampo("reman_itd", "tab_itdoc", "reman_itd<valor_itd and seq_doc='" & !seq_doc & "'") = "" Then
                If MsgBox("Deseja realmente excluir o documento e suas parcelas?", vbYesNo + vbCritical) = vbYes Then
                    strSQL = "delete from tab_doc where seq_doc='" & !seq_doc & "' "
                    Banco.Execute strSQL
                    strSQL = "delete from tab_itdoc where seq_doc='" & !seq_doc & "' "
                    Banco.Execute strSQL
                    strSQL = "delete from tab_cheques where seq_doc='" & !seq_doc & "' "
                    Banco.Execute strSQL
                    dtNot.Refresh
                End If
            Else
                MsgBox "Este documento não pode ser excluído por que já teve uma parcela baixada!", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "Este documento não pode ser excluído pois se refere a uma venda!", vbExclamation
            
        End If
    End With
    
    
    
    
    
    
End Sub

Private Sub cmndbtnGru_Click()
    Dim g As grupoCli
    g = F_BuscaGrupoCli
    txtbxGrupo.Text = g.nome_gru
    txtbxGrupo.tag = g.cod_gru
    
End Sub

Private Sub cmndbtnImpr_Click()
    
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.Show 0
End Sub

Private Sub cmndbtnImprimir_Click()
    P_MsflxToHtml "Resumo de Carnês a Receber", msflxgrd, "De: " & txtbxDataIni.Text & " Até " & txtbxDataFin.Text
End Sub

Private Sub cmndbtnOrg_Click()
    Dim o As orgao.orgao
    o = F_BuscaOrgao(txtbxOrgao.Text, txtbxGrupo.tag)
    
    txtbxOrgao.Text = o.nome_org
    txtbxOrgao.tag = o.cod_org
    
    
    
End Sub

Private Sub cmndbtnRep_Click()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxRepr.Text)
        If r.cod_Rep <> "" Then
            txtbxRepr.Text = r.nome_rep
            txtbxRepr.tag = r.cod_Rep
            
        End If
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.BorderStyle = 1
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.BorderStyle = 0
End Sub

Private Sub cmndbtnTerc_Click()
    
    
        If cmbxTipoTer.ListIndex = 0 Then
            buscaCli
        Else
            buscaRep
        End If
    
    
    cmndbtnBuscar_Click
End Sub
Sub buscaCli()
    Dim c As Cliente
    c = F_BuscaCliente(txtbxTerceiro.Text)
        If c.cod_cli <> "" Then
            txtbxTerceiro.Text = c.nome_Cli
            txtbxTerceiro.tag = c.cod_cli
            lblCPF.caption = c.cnpj_Cli
        End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
        f = F_SelectForn(txtbxTerceiro.Text)
        If f.cod_for <> "" Then
            txtbxTerceiro.Text = f.nome_for
            txtbxTerceiro.tag = f.cod_for
            lblCPF.caption = f.cnpj_for
        End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxTerceiro.Text)
        If r.cod_Rep <> "" Then
            txtbxTerceiro.Text = r.nome_rep
            txtbxTerceiro.tag = r.cod_Rep
            lblCPF.caption = r.cpf_rep
        End If
    
End Sub


Private Sub cmndbtnTerceiro_Click()
    
End Sub

Private Sub cmndbtnTipoCob_Click()
    buscaTipoCob
End Sub
Sub buscaTipoCob()
    Dim t As tipocob
    t = F_BuscaTipoCob(txtbxTipoCob.Text)
    If t.cod_tip <> "" Then
        txtbxTipoCob.Text = t.nome_tip
        txtbxTipoCob.tag = t.cod_tip
    End If
End Sub


'Private Sub dbgrd_DblClick1()
'        Dim cJuros As Currency
'        With dtNot.Recordset
'        Dim sJur As String
'        If !reman_itd <= 0 Then
'            Exit Sub
'        End If
'        On Error Resume Next
'        cJuros = F_CalculaJuros(!venc_itd, Date, !reman_itd)
'        If cJuros < 0 Then cJuros = 0
'
'        If IsLoad("frmListaDoc") Then
'            If frmListaDoc.tag <> Me.Name And frmListaDoc.msflxGrd.rows > 0 Then
'                MsgBox "A lista está sendo usada por outro formulário!", vbExclamation
'                Unload frmListaDoc
'                Exit Sub
'
'            End If
'            If frmListaDoc.lblCliente.tag <> "" And _
'                frmListaDoc.lblCliente.tag <> !cod_ter Then
'                MsgBox "Você só pode selecionar parcelas do mesmo Terceiro!", vbExclamation
'                Exit Sub
'            End If
'            frmListaDoc.lblCliente.tag = !cod_ter
'            frmListaDoc.lblCliente.caption = !Nome
'
'            If F_InFlxgrd(frmListaDoc.msflxGrd, !seq_itd, 0) = -1 Then
'                frmListaDoc.msflxGrd.AddItem !seq_itd & Chr(9) & FormatVal(!reman_itd) & Chr(9) & FormatVal(cJuros) _
'                & Chr(9) & !num_doc & Chr(9) & !cod_pla & Chr(9) & FormatVal(!reman_itd) & Chr(9) & !seq_doc
'            End If
'            frmListaDoc.soma
'        Else
'
'            frmListaDoc.tag = Me.Name
'            frmListaDoc.lblCliente.tag = !cod_ter
'            frmListaDoc.lblCliente.caption = !Nome
'            frmListaDoc.caption = "Documentos"
'            frmListaDoc.msflxGrd.Cols = 7
'            frmListaDoc.msflxGrd.rows = 0
'            frmListaDoc.msflxGrd.AddItem !seq_itd & Chr(9) & FormatVal(!reman_itd) & Chr(9) & FormatVal(cJuros) _
'                & Chr(9) & !num_doc & Chr(9) & !cod_pla & Chr(9) & FormatVal(!reman_itd) & Chr(9) & !seq_doc
'            frmListaDoc.Show 0, mdiFrm
'            frmListaDoc.soma
'        End If
'
'    End With
'
'End Sub

Private Sub dbgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        mnSelec_Click
    End If
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dbgrd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnParcela, , X + dbgrd.Left, Y + dbgrd.Top
    End If
End Sub

Private Sub dbgrdChe_GotFocus()
    dbgrdChe.MarqueeStyle = 3
End Sub

Private Sub dbgrdChe_LostFocus()
    dbgrdChe.MarqueeStyle = 0
End Sub

Private Sub dtNot_Reposition()
    Dim strSQL As String
    On Error Resume Next
    If bActive = False Then Exit Sub
    With dtNot.Recordset
        
        If .RecordCount > 0 Then
            If !reman_itd = 0 Then
                cmndbtnEdit.Enabled = False
                cmndbtnBaixar.Enabled = False
            Else
                cmndbtnEdit.Enabled = True
                cmndbtnBaixar.Enabled = True
            End If
            
        End If
        If !tipo_doc = "R" Then
            cmndbtnCarne.Enabled = True
        Else
            cmndbtnCarne.Enabled = False
        End If
        lblTel.caption = F_BuscaTel(!cod_ter, !tipo_ter)
        buscaBaixas
        buscaOrg
        buscaVias
    End With
End Sub
Sub buscaVias()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error Resume Next
    lblVias.caption = F_BuscaUmCampo("count(*)", "tab_docimpress", "seq_doc='" & dtNot.Recordset!seq_doc & "'")
    
    
    
End Sub
Sub buscaOrg()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro:
    strSQL = orgao.F_SelectOrgaoCli(dtNot.Recordset!cod_ter)
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            lblNomeOrg.caption = !nome_org
            lblEnd.caption = !endereco_org & " " & !nome_set & " " & !nome_cid & "/" & !uf_cid & "(" & !fone_org & ")"
            
        Else
        
        End If
        .Close
    End With
    
    Exit Sub
trata_Erro:
End Sub
Sub buscaBaixas()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim r2 As Recordset
    On Error GoTo trata_Erro
    Dim sMsg As String
    
    With dtNot.Recordset
        sMsg = "Parcelas" & Chr(13)
        strSQL = "select * from tab_baixaitd2 where seq_itd='" & !seq_itd & "' "
'        Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
'        With rPesq
'            Do While Not .EOF
'                sMsg = sMsg & "Data: " & !data_bai & "Cobrança: " & !cod_cob & " Vlr Baixado: " & FormatVal(!valor_bai) & " Usu: " & !cod_usu & Chr(13)
'               ' sMsg = sMsg & "Formas Pagamento" & Chr(13)
'                strSQL = "select * from tab_baixafin where seq_cab='" & !seq_cab & "' "
'                Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
'                With rPesq
'                    Do While Not .EOF
'                        sMsg = sMsg & "Espécie: " & !esp_bai & "Valor: " & FormatVal(!valor_bai)
'
'                        .MoveNext
'                    Loop
'                    .Close
'                End With
'                .MoveNext
'            Loop
'            .Close
'        End With
'        txtbxBaixas.Text = sMsg
    dtBaixas.RecordSource = strSQL
    dtBaixas.Refresh
    End With
    Exit Sub
trata_Erro:
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    
    
    
    bActive = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    On Error Resume Next
    bActive = False
    resizeForm Me
    txtbxDataIni.Text = Date
    
    txtbxDataFin.Text = Date
    dtNot.Connect = Banco.Connect
    cmbxTipo.ListIndex = 0
    cmbxOrder.ListIndex = 0
    dtNot.Connect = Banco.Connect
    dtBaixas.Connect = Banco.Connect
    cmbxBuscar.ListIndex = 1
    cmbxOrder.ListIndex = 0
    cmbxTipo.ListIndex = 0
    cmbxTipoTer.ListIndex = 0
    
    ajusta_Grid
    
End Sub
Sub ajusta_Grid()
    With msflxgrd
        .Cols = 11
        .FixedCols = 0
        .rows = 1
        .Cols = 11
        .Col = 0: .Text = "..."
        .Col = 1: .Text = "Id.Parc."
        .Col = 2: .Text = "Id.Doc."
        .Col = 3: .Text = "Documento"
        .Col = 4: .Text = "Histórico"
        .Col = 5: .Text = "T.Doc."
        .Col = 6: .Text = "T.Terc."
        .Col = 7: .Text = "Terceiro"
        .Col = 8: .Text = "Parcela"
        .Col = 9: .Text = "Reman."
        .Col = 10: .Text = "Vencimento"
        .ColWidth(4) = 3500
        .ColWidth(7) = 2500
        
    End With
End Sub
Private Sub optnCli_Click()
    frExpr.Enabled = True
End Sub

Private Sub optnTodas_Click()
    frExpr.Enabled = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    msflxgrd.Width = Me.Width - msflxgrd.Left - 100
    msflxgrd.Height = pcbot.Top - msflxgrd.Top - 300
    
End Sub

Private Sub mnBaixar_Click()
    cmndbtnBaixar_Click
End Sub

Private Sub mnBaixas_Click()
    On Error Resume Next
    frmRelCarnesRecebidos.txtbxSeqItd.Text = dtNot.Recordset!seq_ite
    frmRelCarnesRecebidos.tag = "x"
    frmRelCarnesRecebidos.Show 0
    frmRelCarnesRecebidos.SetFocus

End Sub

Private Sub mnEdiPar_Click()
    cmndbtnEdit_Click
End Sub

Private Sub mnExclDoc_Click()
    cmndbtnExcl_Click
End Sub

Private Sub mnImpr_Click()
    cmndbtnImpr_Click
End Sub

Private Sub mnImprCarne_Click()
    cmndbtnCarne_Click
End Sub

Private Sub mnImprContr_Click()
    cmndbtnContrato_Click
End Sub

Private Sub mnNova_Click()
    mnNova_Click
End Sub

Private Sub mnSair_Click()
    Unload Me
End Sub

Private Sub mnSelec_Click()
    If Me.tag <> "" Then
        Me.tag = "X"
        Me.Hide
    End If
End Sub

Private Sub mnSelecPar_Click()
    dbgrd_DblClick
End Sub

Private Sub mnSPC_Click()
    On Error Resume Next
    P_ParcelaSpc dtNot.Recordset!cod_ter, dtNot.Recordset!seq_itd, dtNot.Recordset!reman_itd
End Sub

Private Sub mnVisAcertoVenda_Click()
    On Error GoTo trata_Erro:
    With dtNot.Recordset
        If !esp_Doc = "A" Then
            P_VisualizaAcerto !num_ven
        ElseIf !esp_Doc = "V" Then
            ImprVenda !num_ven
        End If
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub mnVisDoc_Click()
    On Error GoTo trata_Erro
    With dtNot.Recordset
        txtbxSeqDoc.Text = !seq_doc
        
        txtbxSeqItd.Text = ""
        'txtbx.Text = ""
        txtbxDataFin.Text = ""
        txtbxDataIni.Text = ""
        cmndbtnBuscar_Click
        txtbxSeqDoc.Text = ""
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub pcTop_GotFocus()
    SendKeys Chr(9)
End Sub

Private Sub txtbxContaCob_Change()
    txtbxContaCob.tag = ""
End Sub

Private Sub txtbxdatafin_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxdatafin_LostFocus
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxdatafin_LostFocus()
    P_AjustaData txtbxDataFin
    If txtbxDataFin.Text <> "" Then
        
    End If
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_GotFocus()

End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        cmndbtnTerc_Click
    End If
End Sub


Private Sub txtbxGrupo_Change()
    txtbxGrupo.tag = ""
End Sub

Private Sub txtbxGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnGru_Click
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
Private Sub txtbxNumDoc_GotFocus()
    P_SelecTexto txtbxNumDoc
End Sub

Private Sub txtbxNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNumDoc_LostFocus()
    P_AjustaTexto txtbxNumDoc
End Sub
Private Sub txtbxValor_GotFocus()
    P_SelecTexto txtbxValor
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    P_ValorLostFocus txtbxValor
End Sub
Private Sub txtbxRepr_Change()
    txtbxRepr.tag = ""
End Sub

Private Sub txtbxRepr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnRep_Click
    End If
End Sub

Private Sub txtbxSeqDoc_GotFocus()
    P_SelecTexto txtbxSeqDoc
End Sub

Private Sub txtbxSeqDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSeqDoc_LostFocus()
    txtbxSeqDoc.Text = Format(txtbxSeqDoc.Text, "000000#")
End Sub

Private Sub txtbxSeqItd_GotFocus()
    P_SelecTexto txtbxSeqItd
End Sub

Private Sub txtbxSeqItd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSeqItd_LostFocus()
    txtbxSeqItd.Text = Format(txtbxSeqItd.Text, "000000#")
End Sub

Private Sub txtbxTerceiro_Change()
    txtbxTerceiro.tag = ""
    lblCPF.caption = ""
    lblCPF.tag = ""

End Sub

Private Sub txtbxTerceiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxTipoCob_Change()
    txtbxTipoCob.tag = ""
End Sub


