VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmNotaEntrada 
   Caption         =   "Escolha uma Nota Fiscal"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   Icon            =   "frmNotaEntrada.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   12300
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4455
      Left            =   240
      TabIndex        =   29
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7858
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "num_not"
         Caption         =   "Número"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nome_for"
         Caption         =   "Fornecedor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "uf_for"
         Caption         =   "uf"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "serie_not"
         Caption         =   "serie"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "valor_not"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "emissao_not"
         Caption         =   "Emissao"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "status_not"
         Caption         =   "Status"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "seq_not"
         Caption         =   "Seq"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4094,929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   12270
      TabIndex        =   25
      Top             =   7170
      Width           =   12300
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar (ESC)"
         Height          =   550
         Left            =   9720
         Picture         =   "frmNotaEntrada.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   555
         Left            =   3720
         Picture         =   "frmNotaEntrada.frx":0858
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdTodas 
         Caption         =   "&Efetivar Todas"
         Height          =   555
         Left            =   7320
         Picture         =   "frmNotaEntrada.frx":0C65
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Cancelar"
         Height          =   550
         Left            =   8520
         Picture         =   "frmNotaEntrada.frx":0DAF
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnEfetivar 
         Caption         =   "&Efetivar"
         Height          =   550
         Left            =   6120
         Picture         =   "frmNotaEntrada.frx":0EF9
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnReabrir 
         Caption         =   "&Reabrir"
         Height          =   550
         Left            =   4920
         Picture         =   "frmNotaEntrada.frx":1043
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnEditar 
         Caption         =   "&Editar"
         Height          =   550
         Left            =   2520
         Picture         =   "frmNotaEntrada.frx":118D
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "&Transferir"
         Height          =   550
         Left            =   1320
         Picture         =   "frmNotaEntrada.frx":12D7
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnNova 
         Caption         =   "&Nova"
         Height          =   550
         Left            =   120
         Picture         =   "frmNotaEntrada.frx":1421
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   12300
      TabIndex        =   0
      Top             =   6435
      Width           =   12300
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   9720
         ScaleHeight     =   705
         ScaleWidth      =   2265
         TabIndex        =   26
         Top             =   0
         Width           =   2295
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   390
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   27
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Pesquisar"
         Height          =   525
         Left            =   8520
         Picture         =   "frmNotaEntrada.frx":156B
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtNum 
         Height          =   285
         Left            =   7320
         MaxLength       =   40
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbxSitua 
         Height          =   315
         ItemData        =   "frmNotaEntrada.frx":1664
         Left            =   120
         List            =   "frmNotaEntrada.frx":1674
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtForn 
         Height          =   285
         Left            =   5160
         MaxLength       =   40
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         Picture         =   "frmNotaEntrada.frx":16D6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmNotaEntrada.frx":17CF
         Left            =   3720
         List            =   "frmNotaEntrada.frx":17DF
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   7320
         TabIndex        =   12
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor:"
         Height          =   195
         Left            =   5160
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   3720
         TabIndex        =   7
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   24
      Top             =   8040
      Width           =   390
   End
End
Attribute VB_Name = "frmNotaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmbxOrder_Click()
    busca
End Sub

Private Sub cmbxSitua_Click()
    busca
End Sub

Private Sub cmdTodas_Click()
    On Error GoTo trata_Erro
    With ado
        Do While Not .EOF
            efetiva
            .MoveNext
        Loop
    End With
        
    MsgBox "Processo concluido!", vbInformation
    busca
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnCancelar_Click()
    Dim strSQL As String
    With ado
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !status_not = "4" Then
            If MsgBox("Deseja realmente cancelar a nota fiscal e voltar todo o estoque dos produtos envolvidos?", vbYesNo + vbCritical) = vbNo Then
                    Exit Sub
            End If
                'P_cancelanotaFechada !seq_not, "R"
                busca
        Else
            If MsgBox("Deseja Realmente Excluir a Nota Fiscal " & !Num_not & _
                " ?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "delete from tab_cmp where num_not = '" & !Num_not & "'"
                strSQL = strSQL & " and cod_for='" & !cod_For & "' "
                strSQL = strSQL & " and serie_not ='" & !serie_not & "' "
                Conn.Execute strSQL
                strSQL = "delete from tab_itcmp where seq_not = '" & !seq_not & "'"
            
            
                Conn.Execute strSQL
                strSQL = "update tab_compras set status_not='2' where num_not = '" & !Num_not & "'"
                strSQL = strSQL & " and cod_for='" & !cod_For & "' "
                strSQL = strSQL & " and serie_not ='" & !serie_not & "' "
                Conn.Execute strSQL
                ado.Requery
            End If
        End If
    End With

End Sub

Private Sub cmndbtnEditar_Click()
    Dim fr As New frmEntradas
    fr.operacao = "I"
    If ado.RecordCount <= 0 Or dbgrd.row = -1 Then
        MsgBox "Não há notas fiscais nesta pesquisa!", vbExclamation
        Exit Sub
    End If
    
    With ado
        If !status_not <> "1" Then
            MsgBox "A Nota Não está Aberta!", vbExclamation
            Exit Sub
        End If
        fr.operacao = "F"
        fr.tag = !seq_not
        fr.grid.tag = fr.tag
        fr.caption = "Finalização de Nota Fiscal: " & !Num_not & " de " & !nome_for
        
        fr.txtbxDesconto.text = FormatVal(!desconto_not)
        fr.txtbxNum_Not.text = !Num_not
        fr.txtbxNum_Not.tag = !Num_not
        fr.txtbxSerie_Not.text = !serie_not
        fr.txtbxSerie_Not.tag = !serie_not
        fr.txtbxBC_Not.text = Format(!bc_not, "##,###,##0.00")
        fr.txtbxValorProd.text = FormatVal(!valorprod_not)
        fr.txtbxDepesas_Not.text = Format(!despesas_not, "###,###,##0.00")
        fr.txtbxFrete.text = Format(!frete_not, "###,###,##0.00")
        fr.txtbxEmissao_Not.text = !emissao_not
        fr.txtbxEntrada_Not.text = !entrada_not
        fr.txtbxICMS_Not.text = Format(!icms_not, "###,###,##0.00")
        fr.txtbxIPI_Not.text = Format(!ipi_not, "###,###,##0.00")
        fr.txtbxNOme_Ter.text = !nome_for
        fr.txtbxNOme_Ter.tag = !cod_For
        fr.lblUF_Ter.caption = Format(!uf_for)
        
        fr.txtbxValor_Not.text = FormatVal(!valor_not)
        
        P_PreencheOperacoes fr.cmbxTipo_Op, "E"
        
         
        P_LocalizaOperacao !operacao_not, fr.cmbxTipo_Op
        
        If fr.operacao = "F" Then
            fr.Show 0
        Else
            fr.Show 0, mdiFrm
        End If
        
    End With
End Sub
Sub efetiva()
      Dim rPesq As New ADODB.Recordset
    Dim rForn As New ADODB.Recordset
    Dim strSQL As String
    Dim cCustoMed As Currency
    Dim cEst As Currency 'o estoque antigo
    Dim cCusto As Currency 'o custo antigo
    Dim sPro As String
    Dim sPreco As String
    Dim cNovoPreco As Currency
    'On Error Resume Next
    On Error GoTo trata_Erro
    With ado
        
        If !status_not <> "2" Then
            MsgBox "Esta nota não pode ser efetivada!", vbExclamation
            Exit Sub
        End If
        
        
        
        strSQL = "update tab_cmp set status_not = '4' "
        strSQL = strSQL & " where cod_for='" & !cod_For & "' "
        strSQL = strSQL & " and serie_not='" & !serie_not & "'"
        strSQL = strSQL & " and num_not='" & !Num_not & "' "
        
        Conn.Execute strSQL
        
            strSQL = "update tab_compras set status_not = '4' where "
        
        strSQL = strSQL & " seq_not ='" & !seq_not & "' "
        Conn.Execute strSQL
        'itens
        strSQL = "select * from tab_itcompras "
        strSQL = strSQL & " where seq_not='" & F_BuscaUmCampo("seq_not", "tab_compras", "cod_for='" & !cod_For & "' and serie_not='" & !serie_not & "' and num_not='" & !Num_not & "'") & "' "
        'strSQL = strSQL & " where seq_not='" & ado!seq_not & "'"
        
        abreObjetoAdo rPesq, strSQL
        With rPesq
            Do While Not .EOF
                cEst = CCur((F_BuscaUmCampo("ef_Pro", "tab_produtos", " cod_pro='" & !cod_pro & "'")))
                cCusto = CCur((F_BuscaUmCampo("cf_pro ", "tab_produtos", " cod_pro='" & !cod_pro & "'")))
                
                
                
                strSQL = "update tab_produtos set ef_pro=ef_pro+" & F_GravaDec(CCur(!qtd_ite) * CCur(!qtd_emb))
                strSQL = strSQL & ", cf_pro=" & F_GravaDec(CCur(!custo_ite))
                
                
                

                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                Conn.Execute strSQL
                
                
                .MoveNext
            Loop
        
        End With
        
        'itens cmp
        strSQL = "select ti.*, tp.pv_pro,tp.nome_pro,tp.er_pro,tp.cr_pro from tab_itcmp as ti , tab_produtos as tp "
        strSQL = strSQL & " where ti.seq_not='" & F_BuscaUmCampo("seq_not", "tab_cmp", "cod_for='" & ado!cod_For & "' and serie_not='" & ado!serie_not & "' and num_not='" & ado!Num_not & "'") & "' "
        strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
        
        abreObjetoAdo rPesq, strSQL
    End With

       
        With rPesq
            Do While Not .EOF
                cEst = !er_pro
                cCusto = !cr_pro
                If (CCur(cEst) + CCur(!qtd_ite) * CCur(!qtd_emb)) Then
                    cCustoMed = ((cEst * cCusto) + (CCur(!qtd_ite) * CCur(!qtd_emb) * CCur(!custo_ite))) / (cEst + CCur(!qtd_ite) * CCur(!qtd_emb))
                End If
                strSQL = "update tab_produtos set er_pro=er_pro+" & F_GravaDec(CCur(!qtd_ite) * CCur(!qtd_emb))
               
                
                strSQL = strSQL & ", cr_pro=" & F_GravaDec(!custo_ite)
                
                
                
                
                cNovoPreco = Round(!precofuturo_ite, 2)
                
                 
                    'strSQL = strSQL & ", pv_pro=" & F_GravaDec(cNovoPreco)
                    
                
                

                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                Conn.Execute strSQL
                P_AltPreco !cod_pro, cNovoPreco, !pv_pro
                .MoveNext
            Loop
            
        End With
        Exit Sub
trata_Erro:
        P_MostraErro
End Sub

Private Sub cmndbtnEfetivar_Click()
  efetiva
        
        MsgBox "Nota Efetivada com Sucesso!", vbInformation
        
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnImpr_Click()
    On Error GoTo trata_Erro
    P_DtgrdToHtml "Notas", grid, ado, Date, "4,8"
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNova_Click()
    Dim fr As New frmEntradas
    fr.operacao = "F"
    fr.Show 0, mdiFrm
End Sub

Private Sub cmndbtnReabrir_Click()
    Dim strSQL As String
    With ado
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If MsgBox("Deseja reabrir a Nota Fiscal?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "update tab_compras set status_not='1' where "
            strSQL = strSQL & "num_not='" & !Num_not & "' "
            strSQL = strSQL & " and cod_for='" & !cod_For & "' "
            strSQL = strSQL & " and serie_not = '" & !serie_not & "' "
            Conn.Execute strSQL
            If Banco.RecordsAffected = 0 Then
                MsgBox "Esta não é uma nota complementada!", vbExclamation
                Exit Sub
            End If
            adc.Refresh
        End If
    End With
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnTerc_Click()
    Dim f As fornecedor.fornecedor
    
    f = fornecedor.F_SelectForn(txtForn.text)
    If f.cod_For <> "" Then
        txtForn.text = f.nome_for
        txtForn.tag = f.cod_For
        busca
        
        
    End If
    
End Sub

Private Sub cmndbtnTransf_Click()
    
'        If MsgBox("Confirma Transferência?", vbYesNo + vbQuestion) = vbYes Then
            transfere
'        End If
    
End Sub
Sub transfere()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sSeq As String
    With ado
        Dim fr As New frmEntradas
        fr.operacao = "F"
        fr.tag = !seq_not
        fr.grid.tag = fr.tag
        fr.caption = "Nota Fiscal: " & !Num_not & " de " & !nome_for
        fr.txtbxDesconto.text = FormatVal(!desconto_not)
        fr.txtbxNum_Not.text = !Num_not
        fr.txtbxNum_Not.tag = !Num_not
        fr.txtbxSerie_Not.text = !serie_not
        fr.txtbxSerie_Not.tag = !serie_not
        fr.txtbxBC_Not.text = Format(!bc_not, "##,###,##0.00")
        fr.txtbxValorProd.text = FormatVal(!valorprod_not)
        fr.txtbxDepesas_Not.text = Format(!despesas_not, "###,###,##0.00")
        fr.txtbxFrete.text = Format(!frete_not, "###,###,##0.00")
        fr.txtbxEmissao_Not.text = !emissao_not
        fr.txtbxEntrada_Not.text = !entrada_not
        fr.txtbxICMS_Not.text = Format(!icms_not, "###,###,##0.00")
        fr.txtbxIPI_Not.text = Format(!ipi_not, "###,###,##0.00")
        fr.txtbxNOme_Ter.text = !nome_for
        fr.txtbxNOme_Ter.tag = !cod_For
        fr.lblUF_Ter.caption = Format(!uf_for)
        
        fr.txtbxValor_Not.text = FormatVal(!valor_not)
        
        P_PreencheOperacoes fr.cmbxTipo_Op, "E"
        
         
        P_LocalizaOperacao !operacao_not, fr.cmbxTipo_Op
        
        sSeq = MaxKeySQL("tab_cmp", "seq_not", "", "000000#")
        strSQL = "insert into tab_cmp(seq_not,Cod_For,Num_not,Serie_not,desconto_not,valorprod_not,Valor_not,Bc_Not,"
        strSQL = strSQL & "frete_not,icms_not,ipi_not,despesas_not,operacao_not,emissao_not,"
        strSQL = strSQL & "entrada_not,cod_usu,status_not,seq_ent)Values("
        strSQL = strSQL & "'" & sSeq & "',"
        strSQL = strSQL & "'" & !cod_For & "', "
        strSQL = strSQL & "'" & !Num_not & "',"
        strSQL = strSQL & "'" & !serie_not & "',"
        strSQL = strSQL & F_GravaDec(!desconto_not) & ","
        strSQL = strSQL & F_GravaDec(!valorprod_not) & ","
        strSQL = strSQL & F_GravaDec(!valor_not) & ","
        strSQL = strSQL & F_GravaDec(!bc_not) & ","
        strSQL = strSQL & F_GravaDec(!frete_not) & ","
        strSQL = strSQL & F_GravaDec(!icms_not) & ","
        strSQL = strSQL & F_GravaDec(!ipi_not) & ","
        strSQL = strSQL & F_GravaDec(!despesas_not) & ","
        strSQL = strSQL & "'" & (!operacao_not) & "',"
        strSQL = strSQL & F_GravaData(!emissao_not) & ","
        strSQL = strSQL & F_GravaData(!entrada_not) & ","
        strSQL = strSQL & "'" & strCod_Usu & "','1', " 'em aberto
        strSQL = strSQL & "'" & !seq_not & "') "
        Conn.Execute strSQL

        
        strSQL = "select * from tab_itcompras where seq_not='" & !seq_not & "' "
        strSQL = strSQL & " order by seq_ite"
        abreObjetoAdo rPesq, strSQL
        With rPesq
            
            Do While Not .EOF
                
                strSQL = "Insert into tab_itcmp(seq_not,seq_ite,cod_pro,unit_ite,qtd_ite,qtd_emb,ipi_ite,"
                strSQL = strSQL & "gastos_ite,custo_ite,sujerido_ite,precofuturo_ite,lucrosuj_ite,"
                strSQL = strSQL & "lucroliq_ite,custosai_ite,icms_ite)values("
                strSQL = strSQL & "'" & sSeq & "',"
                strSQL = strSQL & !seq_ite & ","
                strSQL = strSQL & "'" & !cod_pro & "',"
                strSQL = strSQL & F_GravaDec(!unit_ite) & ","
                strSQL = strSQL & F_GravaDec(!qtd_ite) & ","
                strSQL = strSQL & F_GravaDec(!qtd_emb) & ","
                strSQL = strSQL & F_GravaDec(!ipi_ite) & ","
                strSQL = strSQL & F_GravaDec(!gastos_ite) & ","
                strSQL = strSQL & F_GravaDec(!custo_ite) & ","
                strSQL = strSQL & F_GravaDec(!sujerido_ite) & ","
                strSQL = strSQL & F_GravaDec(!precofuturo_ite) & ","
                strSQL = strSQL & F_GravaDec(!lucrosuj_ite) & ","
                strSQL = strSQL & F_GravaDec(!lucroliq_ite) & ","
                strSQL = strSQL & F_GravaDec(!custosai_ite) & ","
                strSQL = strSQL & F_GravaDec(!icms_ite) & ")"
                Conn.Execute strSQL
                
                .MoveNext
            Loop
            
        End With
        
    End With
    
     With ado
        
        strSQL = "update tab_compras set status_not = '3' where seq_not='" & !seq_not & "' "
        
        Conn.Execute strSQL
        
        fr.caption = "Finalização de NF: " & !Num_not & " - " & !nome_for
        fr.tag = sSeq
        fr.Show 0
    End With
    
End Sub

Private Sub cmdBuscar_Click()
    busca
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If ado.RecordCount = 0 Then Exit Sub
    With ado
        If KeyCode = vbKeyP And Shift = 2 Then
        
            P_ImprNotaFiscal frmVisualiza.Pic, !seq_not, 0
         frmVisualiza.Show 1
            Unload frmVisualiza
        End If
        If KeyCode = vbKeyE And Shift = vbCtrlMask Then
            P_EspelhoNota !seq_not
            
            
        End If
    End With
End Sub

Private Sub adc_Reposition()
    With ado
            cmndbtnCancelar.Enabled = False
            cmndbtnEditar.Enabled = False
            cmndbtnEfetivar.Enabled = False
            cmndbtnReabrir.Enabled = False
            cmndbtnTransf.Enabled = False
        If .RecordCount = 0 Then
            
            
        Else
            If cmbxSitua.ListIndex = 0 Then
                cmndbtnCancelar.Enabled = False
                cmndbtnEditar.Enabled = False
                cmndbtnEfetivar.Enabled = True
                cmndbtnReabrir.Enabled = True
                cmndbtnTransf.Enabled = True
            ElseIf cmbxSitua.ListIndex = 1 Then
                cmndbtnCancelar.Enabled = True
                cmndbtnEditar.Enabled = True
                cmndbtnEfetivar.Enabled = False
                cmndbtnReabrir.Enabled = False
                cmndbtnTransf.Enabled = False
            ElseIf cmbxSitua.ListIndex = 2 Then
                cmndbtnCancelar.Enabled = True
                cmndbtnEditar.Enabled = False
                cmndbtnEfetivar.Enabled = True
                cmndbtnReabrir.Enabled = False
                cmndbtnTransf.Enabled = False
            ElseIf cmbxSitua.ListIndex = 3 Then
                cmndbtnCancelar.Enabled = False
                cmndbtnEditar.Enabled = False
                cmndbtnEfetivar.Enabled = False
                cmndbtnReabrir.Enabled = False
                cmndbtnTransf.Enabled = False
            
            End If
        End If
        
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If kecyode = vbKeyF7 Then
        cmndbtnTransf_Click
    End If
End Sub

Private Sub Form_Load()
        Set ado = New ADODB.Recordset
End Sub
Sub busca()
    Dim strSQL As String
    If cmbxSitua.ListIndex = 0 Then
        strSQL = "select tc.*,tf.nome_for,tf.uf_for from tab_fornecedores as tf, "
        strSQL = strSQL & "tab_compras as tc "
        strSQL = strSQL & " where tc.cod_for=tf.cod_for "
    
        If txtbxdataFin.text <> "" Then
            strSQL = strSQL & " and tc.entrada_not <= " & F_PesqData(txtbxdataFin.text)
        End If
        If txtbxDataIni.text <> "" Then
            strSQL = strSQL & " and tc.entrada_not >= " & F_PesqData(txtbxDataIni.text)
        End If
    
        strSQL = strSQL & " and tc.status_not='2' "
    Else
        strSQL = "select tc.*,tf.nome_for,tf.uf_for from tab_fornecedores as tf, "
        strSQL = strSQL & "tab_cmp as tc "
        strSQL = strSQL & " where tc.cod_for=tf.cod_for "
    
        If txtbxdataFin.text <> "" Then
            strSQL = strSQL & " and tc.entrada_not < " & F_PesqData(txtbxdataFin.text)
        End If
        If txtbxDataIni.text <> "" Then
            strSQL = strSQL & " and tc.entrada_not > " & F_PesqData(txtbxDataIni.text)
        End If
        If cmbxSitua.ListIndex = 1 Then
            strSQL = strSQL & " and tc.status_not='1' "
        ElseIf cmbxSitua.ListIndex = 2 Then
            strSQL = strSQL & " and tc.status_not='2' "
        ElseIf cmbxSitua.ListIndex = 3 Then
            strSQL = strSQL & " and tc.status_not='3' "
        End If
    End If
    
    If txtNum.text <> "" Then
        strSQL = strSQL & " and num_not='" & txtNum.text & "'"
        
    End If
    If txtForn.tag <> "" Then
        strSQL = strSQL & " and tf.cod_for='" & txtForn.tag & "' "
    End If
    
    If cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by num_not,nome_for "
    ElseIf cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by nome_for, num_not "
    ElseIf cmbxOrder.ListIndex = 2 Then
        strSQL = strSQL & " order by emissao_not,nome_for, num_not "
    ElseIf cmbxOrder.ListIndex = 3 Then
        strSQL = strSQL & " order by entrada_not,nome_for, num_not "
    Else
        strSQL = strSQL & " order by num_not,nome_for "
    End If
    
    
    abreObjetoAdo ado, strSQL, grid
    soma
    
End Sub
Sub soma()
    Dim cTotal As Currency
    With ado
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            cTotal = cTotal + !valor_not
            
            .MoveNext
        Loop
        
        
        
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    lblTotal.caption = FormatVal(cTotal)
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
