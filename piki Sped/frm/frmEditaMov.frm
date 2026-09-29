VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEditaMov 
   Caption         =   "Movimentações"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13005
   Icon            =   "frmEditaMov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   13005
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   12975
      TabIndex        =   21
      Top             =   7755
      Width           =   13005
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   6840
         Picture         =   "frmEditaMov.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnDev 
         Caption         =   "&Cheque Devolvido"
         Height          =   735
         Left            =   5160
         Picture         =   "frmEditaMov.frx":0717
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   10560
         Picture         =   "frmEditaMov.frx":0861
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmEditaMov.frx":0DAF
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEditaMov.frx":0EF9
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmEditaMov.frx":1283
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9240
         TabIndex        =   29
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   8760
         TabIndex        =   28
         Top             =   180
         Width           =   405
      End
   End
   Begin VB.PictureBox pcBusca 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12975
      TabIndex        =   2
      Top             =   0
      Width           =   13005
      Begin VB.ComboBox cmbxSituacao 
         Height          =   315
         ItemData        =   "frmEditaMov.frx":13CD
         Left            =   8640
         List            =   "frmEditaMov.frx":13DA
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEditaMov.frx":13FB
         Left            =   9960
         List            =   "frmEditaMov.frx":140E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxSeqDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxHistoricoDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Height          =   525
         Left            =   11640
         Picture         =   "frmEditaMov.frx":1452
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         Picture         =   "frmEditaMov.frx":17DC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   8640
         TabIndex        =   20
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar Por:"
         Height          =   195
         Left            =   9960
         TabIndex        =   19
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente"
         Height          =   195
         Left            =   6000
         TabIndex        =   18
         Tag             =   "!"
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   3360
         TabIndex        =   17
         Tag             =   "!"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         Tag             =   "!"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº Doc:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Tag             =   "!"
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Id. Doc:"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Tag             =   "!"
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Histórico:"
         Height          =   195
         Left            =   4440
         TabIndex        =   13
         Tag             =   "!"
         Top             =   0
         Width           =   660
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   3480
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEditaMov.frx":1B66
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   11456
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adcch 
      Height          =   375
      Left            =   3120
      Top             =   7320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dtgrdCh 
      Bindings        =   "frmEditaMov.frx":1B78
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1508
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEditaMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bActive As Boolean

Private Sub chkAberto_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxOrder_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxTipo_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnBaixar_Click()
    Dim sSeq As String
    Dim iCount As Integer
    Dim c As cheque
    Dim sChe As String
    Dim sEsp As String
    Dim cValor As Currency
    Dim cJuros As Currency
    Dim strSQL As String
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para baixar documentos!", vbExclamation
        Exit Sub
    End If
    
'    If IsLoad("frmListaDoc") Then
'        If frmListaDoc.Tag <> Me.Name Then
'            MsgBox "Por favor, escolha as parcelas a serem baixadas!", vbExclamation
'            Exit Sub
'        End If
'    Else
'        MsgBox "Por favor, escolha as parcelas a serem baixadas!", vbExclamation
'        Exit Sub
'    End If
    
    With adc.Recordset
    
        If .RecordCount = 0 Then
            MsgBox "Não há o que baixar!", vbExclamation
            Exit Sub
        End If
            cJuros = F_CalculaJuros(!Venc_Itd, Date, !reman_itd)
            
            frmBaixaPar.tag = adc.Recordset!seq_itd
            frmBaixaPar.lblNome_Ter.caption = !nome
            frmBaixaPar.txtbxJuros.Text = FormatVal(cJuros)
            frmBaixaPar.txtbxReman.Text = FormatVal(!reman_itd)
            frmBaixaPar.txtbxValor_Par.Text = FormatVal(!Valor_Itd)
            frmBaixaPar.lblParc.caption = "Parcela: " & !Parc_Itd & " de: " & !Parc_Doc
            frmBaixaPar.cmndbtnConfirma.tag = "R"
            frmBaixaPar.Show 1
            
            If frmBaixaPar.tag = "" Then
                Unload frmBaixaPar
                Exit Sub
            Else
                If MsgBox("Confirma baixa da Parcela?", vbYesNo + vbQuestion) = vbYes Then
                    With frmBaixaPar.msflxgrd
                        For iCount = 1 To .rows - 1
                            .Row = iCount
                            
                            
                            .col = 2
                            sEsp = Left(.Text, 1)
                            
                            .col = 0: c.Venc_che = .Text
                            .col = 1: c.Valor_che = .Text
                            .col = 3: c.num_che = .Text
                            .col = 4: c.cod_bco = .Text
                            .col = 5: c.age_che = .Text
                            .col = 6: c.cco_che = .Text
                            .col = 7: c.nome_che = .Text
                            .col = 8: c.cpf_che = .Text
                            If sEsp = "1" Then
                                sSeq = F_GravaFinanceiro("R", Date, CStr(c.Valor_che), "Baixa doc " & _
                                    adc.Recordset!num_doc & "-" & adc.Recordset!nome, sEsp, frmBaixaPar.cmbxConta.tag, _
                                    adc.Recordset!seq_itd, "", "", strPlaRecebDupl)
                                
                                
                                If F_GravaClass(sChe, strPlaRecebDupl, CStr(c.Valor_che)) Then
                                End If
                                
                            Else
                            
                                sSeq = F_GravaDoc("BAIXA DOC:" & adc.Recordset!num_doc & "-" & adc.Recordset!nome, _
                                     adc.Recordset!num_doc, CStr(c.Valor_che), Date, _
                                    adc.Recordset!codigo, "R", "", "1", strPlaRecebDupl)
                                                                    sChe = F_GravaCh(sSeq, c.num_che, c.cod_bco, c.age_che, c.cco_che, _
                                            c.nome_che, CCur(c.Valor_che), c.Venc_che, c.cod_bco, "R", _
                                            "", "", adc.Recordset!Parc_Itd, "", c.cpf_che, Time)
                                
                            End If
                            cValor = cValor + c.Valor_che
                        Next
                        
                        strSQL = "update tab_itdoc1 set reman_itd = reman_itd-" & F_GravaDec(cValor - CCur(frmBaixaPar.txtbxJuros.Text) + CCur(frmBaixaPar.txtbxDesconto.Text))
                        strSQL = strSQL & ",juros_itd=" & F_GravaDec(frmBaixaPar.txtbxJuros.Text)
                        strSQL = strSQL & ",Desconto_itd=" & F_GravaDec(frmBaixaPar.txtbxDesconto.Text)
                        strSQL = strSQL & ", baixa_itd=" & F_GravaData(Date)
                        strSQL = strSQL & " where seq_itd='" & adc.Recordset!seq_itd & "' "
                        conn.Execute strSQL
                        Unload frmBaixaPar
                        MsgBox "Documento Baixado!", vbInformation
                        adc.Refresh
                    End With
                End If
            End If
        
    End With
End Sub

Private Sub busca()
        
    Dim strSQL As String
    Dim sOrderby As String
    Dim sCampo As String
    On Error GoTo trata_erro
    
    
    If cmbxOrder.ListIndex = 0 Then
        sOrderby = "venc_Itd"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderby = "Nome"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderby = "num_doc,venc_itd"
    Else
        sOrderby = "reman_itd"
    End If
        
    
    If "R" = "R" Then
        strSQL = "select historico_doc,td.tipo_doc,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, tc.cod_cli as codigo, tc.nome_cli as nome,td.historico_doc,td.valor_Doc,td.parc_doc,td.num_doc,td.data_doc, td.valor_doc,tl.nome_pla,"
        strSQL = strSQL & " ((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, ((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, "
        strSQL = strSQL & "ti.* from tab_clientes as tc, tab_doc1 as td,tab_planocontas as tl, tab_itdoc1 as ti,tab_parametros as tp  "
        strSQL = strSQL & " where tc.cod_cli=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & " and td.tipo_doc='R' "
        If cmbxSituacao.ListIndex = 1 Then
            strSQL = strSQL & " and ti.reman_itd > 0 "
        ElseIf cmbxSituacao.ListIndex = 2 Then
            strSQL = strSQL & " and ti.reman_itd = 0 "
        End If
        If txtbxDAtaFIn.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd <= " & F_GravaData(txtbxDAtaFIn.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd >= " & F_GravaData(txtbxDataIni.Text)
        End If
        
        If txtbxNumDoc.Text <> "" Then
            strSQL = strSQL & " and td.num_doc like  '" & txtbxNumDoc.Text & "%' "
        End If
        If txtCliente.tag <> "" Then
            
            strSQL = strSQL & " and tc.cod_Cli =  '" & txtCliente.tag & "' "
        End If
        
        strSQL = strSQL & " and tl.cod_pla=td.cod_pla "
        strSQL = strSQL & " order by " & sOrderby
            
        
    Else
'        strSQL = "select  historico_doc,td.tipo_doc,tc.cod_for as codigo,tp.nome_pla,tc.nome_for as nome,td.valor_Doc,td.parc_doc,td.num_doc,td.data_doc, td.valor_doc,"
'        strSQL = strSQL & "ti.* from tab_fornecedores as tc, tab_doc1 as td,tab_itdoc1 as ti ,tab_planocontas as tp  "
'        strSQL = strSQL & " where tc.cod_for=td.cod_ter and td.seq_doc=ti.seq_doc "
'        strSQL = strSQL & " and td.tipo_doc='P' "
'        strSQL = strSQL & " and tp.cod_pla=td.cod_pla "
'        If chkAberto.Value = 1 Then
'            strSQL = strSQL & " and ti.reman_itd > 0 "
'        End If
'        If txtbxdataFin.Text <> "" Then
'            strSQL = strSQL & " and ti.venc_itd <= " & F_GravaData(txtbxdataFin.Text)
'        End If
'        If txtbxDataIni.Text <> "" Then
'            strSQL = strSQL & " and ti.venc_itd >= " & F_GravaData(txtbxDataIni.Text)
'        End If
'        If cmbxBuscar.ListIndex = 0 Then
'            strSQL = strSQL & " and td.num_doc like  '" & txtbxExpr.Text & "%' "
'        ElseIf cmbxBuscar.ListIndex = 1 Then
'            strSQL = strSQL & " and tc.cod_for =  '" & txtbxExpr.Tag & "' "
'        ElseIf cmbxBuscar.ListIndex = 2 Then
'            sCampo = "ti.valor_itd = " & str(txtbxExpr.Text)
'
'        End If
'        strSQL = strSQL & " order by " & sOrderby
    End If
    
    
    
    
    
    
    
    adc.RecordSource = strSQL
    adc.Refresh
    
    soma
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnLocaliza_Click()
    
End Sub

Private Sub cmndbtnEdit_Click()
        
    Dim fr As New frmEdiLan
    'On Error Resume Next
    If cmndbtnEdit.Enabled = False Then Exit Sub
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    
    With adc.Recordset
    
        If .RecordCount < 0 Then
            Exit Sub
        End If
        fr.tag = !seq_itd
        fr.lblData.caption = !Data_Doc
        fr.lblNUM_Doc.caption = !num_doc
        fr.lblParc.caption = !Parc_Itd & "/" & !Parc_Doc
        fr.lblSeq_Doc.caption = !seq_doc
        fr.lblSeq_Itd.caption = !seq_itd
        fr.lblValor.caption = FormatVal(!Valor_Itd)
        fr.txtbxHist.Text = !historico_doc
        fr.txtbxNome.Text = !nome
        fr.txtbxNome.tag = !codigo
        fr.txtbxVenc.Text = !Venc_Itd
        If !tipo_doc = "R" Then
            P_PreenchePlanoContas fr.cmbxPlanoConta, True
        Else
            P_PreenchePlanoContas fr.cmbxPlanoConta, False
        End If
        fr.cmbxPlanoConta.Text = !nome_pla
        
        fr.txtbxValor.Text = FormatVal(!reman_itd)
        fr.Show 1
        adc.Refresh
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnNew_Click()
    If strLandoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    Dim fr As New frmLanDoc
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    
    adcch.Recordset.Close
    adcch.Recordset.Close
    Unload Me
    
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub DBGrid1_GotFocus()
    dtgrd.BorderStyle = 1
End Sub

Private Sub DBGrid1_LostFocus()
    dtgrd.BorderStyle = 0
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnTerc_Click()
    
    buscaCli
End Sub


Private Sub dtgrd_DblClick()
    With adc.Recordset
        cmndbtnEdit_Click
        If !reman_itd <= 0 Then
            Exit Sub
        End If
        If IsLoad("frmListaDoc") Then
            If frmListaDoc.tag <> Me.Name Then
                MsgBox "A lista está sendo usada por outro formulário!", vbExclamation
                Exit Sub
            End If
            If F_InFlxgrd(frmListaDoc.msflxgrd, !num_doc, 0) = -1 Then
                frmListaDoc.msflxgrd.AddItem !num_doc & Chr(9) & !reman_itd
            End If
            frmListaDoc.soma
        Else
            frmListaDoc.tag = Me.Name
            frmListaDoc.caption = "Documentos de Financeiras"
            frmListaDoc.msflxgrd.Cols = 2
            frmListaDoc.msflxgrd.rows = 0
            frmListaDoc.msflxgrd.AddItem !num_doc & Chr(9) & !reman_itd
            frmListaDoc.Show 0, mdiFrm
            frmListaDoc.soma
        End If
    
    End With
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub dtgrdCh_GotFocus()
    dtgrdCh.MarqueeStyle = 3
End Sub

Private Sub dtgrdChe_LostFocus()
    dtgrdCh.MarqueeStyle = 0
End Sub

Private Sub adc_Reposition()
    Dim strSQL As String
    On Error Resume Next
    If bActive = False Then Exit Sub
    With adc.Recordset
        If .RecordCount = 0 Then Exit Sub
        strSQL = "select * from tab_cheques1 where seq_doc='" & !seq_doc & "' "
        adcch.RecordSource = strSQL
        adcch.Refresh
        
        If adcch.Recordset.RecordCount > 0 Then
            dtgrdCh.caption = "Cheques referentes ao documento " & !num_doc
            dtgrdCh.Visible = True
        Else
            dtgrdCh.Visible = False
        End If
        If .RecordCount > 0 Then
            If !reman_itd = 0 Then
                cmndbtnEdit.Enabled = False
                cmndbtnBaixar.Enabled = False
            Else
                cmndbtnEdit.Enabled = True
                cmndbtnBaixar.Enabled = True
            End If
            Exit Sub
        
        End If
        
    End With
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
Sub soma()
    Dim cTotal As Currency
    
    On Error GoTo trata_erro
    bActive = False
    With adc.Recordset
        .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !reman_itd
            .MoveNext
        Loop
        .MoveFirst
    End With
    lblTOtal.caption = FormatVal(cTotal)
    bActive = True
    Exit Sub
trata_erro:
    lblTOtal.caption = "0,00"
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    On Error Resume Next
    txtbxDataIni.Text = "01" & Mid(Format(Date, "DD/MM/YYYY"), 3, 8)
    strSQL = DateAdd("M", 1, CDate(txtbxDataIni.Text))
    txtbxDAtaFIn.Text = DateAdd("D", -1, strSQL)
    adc.ConnectionString = strConnect
    
    cmbxOrder.ListIndex = 0
    adc.ConnectionString = strConnect
    adcch.ConnectionString = strConnect
    
    resizeForm Me
    
    
    cmbxOrder.ListIndex = 0
    
    cmndbtnBuscar_Click
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    dtgrd.Width = Me.Width - 300
    
    dtgrd.Height = pcBot.Top - dtgrd.Top - 100
    
End Sub

Private Sub txtbxdataFin_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxdataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
        cmndbtnBuscar_Click
    End If
    
End Sub

Private Sub txtbxdataFin_LostFocus()
    P_AjustaData txtbxDAtaFIn
    cmndbtnBuscar_Click
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub


Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        cmndbtnTerc_Click
    End If
End Sub



Private Sub txtbxHistoricoDoc_GotFocus()
    ganhaFoco txtbxHistoricoDoc
End Sub

Private Sub txtCliente_GotFocus()
    ganhaFoco txtCliente
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCli
    End If
End Sub
Sub buscaCli()
    Dim c As tCliente
    c = Clientes.F_BuscaCliente("C", txtCliente.Text)
    If c.cod_cli <> "" Then
        txtCliente.Text = c.nome_Cli
        txtCliente.tag = c.cod_cli
        busca
    End If
    
End Sub
