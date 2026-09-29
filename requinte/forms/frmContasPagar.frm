VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmContasPagar 
   Caption         =   "Movimentações"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15765
   Icon            =   "frmContasPagar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   15765
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmContasPagar.frx":030A
      Height          =   6495
      Left            =   120
      TabIndex        =   28
      Top             =   720
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11456
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
      Caption         =   "Documentos a Pagar"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "seq_doc"
         Caption         =   "Id.Doc"
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
         DataField       =   "seq_itd"
         Caption         =   "Id.Parc"
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
         DataField       =   "num_doc"
         Caption         =   "Documento"
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
         DataField       =   "restante_itd"
         Caption         =   "Restam"
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
         DataField       =   "valor_itd"
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
         DataField       =   "vencimento_itd"
         Caption         =   "Vencimento"
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
         DataField       =   "data_doc"
         Caption         =   "Data"
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
         DataField       =   "historico_itd"
         Caption         =   "Histórico"
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
      BeginProperty Column08 
         DataField       =   "nome_pla"
         Caption         =   "Classificação"
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
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   4080,189
         EndProperty
         BeginProperty Column08 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   15705
      TabIndex        =   17
      Top             =   8295
      Width           =   15765
      Begin VB.CommandButton cmndbtnCarne 
         Caption         =   "Imprimir Grade"
         Height          =   595
         Left            =   7200
         Picture         =   "frmContasPagar.frx":031C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   595
         Left            =   3600
         Picture         =   "frmContasPagar.frx":0418
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   595
         Left            =   9000
         Picture         =   "frmContasPagar.frx":0562
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   595
         Left            =   1800
         Picture         =   "frmContasPagar.frx":06AC
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   595
         Left            =   0
         Picture         =   "frmContasPagar.frx":0A36
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   595
         Left            =   10800
         Picture         =   "frmContasPagar.frx":0B80
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Controle de Cheques"
         Height          =   595
         Left            =   5400
         Picture         =   "frmContasPagar.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   15765
      TabIndex        =   16
      Top             =   0
      Width           =   15765
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         MaxLength       =   60
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDocumento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         MaxLength       =   60
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxPlanoContas 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   10440
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtbxTerceiro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6720
         TabIndex        =   7
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton cmndbtnTerc 
         Height          =   285
         Left            =   10080
         Picture         =   "frmContasPagar.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmContasPagar.frx":1D1E
         Left            =   12840
         List            =   "frmContasPagar.frx":1D2E
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   60
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "&Pesquisar"
         Height          =   480
         Left            =   14400
         Picture         =   "frmContasPagar.frx":1D5A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnPlanoContas 
         Height          =   285
         Left            =   12480
         Picture         =   "frmContasPagar.frx":2105
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   5640
         TabIndex        =   32
         Tag             =   "!"
         Top             =   0
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
         Height          =   195
         Left            =   4560
         TabIndex        =   30
         Tag             =   "!"
         Top             =   0
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de Custo:"
         Height          =   195
         Left            =   10440
         TabIndex        =   27
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Participante"
         Height          =   195
         Left            =   6720
         TabIndex        =   6
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar Por:"
         Height          =   195
         Left            =   12840
         TabIndex        =   9
         Tag             =   "!"
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Histórico"
         Height          =   195
         Left            =   2280
         TabIndex        =   0
         Tag             =   "!"
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Tag             =   "!"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Tag             =   "!"
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox pcTotal 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   15735
      TabIndex        =   12
      Top             =   8040
      Width           =   15765
      Begin VB.CheckBox chkAberto 
         Caption         =   "Apenas Parc. em Aberto"
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   2280
         TabIndex        =   15
         Top             =   0
         Width           =   405
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
         Left            =   2760
         TabIndex        =   14
         Top             =   -60
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmContasPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim WithEvents adoR As ADODB.Recordset
Attribute adoR.VB_VarHelpID = -1
Private Sub chkAberto_Click()
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxOrder_Click()
    If Not bActive Then Exit Sub
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
End Sub

'Private Sub cmndbtnBaixar_Click1()
'    Dim sSEQ As String
'    Dim iCount As Integer
'    Dim c As cheque
'    Dim sChe As String
'    Dim sEsp As String
'    Dim cValor As Currency
'    Dim cJuros As Currency
'    With ador
'        If .RecordCount = 0 Then
'            MsgBox "Não há o que baixar!", vbExclamation
'            Exit Sub
'        End If
'            cJuros = F_CalculaJuros(!venc_itd, Date, !reman_itd)
'            frmBaixaPar.tag = ador!seq_itd
'            frmBaixaPar.lblNome_Ter.caption = !Nome
'            If cmbxBuscar.ListIndex = 0 Then
'                frmBaixaPar.txtbxJuros.Text = FormatVal(cJuros)
'            End If
'            frmBaixaPar.txtbxReman.Text = FormatVal(!reman_itd)
'            frmBaixaPar.txtbxValor_Par.Text = FormatVal(!valor_itd)
'            frmBaixaPar.lblParc.caption = "Parcela: " & !parcela_itd & " de: " & !parcelas_doc
'            frmBaixaPar.cmndbtnConfirma.tag = "P"
'            frmBaixaPar.Show 1
'
'            If frmBaixaPar.tag = "" Then
'                Unload frmBaixaPar
'                Exit Sub
'            Else
'                If MsgBox("Confirma baixa da Parcela?", vbYesNo + vbQuestion) = vbYes Then
'                    With frmBaixaPar.msflxgrd
'                        For iCount = 1 To .rows - 1
'                            .Row = iCount
'                            .Col = 2
'                            sEsp = Left(.Text, 1)
'                            .Col = 0: c.Venc_che = .Text
'                            .Col = 1: c.Valor_che = .Text
'                            .Col = 3: c.num_che = .Text
'                            .Col = 4: c.cod_bco = .Text
'                            .Col = 5: c.age_che = .Text
'                            .Col = 6: c.cco_che = .Text
'                            .Col = 7: c.nome_che = .Text
'                            If sEsp = "1" Then
'                                With ador
'                                    If frmBaixaPar.tag = 2 Then
'                                    sSEQ = F_GravaFinanceiro("P", CStr(Date), CStr(c.Valor_che), _
'                                        "Baixa doc " & !num_doc & "-" & "", _
'                                        frmBaixaPar.cmbxConta.tag, "S", ador!cod_pla)
'                                    Else
'                                        sSEQ = F_GravaCai(Left(cmbxTIpo.Text, 1), CStr(Date), CStr(c.Valor_che), _
'                                        "Baixa doc " & !num_doc & "-" & !Nome, "", _
'                                            "DIN", strCod_Cai, CStr(c.Valor_che), "0", strCod_Cai, strCod_Tur, ador!cod_pla)
'                                    End If
'
'                                End With
'                                 'sSeq = F_GravaFinanceiro(Left(cmbxTipo.Text, 1), Date, CStr(c.Valor_che), "HISTORICO", "DIN", _
'                                        frmBaixaPar.cmbxConta.Tag, ador!seq_itd, "1", "")
'
'                                        If F_GravaBaixaItd(F_BuscaChaveBaixa, ador!seq_itd, CStr(c.Valor_che), "DIN", sSEQ, "") = "" Then
'                                            MsgBox "Erro na gravação da baixa!", vbExclamation
'                                        End If
'
'                                'P_Class c.Valor_che, Left(cmbxTipo.Text, 1), sSEQ, "F", ador!historico_doc
'                            ElseIf sEsp = "2" Then
'                                sSEQ = F_GravaDoc("BAIXA DOC:" & ador!num_doc & "-" & ador!Nome, _
'                                     ador!num_doc, CStr(c.Valor_che), Date, _
'                                    ador!Codigo, Left(cmbxTIpo.Text, 1), "", "1", "F", "C", ador!cod_pla)
'                                    If F_GravaBaixaItd(F_BuscaChaveBaixa, ador!seq_itd, CStr(c.Valor_che), "DUP", sSEQ, "") = "" Then
'                                            MsgBox "Erro na gravação da baixa!", vbExclamation
'                                        End If
'
'                                If Left(cmbxTIpo.Text, 1) = "R" Then
'                                    sChe = F_GravaCh(sSEQ, c.num_che, c.cod_bco, c.age_che, c.cco_che, _
'                                            c.nome_che, CCur(c.Valor_che), c.Venc_che, c.cod_bco, Left(cmbxTIpo.Text, 1), _
'                                            "", "", ador!parcela_itd, "", "")
'
'                                    If F_GravaBaixaItd(F_BuscaChaveBaixa, ador!seq_itd, CStr(c.Valor_che), "CHP", sChe, "") = "" Then
'                                            MsgBox "Erro na gravação da baixa!", vbExclamation
'                                        End If
'                                Else
'                                    sChe = F_GravaChePag(sSEQ, c.num_che, c.cod_bco, _
'                                            CCur(c.Valor_che), c.Venc_che, ador!Codigo, ador!parcela_itd)
'                                End If
'                            Else
'                                sSEQ = MaxKeySQL("tab_venfin", "num_ven", "num_ven like 'B%'", "B00000#")
'                                'sSEQ = F_GravaVenFin(sSEQ, c.num_che, c.Valor_che, Date, strCod_Usu, ador!Codigo, c.cod_bco)
'                                If F_GravaBaixaItd(F_BuscaChaveBaixa, ador!seq_itd, CStr(c.Valor_che), "FIN", sSEQ, "") = "" Then
'                                            MsgBox "Erro na gravação da baixa!", vbExclamation
'                                        End If
'                            End If
'
'                            cValor = cValor + c.Valor_che
'                        Next
'                        strSQL = "update tab_itdoc set reman_itd = reman_itd-" & F_GravaDec(cValor - CCur(frmBaixaPar.txtbxJuros.Text) + CCur(frmBaixaPar.txtbxDesconto.Text))
'                        strSQL = strSQL & ",juros_itd=" & F_GravaDec(frmBaixaPar.txtbxJuros.Text)
'                        strSQL = strSQL & ",Desconto_itd=" & F_GravaDec(frmBaixaPar.txtbxDesconto.Text)
'                        strSQL = strSQL & ", baixa_itd=" & F_GravaData(Date)
'                        strSQL = strSQL & " where seq_itd='" & ador!seq_itd & "' "
'                        conn.execute strSQL
'                        Unload frmBaixaPar
'                        MsgBox "Documento Baixado!", vbInformation
'                        adc.Refresh
'                    End With
'                End If
'            End If
'    End With
'End Sub

Private Sub cmndbtnBaixar_Click()
    
    With adoR
        cJuros = F_CalculaJuros(!venc_itd, Date, !reman_itd)
        frmBaixaPar.tag = adoR!seq_itd
        frmBaixaPar.lblNome_Ter.caption = !nome
        If cmbxBuscar.ListIndex = 0 Then
            frmBaixaPar.txtbxJuros.Text = FormatVal(cJuros)
        End If
        frmBaixaPar.txtbxReman.Text = FormatVal(!reman_itd)
        frmBaixaPar.txtbxValor_Par.Text = FormatVal(!valor_itd)
        frmBaixaPar.lblParc.caption = "Parcela: " & !parcela_itd & " de: " & !parcelas_doc
        frmBaixaPar.cmndbtnConfirma.tag = "P"
        
        frmBaixaPar.Show 1
    End With
    Unload frmBaixaPar
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    Dim sCampo As String
    On Error GoTo trata_Erro
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sTipoTer As String
    
    
    If cmbxOrder.ListIndex = 0 Then
        sOrderBy = "venc_Itd"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderBy = "Nome"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderBy = "num_doc,venc_itd"
    Else
        sOrderBy = "reman_itd"
    End If
    
    strSQL = "select td.cod_pla,  tpc.nome_pla,tc.cod_for as codigo,td.cod_ter,td.parcelas_doc,td.num_ven,historico_doc,"
    strSQL = strSQL & " (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, "
    strSQL = strSQL & "tc.nome_for as nome,td.valor_Doc,td.parcelas_doc,td.num_doc,"
    strSQL = strSQL & " td.data_doc, td.valor_doc,"
    strSQL = strSQL & "ti.* from tab_fornecedores as tc, tab_doc as td,tab_itdoc as ti ,tab_planocontas as tpc  "
        strSQL = strSQL & " where tc.cod_for=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & " and td.tipo_doc='P' "
        strSQL = strSQL & " and tpc.cod_pla=td.cod_pla and tpc.tipo_pla='D' "
        If chkAberto.Value = 1 Then
            strSQL = strSQL & " and ti.reman_itd > 0 "
        End If
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd <= " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd >= " & F_GravaData(txtbxDataIni.Text)
        End If
        
            strSQL = strSQL & " and (td.num_doc like  '" & txtbxExpr.Text & "%' "
        
            strSQL = strSQL & " or  historico_doc like  '%" & txtbxExpr.Text & "%' )"
        
            
        
            
        
        
        If txtbxTerceiro.tag <> "" Then
            strSQL = strSQL & " and cod_ter='" & txtbxTerceiro.tag & "' "
        End If
        If txtbxPlanoContas.tag <> "" Then
            strSQL = strSQL & " and cod_pla='" & txtbxPlanoContas.tag & "'"
        End If
        strSQL = strSQL & " and tipo_ter='F' "
        strSQL = strSQL & " order by " & sOrderBy
        
    FechaObjeto adoR
    adoR.Open strSQL, Conn_Client
    Set dtgrd.DataSource = adoR
    soma
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
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

Private Sub cmndbtnCarne_Click()
    P_DbgrdToHtml "Movimentação " & cmbxTipo.Text & " de " & txtbxDataIni.Text & " Até " & txtbxdataFin.Text, dtgrd, adoR
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmEdiLan
    On Error GoTo trata_Erro
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    With adoR
        If .RecordCount < 0 Then
            Exit Sub
        End If
        fr.tag = !seq_itd
        fr.lblData.caption = !data_doc
        fr.lblNUM_Doc.caption = !num_doc
        fr.lblParc.caption = !parcela_itd & "/" & !parcelas_doc
        fr.lblSeq_Doc.caption = !seq_doc
        fr.lblSeq_Itd.caption = !seq_itd
        fr.lblValor.caption = FormatVal(!valor_itd)
        fr.txtbxHist.Text = !historico_doc
        fr.txtbxNOme.Text = !nome
        fr.txtbxNOme.tag = !codigo
        fr.txtbxVenc.Text = !venc_itd
        fr.txtbxValor.Text = FormatVal(!reman_itd)
        fr.txtbxBoleto.Text = !boleto_itd
        fr.txtbxClass.Text = !nome_pla
        fr.txtbxClass.tag = !cod_pla
        'fr.lblTipo.caption = "Débito"
        fr.Show 1
        adc.Refresh
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
    
    With adoR
        If !num_ven = "" Then
            If F_BuscaUmCampo("reman_itd", "tab_itdoc", "reman_itd<valor_itd and seq_doc='" & !seq_doc & "'") = "" Then
                If MsgBox("Deseja realmente excluir o documento e suas parcelas?", vbYesNo + vbCritical) = vbYes Then
                    strSQL = "delete from tab_doc where seq_doc='" & !seq_doc & "' "
                    Conn.Execute strSQL
                    strSQL = "delete from tab_itdoc where seq_doc='" & !seq_doc & "' "
                    Conn.Execute strSQL
                    strSQL = "delete from tab_cheques where seq_doc='" & !seq_doc & "' "
                    Conn.Execute strSQL
                    adc.Refresh
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

Private Sub cmndbtnImpr_Click()
    frmControleChequesPagar.Show 0
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.optnCh.Value = False
    fr.optnPag.Value = True
    
    fr.Show 1
End Sub

Private Sub cmndbtnPlanoContas_Click()
    Dim p As PlanoContas
    p = F_BuscaPlanoContas("D", txtbxPlanoContas.Text)
    If p.cod_pla <> "" Then
        txtbxPlanoContas.tag = p.cod_pla
        txtbxPlanoContas.Text = p.nome_pla
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.BorderStyle = 1
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.BorderStyle = 0
End Sub

Private Sub cmndbtnTerc_Click()
    
    
    
    buscaFor
    
    
    cmndbtnBuscar_Click
End Sub
Sub buscaCli()
    Dim c As cliente
    c = F_BuscaCliente(txtbxTerceiro.Text)
        If c.cod_cli <> "" Then
            txtbxTerceiro.Text = c.nome_cli
            txtbxTerceiro.tag = c.cod_cli
        End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
        f = F_SelectForn(txtbxTerceiro.Text)
        If f.cod_For <> "" Then
            txtbxTerceiro.Text = f.nome_for
            txtbxTerceiro.tag = f.cod_For
        End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxTerceiro.Text)
        If r.cod_Rep <> "" Then
            txtbxTerceiro.Text = r.nome_rep
            txtbxTerceiro.tag = r.cod_Rep
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


Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dbgrdChe_GotFocus()
    dbgrdChe.MarqueeStyle = 3
End Sub

Private Sub dbgrdChe_LostFocus()
    dbgrdChe.MarqueeStyle = 0
End Sub

Private Sub adc_Reposition()
    Dim strSQL As String
    On Error Resume Next
    If bActive = False Then Exit Sub
    With adoR
        If .RecordCount = 0 Then Exit Sub
        strSQL = "select * from tab_cheques where seq_doc='" & !seq_doc & "' "
        dtChe.RecordSource = strSQL
        dtChe.Refresh
        If dtChe.Recordset.RecordCount > 0 Then
            dbgrdChe.caption = "Cheques referentes ao documento " & !num_doc
            dbgrdChe.Visible = True
        Else
            dbgrdChe.Visible = False
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
        If !tipo_doc = "R" Then
            cmndbtnCarne.Enabled = True
        Else
            cmndbtnCarne.Enabled = False
        End If
    End With
End Sub

Private Sub Form_Activate()
   ' MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
Sub soma()
    Dim cTotal As Currency
    On Error GoTo trata_Erro
    bActive = False
    With adoR
        .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !reman_itd
            .MoveNext
        Loop
        .MoveFirst
    End With
    lblTotal.caption = FormatVal(cTotal)
    bActive = True
    Exit Sub
trata_Erro:
    lblTotal.caption = "0,00"
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    On Error Resume Next
    resizeForm Me
   ' txtbxDataIni.Text = F_PrimDiaMes(Date)
    
    txtbxdataFin.Text = F_UltDiaMes(Date)
    adc.ConnectionString = Conn.ConnectionString
    cmbxTipo.ListIndex = 0
    cmbxOrder.ListIndex = 0
    Set adoR = New ADODB.Recordset
    'P_CentraForm Me
'    strSQL = "select tn.*,tt.* from tab_compras as tn, tab_fornecedores as tt "
'    strSQL = strSQL & " where tt.cod_for=tn.cod_for "
'
'    strSQL = strSQL & " order by num_not "
'    adc.RecordSource = strSQL
'    adc.Refresh
    cmbxBuscar.ListIndex = 1
    cmbxOrder.ListIndex = 0
    cmbxTipo.ListIndex = 0
    cmndbtnBuscar_Click
End Sub

Private Sub optnCli_Click()
    frExpr.Enabled = True
End Sub

Private Sub optnTodas_Click()
    frExpr.Enabled = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dtgrd.width = Me.width - dtgrd.Left - 300
    
    dtgrd.height = pcTotal.top - dtgrd.top - 100
    
    
End Sub

Private Sub pcTop_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub txtbxContaCob_Change()
    txtbxContaCob.tag = ""
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
    cmndbtnBuscar_Click
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

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    
        P_AjustaTexto txtbxExpr
    
End Sub

Private Sub txtbxPlanoContas_Change()
    txtbxPlanoContas.tag = ""
End Sub

Private Sub txtbxPlanoContas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnPlanoContas_Click
    End If
End Sub

Private Sub txtbxTerceiro_Change()
    txtbxTerceiro.tag = ""
End Sub

Private Sub txtbxTerceiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxTipoCob_Change()
    txtbxTipoCob.tag = ""
End Sub
