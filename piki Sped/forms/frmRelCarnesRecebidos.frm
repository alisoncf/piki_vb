VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelCarnesRecebidos 
   Caption         =   "Relatório de Carnês Recebidos"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16350
   Icon            =   "frmRelCarnesRecebidos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   16350
   Begin MSDataGridLib.DataGrid GridBaixas 
      Height          =   1215
      Left            =   120
      TabIndex        =   30
      Top             =   1320
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   2143
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
      Caption         =   "Relação de Baixas"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "data_bai"
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
      BeginProperty Column01 
         DataField       =   "num_Doc"
         Caption         =   "Nº Documento"
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
         DataField       =   "seq_doc"
         Caption         =   "Id. Documento"
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
         DataField       =   "seq_itd"
         Caption         =   "Id. Parcela"
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
         DataField       =   "terceiro"
         Caption         =   "Cliente"
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
         DataField       =   "valor_bai"
         Caption         =   "Baixado"
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
         DataField       =   "Cod_usu"
         Caption         =   "Cod Usu."
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
         DataField       =   "data_bai"
         Caption         =   "Data Baixa"
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
         DataField       =   "hora_bai"
         Caption         =   "Hora"
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
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4350,047
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1110,047
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   705,26
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   16290
      TabIndex        =   1
      Top             =   0
      Width           =   16350
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmRelCarnesRecebidos.frx":0442
         Left            =   4080
         List            =   "frmRelCarnesRecebidos.frx":0455
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxSeqDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxSeqItd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmRelCarnesRecebidos.frx":0480
         Left            =   5280
         List            =   "frmRelCarnesRecebidos.frx":048A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   11520
         Picture         =   "frmRelCarnesRecebidos.frx":04A6
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1335
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxContaCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnTerc 
         Height          =   285
         Left            =   2400
         Picture         =   "frmRelCarnesRecebidos.frx":0830
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         TabIndex        =   25
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton cmndbtnBuscaCli 
         Height          =   285
         Left            =   11040
         Picture         =   "frmRelCarnesRecebidos.frx":0BBA
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordem:"
         Height          =   195
         Left            =   4080
         TabIndex        =   29
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   7080
         TabIndex        =   27
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Num Doc:"
         Height          =   195
         Left            =   2880
         TabIndex        =   24
         Tag             =   "!"
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seq.Doc:"
         Height          =   195
         Left            =   1320
         TabIndex        =   21
         Tag             =   "!"
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seq.Itd:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Tag             =   "!"
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Terceiro:"
         Height          =   195
         Left            =   5280
         TabIndex        =   11
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   4080
         TabIndex        =   7
         Tag             =   "!"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   2880
         TabIndex        =   6
         Tag             =   "!"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Cobrança(Da Baixa):"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   16320
      TabIndex        =   0
      Top             =   8160
      Width           =   16350
      Begin VB.PictureBox pcBotoes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   3240
         ScaleHeight     =   855
         ScaleWidth      =   11055
         TabIndex        =   14
         Top             =   0
         Width           =   11055
         Begin VB.CommandButton cmdReciboEmail 
            Caption         =   "Recibo pro Email"
            Height          =   735
            Left            =   4200
            Picture         =   "frmRelCarnesRecebidos.frx":0F44
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   0
            Width           =   1935
         End
         Begin VB.CommandButton cmndbnCancelarBaixa 
            Caption         =   "&Cancelar Baixa"
            Height          =   735
            Left            =   120
            Picture         =   "frmRelCarnesRecebidos.frx":1040
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   0
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnImprimir 
            Caption         =   "Im&primir"
            Height          =   735
            Left            =   6360
            Picture         =   "frmRelCarnesRecebidos.frx":13CA
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnImprimirRecibo 
            Caption         =   "&Recibo"
            Height          =   735
            Left            =   2160
            Picture         =   "frmRelCarnesRecebidos.frx":17D7
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   0
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Sair"
            Height          =   735
            Left            =   8520
            Picture         =   "frmRelCarnesRecebidos.frx":18D3
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   960
         TabIndex        =   12
         Top             =   120
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmRelCarnesRecebidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmbxFunc_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTipoTer_Click()
    txtbxCliente.tag = ""
    txtbxCliente.text = ""
End Sub

Private Sub cmdReciboEmail_Click()
    On Error GoTo trata_Erro
    EnviarEmailRecibo ado!seq_cab, ado!codigo
trata_Erro:
    
End Sub

Private Sub cmndbnCancelarBaixa_Click()
    If Not F_senha(True) Then
        Exit Sub
    End If
    If strExcldoc_usu <> "S" Then
        MsgBox "VocÊ não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma cancelamento da baixa?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    P_CancelaBaixaiItd ado!seq_cab
End Sub

Private Sub cmndbtnBuscaCli_Click()
    If cmbxTipoTer.ListIndex = 0 Then
        buscaCli
    Else
        buscaRep
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sCampo As String
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sTipoTer As String
    
    On Error GoTo trata_Erro
       
       
       
       
    
    P_Status "Aguarde..."
    If cmbxTipoTer.ListIndex = 0 Then
            sTab = TAB_CLIENTES
            sCod = "cod_cli"
            sNome = "nome_cli"
            sTipoTer = "C"
    Else
            sTab = "tab_representantes"
            sNome = "nome_rep"
            sCod = "cod_rep"
            sTipoTer = "R"
    End If
    
    strSQL = "select tc." & sNome & " as terceiro, tb.cod_cob as cobranca,tb.seq_bai,tb.seq_cab,tb.valor_bai,tb.hora_bai,tb.esp_bai,tb.data_bai,tb.cod_usu,ti.*,"
    strSQL = strSQL & "td.num_doc,"
    strSQL = strSQL & "td.historico_doc,td.data_doc,td.parcelas_doc,td.cod_ter as codigo "
    strSQL = strSQL & " from " & sTab & " as tc, " & TAB_BAIXAITD & "  as tb,  "
    strSQL = strSQL & "  " & TAB_ITDOC & "  as ti,  " & TAB_DOC & "  as td "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc and ti.seq_itd=tb.seq_itd "
    
    strSQL = strSQL & " and tc." & sCod & "=td.cod_ter "
    If txtbxContaCob.tag <> "" Then
        strSQL = strSQL & " and tb.cod_cob='" & txtbxContaCob.tag & "'"
    
    End If
    If txtbxCliente.tag <> "" Then
        strSQL = strSQL & " and td.cod_ter='" & txtbxCliente.tag & "' "
    End If
    If txtbxDataFin.text <> "" Then
        strSQL = strSQL & " and data_bai<=" & F_GravaData(txtbxDataFin.text)
    End If
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and data_bai>=" & F_GravaData(txtbxDataIni.text)
    End If
    If txtbxSeqDoc.text <> "" Then
        strSQL = strSQL & " and td.seq_doc='" & txtbxSeqDoc.text & "' "
    End If
    If txtbxSeqItd.text <> "" Then
        strSQL = strSQL & " and ti.seq_itd='" & txtbxSeqItd.text & "' "
    End If
    If txtbxNumDoc.text <> "" Then
        strSQL = strSQL & " and num_doc='" & txtbxNumDoc.text & "' "
    End If
    strSQL = strSQL & " and td.tipo_ter='" & sTipoTer & "' "
    If cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by data_bai,terceiro"
    ElseIf cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by terceiro, data_bai"
    ElseIf cmbxOrder.ListIndex = 2 Then
        strSQL = strSQL & " order by td.seq_doc,terceiro, data_bai"
    ElseIf cmbxOrder.ListIndex = 3 Then
        strSQL = strSQL & " order by ti.seq_itd,terceiro, data_bai"
    ElseIf cmbxOrder.ListIndex = 4 Then
        strSQL = strSQL & " order by cobranca, ti.seq_itd,terceiro, data_bai"
    Else
    
        strSQL = strSQL & " order by data_bai,venc_itd "
    End If
    abreObjetoAdo ado, strSQL, gridBaixas
    soma
    P_Status ""
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub soma()
    Dim cTotal As Currency
    
    gridBaixas.Visible = False
    With ado
        Do While Not .EOF
            cTotal = cTotal + !valor_bai
            .MoveNext
        Loop
        lblTotal.caption = FormatVal(cTotal)
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        
    End With
    gridBaixas.Visible = True
End Sub

Private Sub cmndbtnCliente_Click()
    buscaRep
End Sub
Sub buscaRep()
    Dim r As representante.representante
    r = representante.F_BuscaRep(txtbxCliente.text)
    If r.cod_Rep <> "" Then
        
        txtbxCliente.text = r.nome_rep
        txtbxCliente.tag = r.cod_Rep
        
        
        
        
    End If
    
    
End Sub
Sub buscaCli()
    Dim c As cliente
    c = Clientes.F_BuscaCliente(txtbxCliente.text)
    If c.cod_cli <> "" Then
        txtbxCliente.text = c.nome_Cli
        txtbxCliente.tag = c.cod_cli
    End If
    
    
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DtgrdToHtml "Relatório de Carnês Recebidos", gridBaixas, ado, "Conta: " & txtbxContaCob.text & _
    " Venc. Inicial: " & txtbxDataIni.text & " Venc. Final: " & txtbxDataFin.text & " - Total: " & lblTotal.caption
End Sub

Private Sub cmndbtnImprimirRecibo_Click()
   
   ImprReciboBai ado!seq_cab
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub
Sub imprime(obj As Object)
    P_CabRelTimbr "Relatório de Comissões de " & txtbxDataIni.text & " até " & txtbxDataFin.text, obj
    P_ImprLinha obj, "-"
    P_ImprX obj, "Vendedor", 10
    P_ImprX obj, "Vendido", 70, True, True
    P_ImprX obj, "Comissão", 100, True, True
    P_ImprX obj, "Perc.", 115, False, 3
    P_ImprLinha obj, "-"
    With adc.Recordset
        .MoveFirst
        Do While Not .EOF
            P_ImprX obj, !guerra_fun, 10
            P_ImprX obj, FormatVal(!total), 70, True, True
            P_ImprX obj, FormatVal(!comiss), 100, True, True
            P_ImprX obj, FormatVal(!total / CCur(lblVendas.caption) * 100), 115, False, 3

            .MoveNext
        Loop
        P_RodapeVis obj, strNomeSis
    End With
    
    
End Sub

Private Sub cmndbtnTerc_Click()
    Dim c As ContaCob
    c = F_BuscaContaCob("")
    txtbxContaCob.text = c.nome_cob
    txtbxContaCob.tag = c.cod_cob
    
    
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    'P_PreencheTipoCob cmbxTipoCob
    resizeForm Me
    cmbxTipoTer.ListIndex = 0
    
End Sub

Private Sub Form_Resize()
    pcBotoes.Left = pcBot.width - pcBotoes.width - 100
    
    PosicionaGrade gridBaixas, Me, pcBot
    
    
End Sub

Private Sub txtbxCliente_Change()
    txtbxCliente.tag = ""
End Sub

Private Sub txtbxCliente_GotFocus()
    P_SelecTexto txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaCli_Click
    End If
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDataFin
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

Private Sub txtbxNumDoc_GotFocus()
    P_SelecTexto txtbxNumDoc
End Sub

Private Sub txtbxNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNumDoc_LostFocus()
    txtbxNumDoc.text = Format(txtbxNumDoc.text, "000000#")
End Sub
