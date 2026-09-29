VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiCai 
   Caption         =   "Edição de Fluxo de Caixa"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   Icon            =   "frmEdiCai.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11070
   Begin MSDataGridLib.DataGrid dtg 
      Height          =   5055
      Left            =   0
      TabIndex        =   18
      Top             =   840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8916
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "seq_cai"
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
      BeginProperty Column01 
         DataField       =   "cod_cai"
         Caption         =   "Código"
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
         DataField       =   "nome_cai"
         Caption         =   "Caixa"
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
         DataField       =   "data_cai"
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
      BeginProperty Column04 
         DataField       =   "valor_cai"
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
         DataField       =   "tipo_cai"
         Caption         =   "Tipo"
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
         DataField       =   "historico_cai"
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
      BeginProperty Column07 
         DataField       =   "esp_cai"
         Caption         =   "Espécie"
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
         DataField       =   "num_ven"
         Caption         =   "venda"
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
      BeginProperty Column09 
         DataField       =   "cod_usu"
         Caption         =   "Usuário"
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
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2250,142
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcbot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11040
      TabIndex        =   19
      Top             =   6150
      Width           =   11070
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   9240
         Picture         =   "frmEdiCai.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   7440
         Picture         =   "frmEdiCai.frx":02B5
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5640
         Picture         =   "frmEdiCai.frx":06C2
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3840
         Picture         =   "frmEdiCai.frx":080C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   2040
         Picture         =   "frmEdiCai.frx":0956
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblTotal_Din 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
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
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   660
         TabIndex        =   30
         Tag             =   "!"
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   0
         TabIndex        =   29
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Débito:"
         Height          =   195
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblDeb_Din 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   660
         TabIndex        =   27
         Tag             =   "!"
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito:"
         Height          =   195
         Left            =   0
         TabIndex        =   26
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblCred_din 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   660
         TabIndex        =   25
         Tag             =   "!"
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.PictureBox pcLeft 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   11040
      TabIndex        =   0
      Top             =   0
      Width           =   11070
      Begin VB.ComboBox cmbxTurno 
         Height          =   315
         ItemData        =   "frmEdiCai.frx":0CE0
         Left            =   2760
         List            =   "frmEdiCai.frx":0CED
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbxCai 
         Height          =   315
         ItemData        =   "frmEdiCai.frx":0D0B
         Left            =   1440
         List            =   "frmEdiCai.frx":0D18
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8640
         MaxLength       =   40
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   405
         Left            =   10200
         Picture         =   "frmEdiCai.frx":0D36
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton optnDesc 
         Caption         =   "Desc"
         Height          =   255
         Left            =   7800
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optnAsc 
         Caption         =   "Asc"
         Height          =   255
         Left            =   7800
         TabIndex        =   13
         Top             =   120
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiCai.frx":10C0
         Left            =   4200
         List            =   "frmEdiCai.frx":10D0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbxTIpo 
         Height          =   315
         ItemData        =   "frmEdiCai.frx":10F8
         Left            =   120
         List            =   "frmEdiCai.frx":1105
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         Height          =   195
         Left            =   2760
         TabIndex        =   5
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   0
         Width           =   435
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   8640
         TabIndex        =   15
         Top             =   0
         Width           =   780
      End
      Begin VB.Label lbldataFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   6600
         TabIndex        =   11
         Top             =   0
         Width           =   765
      End
      Begin VB.Label lblDataIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   5520
         TabIndex        =   9
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   4200
         TabIndex        =   7
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmEdiCai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub Text1_Change()

End Sub

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 '
    movedata
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    If txtbxDAtaFIn.Visible = True Then
        If F_DataMaior(txtbxDataIni.Text, txtbxDAtaFIn.Text) Then
            MsgBox "A data final deve ser maior que a inicial!", vbExclamation
            txtbxDAtaFIn.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "select tc.*,tb.nome_cai from tab_caixa as tc, tab_caixas as tb where "
    strSQL = strSQL & "tc.cod_cai=tb.cod_cai  and "
    
        strSQL = strSQL & "( historico_cai like '%" & txtbxExpr.Text & "%'"
    
        strSQL = strSQL & " or num_ven like '" & txtbxExpr.Text & "%'"
        If F_GravaDec(txtbxExpr.Text) > 0 Then
            strSQL = strSQL & " or valor_cai = " & F_GravaDec(txtbxExpr.Text)
        End If
    strSQL = strSQL & ") "
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_cai >= " & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxDAtaFIn.Text <> "" Then
        strSQL = strSQL & " and data_cai <= " & F_GravaData(txtbxDAtaFIn.Text)
    End If
    If cmbxTipo.Text <> "" Then
        strSQL = strSQL & " and tipo_cai='" & Left(cmbxTipo.Text, 1) & "' "
    End If
    
    If cmbxCai.Text <> "" Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCai.Text, 2) & "'"
    End If
    If cmbxTurno.Text <> "" Then
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    End If
    
    If cmbxOrd.ListIndex = 0 Then
        strSQL = strSQL & " order by data_cai,seq_cai"
    ElseIf cmbxOrd.ListIndex = 1 Then
        strSQL = strSQL & " order by seq_cai "
    ElseIf cmbxOrd.ListIndex = 2 Then
        strSQL = strSQL & " order by historico_cai, seq_cai"
    ElseIf cmbxOrd.ListIndex = 3 Then
        strSQL = strSQL & " order by valor_cai,seq_cai"
    Else
        strSQL = strSQL & " order by seq_cai  "
    End If
    
    If optnAsc.Value Then
        strSQL = strSQL & " asc "
    Else
        strSQL = strSQL & " desc "
    End If
    abreObjetoAdo ado, strSQL, dtg
    P_Resumo
    Exit Sub
trata_erro:
    MostraErro
End Sub
Sub P_Resumo()
Dim cCheDeb As Currency
    Dim cCheCred As Currency
    Dim cDinDeb As Currency
    Dim cDinCred As Currency
    
    With ado
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            
            
                If !tipo_cai = "C" Then
                    cDinCred = cDinCred + !valor_Cai
                Else
                    cDinDeb = cDinDeb + !valor_Cai
                End If
            
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    'lblCred_Che.caption = FormatVal(cCheCred)
    lblCred_din.caption = FormatVal(cDinCred)
    lblDeb_Din.caption = FormatVal(cDinDeb)
 '   lblDeb_Che.caption = FormatVal(cCheDeb)
  '  lblTotal_Che.caption = FormatVal(cCheCred - cCheDeb)
    lblTotal_Din.caption = FormatVal(cDinCred - cDinDeb)
'    lblSaldo.caption = FormatVal(CCur(lblTotal_Che.caption) + CCur(lblTotal_Din.caption))
End Sub
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmLanCaixa
    With ado
        fr.tag = !seq_cai
        fr.txtbxData_Cai.Text = !data_cai
        fr.txtbxHist.Text = !historico_cai
        fr.txtbxValor_Cai.Text = FormatVal(!valor_Cai)
        If !tipo_cai = "C" Then
            fr.optnCerd.Value = True
        Else
            fr.optnDeb.Value = True
        End If
        fr.optnCerd.Enabled = False
        fr.optnDeb.Enabled = False
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    If ado.RecordCount = 0 Then
        cmndbtnExcl.Enabled = False
        Exit Sub
    End If
    If MsgBox("Confirma exclusão do Lançamento? ", vbYesNo + vbCritical) = vbYes Then
        strSQL = "Delete from tab_caixa where seq_cai ='" & ado!seq_cai & "'"
        conn.Execute strSQL
        adc.Refresh
    End If
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DbgrdToHtml "Relatório de Fluxo de Caixa de " & _
    txtbxDataIni.Text & " até " & txtbxDAtaFIn.Text, dtg, ado
    
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanCaixa
    fr.Show 0
 End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    
    FechaObjeto ado
    Unload Me
End Sub

Private Sub dbgrd_GotFocus()
    
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub movedata()
    On Error GoTo trata_erro
    If ado.RecordCount = 0 Or dtg.Row = -1 Then
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
        Exit Sub
    
    End If
    
    If ado!num_ven _
     <> "" Then
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
    Else
        cmndbtnEdit.Enabled = True
        cmndbtnExcl.Enabled = True
    End If
    Exit Sub
trata_erro:
    cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
End Sub

Private Sub Form_Activate()
    bActive = True
End Sub

Private Sub Form_Load()
    resizeForm Me
    
    Set ado = New ADODB.Recordset
    txtbxDAtaFIn.Text = Date
    txtbxDataIni.Text = Date
    
    
    P_PreencheCaixas cmbxCai
    P_PreencheTurnos cmbxTurno
    
End Sub

Private Sub Form_Resize()
    PosicionaGrade dtg, Me, pcBot
    posicionaADireita cmndbtnSair, pcBot
End Sub

Private Sub txtbxdataFin_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxdataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxdataFin_LostFocus()
    P_AjustaData txtbxDAtaFIn
    If F_DataMaior(txtbxDataIni.Text, txtbxDAtaFIn.Text) Then
        MsgBox "A data final deve ser maior que a inicial!", vbExclamation
        txtbxDAtaFIn.SetFocus
        Exit Sub
    End If
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

Private Sub txtbxExpr_GotFocus()
    ganhaFoco txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    If cmbxBusca.ListIndex = 1 Then
        txtbxExpr.Text = Format(F_SoNumero(txtbxExpr.Text), "000000#")
    ElseIf cmbxBusca.ListIndex = 2 Then
        PerdeFoco txtbxExpr
    Else
        AjustaTexto txtbxExpr
    End If
End Sub
