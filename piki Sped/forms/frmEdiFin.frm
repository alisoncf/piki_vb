VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiFin 
   Caption         =   "Financeiro"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15195
   Icon            =   "frmEdiFin.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   15195
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   15165
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      Begin VB.TextBox txtConta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   29
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10320
         MaxLength       =   60
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Pesquisar"
         Height          =   525
         Left            =   13080
         Picture         =   "frmEdiFin.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7320
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   11
         Top             =   240
         Width           =   990
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         MaxLength       =   60
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiFin.frx":04F5
         Left            =   1680
         List            =   "frmEdiFin.frx":0505
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxTIpo 
         Height          =   315
         ItemData        =   "frmEdiFin.frx":052D
         Left            =   120
         List            =   "frmEdiFin.frx":053A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdConta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         Picture         =   "frmEdiFin.frx":0551
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&onta:"
         Height          =   195
         Left            =   3360
         TabIndex        =   31
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   10320
         TabIndex        =   21
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lbldataFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   7320
         TabIndex        =   17
         Top             =   0
         Width           =   765
      End
      Begin VB.Label lblDataIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   6240
         TabIndex        =   16
         Top             =   0
         Width           =   840
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Histórico"
         Height          =   195
         Left            =   8400
         TabIndex        =   15
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox pcBotoes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   15195
      TabIndex        =   1
      Top             =   6750
      Width           =   15195
      Begin VB.CommandButton cmdGerarArquivo 
         Caption         =   "Exportar"
         Height          =   735
         Left            =   12120
         Picture         =   "frmEdiFin.frx":08FC
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   0
         Width           =   1590
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtDebito 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnAjustaSaldo 
         Caption         =   "C&alcular Saldo"
         Height          =   735
         Left            =   8760
         Picture         =   "frmEdiFin.frx":0D2B
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1590
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   13800
         Picture         =   "frmEdiFin.frx":0E75
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1590
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   3720
         Picture         =   "frmEdiFin.frx":0F6A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1590
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   735
         Left            =   7080
         Picture         =   "frmEdiFin.frx":12F4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1590
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   5400
         Picture         =   "frmEdiFin.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1590
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   10440
         Picture         =   "frmEdiFin.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   2520
         TabIndex        =   24
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito"
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Débito"
         Height          =   195
         Left            =   1440
         TabIndex        =   22
         Top             =   120
         Width           =   465
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiFin.frx":1912
      Height          =   5535
      Left            =   0
      TabIndex        =   19
      Top             =   960
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
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
      Caption         =   "Lançamentos Financeiros"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "nome_con"
         Caption         =   "Conta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cod_con"
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
         DataField       =   "data_fin"
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
      BeginProperty Column03 
         DataField       =   "historico_fin"
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
      BeginProperty Column04 
         DataField       =   "cred"
         Caption         =   "Crédito"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "deb"
         Caption         =   "Débito"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "esp_fin"
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
      BeginProperty Column07 
         DataField       =   "tipo_fin"
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
      BeginProperty Column09 
         DataField       =   "cod_pla"
         Caption         =   "Código PC"
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
      BeginProperty Column10 
         DataField       =   "usuario"
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
      BeginProperty Column11 
         DataField       =   "seq_fin"
         Caption         =   "Id.Fin."
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
            ColumnWidth     =   1904,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2849,953
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2399,811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2115,213
         EndProperty
         BeginProperty Column11 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub cmbxConta_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxConta.ListIndex = -1
    End If
End Sub

Private Sub cmbxOrd_Click()
    If bActive = False Then Exit Sub
    buscar
End Sub

Private Sub cmdConta_Click()
    buscaConta
End Sub
Sub buscaConta()
    Dim c As ContaCorrente
    c = F_BuscaContaCorrente(txtConta.text)
    If c.cod_con <> "" Then
        txtConta.text = c.nome_con
        txtConta.tag = c.cod_con
        buscar
    End If
    
    
    
    
    
End Sub

Private Sub cmdGerarArquivo_Click()
    Dim caminho As String
    caminho = BuscaNomeArquivoExterno("*.csv")
    
    If caminho = "" Then
        Exit Sub
    End If
    
    ExportaGridToCSV dtgrd, ado, ";"
     
     
End Sub

Private Sub cmndbtnAjustaSaldo_Click()
    Dim cSaldo As Currency
    'cSaldo = F_CalculaSaldo("01/01/1900", Left(cmbxConta.Text, 3))
    
End Sub

Private Sub cmndbtnBuscar_Click_1()
    Dim strSQL As String
    If txtbxDAtaFIn.Visible = True Then
        If F_DataMaior(txtbxDataIni.text, txtbxDAtaFIn.text) Then
            MsgBox "A data final deve ser maior que a inicial!", vbExclamation
            txtbxDAtaFIn.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "select tf.*,tb.nome_con from " & TAB_FINANCEIRO & " as tf, tab_contas as tb where "
    strSQL = strSQL & "tf.cod_con=tb.cod_con  and "
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " historico_fin like '" & txtbxExpr.text & "%'"
    
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " valor_fin = " & F_GravaDec(txtbxExpr.text)
    End If
    If cmbxConta.ListIndex > -1 Then
        strSQL = strSQL & " and tb.cod_con='" & Left(cmbxConta.text, 3) & "'"
    End If
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and data_fin >= " & F_GravaData(txtbxDataIni.text)
    End If
    If txtbxDAtaFIn.text <> "" Then
        strSQL = strSQL & " and data_fin <= " & F_GravaData(txtbxDAtaFIn.text)
    End If
    If cmbxTIpo.ListIndex <> 0 Then
        strSQL = strSQL & " and tipo_fin='" & Left(cmbxTIpo.text, 1) & "' "
    End If
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " order by data_fin,seq_fin"
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & " order by seq_fin"
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " order by historico_fin, seq_fin"
    ElseIf cmbxBusca.ListIndex = 3 Then
        strSQL = strSQL & " order by valor_Fin,seq_Fin"
    End If
    If optnAsc.Value Then
        strSQL = strSQL & " asc "
    Else
        strSQL = strSQL & " desc "
    End If
    
    FechaObjeto ado
    ado.Open strSQL, conn_client
    
    
    Set dtgrd.DataSource = ado
    
  
  '
    P_Resumo
End Sub
Sub P_Resumo()
Dim cCheDeb As Currency
    Dim cCheCred As Currency
    Dim cDinDeb As Currency
    Dim cDinCred As Currency
    
    With adc.Recordset
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            
            If !Esp_fin = "1" Or !Esp_fin = "DIN" Then
                If !tipo_fin = "C" Then
                    cDinCred = cDinCred + !valor_fin
                Else
                    cDinDeb = cDinDeb + !valor_fin
                End If
            Else
                If !tipo_fin = "C" Then
                    cCheCred = cCheCred + !valor_fin
                Else
                    cCheDeb = cCheDeb + !valor_fin
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    'lblCred_Che.caption = FormatVal(cCheCred)
    lblCred_din.caption = FormatVal(cDinCred)
    lblDeb_Din.caption = FormatVal(cDinDeb)
'    lblDeb_Che.caption = FormatVal(cCheDeb)
'    lblTotal_Che.caption = FormatVal(cCheCred - cCheDeb)
    lblTotal_Din.caption = FormatVal(cDinCred - cDinDeb)
    'lblSaldo.caption = lblTotal_Din.caption
End Sub

Private Sub cmndbtnBuscar_Click()
    buscar
End Sub
Sub buscar()
        Dim strSQL As String
    On Error GoTo trata_Erro
    Dim order As String
    If cmbxOrd.ListIndex = 0 Then
        order = "data_fin,seq_fin"
    ElseIf cmbxOrd.ListIndex = 3 Then
        order = "historico_fin"
    ElseIf cmbxOrd.ListIndex = 2 Then
        order = "historico_fin"
    
        
    End If
    
    If strVisrelfin_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    strSQL = financeiro.F_SelectFinanceiro(txtbxDataIni.text, txtbxDAtaFIn.text, txtConta.tag, _
        "", txtbxExpr.text, txtValor.text, Left(cmbxTIpo.text, 1), order)
    abreObjetoAdo ado, strSQL, dtgrd
    
    
     
        soma
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub soma()
    Dim cred As Currency
    Dim deb As Currency
    On Error GoTo trata_Erro
    dtgrd.Visible = False
    With ado
        .MoveFirst
        Do While Not .EOF
            If !tipo_fin = "C" Then
                cred = cred + !valor_fin
            Else
                deb = deb + !valor_fin
            End If
            .MoveNext
        Loop
        .MoveFirst
    End With
    dtgrd.Visible = True
    txtCredito.text = FormatVal(cred)
    txtDebito.text = FormatVal(deb)
    txtTotal.text = FormatVal(cred - deb)
    If CCur(txtTotal.text) > 0 Then
        txtTotal.BackColor = vbGreen
    Else
        txtTotal.BackColor = vbRed
    End If
    Exit Sub
    
trata_Erro:
dtgrd.Visible = True
txtTotal.text = "0"
txtDebito.text = "0"
txtCredito.text = "0"
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmLanFin
    
    On Error GoTo trata_Erro
    
    With ado
        fr.tag = !seq_fin
        fr.txtbxData_Cai.text = !data_fin
        fr.txtbxHist.text = !historico_fin
        fr.txtbxValor_Cai.text = FormatVal(!valor_fin)
        If !tipo_fin = "C" Then
            fr.optnCred.Value = True
        Else
            fr.optnDeb.Value = True
        End If
        fr.optnCred.Enabled = False
        fr.optnDeb.Enabled = False
        fr.txtCCo1.text = !nome_con
        fr.txtCCo1.tag = !cod_con
        fr.txtbxClass.text = !nome_pla
        fr.txtbxClass.tag = !cod_pla
        If TiraNull(!seq_tra) <> "" Then
            fr.optTransf = True
            
        End If
        If !vinculado_fin <> "" Then
            fr.txtbxValor_Cai.Enabled = False
        End If
        fr.optTransf.Enabled = False
        
        fr.Show 1
    End With
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If ado.RecordCount = 0 Then
        cmndbtnExcl.Enabled = False
        Exit Sub
    End If
    If MsgBox("Confirma exclusão do Lançamento? ", vbYesNo + vbCritical) = vbYes Then
        strSQL = "Delete from  " & TAB_FINANCEIRO & "  where seq_fin ='" & ado!seq_fin & "'"
        conn.Execute strSQL
        
    End If
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnImpr_Click()
    'imprime frmVisualiza.Pic
    'frmVisualiza.Show 1
    'If frmVisualiza.tag <> "" Then
    '    imprime Printer
    '    Printer.EndDoc
    'End If
    'Unload frmVisualiza
On Error GoTo trata_Erro
    
    P_DtgrdToHtml "Movimento Financeiro - Conta " & cmbxConta.text, dtgrd, ado
    
    Exit Sub
trata_Erro:
End Sub
Sub imprime(obj As Object)
    P_CabRelTimbr "Movimento de Contas Financeiras", obj
    P_ImprX obj, "Período de " & txtbxDataIni.text, 10
    P_ImprX obj, "até " & txtbxDAtaFIn.text, 40, True
    P_ImprX obj, "Conta:" & cmbxConta.text, 100, False
    P_line1 obj
    With adc.Recordset
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        P_ImprX obj, "...", 2
        P_ImprX obj, "data", 15
        P_ImprX obj, "Valor", 35
        P_ImprX obj, "Tipo", 50
        P_ImprX obj, "Histórico", 60
        P_ImprX obj, "Conta", 140, False
        
        
        
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                If TypeName(obj) = "Printer" Then
                    P_RodapeVis obj
                    
                Else
                    P_AddPicVis obj
                End If
            End If
            P_ImprX obj, !seq_fin, 2
            P_ImprX obj, !data_fin, 15
            P_ImprX obj, FormatVal(!valor_fin), 35
            P_ImprX obj, !tipo_fin, 50
            P_ImprX obj, Left(!historico_fin, 35), 60
            P_ImprX obj, !nome_con, 140, False
            
            .MoveNext
        Loop
        
        P_line1 obj
        P_ImprX obj, "Crédito...:" & lblCred_din.caption, 2
        P_ImprX obj, "Débito...:" & lblDeb_Din.caption, 40
        P_ImprX obj, "Total...:" & lblTotal_Din.caption, 80, False
    End With
    
    
End Sub
Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanFin
    fr.Show 0, mdiFrm
 End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnTrasf_Click()

End Sub

Private Sub dbgrd_DblClick()
    If cmndbtnEdit.Enabled Then cmndbtnEdit_Click
End Sub

Private Sub dbgrd_GotFocus()
    
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtServ_Reposition()
    On Error Resume Next
    If adc.Recordset.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
        Exit Sub
    Else
        If adc.Recordset!vinculado_fin = "N" Then
            cmndbtnEdit.Enabled = True
            cmndbtnExcl.Enabled = True
        Else
            cmndbtnEdit.Enabled = False
            cmndbtnExcl.Enabled = False
        End If
    End If
    
    
End Sub

Private Sub Form_Activate()
    bActive = True
End Sub

Private Sub Form_Load()
    bActive = False
    resizeForm Me
'    cmbxBusca.ListIndex = 0
    'cmbxOrd.ListIndex = 0
    'cmbxTIpo.ListIndex = 0
    txtbxDAtaFIn.text = Date
    txtbxDataIni.text = Date
    
    
    
    Set ado = New ADODB.Recordset
    
    
    
End Sub

Private Sub Form_Resize()
    On Error GoTo trata_Erro
    PosicionaGrade dtgrd, Me, pcBotoes
    
    posicionaADireita cmndbtnSair, pcBotoes
    Exit Sub
trata_Erro:
    
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
    If F_DataMaior(txtbxDataIni.text, txtbxDAtaFIn.text) Then
        MsgBox "A data final deve ser maior que a inicial!", vbExclamation
        txtbxDAtaFIn.SetFocus
        Exit Sub
    End If
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    'If cmbxBusca.ListIndex = 1 Then
    '    txtbxExpr.Text = Format(F_SoNumero(txtbxExpr.Text), "000000#")
    'ElseIf cmbxBusca.ListIndex = 2 Then
    '    P_ValorLostFocus txtbxExpr
    'Else
        P_AjustaTexto txtbxExpr
    'E'nd If

End Sub

Private Sub txtConta_Change()
    txtConta.tag = ""
End Sub

Private Sub txtConta_GotFocus()
    P_SelecTexto txtConta
End Sub

Private Sub txtConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaConta
    End If
End Sub
