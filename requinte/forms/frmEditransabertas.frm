VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiTransAbertas 
   Caption         =   "Edição de Transferências"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   18525
   Icon            =   "frmEditransabertas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   18525
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   18495
      TabIndex        =   20
      Top             =   7710
      Width           =   18525
      Begin VB.TextBox txtUsu 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkRetif 
         Caption         =   "Retificadora"
         Height          =   195
         Left            =   1680
         TabIndex        =   23
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   285
         Left            =   6360
         Picture         =   "frmEditransabertas.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   4680
         TabIndex        =   26
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   405
      End
      Begin VB.Label lblVendas 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   720
         TabIndex        =   21
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   18495
      TabIndex        =   19
      Top             =   0
      Width           =   18525
      Begin VB.TextBox txtbxSeqAce 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnEdita 
         Caption         =   "Imprimir Transferência"
         Height          =   735
         Left            =   11640
         Picture         =   "frmEditransabertas.frx":0D8C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir Acerto"
         Height          =   735
         Left            =   15360
         Picture         =   "frmEditransabertas.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   16800
         Picture         =   "frmEditransabertas.frx":14A0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir Lista"
         Height          =   735
         Left            =   13920
         Picture         =   "frmEditransabertas.frx":1595
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   735
         Left            =   10200
         Picture         =   "frmEditransabertas.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmEditransabertas.frx":1D2C
         Left            =   7200
         List            =   "frmEditransabertas.frx":1D3C
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEditransabertas.frx":1D6E
         Left            =   8640
         List            =   "frmEditransabertas.frx":1D78
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   285
         Left            =   6840
         Picture         =   "frmEditransabertas.frx":1D96
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   3480
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   7200
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   8640
         TabIndex        =   11
         Top             =   120
         Width           =   510
      End
   End
   Begin MSDataGridLib.DataGrid dtg 
      Bindings        =   "frmEditransabertas.frx":2120
      Height          =   6255
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11033
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
      Caption         =   "Transferências Encontradas"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "data_con"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "hora"
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
      BeginProperty Column02 
         DataField       =   "cod_rep"
         Caption         =   "Cód."
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
         DataField       =   "nome_rep"
         Caption         =   "Representante"
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
         DataField       =   "guerra_fun"
         Caption         =   "Usuario"
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
         DataField       =   "retif_con"
         Caption         =   "Retif?"
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
         DataField       =   "valor_tra"
         Caption         =   "Valor"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4680
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   840
      Top             =   5400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
End
Attribute VB_Name = "frmEdiTransAbertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sTab As String
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmbxOrder_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmdUsu_Click()
    Dim f As funcionario.funcionario
    
    f = funcionario.F_SelectFuncionario(txtUsu.Text)
    If f.cod_fun <> "" Then
        txtUsu.Text = f.nome_fun
        txtUsu.tag = f.cod_fun
    End If
    
    
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim sCampo As String
    Dim sCod As String
    
    On Error GoTo trata_Erro
    If cmbxStatus.ListIndex = 1 Then
        sTab = "tab_consignabertas"
        
    Else
        sTab = "tab_consignacoes"
        
        If strCorrigeEstCons_Usu <> "S" Then
        
            MsgBox "Você não tem permissão para esta consulta!", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    
    If txtbxdataFin.Text = "" And txtbxDataIni.Text = "" Then
        MsgBox "Informe as datas!", vbInformation
        Exit Sub
    End If
    strSQL = "select distinct ta.seq_con,tr.nome_rep, ta.cod_rep,left(ta.data_con,10) as data_con,"
    strSQL = strSQL & " mid(data_con,12,5) as hora, sum(qtd_con*unit_con) as valor_tra, ta.retif_con,tu.guerra_fun,tr.nome_rep "
    strSQL = strSQL & " from " & sTab & " as ta, "
    strSQL = strSQL & "tab_funcionarios as tu, tab_representantes as tr "
    
    strSQL = strSQL & " where tu.cod_fun=ta.cod_usu "
    strSQL = strSQL & " and tr.cod_rep=ta.cod_rep "
    
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_con<=" & F_GravaDataTempo(txtbxdataFin.Text & " 23:59")
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_con>=" & F_GravaDataTempo(txtbxDataIni.Text & " 00:00")
    End If
    If txtbxRepresentante.tag <> "" Then
        strSQL = strSQL & " and ta.cod_rep='" & txtbxRepresentante.tag & "' "
    End If
    'If cmbxStatus.Text <> "" Then
    '    strSQL = strSQL & " and status_CON='" & Left(cmbxStatus.Text, 1) & "' "
    'End If
    If txtUsu.tag <> "" Then
        strSQL = strSQL & " and cod_usu='" & txtUsu.tag & "' "
    End If
    If txtbxSeqAce.Text <> "" Then
        strSQL = strSQL & " and ta.seq_con='" & txtbxSeqAce.Text & "' "
    End If
    If chkRetif.Value = 1 Then
        strSQL = strSQL & " and retif_con='S' "
    End If
    strSQL = strSQL & " group by ta.seq_con "
    If cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by nome_rep,data_con "
    Else
        strSQL = strSQL & " order by data_con desc, ta.seq_con,nome_rep "
    End If
    
    
    abreObjetoAdo ado, strSQL, dtg
    
    soma
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub soma()
    Dim cTotal As Currency
    dtg.Visible = False
    With ado
        
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
           ' cTotal = cTotal + !valor_tra
            
            .MoveNext
        Loop
        lblVendas.caption = FormatVal(cTotal)
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    dtg.Visible = True
End Sub

Private Sub cmndbtnBuscaRep_Click()
    buscaRep
End Sub

Private Sub cmndbtnEdita_Click()
  On Error GoTo trata_Erro
  
    If sTab = "tab_consignacoes" Then
        imprimeConsign ado!seq_con
    Else
        frmTransfConsign.tag = ado!cod_Rep
        frmTransfConsign.Show 0
        
        
    End If
  
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnExcl_Click()
    On Error GoTo trata_Erro
    If MsgBox("Confirma exclusão do Acerto?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    If strExcldoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta exclusão!", vbExclamation
        Exit Sub
    End If
    If Not F_senha(True) Then
        Exit Sub
    End If
    
    
    If Acerto.F_DeletaAcerto(dt.Recordset!seq_ace) Then
        cmndbtnBuscar_Click
    End If
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DtgrdToHtml "Relatório de Acertos de " & txtbxDataIni.Text & " até " & _
        txtbxdataFin.Text, dtg, ado, txtbxRepresentante.Text, "3"
    
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVisualizaAcerto_Click()
        On Error GoTo trata_Erro
    P_VisualizaAcerto dt.Recordset!seq_ace
    Exit Sub
trata_Erro:

End Sub


Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
     Set ado = New ADODB.Recordset
    
    txtbxdataFin.Text = Date
    txtbxDataIni.Text = Date
    resizeForm Me
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    PosicionaGrade dtg, Me, pcBot
    posicionaADireita cmndbtnSair, pcTop
    
    
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

Private Sub txtbxRepresentante_GotFocus()
    P_SelecTexto txtbxRepresentante
End Sub

Private Sub txtbxRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaRep
    End If
End Sub

Sub buscaRep()
    Dim r As representante.representante
    r = F_BuscaRep(txtbxRepresentante.Text)
    txtbxRepresentante.Text = r.nome_rep
    txtbxRepresentante.tag = r.cod_Rep
    busca
End Sub

Private Sub txtbxSeqAce_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxSeqAce_LostFocus()
    txtbxSeqAce.Text = Format(txtbxSeqAce.Text, "000000#")
End Sub

Private Sub txtUsu_Change()
    txtUsu.tag = ""
End Sub

Private Sub txtUsu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdUsu_Click
    End If
End Sub
