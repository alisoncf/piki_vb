VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiDevConsign 
   Caption         =   "Edição de Devoluções"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16305
   Icon            =   "frmEdiDevConsign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   16305
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   16275
      TabIndex        =   19
      Top             =   7710
      Width           =   16305
      Begin VB.TextBox txtUsu 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkRetif 
         Caption         =   "Retificadora"
         Height          =   195
         Left            =   1680
         TabIndex        =   22
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   285
         Left            =   6360
         Picture         =   "frmEdiDevConsign.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   4680
         TabIndex        =   25
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   16275
      TabIndex        =   18
      Top             =   0
      Width           =   16305
      Begin VB.CommandButton cmdPrintDev 
         Caption         =   "Im&primir Dev."
         Height          =   615
         Left            =   12840
         Picture         =   "frmEdiDevConsign.frx":0D8C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtbxSeqAce 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnEdita 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   14160
         Picture         =   "frmEdiDevConsign.frx":1199
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   15480
         Picture         =   "frmEdiDevConsign.frx":1523
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir Grade"
         Height          =   615
         Left            =   11520
         Picture         =   "frmEdiDevConsign.frx":1618
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   10200
         Picture         =   "frmEdiDevConsign.frx":1A25
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1215
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
         ItemData        =   "frmEdiDevConsign.frx":1DAF
         Left            =   7200
         List            =   "frmEdiDevConsign.frx":1DBF
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEdiDevConsign.frx":1DF1
         Left            =   8640
         List            =   "frmEdiDevConsign.frx":1DFB
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   285
         Left            =   6840
         Picture         =   "frmEdiDevConsign.frx":1E19
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
      Bindings        =   "frmEdiDevConsign.frx":21A3
      Height          =   6255
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   14175
      _ExtentX        =   25003
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
         DataField       =   "seq_dev"
         Caption         =   "seq"
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
         DataField       =   "data_dev"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "hora_dev"
         Caption         =   "Hora"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   4
         EndProperty
      EndProperty
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "retif_dev"
         Caption         =   "Retificadora?"
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
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4680
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
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
Attribute VB_Name = "frmEdiDevConsign"
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

Private Sub cmdPrintDev_Click()
    imprimeDevConsign ado!seq_dev
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
    
    On Error GoTo trata_erro
    If cmbxStatus.ListIndex = 1 Then
        sTab = "tab_devconsignabertas"
        
    Else
        sTab = "tab_devconsign"
    End If
    
    
    
    
    
    If txtbxDataFin.Text = "" And txtbxDataIni.Text = "" Then
        MsgBox "Informe as datas!", vbInformation
        Exit Sub
    End If
    strSQL = "select distinct ta.seq_dev,tr.nome_rep, ta.cod_rep,left(ta.data_dev,10) as data_dev,right(data_dev,8) as hora_dev,ta.*,tu.guerra_fun,tr.nome_rep "
    strSQL = strSQL & " from " & sTab & " as ta, "
    strSQL = strSQL & "tab_funcionarios as tu, tab_representantes as tr "
    
    strSQL = strSQL & " where tu.cod_fun=ta.cod_usu "
    strSQL = strSQL & " and tr.cod_rep=ta.cod_rep "
    
    If txtbxDataFin.Text <> "" Then
        strSQL = strSQL & " and data_dev<=" & F_GravaDataTempo(txtbxDataFin.Text & " 23:59")
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_dev>=" & F_GravaDataTempo(txtbxDataIni.Text & " 00:00")
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
        strSQL = strSQL & " and ta.seq_dev='" & txtbxSeqAce.Text & "' "
    End If
    If chkRetif.Value = 1 And sTab = "tab_devconsignabertas" Then
        strSQL = strSQL & " and retif_dev='S' "
    End If
    strSQL = strSQL & " group by ta.seq_dev "
    If cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by nome_rep,data_con "
    Else
        strSQL = strSQL & " order by data_dev desc, ta.seq_dev,nome_rep "
    End If
    
    
    FechaObjeto ado
    ado.Open strSQL, Conn_Client
    Set dtg.DataSource = ado
    
    soma
    Exit Sub
trata_erro:
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
  On Error GoTo trata_erro
  
    If sTab = "tab_devconsign" Then
        imprimeDevConsign ado!seq_con
    Else
        frmDevolucaoConsign.tag = ado!cod_Rep
        frmDevolucaoConsign.Show 0
        
        
    End If
  
    
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub cmndbtnExcl_Click()
    On Error GoTo trata_erro
    
    Exit Sub
trata_erro:
    
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DtgrdToHtml "Relatório de Devoluções de " & txtbxDataIni.Text & " até " & _
        txtbxDataFin.Text, dtg, ado, txtbxRepresentante.Text, "1"
    
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVisualizaAcerto_Click()
        On Error GoTo trata_erro
    P_VisualizaAcerto dt.Recordset!seq_ace
    Exit Sub
trata_erro:

End Sub


Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
     Set ado = New ADODB.Recordset
    
    txtbxDataFin.Text = Date
    txtbxDataIni.Text = Date
    resizeForm Me
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    PosicionaGrade dtg, Me, pcBot
    posicionaADireita cmndbtnSair, pcTop
    
    
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

Private Sub txtbxRepresentante_Change()
    txtbxRepresentante.tag = ""
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
