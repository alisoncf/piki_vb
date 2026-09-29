VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiAcertos 
   Caption         =   "Edição de Acertos"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15195
   Icon            =   "frmEdiAcertos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   15195
   Begin MSDataGridLib.DataGrid gridAcertos 
      Height          =   5655
      Left            =   0
      TabIndex        =   30
      Top             =   1080
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   9975
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
      Caption         =   "Acertos no Período"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "seq_ace"
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
         DataField       =   "cod_rep"
         Caption         =   "Cód.Rep"
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
      BeginProperty Column03 
         DataField       =   "nome_for"
         Caption         =   "Forma Pgto"
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
         DataField       =   "valor_ace"
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
         DataField       =   "diferenca"
         Caption         =   "Diferença"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "comissao_ace"
         Caption         =   "comiss"
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
         DataField       =   "data_ace"
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
      BeginProperty Column08 
         DataField       =   "hora_ace"
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
      BeginProperty Column09 
         DataField       =   "guerra_fun"
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
      BeginProperty Column10 
         DataField       =   "status_ace"
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
      BeginProperty Column11 
         DataField       =   "queima"
         Caption         =   "Queima"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   15165
      TabIndex        =   21
      Top             =   0
      Width           =   15195
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEdiAcertos.frx":0A22
         Left            =   8040
         List            =   "frmEdiAcertos.frx":0A2C
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmEdiAcertos.frx":0A4A
         Left            =   6600
         List            =   "frmEdiAcertos.frx":0A5A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   9600
         Picture         =   "frmEdiAcertos.frx":0A8C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   10920
         Picture         =   "frmEdiAcertos.frx":0E16
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   14880
         Picture         =   "frmEdiAcertos.frx":1223
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir Acerto"
         Height          =   615
         Left            =   13560
         Picture         =   "frmEdiAcertos.frx":1318
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmndbtnEdita 
         Caption         =   "Editar"
         Height          =   615
         Left            =   12240
         Picture         =   "frmEdiAcertos.frx":16A2
         Style           =   1  'Graphical
         TabIndex        =   15
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
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   285
         Left            =   6240
         Picture         =   "frmEdiAcertos.frx":1A2C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   8040
         TabIndex        =   11
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   6600
         TabIndex        =   9
         Top             =   120
         Width           =   495
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
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   15165
      TabIndex        =   18
      Top             =   7710
      Width           =   15195
      Begin VB.TextBox txtUsu 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   285
         Left            =   2400
         Picture         =   "frmEdiAcertos.frx":1DB6
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comissão:"
         Height          =   195
         Left            =   7320
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblComiss 
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
         Left            =   8160
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença:"
         Height          =   195
         Left            =   4800
         TabIndex        =   24
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblDif 
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
         Left            =   5640
         TabIndex        =   23
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblDataCancelamento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
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
         Left            =   10560
         TabIndex        =   22
         Top             =   120
         Width           =   1230
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
         Left            =   3360
         TabIndex        =   20
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   2760
         TabIndex        =   19
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   4455
   End
End
Attribute VB_Name = "frmEdiAcertos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error GoTo trata_Erro
    Dim rPesq As New ADODB.Recordset
    lblDataCancelamento.caption = ""
    Dim strSQL As String
    
    If bActive = False Then Exit Sub
    
    With ado
        strSQL = "select * from tab_acertoscanc where seq_ace='" & !seq_ace & "' "
        abreObjetoAdo rPesq, strSQL
        
        If rPesq.RecordCount > 0 Then
            lblDataCancelamento.caption = "Cancelado em " & rPesq!data_can & " Usuário: " & rPesq!cod_usu
        End If
        
    End With
    
    Exit Sub
trata_Erro:
End Sub

Private Sub cmbxOrder_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_DblClick()
    cmndbtnEdita_Click
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
    Dim sTab As String
    gridAcertos.Visible = False
    On Error GoTo trata_Erro:
    If cmbxStatus.ListIndex > 0 Then
        sTab = "tab_acertos"
    Else
        sTab = "tab_acertosabertos"
    End If
    
    strSQL = "select tf.nome_for, ta.*, if(ta.diferenca_ace>0,ta.diferenca_ace*.85,ta.diferenca_ace) as diferenca,tu.guerra_fun,tr.nome_rep "
    strSQL = strSQL & " from " & sTab & " as ta, tab_formaspag as tf,"
    strSQL = strSQL & "tab_funcionarios as tu,tab_representantes as tr "
    strSQL = strSQL & " where tf.cod_for=ta.cod_for "
    strSQL = strSQL & " and tu.cod_fun=ta.cod_usu "
    strSQL = strSQL & " and tr.cod_rep=ta.cod_rep "
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_ace<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_ace>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxRepresentante.tag <> "" Then
        strSQL = strSQL & " and ta.cod_rep='" & txtbxRepresentante.tag & "' "
    End If
    If cmbxStatus.Text <> "" Then
        strSQL = strSQL & " and status_ace='" & Left(cmbxStatus.Text, 1) & "' "
    End If
    If txtbxSeqAce.Text <> "" Then
        strSQL = strSQL & " and ta.seq_ace='" & txtbxSeqAce.Text & "' "
    End If
    
    If txtUsu.tag <> "" Then
        strSQL = strSQL & " and ta.cod_usu='" & txtUsu.tag & "' "
    End If
    If txtbxSeqAce.Text <> "" Then
        strSQL = strSQL & " and ta.seq_ace='" & txtbxSeqAce & "' "
    End If
    If cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by nome_rep,data_ace,hora_ace"
    Else
        strSQL = strSQL & " order by data_ace desc,hora_ace desc, nome_rep "
    End If
    
    
    
    
    abreObjetoAdo ado, strSQL, gridAcertos
    soma
    gridAcertos.Visible = True
Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub soma()
    Dim cTotal As Currency
    Dim cDif As Currency
    Dim cComiss As Currency
    
    
    
    bActive = False
    With ado
    
        Do While Not .EOF
            cTotal = cTotal + !valor_ace
            cDif = cDif + Format(!diferenca, "###,##0")
            cComiss = cComiss + !comissao_ace
            
            .MoveNext
        Loop
        lblVendas.caption = FormatVal(cTotal)
        lblComiss.caption = FormatVal(cComiss)
        lblDif.caption = FormatVal(cDif)
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    bActive = True
End Sub

Private Sub cmndbtnBuscaRep_Click()
    buscaRep
End Sub

Private Sub cmndbtnEdita_Click()
  On Error GoTo trata_Erro
  
    If Left(cmbxStatus.Text, 1) = "2" Or Left(cmbxStatus.Text, 1) = "3" Then
        imprimeacerto ado!seq_ace
    Else
        frmAcerto.tag = ado!seq_ace
        frmAcerto.Show 0
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
    If ado!status_ace <> "2" Then
        MsgBox "Só possível cancelar um acerto concluído!", vbExclamation
        Exit Sub
    End If
    
    If Acerto.F_DeletaAcerto(ado!seq_ace) Then
        cmndbtnBuscar_Click
    End If
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DtgrdToHtml "Relatório de Acertos de " & txtbxDataIni.Text & " até " & _
        txtbxdataFin.Text, gridAcertos, ado, txtbxRepresentante.Text, "4,5,6"
    
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVisualizaAcerto_Click()
        On Error GoTo trata_Erro
    P_VisualizaAcerto ado!seq_ace
    Exit Sub
trata_Erro:

End Sub


Private Sub Form_Activate()
    busca
    
    If strCorrigeEstCons_Usu <> "S" Then
        lblComiss.Visible = False
        lblDif.Visible = False
        
        cmndbtnImprimir.Enabled = False
        gridAcertos.Columns(4).Visible = False
        gridAcertos.Columns(5).Visible = False
        gridAcertos.Columns(6).Visible = False
        
    
    End If
    
End Sub

Private Sub Form_Load()
    
    txtbxdataFin.Text = Date
    txtbxDataIni.Text = Date
    
    Set ado = New ADODB.Recordset

    resizeForm Me
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbgrd.width = Me.width - dbgrd.Left - 100
    dbgrd.height = pcBot.top - dbgrd.top - 100
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

Private Sub txtUsu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdUsu_Click
    End If
    
End Sub
