VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiRep 
   Caption         =   "Edição de Representantes"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17775
   Icon            =   "frmedirep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10275
   ScaleWidth      =   17775
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   17745
      TabIndex        =   5
      Top             =   0
      Width           =   17775
      Begin VB.CommandButton cmdLembreteAcerto 
         Caption         =   "Lembrete Acerto"
         Height          =   615
         Left            =   12000
         Picture         =   "frmedirep.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdLembrete 
         Caption         =   "Lembrete"
         Height          =   615
         Left            =   10440
         Picture         =   "frmedirep.frx":0D44
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   1455
      End
      Begin VB.CheckBox chkBloq 
         Caption         =   "Excluidos"
         Height          =   195
         Left            =   8280
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   1
         Top             =   240
         Width           =   6495
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmedirep.frx":11BE
         Left            =   120
         List            =   "frmedirep.frx":11D1
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar:"
         Height          =   195
         Left            =   1680
         TabIndex        =   0
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox pcBotoes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   870
      ScaleWidth      =   17745
      TabIndex        =   4
      Top             =   9375
      Width           =   17775
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Fe&char"
         Height          =   695
         Left            =   15720
         Picture         =   "frmedirep.frx":11F8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnExcluir 
         Caption         =   "Excluir"
         Height          =   695
         Left            =   12600
         Picture         =   "frmedirep.frx":1363
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   695
         Left            =   4800
         Picture         =   "frmedirep.frx":202D
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnResumoDeb 
         Caption         =   "Resumo Débitos"
         Height          =   695
         Left            =   6360
         Picture         =   "frmedirep.frx":243A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "&Fechar"
         Height          =   695
         Left            =   120
         Picture         =   "frmedirep.frx":2584
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton cmdManifesto 
         Caption         =   "&Gerar Manifesto"
         Height          =   695
         Left            =   7920
         Picture         =   "frmedirep.frx":26CE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnAcerto 
         Caption         =   "&Acerto"
         Height          =   695
         Left            =   9480
         Picture         =   "frmedirep.frx":2C58
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnDeb 
         Caption         =   "&Débitos"
         Height          =   695
         Left            =   11040
         Picture         =   "frmedirep.frx":3522
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnAcertos 
         Caption         =   "Acertos"
         Height          =   695
         Left            =   14160
         Picture         =   "frmedirep.frx":366C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "&Histórico"
         Height          =   695
         Left            =   3240
         Picture         =   "frmedirep.frx":3721
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   695
         Left            =   1680
         Picture         =   "frmedirep.frx":3FEB
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   695
         Left            =   120
         Picture         =   "frmedirep.frx":4135
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmedirep.frx":44BF
      Height          =   3495
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Registros Encontrados"
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "cod_rep"
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
      BeginProperty Column01 
         DataField       =   "cpf_rep"
         Caption         =   "CPF:"
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
         Caption         =   "Nome"
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
         DataField       =   "endereco_rep"
         Caption         =   "Endereço"
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
         DataField       =   "bairro_rep"
         Caption         =   "Bairro"
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
         DataField       =   "Cidade_rep"
         Caption         =   "Cidade"
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
         DataField       =   "excluido_rep"
         Caption         =   "Excluído"
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
         DataField       =   "datacad_rep"
         Caption         =   "Atualização"
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
         DataField       =   "Status_Rep"
         Caption         =   "status"
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
         DataField       =   "lembrete"
         Caption         =   "Lembrete"
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
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2759,811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4140,284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2459,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2940,095
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column09 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   1560
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   1
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
Attribute VB_Name = "frmEdiRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Public WithEvents AdoRep As ADODB.Recordset
Attribute AdoRep.VB_VarHelpID = -1
'Public WithEvents AdoTel As ADODB.Recordset
Private Sub DBGrid1_Click()

End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub AdoRep_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    buscaTel
On Error Resume Next
    With AdoRep
        
        
        If !excluido_rep = "S" Then
            cmndbtnExcluir.caption = "Reativar"
        Else
            cmndbtnExcluir.caption = "Excluir"
        End If
    End With
End Sub

Private Sub chkBloq_Click()
    cmndbtnLocaliza_Click
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnCard_Click()
    'P_Cartao adorep!cod_cli
End Sub

Private Sub cmdLembrete_Click()
    buscaLembrete AdoRep!cod_Rep, True
End Sub

Private Sub cmdLembreteAcerto_Click()
    buscaLembreteAcerto AdoRep!cod_Rep, True
End Sub

Private Sub cmdManifesto_Click()
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para essa operação!", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub cmdSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub cmndbtnAcerto_Click()
    On Error Resume Next
    frmAcerto.txtbxRepresentante.tag = AdoRep!cod_Rep
    frmAcerto.caption = "Acerto de: " & AdoRep!nome_rep
    frmAcerto.txtbxRepresentante.text = AdoRep!nome_rep
    frmAcerto.Show 0
End Sub

Private Sub cmndbtnAcertos_Click()
    On Error Resume Next
    P_Acertos AdoRep!cod_Rep
End Sub

Private Sub cmndbtnDeb_Click()
    On Error GoTo trata_Erro
    frmParcelasTer.txtbxNomeCli.text = AdoRep!nome_rep
    frmParcelasTer.txtbxNomeCli.tag = AdoRep!cod_Rep
    frmParcelasTer.tag = AdoRep!cod_Rep
    frmParcelasTer.Show 0
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadRep
    If AdoRep.RecordCount <= 0 Then
        MsgBox "Não há o que ser editado!", vbCritical
        Exit Sub
    End If
    strHora = Time
    fr.tag = AdoRep!cod_Rep
    
    fr.Show 0
End Sub

Private Sub cmndbtnExcluir_Click()
    On Error GoTo trata_Erro:
    Dim strSQL As String
    If strAlterarComRepr_Usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    If AdoRep!excluido_rep = "N" Then
        If representante.ExisteRegistroAberto(AdoRep!cod_Rep, True) Then
                MsgBox "Não é possível excluir representante com movimentos!", vbExclamation
                
                Exit Sub
            End If
        If MsgBox("Confirma realmente a exclusão do Representante " & AdoRep!nome_rep & "?", vbYesNo + vbQuestion) Then
            
            
            strSQL = "update tab_representantes set excluido_rep='S' where cod_rep='" & AdoRep!cod_Rep & "'"
            
            
            
            
            Conn.Execute strSQL
        'grava_audit "tab_repr
            dtgrd.Refresh
        
        End If
    Else
        If MsgBox("Confirma realmente a reativação do Representante " & AdoRep!nome_rep & "?", vbYesNo + vbQuestion) Then
            strSQL = "update tab_representantes set excluido_rep='N' where cod_rep='" & AdoRep!cod_Rep & "'"
            Conn.Execute strSQL
        'grava_audit "tab_repr
            dtgrd.Refresh
        
        End If
    End If
Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnHist_Click()
'    On Error GoTo trata_erro
    With AdoRep
        If .RecordCount = 0 Then Exit Sub
        P_histCli !cod_Rep
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnLocaliza_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    On Error GoTo trata_Erro
    strSQL = "select tc.cod_rep,status_rep,nome_rep,cidade_rep,bairro_rep,excluido_rep,endereco_rep,cpf_rep,datacad_rep from tab_representantes as tc  "
    
    
    
    strSQL = strSQL & " where "
    
     
          strSQL = strSQL & "(Nome_rep like '" & txtbxExpr.text & "%' "
    
        strSQL = strSQL & " or cod_rep='" & txtbxExpr.text & "' "
        
        strSQL = strSQL & "or cpf_rep like '" & F_FormataCPF(txtbxExpr.text) & "' "
    
        strSQL = strSQL & "or cpf_rep like '" & F_FormataCGC(txtbxExpr.text) & "' "
    
        strSQL = strSQL & "or rg_rep like '" & txtbxExpr.text & "') "
'        strSQL = strSQL & " and tipo_ter='R' "
    If chkBloq.Value = 0 Then
    
        strSQL = strSQL & " and excluido_rep='N'"
    End If
    
    strSQL = strSQL & " group by tc.cod_rep "
    If cmbxOrd.ListIndex = 0 Then
        sOrderBy = "tc.cod_rep"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrderBy = "Nome_rep,cpf_rep"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrderBy = "cpf_rep"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrderBy = "Cidade_rep,uf_rep,NOme_rep"
    ElseIf cmbxOrd.ListIndex = 4 Then
        sOrderBy = "UF_rep,Cidade_rep"
    End If
    strSQL = strSQL & " order by " & sOrderBy
    FechaObjeto AdoRep
    AdoRep.Open strSQL, Conn_Client
    Set dtgrd.DataSource = AdoRep
    
    
    dtgrd.caption = "Cadastrados...: " & F_BuscaUmCampo("count(cod_rep)", "tab_representantes", "")
    dtgrd.caption = dtgrd.caption & " - Listados...: " & AdoRep.RecordCount
Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DtgrdToHtml "Lista de Representantes", dtgrd, AdoRep
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadRep
    fr.Show 0
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmndbtnPremiacao_Click()
    On Error GoTo trata_Erro
    'P_PremiacoesCli adorep!cod_cli
    Exit Sub
trata_Erro:
End Sub

Private Sub cmndbtnPromocoes_Click()
    On Error GoTo trata_Erro
    'P_PontosCli adorep!cod_cli
    Exit Sub
trata_Erro:
End Sub

Private Sub cmndbtnRes_Click()
    frmReserva.Show 0
End Sub

Private Sub cmndbtnResumoDeb_Click()
    frmRelResumoRep.Show
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim c As Clientes.cliente
    c = Clientes.F_BuscaCliente("")
    
    MsgBox c.t(0).num_tel & c.t(1).num_tel
    
    
    
End Sub

Private Sub Command2_Click()

End Sub

Private Sub dtgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub buscaTel()
    Dim strSQL As String
    On Error Resume Next
    With AdoRep
        
        
        strSQL = "select * from tab_telefones where cod_cli='" & !cod_Rep & "' and tipo_ter='R'"
        adctel.RecordSource = strSQL
        adctel.Refresh
        
    
    End With
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    cmndbtnLocaliza_Click
    
    txtbxExpr.SetFocus
    bActive = True
End Sub

Private Sub Form_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    Set AdoRep = New ADODB.Recordset
    
    
    resizeForm Me
    
   
    
    cmbxOrd.ListIndex = 0
    bActive = False
    
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade dtgrd, Me, pcBotoes
    posicionaADireita cmdSair, pcBotoes
    
End Sub

Private Sub pcBotoes_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub txtbxExpr_Change()
    cmndbtnLocaliza_Click
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
    If KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
    
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
    
End Sub
