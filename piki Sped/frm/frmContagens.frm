VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmContagens 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contagens"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9330
   Icon            =   "frmContagens.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   9255
      TabIndex        =   11
      Top             =   6120
      Width           =   9255
   End
   Begin VB.Frame Frame3 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmndbtnEfetiva 
         Caption         =   "&Efetivar Contagem "
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "cria um novo inventário."
         Top             =   2235
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnLançar 
         Caption         =   "&Lançar Produto"
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "cria um novo inventário."
         Top             =   3000
         Width           =   2055
      End
      Begin VB.ComboBox cmbxCont 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cmbxOrderBy 
         BackColor       =   &H80000014&
         Height          =   315
         ItemData        =   "frmContagens.frx":03DE
         Left            =   120
         List            =   "frmContagens.frx":03EB
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Abrir Nova Contagem"
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "cria um novo inventário."
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":0556
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fecha formulário."
         Top             =   4560
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":08E0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Emite Inventário"
         Top             =   3780
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Contagem:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Organizar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1005
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   3600
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
   Begin MSDataGridLib.DataGrid dbgrd 
      Bindings        =   "frmContagens.frx":0A2A
      Height          =   5895
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   10398
      _Version        =   393216
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
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   4080
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   135
      Left            =   5640
      TabIndex        =   10
      Top             =   4920
      Width           =   15
   End
End
Attribute VB_Name = "frmContagens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxCont_Click()
    Dim strSQL As String
    strSQL = "select tc.cod_usu,tc.tipo_con,tc.data_cnt,tc.status_cnt,tp.nome_pro,ti.* from tab_produtos as tp,tab_contagens as tc, "
    strSQL = strSQL & "tab_itcont as ti where ti.cod_pro=tp.cod_pro and tc.seq_cnt=ti.seq_cnt "
    
    strSQL = strSQL & " and tc.seq_cnt='" & Left(cmbxCont.Text, 4) & "' "
    strSQL = strSQL & " order by nome_pro"
    adc.RecordSource = strSQL
    adc.Refresh
    With adc.Recordset
    If .RecordCount > 0 Then
        dbgrd.caption = "Contagem de " & !data_cnt & " |   Usuário: " & F_BuscaUmCampo("guerra_fun", "tab_funcionarios", "cod_fun='" & !cod_usu & "'")
        If !status_cnt = "N" Then
            dbgrd.caption = dbgrd.caption & "   | Status:" & " Em Aberto"
            cmndbtnEfetiva.Enabled = True
            cmndbtnLançar.Enabled = True
        Else
            dbgrd.caption = dbgrd.caption & "   | Status:" & " Efetivada"
            cmndbtnEfetiva.Enabled = False
            cmndbtnLançar.Enabled = False
        End If
    End If
    End With
End Sub

Private Sub cmbxCont_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDelete Then
        busca
    End If
End Sub

Private Sub cmndbtnEfetiva_Click()
    Dim strSQL As String
    strSQL = "Efetivar a Contagem implica em alterar o estoque de todos os Produtos "
    strSQL = strSQL & " nela contidos para suas quantidades contadas, respectivamente."
    strSQL = strSQL & Chr(13) & "Deseja Realmente Efetivar a Contagem: " & cmbxCont.Text
    strSQL = strSQL & " ?"
    If MsgBox(strSQL, vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
        efetivaCont
    End If
    P_AjustaEstoqueReservado
End Sub
Sub efetivaCont()
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim sCampo As String
    With adc.Recordset
        .MoveFirst
        If !tipo_con = "R" Then
            sCampo = "er_pro"
        Else
            sCampo = "ef_pro"
        End If
        Do While Not .EOF
            
            strSQL = "update tab_produtos set " & sCampo & "=" & F_GravaDec(!qtd_itc) & " "
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            conn.Execute strSQL
            .MoveNext
        Loop
        strSQL = "update tab_contagens set status_cnt='S' where seq_cnt='"
        strSQL = strSQL & Left(cmbxCont.Text, 4) & "'"
        conn.Execute strSQL
        MsgBox "Contagem Efetivada!", vbExclamation
        cmbxCont_Click
    End With
Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_Erro
    With adc.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser impresso!", vbExclamation
            Exit Sub
        End If
        imprContagem !seq_cnt
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnLançar_Click()
    Dim strSQL As String
    strSQL = F_BuscaUmCampo("seq_cnt", "tab_contagens", "status_cnt='N'")
        If strSQL = "" Then
            MsgBox "Gere primeiramente uma contagem! ", vbExclamation
            Exit Sub
        End If
    frmLanCont.Show 1
    cmbxCont_Click
End Sub

Private Sub cmndbtnNovo_Click()
    Dim strSQL As String
    If MsgBox("Deseja Realmente abrir nova contagem?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = F_BuscaUmCampo("seq_cnt", "tab_contagens", "status_cnt='N' and tipo_Con='" & pc.Tag & "'")
        If strSQL <> "" Then
            MsgBox "Não é possível gerar um nova contagem pois a contagem " & _
                strSQL & " se encontra em aberto!", vbExclamation
            Exit Sub
        End If
        NovaContagem
    End If
End Sub
Sub NovaContagem()
    Dim strSQL As String
    Dim sCod As String
    
    
    
    sCod = ChaveSQL("Tab_Contagens", "seq_cnt", "000#")
    strSQL = strSQL & "Insert into tab_contagens(seq_cnt,data_cnt,cod_usu,status_cnt,tipo_con)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'N', "
    strSQL = strSQL & "'" & pc.Tag & "') "
    conn.Execute strSQL
    frmLanCont.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        pc.BackColor = vbGreen
        pc.Tag = "F"
        busca
    ElseIf KeyCode = vbKeyF9 Then
        pc.BackColor = vbRed
        pc.Tag = "R"
        busca
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    pc.Tag = "R"
    pc.BackColor = vbRed
    adc.ConnectionString = banco.Connect
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_contagens "
    strSQL = strSQL & " where tipo_con='" & pc.Tag & "' "
    strSQL = strSQL & " order by data_cnt desc "
    rPesq.Open strSQL, conn_client
    With rPesq
        cmbxCont.Clear
        Do While Not .EOF
            cmbxCont.AddItem !seq_cnt & "-" & !data_cnt
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
