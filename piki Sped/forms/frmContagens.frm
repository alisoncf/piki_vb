VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmContagens 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contagens"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10275
   Icon            =   "frmContagens.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar Contagem"
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "cria um novo inventário."
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnExcluir 
         Caption         =   "&Excluir Produto"
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":2511
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "cria um novo inventário."
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnEfetiva 
         Caption         =   "&Efetivar Contagem "
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":31DB
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "cria um novo inventário."
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnLançar 
         Caption         =   "&Lançar Produto"
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":3325
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "cria um novo inventário."
         Top             =   4800
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
         ItemData        =   "frmContagens.frx":346F
         Left            =   120
         List            =   "frmContagens.frx":347C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Abrir Nova Contagem"
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":349D
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
         Picture         =   "frmContagens.frx":3601
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fecha formulário."
         Top             =   6480
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   120
         Picture         =   "frmContagens.frx":376C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Emite Inventário"
         Top             =   5640
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
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   10095
      TabIndex        =   14
      Top             =   7440
      Width           =   10095
   End
   Begin VB.Data dtInvent 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmContagens.frx":38B6
      Height          =   7215
      Left            =   2400
      OleObjectBlob   =   "frmContagens.frx":38CD
      TabIndex        =   12
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   135
      Left            =   5640
      TabIndex        =   13
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
    strSQL = "select ref_pro,tc.cod_usu,tc.tipo_con,tc.data_cnt,tc.status_cnt,tp.nome_pro,ti.* from tab_produtos as tp,tab_contagens as tc, "
    strSQL = strSQL & "tab_itcont as ti where ti.cod_pro=tp.cod_pro and tc.seq_cnt=ti.seq_cnt "
    
    strSQL = strSQL & " and tc.seq_cnt='" & Left(cmbxCont.Text, 5) & "' "
    strSQL = strSQL & " order by nome_pro"
    dtInvent.RecordSource = strSQL
    dtInvent.Refresh
    With dtInvent.Recordset
    If .RecordCount > 0 Then
        dbgrd.caption = "Contagem de " & !data_cnt & " |   Usuário: " & F_BuscaUmCampo("guerra_fun", "tab_funcionarios", "cod_fun='" & !cod_usu & "'")
        If !status_cnt = "N" Then
            dbgrd.caption = dbgrd.caption & "   | Status:" & " Em Aberto"
            cmndbtnEfetiva.Enabled = True
            cmndbtnLançar.Enabled = True
            cmndbtnNovo.Enabled = False
            cmndbtnCancel.Enabled = True
            cmndbtnExcluir.Enabled = True
        Else
            dbgrd.caption = dbgrd.caption & "   | Status:" & " Efetivada"
            cmndbtnEfetiva.Enabled = False
            cmndbtnLançar.Enabled = False
            cmndbtnNovo.Enabled = True
            cmndbtnCancel.Enabled = False
            cmndbtnExcluir.Enabled = False
        End If
    Else
        cmndbtnLançar.Enabled = True
        cmndbtnEfetiva.Enabled = False
        cmndbtnNovo.Enabled = False
        cmndbtnCancel.Enabled = True
        cmndbtnExcluir.Enabled = True
    End If
    End With
End Sub

Private Sub cmndbtnCancel_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro:
    With dtInvent.Recordset
        If MsgBox("Confirma exclusão da Contagem " & Left(cmbxCont.Text, 5) & "?", vbYesNo + vbExclamation) = vbYes Then
            strSQL = "delete from tab_itcont where seq_cnt='"
            strSQL = strSQL & Left(cmbxCont.Text, 5) & "'"
            
            Conn.Execute strSQL
            
            strSQL = "delete from tab_contagens where seq_cnt='"
            strSQL = strSQL & Left(cmbxCont.Text, 5) & "'"
            
            Conn.Execute strSQL
        End If
        busca
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
    
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
    
End Sub
Sub efetivaCont()
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim sCampo As String
    Dim sSEQ As String
    With dtInvent.Recordset
        .MoveFirst
        If !tipo_con = "R" Then
            sCampo = "er_pro"
        Else
            sCampo = "ef_pro"
        End If
        Do While Not .EOF
            
            
            
            strSQL = "update tab_produtos set " & sCampo & "=" & F_GravaDec(!qtd_itc) & " "
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            
            Conn.Execute strSQL
            .MoveNext
        Loop
        strSQL = "update tab_contagens set status_cnt='S' where seq_cnt='"
        strSQL = strSQL & Left(cmbxCont.Text, 5) & "'"
        Conn.Execute strSQL
        MsgBox "Contagem Efetivada!", vbInformation
        cmbxCont_Click
    End With
Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnExcluir_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro:
    With dtInvent.Recordset
        If MsgBox("Confirma exclusão do produto " & dtInvent.Recordset!nome_pro & "?", vbYesNo + vbExclamation) = vbYes Then
            strSQL = "delete from tab_itcont where seq_cnt='"
            strSQL = strSQL & Left(cmbxCont.Text, 5) & "'"
            strSQL = strSQL & " and cod_pro='" & dtInvent.Recordset!cod_pro & "' "
            strSQL = strSQL & " and lote_pro='" & dtInvent.Recordset!lote_pro & "' "
            
            Conn.Execute strSQL
            
        End If
        cmbxCont_Click
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_Erro
    With dtInvent.Recordset
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
        strSQL = F_BuscaUmCampo("seq_cnt", "tab_contagens", "status_cnt='N' and tipo_Con='" & pc.tag & "'")
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
    
    
    
    sCod = MaxKeySQL("Tab_Contagens", "seq_cnt", "", "0000#")
    strSQL = "Insert into tab_contagens(seq_cnt,data_cnt,cod_usu,status_cnt,tipo_con)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'N', "
    strSQL = strSQL & "'" & pc.tag & "') "
    Conn.Execute strSQL
    'frmLanCont.Show 1
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
        pc.tag = "F"
        busca
    ElseIf KeyCode = vbKeyF9 Then
        pc.BackColor = vbRed
        pc.tag = "R"
        busca
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    pc.tag = "R"
    pc.BackColor = vbRed
    dtInvent.Connect = Banco.Connect
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_contagens "
    strSQL = strSQL & " where tipo_con='" & pc.tag & "' "
    strSQL = strSQL & " order by seq_cnt desc "
    rPesq.Open strSQL, Conn
    With rPesq
        cmbxCont.Clear
        Do While Not .EOF
            cmbxCont.AddItem !seq_cnt & "-" & !data_cnt
            .MoveNext
        Loop
        .Close
    End With
    
End Sub

