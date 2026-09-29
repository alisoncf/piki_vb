VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelResumoConsData 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumo de Movimentação do Consignado por Datas"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      ScaleHeight     =   585
      ScaleWidth      =   13065
      TabIndex        =   21
      Top             =   7560
      Width           =   13095
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   14
         Top             =   240
         Width           =   4695
      End
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmRelResumoMovConsData.frx":0000
         Left            =   8640
         List            =   "frmRelResumoMovConsData.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmRelResumoMovConsData.frx":0042
         Left            =   10680
         List            =   "frmRelResumoMovConsData.frx":0055
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnProd 
         Height          =   285
         Left            =   8280
         Picture         =   "frmRelResumoMovConsData.frx":007C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto:"
         Height          =   195
         Left            =   3600
         TabIndex        =   13
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   8640
         TabIndex        =   16
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   10680
         TabIndex        =   18
         Top             =   0
         Width           =   510
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
         TabIndex        =   23
         Top             =   120
         Width           =   480
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
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.CommandButton cmndbtnVisualizaAcerto 
         Caption         =   "Ver Acerto"
         Height          =   615
         Left            =   9360
         Picture         =   "frmRelResumoMovConsData.frx":0406
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir Acerto"
         Height          =   615
         Left            =   10800
         Picture         =   "frmRelResumoMovConsData.frx":0790
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   12240
         Picture         =   "frmRelResumoMovConsData.frx":0B1A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   7920
         Picture         =   "frmRelResumoMovConsData.frx":0C0F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   6480
         Picture         =   "frmRelResumoMovConsData.frx":101C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   285
         Left            =   5880
         Picture         =   "frmRelResumoMovConsData.frx":13A6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
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
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelResumoMovConsData.frx":1730
      Height          =   6495
      Left            =   120
      OleObjectBlob   =   "frmRelResumoMovConsData.frx":1741
      TabIndex        =   20
      Top             =   960
      Width           =   13455
   End
End
Attribute VB_Name = "frmRelResumoConsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxOrder_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim sCampo As String
    Dim sCod As String
    Dim sTa As String
    Dim sTd As String
    Dim sTt As String
    Dim sOrderby As String
    Dim bNaoExec As Boolean
   ' On Error Resume Next
    Dim rPesq As Recordset
    P_Status "Buscando..."
    If cmbxOrder.ListIndex = 0 Then
        sOrderby = " order by cod"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderby = " order by nome"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderby = " order by t desc"
    ElseIf cmbxOrder.ListIndex = 3 Then
        sOrderby = " order by a desc"
    Else
        sOrderby = " order by d desc"
    End If
    
    
    strSQL = "drop table if exists tab_resumotmp"
    Banco.Execute strSQL
    strSQL = "create table tab_resumotmp(cod_rep char(5) not null, valor decimal(14,2) not null, qtd integer not null,tipo char(1) not null)"
    Banco.Execute strSQL
    strSQL = "delete from tab_resumotmp"
    Banco.Execute strSQL
    
    
    strSQL = "insert into tab_resumotmp  select cod_rep,sum(unit_con*qtd_con) as total,count(*) as qtd,'T' as tipo  from tab_Consignacoes "
    strSQL = strSQL & " where seq_con <> '' "
    If txtbxProduto.tag <> "" Then
        strSQL = strSQL & " and cod_pro='" & txtbxProduto.tag & "' "
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_con<=" & F_GravaDataTempo(txtbxdataFin.Text & " 23:59:59")
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_con>=" & F_GravaDataTempo(txtbxDataIni.Text & " 00:00:00")
    End If
    If txtbxRepresentante.tag <> "" Then
        strSQL = strSQL & " and cod_rep='" & txtbxRepresentante.tag & "' "
    End If
    'If cmbxStatus.Text <> "" Then
    '    strSQL = strSQL & " and status_ace='" & Left(cmbxStatus.Text, 1) & "' "
    'End If
    strSQL = strSQL & " group by cod_rep "
    
    Banco.Execute strSQL
    
    
    strSQL = "insert into tab_resumotmp  select cod_rep,"
    If txtbxProduto.tag = "" Then
        strSQL = strSQL & " sum(valor_ace) as total,"
        strSQL = strSQL & " count(*) as qtd,"
        strSQL = strSQL & " 'A' as tipo  from tab_Acertos ta "
        strSQL = strSQL & " where status_ace='2' "
    Else
        strSQL = strSQL & " sum(ti.unit_ite*qtd_ite) as total,"
        strSQL = strSQL & " count(*) as qtd,"
        strSQL = strSQL & " 'A' as tipo  from tab_Acertos ta, tab_itacertos as ti "
        strSQL = strSQL & " where ta.seq_ace=ti.seq_ace "
        strSQL = strSQL & " and status_ace='2' "
        strSQL = strSQL & " and cod_pro='" & txtbxProduto.tag & "' "
    End If
    
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_ace<=" & F_GravaData(txtbxdataFin.Text)
    End If
    
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_ace>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxRepresentante.tag <> "" Then
        strSQL = strSQL & " and cod_rep='" & txtbxRepresentante.tag & "' "
    End If
    
    strSQL = strSQL & " group by cod_rep "
    
    Banco.Execute strSQL
    
    
    
    strSQL = "insert into tab_resumotmp  select cod_rep,"
    strSQL = strSQL & " sum(pv_pro*qtd_dev) as total,count(*) as qtd,'D' as tipo  "
    strSQL = strSQL & " from tab_devconsign as td,"
    strSQL = strSQL & " tab_itdevconsign as ti, tab_produtos as tp "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro and ti.seq_dev=td.seq_dev "
    
    
     If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_dev<=" & F_GravaDataTempo(txtbxdataFin.Text & " 23:59:59")
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_dev>=" & F_GravaDataTempo(txtbxDataIni.Text & " 00:00:00")
    End If
    If txtbxRepresentante.tag <> "" Then
        strSQL = strSQL & " and cod_rep='" & txtbxRepresentante.tag & "' "
    End If
    If txtbxProduto.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pro='" & txtbxProduto.tag & "' "
    End If
    strSQL = strSQL & " group by cod_rep "
    
    Banco.Execute strSQL
    
    strSQL = "select tt.cod_rep,tr.nome_rep as nome,sum(valor) as val, sum(qtd) as q, tipo from tab_resumotmp as tt,"
    strSQL = strSQL & " tab_representantes As tr"
    strSQL = strSQL & " where tr.cod_rep = tt.cod_rep"
    strSQL = strSQL & " group by tt.cod_rep,tipo"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    strSQL = "repair table tab_resumotmpresult"
    Banco.Execute strSQL
    strSQL = "drop table if exists tab_resumotmpresult "
    Banco.Execute strSQL
    strSQL = "create table tab_resumotmpresult(cod char(5) not null,nome char(60) not null, "
    strSQL = strSQL & " a decimal(14,2),qa integer,"
    strSQL = strSQL & " d decimal(14,2),qd integer,"
    strSQL = strSQL & " t decimal(14,2),qt integer) "
    Banco.Execute strSQL
    strSQL = ""
    With rPesq
        '.MoveFirst
        Do While Not .EOF
            If sCod <> !cod_Rep Then
                If strSQL <> "" Then
                    'strSQL = strSQL & sTa & sTd & sTt
                    
                    'Banco.Execute strSQL
                    'bNaoExec = True
                End If
                sCod = !cod_Rep
                strSQL = "insert into tab_resumotmpresult (cod,nome,a,qa,d,qd,t,qt)values("
                strSQL = strSQL & "'" & !cod_Rep & "',"
                strSQL = strSQL & "'" & !Nome & "',"
                sTa = "0,0,"
                sTd = "0,0,"
                sTt = "0,0)"
                Banco.Execute strSQL & sTa & sTd & sTt
            End If
            If !tipo = "A" Then
                sTa = F_GravaDec(!val) & "," & !q & ","
            ElseIf !tipo = "D" Then
                sTd = F_GravaDec(!val) & "," & !q & ","
            ElseIf !tipo = "T" Then
                sTt = F_GravaDec(!val) & "," & !q & " )"
                Banco.Execute strSQL & sTa & sTd & sTt
            End If
            .MoveNext
        Loop
        
        .Close
    End With
    strSQL = "select * from tab_Resumotmpresult where (a>0 or d>0 or t>0)"
    If sOrderby <> "" Then
        strSQL = strSQL & sOrderby
    End If
    dt.RecordSource = strSQL
    dt.Refresh
    P_Status ""
    
End Sub
Sub soma()
    Dim cTotal As Currency
    With dt.Recordset
        Do While Not .EOF
            cTotal = cTotal + !valor_ace
            
            .MoveNext
        Loop
        lblVendas.caption = FormatVal(cTotal)
    End With
End Sub

Private Sub cmndbtnBuscaRep_Click()
    buscaRep
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
    P_DbgrdToHtml "Relatório de Resumo de Movimentação do Consignado " & txtbxDataIni.Text & " até " & _
        txtbxdataFin.Text, dbgrd, dt, txtbxRepresentante.Text, "1"
    
    
    
End Sub

Private Sub cmndbtnProd_Click()
    Dim p As Produto
    p = F_BuscaProdCabecalho(txtbxProduto.Text)
    If p.cod_pro <> "" Then
        txtbxProduto.Text = p.nome_pro
        txtbxProduto.tag = p.cod_pro
        cmndbtnBuscar_Click
    End If
    
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

Private Sub Form_Load()
    dt.Connect = Banco.Connect
    
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxProduto_Change()
    txtbxProduto.tag = ""
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnProd_Click
    End If
    
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
