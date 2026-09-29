VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelGuias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Guias:"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   Icon            =   "frmRelGuias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9525
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   2520
      TabIndex        =   14
      Top             =   0
      Width           =   6975
      Begin MSDBGrid.DBGrid dbgrdRel 
         Bindings        =   "frmRelGuias.frx":08CA
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "frmRelGuias.frx":08DE
         TabIndex        =   15
         Top             =   120
         Width           =   6795
      End
      Begin VB.Data dtRel 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   ""
         Top             =   840
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblSoma 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3600
         TabIndex        =   19
         Top             =   5520
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soma:"
         Height          =   195
         Left            =   3000
         TabIndex        =   18
         Top             =   5520
         Width           =   450
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1320
         TabIndex        =   17
         Top             =   5520
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guias Listadas:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   5520
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.ComboBox cmbxRelat 
         Height          =   315
         ItemData        =   "frmRelGuias.frx":17B9
         Left            =   120
         List            =   "frmRelGuias.frx":17C9
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtbxdataFin 
         BackColor       =   &H000000C0&
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         BackColor       =   &H000000C0&
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmbxOrderBy 
         Height          =   315
         ItemData        =   "frmRelGuias.frx":180E
         Left            =   120
         List            =   "frmRelGuias.frx":181B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1620
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnPesq 
         Caption         =   "&Pesquisar"
         Height          =   615
         Left            =   120
         Picture         =   "frmRelGuias.frx":183C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2100
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   120
         Picture         =   "frmRelGuias.frx":1BC6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4260
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   120
         Picture         =   "frmRelGuias.frx":1F50
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4980
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnVisualizar 
         Caption         =   "&Visualizar"
         Height          =   615
         Left            =   120
         Picture         =   "frmRelGuias.frx":2392
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3540
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnConfig 
         Caption         =   "&Configurar Impressão"
         Height          =   615
         Left            =   120
         Picture         =   "frmRelGuias.frx":27D4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2820
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Relatório:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lbldataF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   840
         Width           =   765
      End
      Begin VB.Label lbldataI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1380
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRelGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub P_Soma()
    On Error Resume Next
    P_Status "Somando..."
    Dim cTotal As Currency
    With dtRel.Recordset
        Do While Not .EOF
            cTotal = cTotal + !Valor_Gui
            .MoveNext
        Loop
        dtRel.Recordset.MoveFirst
        lblCount.caption = .RecordCount
        lblSoma.caption = FormatVal(CStr(cTotal))
    End With
    P_Status ""
    
End Sub

Private Sub cmbxRelat_Click()
    If cmbxRelat.ListIndex = 2 Then
        txtbxdataFin.Enabled = False
        lbldataF.Enabled = False
        txtbxDataIni.Enabled = False
        lbldataI.Enabled = False
    Else
        txtbxdataFin.Enabled = True
        lbldataF.Enabled = True
        txtbxDataIni.Enabled = True
        lbldataI.Enabled = True
    End If
End Sub

Private Sub cmndbtnImprimir_Click()
    If dtRel.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que ser impresso!", vbExclamation
        Exit Sub
    Else
        If MsgBox("Confirma impressão do Relatório?", vbYesNo + vbQuestion) = vbYes Then
            P_Status "Imprimindo..."
            P_Rel Printer
            P_Status ""
        End If
    End If
End Sub
Sub P_Rel(obj As Object, Optional bobj As Boolean = True)
    On Error GoTo trata_erro
    If bobj = True Then
        obj.PrintQuality = -1
    End If
    P_CabRelTimbr "Relatório de Guias(" & cmbxRelat.Text & ")", obj
    obj.CurrentY = obj.CurrentY + 2
    P_ImprX obj, "Número", 10, True
    P_ImprX obj, "Imposto", 30, True
    P_ImprX obj, "Vencimento", 80, True
    P_ImprX obj, "S", 100, True
    P_ImprX obj, "Valor", 120, False, True
    P_ImprLinha obj
    With dtRel.Recordset
        Do While Not .EOF
            P_ImprX obj, !num_Gui, 10, True
            P_ImprX obj, !nome_Imp, 30, True
            P_ImprX obj, !Venc_Gui, 80, True
            P_ImprX obj, !Status_Gui, 100, True
            P_ImprX obj, FormatVal(!Valor_Gui), 120, False, True
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                If bobj = True Then
                    P_Rodape "Continua...", True, "Pág.:" & obj.Page
                Else
                    P_RodapeVis obj, "Continua...", "Pág.:XXX:"
                End If
                P_ImprX obj, "Número", 10, True
                P_ImprX obj, "Imposto", 30, True
                P_ImprX obj, "Vencimento", 80, True
                P_ImprX obj, "S", 100, True
                P_ImprX obj, "Valor", 120, False, True
                P_ImprLinha obj
            End If
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "Guias listadas:" & lblCount, 20
        P_ImprX obj, "Soma..:" & lblSoma, 80
        If bobj = True Then
            P_Rodape ConstEmpresa, True, "Pág.:" & obj.Page
            obj.EndDoc
        Else
            P_RodapeVis obj, ConstEmpresa, "Pág.:XXX"
        End If
    End With
    
Exit Sub
trata_erro:
    If bobj = True Then
        obj.KillDoc
    End If
    MsgBox Err.Description, vbCritical, constNome_sis
End Sub
Private Sub cmndbtnPesq_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    If cmbxOrderBy.ListIndex = 0 Then
        sOrderBy = "tg.num_gui,tg.venc_gui "
    ElseIf cmbxOrderBy.ListIndex = 1 Then
        sOrderBy = "ti.nome_imp,tg.venc_gui"
    Else
        sOrderBy = "tg.venc_gui,ti.nome_imp"
    End If
    strSQL = "select tg.*,ti.nome_imp from tab_guias as tg, tab_impostos as ti "
    strSQL = strSQL & " where tg.cod_imp=ti.cod_imp and tg.cod_Cli='" & strCOD_CLI & "' "
    If cmbxRelat.ListIndex = 0 Then
        strSQL = strSQL & " and tg.status_gui='1'"
    ElseIf cmbxRelat.ListIndex = 1 Then
        strSQL = strSQL & " and tg.status_gui='2'"
    ElseIf cmbxRelat.ListIndex = 2 Then
        strSQL = strSQL & " and tg.status_gui='1'"
        strSQL = strSQL & " and venc_gui<=" & F_PesqData(Date)
    End If
    If cmbxRelat.ListIndex <> 2 Then
        strSQL = strSQL & " and venc_gui>=" & F_PesqData(txtbxDataIni.Text)
        strSQL = strSQL & " and venc_gui<=" & F_PesqData(txtbxdataFin.Text)
    End If
    strSQL = strSQL & " order by " & sOrderBy
    dtRel.RecordSource = strSQL
    dtRel.Refresh
    If dtRel.Recordset.RecordCount = 0 Then
        MsgBox "Não há registros na sua pesquisa!", vbExclamation
    Else
        P_Soma
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVisualizar_Click()
    On Error Resume Next
    frmVisualiza.Show 1
    P_Rel frmVisualiza.Pic, False
End Sub

Private Sub dbgrdRel_GotFocus()
    dbgrdRel.BorderStyle = 1
End Sub

Private Sub dbgrdRel_LostFocus()
    dbgrdRel.BorderStyle = 0
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    cmbxOrderBy.ListIndex = 0
    cmbxRelat.ListIndex = 0
    dtRel.Connect = banco.Connect
End Sub

Private Sub txtbxdataFin_GotFocus()
        P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxdataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxdataFin_LostFocus()
    P_AjustaData txtbxdataFin
    If F_DataMaior(txtbxDataIni.Text, txtbxdataFin.Text) Then
        MsgBox "A data Inicial não pode ser maior que a Data Final!", vbExclamation
    End If
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
