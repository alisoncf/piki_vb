VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiGui 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Guias de:"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmEdiGui.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5745
   Begin MSDBGrid.DBGrid dbgrdGui 
      Bindings        =   "frmEdiGui.frx":0442
      Height          =   4215
      Left            =   120
      OleObjectBlob   =   "frmEdiGui.frx":0456
      TabIndex        =   11
      Top             =   1200
      Width           =   5535
   End
   Begin VB.CommandButton cmndbtnSair 
      BackColor       =   &H8000000A&
      Caption         =   "&Sair"
      Height          =   735
      Left            =   4800
      Picture         =   "frmEdiGui.frx":14D9
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnBaixar 
      Caption         =   "&Baixar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3840
      Picture         =   "frmEdiGui.frx":191B
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnExcluir 
      Caption         =   "E&xcluir"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2880
      Picture         =   "frmEdiGui.frx":1D5D
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnEditar 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1920
      Picture         =   "frmEdiGui.frx":20E7
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   855
   End
   Begin VB.ComboBox cmbxStatus 
      Height          =   315
      ItemData        =   "frmEdiGui.frx":2231
      Left            =   2040
      List            =   "frmEdiGui.frx":223E
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox cmbxImposto 
      Height          =   315
      ItemData        =   "frmEdiGui.frx":227F
      Left            =   120
      List            =   "frmEdiGui.frx":2281
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Data dtGui 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtbxDataIni 
      Height          =   285
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   3
      Top             =   270
      Width           =   975
   End
   Begin VB.TextBox txtbxdataFin 
      Height          =   285
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   5
      Top             =   270
      Width           =   975
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Pesquisar"
      Height          =   855
      Left            =   4440
      Picture         =   "frmEdiGui.frx":2283
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox cmbxOrder 
      Height          =   315
      ItemData        =   "frmEdiGui.frx":260D
      Left            =   120
      List            =   "frmEdiGui.frx":261D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Status:"
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Imposto:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Venc. Inicial:"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Venc. Final:"
      Height          =   195
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   840
   End
   Begin VB.Label Label4 
      Caption         =   "&Ordernar Por"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmEdiGui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBaixar_Click()
    
    Dim strSQL As String
    On Error GoTo trata_erro
    If MsgBox("Confirma baixa da gui: " & dtGui.Recordset!num_Gui & " ?", vbYesNo _
                + vbQuestion) = vbYes Then
        strSQL = "update tab_guias set status_gui='2' where "
        strSQL = strSQL & "cod_gui='" & dtGui.Recordset!Cod_Gui & "' "
        banco.Execute strSQL
        MsgBox "Guia Baixada com sucesso!", vbInformation
        dtGui.Refresh
    End If
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbYesNo + vbQuestion
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sOrder As String
    If cmbxImposto.Text = "" Then
        MsgBox "Informe o Imposto pelo qual deseja buscar guias!", vbExclamation
        cmbxImposto.SetFocus
        Exit Sub
    End If
    If cmbxOrder.ListIndex = 0 Then
        sOrder = "num_Gui,venc_gui,cod_imp "
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrder = "cod_imp,num_Gui,venc_gui "
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrder = "venc_gui,cod_imp,num_gui "
    Else
        sOrder = "status_gui,cod_imp,venc_gui"
    End If
    strSQL = "select tg.*, ti.nome_imp from "
    strSQL = strSQL & "Tab_guias as tg, "
    strSQL = strSQL & "tab_impostos as ti where "
    strSQL = strSQL & " ti.cod_imp=tg.cod_imp and "
    strSQL = strSQL & " tg.cod_cli='" & strCOD_CLI & "' "
    strSQL = strSQL & " and tg.cod_imp='" & Left(cmbxImposto.Text, 2) & "'"
    If cmbxStatus.ListIndex <> 0 Then
        strSQL = strSQL & " and tg.status_gui='" & Left(cmbxStatus.Text, 1) & "' "
    End If
    If txtbxDataFin.Text <> "" Then
        strSQL = strSQL & " and tg.venc_gui <= " & F_GravaData(txtbxDataFin.Text)
    End If
    
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tg.venc_gui >= " & F_GravaData(txtbxDataIni.Text)
    End If
    strslq = strSQL & "order by " & sOrder
    dtGui.RecordSource = strSQL
    dtGui.Refresh
End Sub

Private Sub cmndbtnEditar_Click()
    Dim iPos As Integer
    On Error GoTo trata_erro
    If dtGui.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que ser editado!", vbExclamation
        Exit Sub
    End If
    With dtGui.Recordset
        frmLanGuias.Tag = !Cod_Gui
        frmLanGuias.caption = "Guia:" & !num_Gui & " - Imposto:" & !nome_Imp
        
        frmLanGuias.txtbxJuros_Gui.Enabled = True
        frmLanGuias.txtbxJuros_Gui.Text = !Juros_Gui
        frmLanGuias.txtbxNum_Gui.Text = !num_Gui
        frmLanGuias.txtbxValor_Gui.Text = !Valor_Gui
        frmLanGuias.txtbxVenc_GUi.Text = !Venc_Gui
        frmLanGuias.cmbxCod_Imp.Text = !cod_Imp & "-" & F_BuscaUmCampo("Nome_Imp", "tab_Impostos", _
                    " cod_imp='" & !cod_Imp & "'")
        iPos = .AbsolutePosition
        On Error Resume Next
        frmLanGuias.Show 0
        
        On Error GoTo trata_erro
        
        
    End With
Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnExcluir_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    If MsgBox("Deseja realmente excluir a guia: " & dtGui.Recordset!num_Gui & " ?", vbYesNo _
                + vbCritical) = vbYes Then
        strSQL = "delete from tab_guias set status_gui='2' where "
        strSQL = strSQL & "cod_gui='" & dtGui.Recordset!Cod_Gui & "' "
        banco.Execute strSQL
        MsgBox "Guia Excluída!", vbInformation
        dtGui.Refresh
    End If
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
    


Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrdGui_GotFocus()
    dbgrdGui.BorderStyle = 1
End Sub

Private Sub dbgrdGui_LostFocus()
    dbgrdGui.BorderStyle = 0
End Sub

Private Sub dtGui_Reposition()
    If dtGui.Recordset.RecordCount = 0 Then
        cmndbtnBaixar.Enabled = False
        cmndbtnEditar.Enabled = False
        cmndbtnExcluir.Enabled = False
        Exit Sub
    End If
    If dtGui.Recordset!Status_Gui = "1" Then
        cmndbtnBaixar.Enabled = True
        cmndbtnEditar.Enabled = True
        cmndbtnExcluir.Enabled = True
    Else
        cmndbtnBaixar.Enabled = False
        cmndbtnEditar.Enabled = False
        cmndbtnExcluir.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    If Not F_SelectCli(Me) Then Unload Me
    Me.caption = "Edição de Guias de: " & mdiFrm.StatusBar1.Panels(1).Text
    cmbxOrder.ListIndex = 0
    cmbxStatus.ListIndex = 0
End Sub

Private Sub Form_Load()
    mdiFrm.mnIncGuias.Enabled = False
    mdiFrm.mnGuias.Enabled = False
    P_CentraForm Me
    Me.caption = "Edição de Guias de: " & mdiFrm.StatusBar1.Panels(1).Text
    P_PreencheImpostos cmbxImposto
    dtGui.Connect = banco.Connect
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiFrm.mnIncGuias.Enabled = True
    mdiFrm.mnGuias.Enabled = True
End Sub

Private Sub txtbxdataFin_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxdataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxdataFin_LostFocus()
    P_AjustaData txtbxDataFin
    If F_DataMaior(txtbxDataIni.Text, txtbxDataFin.Text) Then
        MsgBox "A data Final não pode ser menor que a Inicial!", vbExclamation
        txtbxDataFin.SetFocus
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
     If F_DataMaior(txtbxDataIni.Text, txtbxDataFin.Text) Then
        MsgBox "A data Final não pode ser menor que a Inicial!", vbExclamation
        txtbxDataFin.SetFocus
    End If
End Sub
