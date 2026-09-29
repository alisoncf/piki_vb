VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiCai 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edição de Fluxo de Caixa"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11070
   Icon            =   "frmEdiCai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEdiCai.frx":014A
         Height          =   5895
         Left            =   2400
         OleObjectBlob   =   "frmEdiCai.frx":015F
         TabIndex        =   17
         Top             =   240
         Width           =   8415
      End
      Begin VB.Frame Frame6 
         Caption         =   "&Turno:"
         Height          =   615
         Left            =   4440
         TabIndex        =   32
         Top             =   6120
         Width           =   1935
         Begin VB.ComboBox cmbxTurno 
            Height          =   315
            ItemData        =   "frmEdiCai.frx":16E4
            Left            =   120
            List            =   "frmEdiCai.frx":16F1
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cai&xa:"
         Height          =   615
         Left            =   2400
         TabIndex        =   30
         Top             =   6120
         Width           =   1935
         Begin VB.ComboBox cmbxCai 
            Height          =   315
            ItemData        =   "frmEdiCai.frx":170F
            Left            =   120
            List            =   "frmEdiCai.frx":171C
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2175
         Begin VB.CommandButton cmndbtnBuscar 
            Height          =   285
            Left            =   1800
            Picture         =   "frmEdiCai.frx":173A
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2760
            Width           =   255
         End
         Begin VB.ComboBox cmbxTIpo 
            Height          =   315
            ItemData        =   "frmEdiCai.frx":1AC4
            Left            =   120
            List            =   "frmEdiCai.frx":1AD1
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   1935
         End
         Begin VB.ComboBox cmbxOrd 
            Height          =   315
            ItemData        =   "frmEdiCai.frx":1AEF
            Left            =   120
            List            =   "frmEdiCai.frx":1AFF
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1560
            Width           =   1215
         End
         Begin VB.ComboBox cmbxBusca 
            Height          =   315
            ItemData        =   "frmEdiCai.frx":1B27
            Left            =   120
            List            =   "frmEdiCai.frx":1B34
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtbxExpr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   15
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtbxDataFin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   13
            Top             =   2160
            Width           =   975
         End
         Begin VB.OptionButton optnAsc 
            Caption         =   "Asc"
            Height          =   255
            Left            =   1320
            TabIndex        =   8
            Top             =   1440
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optnDesc 
            Caption         =   "Desc"
            Height          =   255
            Left            =   1320
            TabIndex        =   9
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Ordenar:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblExpr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E&xpressão:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   2520
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Buscar Por:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblDataIni 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Data Inicial:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1920
            Width           =   840
         End
         Begin VB.Label lbldataFin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data &Final:"
            Height          =   195
            Left            =   1200
            TabIndex        =   12
            Top             =   1920
            Width           =   765
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   6480
         TabIndex        =   23
         Top             =   6120
         Width           =   4335
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Crédito:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   0
            Width           =   540
         End
         Begin VB.Label lblCred_din 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Tag             =   "!"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Débito:"
            Height          =   195
            Left            =   1080
            TabIndex        =   27
            Top             =   0
            Width           =   510
         End
         Begin VB.Label lblDeb_Din 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   26
            Tag             =   "!"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   2040
            TabIndex        =   25
            Top             =   0
            Width           =   405
         End
         Begin VB.Label lblTotal_Din 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Tag             =   "!"
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Data dtServ 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2040
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Frame Frame2 
         Height          =   3615
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   2175
         Begin VB.CommandButton cmndbtnImprimir 
            Caption         =   "Imprimir"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiCai.frx":1B5B
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiCai.frx":1F68
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   2835
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiCai.frx":20D3
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Novo"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiCai.frx":221D
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   160
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnExcl 
            Caption         =   "E&xcluir"
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiCai.frx":25A7
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1515
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmEdiCai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Change()

End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    If txtbxDAtaFIn.Visible = True Then
        If F_DataMaior(txtbxDataIni.Text, txtbxDAtaFIn.Text) Then
            MsgBox "A data final deve ser maior que a inicial!", vbExclamation
            txtbxDAtaFIn.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "select tc.*,tb.nome_cai from tab_caixa as tc, tab_caixas as tb where "
    strSQL = strSQL & "tc.cod_cai=tb.cod_cai  and "
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " historico_cai like '%" & txtbxExpr.Text & "%'"
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & " num_ven like '" & txtbxExpr.Text & "%'"
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " valor_cai = " & F_GravaDec(txtbxExpr.Text)
    End If
    strSQL = strSQL & " and data_cai >= " & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and data_cai <= " & F_GravaData(txtbxDAtaFIn.Text)
    If cmbxTipo.ListIndex <> 0 Then
        strSQL = strSQL & " and tipo_cai='" & Left(cmbxTipo.Text, 1) & "' "
    End If
    
    If cmbxCai.Text <> "" Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCai.Text, 2) & "'"
    End If
    If cmbxTurno.Text <> "" Then
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    End If
    
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " order by data_cai,seq_cai"
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & " order by seq_cai"
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " order by historico_cai, seq_cai"
    ElseIf cmbxBusca.ListIndex = 3 Then
        strSQL = strSQL & " order by valor_cai,seq_cai"
    End If
    
    If optnAsc.Value Then
        strSQL = strSQL & " asc "
    Else
        strSQL = strSQL & " desc "
    End If
    dtServ.RecordSource = strSQL
    dtServ.Refresh
    P_Resumo
End Sub
Sub P_Resumo()
Dim cCheDeb As Currency
    Dim cCheCred As Currency
    Dim cDinDeb As Currency
    Dim cDinCred As Currency
    
    With dtServ.Recordset
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            
            
                If !tipo_cai = "C" Then
                    cDinCred = cDinCred + !valor_Cai
                Else
                    cDinDeb = cDinDeb + !valor_Cai
                End If
            
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    'lblCred_Che.caption = FormatVal(cCheCred)
    lblCred_din.caption = FormatVal(cDinCred)
    lblDeb_Din.caption = FormatVal(cDinDeb)
 '   lblDeb_Che.caption = FormatVal(cCheDeb)
  '  lblTotal_Che.caption = FormatVal(cCheCred - cCheDeb)
    lblTotal_Din.caption = FormatVal(cDinCred - cDinDeb)
'    lblSaldo.caption = FormatVal(CCur(lblTotal_Che.caption) + CCur(lblTotal_Din.caption))
End Sub
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmLanCaixa
    With dtServ.Recordset
        fr.Tag = !seq_Cai
        fr.txtbxData_Cai.Text = !data_cai
        fr.txtbxHist.Text = !historico_cai
        fr.txtbxValor_Cai.Text = FormatVal(!valor_Cai)
        If !tipo_cai = "C" Then
            fr.optnCerd.Value = True
        Else
            fr.optnDeb.Value = True
        End If
        fr.optnCerd.Enabled = False
        fr.optnDeb.Enabled = False
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    If dtServ.Recordset.RecordCount = 0 Then
        cmndbtnExcl.Enabled = False
        Exit Sub
    End If
    If MsgBox("Confirma exclusão do Lançamento? ", vbYesNo + vbCritical) = vbYes Then
        strSQL = "Delete from tab_caixa where seq_cai ='" & dtServ.Recordset!seq_Cai & "'"
        Banco.Execute strSQL
        RefreshDataPosition dtServ
    End If
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DbgrdToHtml "Relatório de Fluxo de Caixa de " & _
    txtbxDataIni.Text & " até " & txtbxDAtaFIn.Text, dbgrd, dtServ
    
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanCaixa
    fr.Show 0
 End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    dtServ.Recordset.Close
    Unload Me
End Sub

Private Sub dbgrd_GotFocus()
    
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtServ_Reposition()
    'On Error Resume Next
    If dtServ.Recordset.RecordCount = 0 Or dbgrd.Row = -1 Then
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
        Exit Sub
    
    End If
    
    If dtServ.Recordset!num_ven _
     <> "" Then
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
    Else
        cmndbtnEdit.Enabled = True
        cmndbtnExcl.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
   ' Me.Width = 11730
   ' Me.Height = 6705
    P_CentraForm Me
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 0
    cmbxTipo.ListIndex = 0
    txtbxDAtaFIn.Text = Date
    txtbxDataIni.Text = Date
    dtServ.Connect = Banco.Connect
    cmndbtnBuscar_Click
    P_PreencheCaixas cmbxCai
    P_PreencheTurnos cmbxTurno
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
    If F_DataMaior(txtbxDataIni.Text, txtbxDAtaFIn.Text) Then
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    If cmbxBusca.ListIndex = 1 Then
        txtbxExpr.Text = Format(F_SoNumero(txtbxExpr.Text), "000000#")
    ElseIf cmbxBusca.ListIndex = 2 Then
        P_ValorLostFocus txtbxExpr
    Else
        P_AjustaTexto txtbxExpr
    End If
End Sub
