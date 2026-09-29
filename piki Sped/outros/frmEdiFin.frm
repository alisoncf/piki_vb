VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiFin 
   Caption         =   "Edição de Movimento Financeiro"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   Icon            =   "frmEdiFin.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   11805
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8325
      Left            =   0
      ScaleHeight     =   8295
      ScaleWidth      =   2145
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiFin.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   7560
         Width           =   1950
      End
      Begin VB.CommandButton cmndbtnAjustaSaldo 
         Caption         =   "C&alcular Saldo"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiFin.frx":02B5
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6120
         Width           =   1950
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiFin.frx":03FF
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5400
         Width           =   1950
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiFin.frx":0549
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4680
         Width           =   1950
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiFin.frx":0693
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3960
         Width           =   1950
      End
      Begin VB.OptionButton optnDesc 
         Caption         =   "P/Baixo"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton optnAsc 
         Caption         =   "P/Cima"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtbxDataFin 
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   16
         Top             =   3600
         Width           =   1695
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmEdiFin.frx":0A1D
         Left            =   120
         List            =   "frmEdiFin.frx":0A2A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiFin.frx":0A51
         Left            =   120
         List            =   "frmEdiFin.frx":0A61
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox cmbxTIpo 
         Height          =   315
         ItemData        =   "frmEdiFin.frx":0A89
         Left            =   120
         List            =   "frmEdiFin.frx":0A96
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   285
         Left            =   1800
         Picture         =   "frmEdiFin.frx":0AB4
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3600
         Width           =   255
      End
      Begin VB.ComboBox cmbxConta 
         Height          =   315
         ItemData        =   "frmEdiFin.frx":0E3E
         Left            =   120
         List            =   "frmEdiFin.frx":0E4B
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lbldataFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   2760
         Width           =   765
      End
      Begin VB.Label lblDataIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Conta:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   465
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
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11775
      TabIndex        =   24
      Top             =   8325
      Width           =   11805
      Begin VB.Label lblTotal_Din 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8880
         TabIndex        =   30
         Tag             =   "!"
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   255
         Left            =   8040
         TabIndex        =   29
         Top             =   45
         Width           =   615
      End
      Begin VB.Label lblDeb_Din 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5520
         TabIndex        =   28
         Tag             =   "!"
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Débito:"
         Height          =   255
         Left            =   4800
         TabIndex        =   27
         Top             =   45
         Width           =   615
      End
      Begin VB.Label lblCred_din 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2880
         TabIndex        =   26
         Tag             =   "!"
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito:"
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   45
         Width           =   615
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiFin.frx":0E72
      Height          =   4935
      Left            =   2280
      OleObjectBlob   =   "frmEdiFin.frx":0E87
      TabIndex        =   23
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmEdiFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxConta.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnAjustaSaldo_Click()
    Dim cSaldo As Currency
    cSaldo = F_CalculaSaldo("01/01/1900", Left(cmbxConta.Text, 3))
    
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
    strSQL = "select tf.*,tp.nome_pla,tb.nome_con from tab_financeiro as tf, tab_contas as tb,"
    strSQL = strSQL & " tab_planocontas as tp where "
    strSQL = strSQL & "tf.cod_con=tb.cod_con  and tp.cod_pla=tf.cod_pla and "
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " historico_fin like '" & txtbxExpr.Text & "%'"
    
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " valor_fin = " & F_GravaDec(txtbxExpr.Text)
    End If
    If cmbxConta.ListIndex > -1 Then
        strSQL = strSQL & " and tb.cod_con='" & Left(cmbxConta.Text, 3) & "'"
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_fin >= " & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxDAtaFIn.Text <> "" Then
     strSQL = strSQL & " and data_fin <= " & F_GravaData(txtbxDAtaFIn.Text)
    End If
    If cmbxTipo.ListIndex <> 0 Then
    
        strSQL = strSQL & " and tipo_fin='" & Left(cmbxTipo.Text, 1) & "' "
    End If
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " order by data_fin,seq_fin"
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & " order by seq_fin"
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " order by historico_fin, seq_fin"
    ElseIf cmbxBusca.ListIndex = 3 Then
        strSQL = strSQL & " order by valor_Fin,seq_Fin"
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
            
            If !Esp_fin = "1" Or !Esp_fin = "DIN" Then
                If !tipo_fin = "C" Then
                    cDinCred = cDinCred + !valor_fin
                Else
                    cDinDeb = cDinDeb + !valor_fin
                End If
            Else
                If !tipo_fin = "C" Then
                    cCheCred = cCheCred + !valor_fin
                Else
                    cCheDeb = cCheDeb + !valor_fin
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    'lblCred_Che.caption = FormatVal(cCheCred)
    lblCred_din.caption = FormatVal(cDinCred)
    lblDeb_Din.caption = FormatVal(cDinDeb)
'    lblDeb_Che.caption = FormatVal(cCheDeb)
'    lblTotal_Che.caption = FormatVal(cCheCred - cCheDeb)
    lblTotal_Din.caption = FormatVal(cDinCred - cDinDeb)
    'lblSaldo.caption = lblTotal_Din.caption
End Sub
Private Sub cmndbtnEdit_Click()
    On Error Resume Next
    Dim fr As New frmLanFin
    With dtServ.Recordset
        fr.Tag = !seq_fin
        fr.txtbxData_Cai.Text = !data_fin
        fr.txtbxHist.Text = !historico_fin
        fr.txtbxValor_Cai.Text = FormatVal(!valor_fin)
        
        If !tipo_fin = "C" Then
            fr.optnCerd.Value = True
            P_PreenchePlanoContas fr.cmbxCentro, True
        Else
            fr.optnDeb.Value = True
            P_PreenchePlanoContas fr.cmbxCentro, False
        End If
        fr.cmbxCentro.Text = !nome_pla
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
        strSQL = "Delete from tab_financeiro where seq_fin ='" & dtServ.Recordset!seq_fin & "'"
        Banco.Execute strSQL
        RefreshDataPosition dtServ
    End If
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanFin
    fr.Show 0
 End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    Unload Me
    dtServ.Recordset.Close
End Sub

Private Sub dbgrd_DblClick()
    If cmndbtnEdit.Enabled Then cmndbtnEdit_Click
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
    Else
        If dtServ.Recordset!seq_lan = "" Or IsNull(dtServ.Recordset!seq_lan) Then
            cmndbtnEdit.Enabled = True
            cmndbtnExcl.Enabled = True
        Else
            cmndbtnEdit.Enabled = False
            cmndbtnExcl.Enabled = False
        End If
    End If
    
    
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
  resizeForm Me
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 0
    cmbxTipo.ListIndex = 0
    txtbxDAtaFIn.Text = Date
    txtbxDataIni.Text = Date
    P_PreencheContas cmbxConta
    dtServ.Connect = Banco.Connect
    cmndbtnBuscar_Click
    
End Sub

Private Sub Form_Resize()
    dbgrd.Width = Me.Width - dbgrd.Left - 300
    dbgrd.Height = pcBot.Top - 300
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
