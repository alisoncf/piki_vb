VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiFin 
   Caption         =   "Edição de Movimento Financeiro"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14295
   Icon            =   "frmEdiFin.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   14295
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   14265
      TabIndex        =   7
      Top             =   0
      Width           =   14295
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10560
         MaxLength       =   60
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbxConta 
         Height          =   315
         ItemData        =   "frmEdiFin.frx":014A
         Left            =   5040
         List            =   "frmEdiFin.frx":0157
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Pesquisar"
         Height          =   525
         Left            =   12240
         Picture         =   "frmEdiFin.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         MaxLength       =   60
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmEdiFin.frx":0529
         Left            =   1680
         List            =   "frmEdiFin.frx":0536
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiFin.frx":055D
         Left            =   3360
         List            =   "frmEdiFin.frx":056D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxTIpo 
         Height          =   315
         ItemData        =   "frmEdiFin.frx":0595
         Left            =   120
         List            =   "frmEdiFin.frx":05A2
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   10560
         TabIndex        =   25
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lbldataFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   7800
         TabIndex        =   20
         Top             =   0
         Width           =   765
      End
      Begin VB.Label lblDataIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   6720
         TabIndex        =   19
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   1680
         TabIndex        =   18
         Top             =   0
         Width           =   825
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Histórico"
         Height          =   195
         Left            =   8880
         TabIndex        =   17
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   3360
         TabIndex        =   16
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Conta:"
         Height          =   195
         Left            =   5040
         TabIndex        =   14
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.PictureBox pcBotoes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   14295
      TabIndex        =   0
      Top             =   6540
      Width           =   14295
      Begin VB.CommandButton cmndbtnAjustaSaldo 
         Caption         =   "C&alcular Saldo"
         Height          =   735
         Left            =   5040
         Picture         =   "frmEdiFin.frx":05C0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1590
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   8520
         Picture         =   "frmEdiFin.frx":070A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1590
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   0
         Picture         =   "frmEdiFin.frx":07FF
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1590
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   735
         Left            =   3360
         Picture         =   "frmEdiFin.frx":0B89
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1590
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1680
         Picture         =   "frmEdiFin.frx":0F13
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1590
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   6720
         Picture         =   "frmEdiFin.frx":105D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1590
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiFin.frx":11A7
      Height          =   5535
      Left            =   0
      TabIndex        =   23
      Top             =   840
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
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
      Caption         =   "Lançamentos Financeiros"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "nome_con"
         Caption         =   "Conta"
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
         DataField       =   "cod_con"
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
      BeginProperty Column02 
         DataField       =   "data_fin"
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
      BeginProperty Column03 
         DataField       =   "historico_fin"
         Caption         =   "Histórico"
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
         DataField       =   "valor_fin"
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
         DataField       =   "esp_fin"
         Caption         =   "Espécie"
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
         DataField       =   "nome_pla"
         Caption         =   "Classificação"
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
         DataField       =   "cod_pla"
         Caption         =   "Código PC"
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
         DataField       =   "cod_usu"
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
      BeginProperty Column09 
         DataField       =   "seq_fin"
         Caption         =   "Id.Fin."
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
            ColumnWidth     =   1904,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2849,953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2399,811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmbxConta_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxConta.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnAjustaSaldo_Click()
    Dim cSaldo As Currency
    cSaldo = F_CalculaSaldo("01/01/1900", Left(cmbxConta.Text, 3))
    
End Sub

Private Sub cmndbtnBuscar_Click_1()
    Dim strSQL As String
    If txtbxDataFin.Visible = True Then
        If F_DataMaior(txtbxDataIni.Text, txtbxDataFin.Text) Then
            MsgBox "A data final deve ser maior que a inicial!", vbExclamation
            txtbxDataFin.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "select tf.*,tb.nome_con from tab_financeiro as tf, tab_contas as tb where "
    strSQL = strSQL & "tf.cod_con=tb.cod_con  and "
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
    If txtbxDataFin.Text <> "" Then
        strSQL = strSQL & " and data_fin <= " & F_GravaData(txtbxDataFin.Text)
    End If
    If cmbxTIpo.ListIndex <> 0 Then
        strSQL = strSQL & " and tipo_fin='" & Left(cmbxTIpo.Text, 1) & "' "
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
    
    FechaObjeto ado
    ado.Open strSQL, Conn_Client
    
    
    Set dtgrd.DataSource = ado
    
  
  '
    P_Resumo
End Sub
Sub P_Resumo()
Dim cCheDeb As Currency
    Dim cCheCred As Currency
    Dim cDinDeb As Currency
    Dim cDinCred As Currency
    
    With adc.Recordset
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

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = financeiro.F_SelectFinanceiro(txtbxDataIni.Text, txtbxDataFin.Text, Left(cmbxConta.Text, 3), _
        "", txtbxExpr.Text, txtValor.Text)
    FechaObjeto ado
    ado.Open strSQL, Conn_Client
    
     Set dtgrd.DataSource = ado
    
    Exit Sub
trata_erro:
    P_MostraErro
End Sub


Private Sub cmndbtnEdit_Click()
    Dim fr As New frmLanFin
    With ado
        fr.tag = !seq_fin
        fr.txtbxData_Cai.Text = !data_fin
        fr.txtbxHist.Text = !historico_fin
        fr.txtbxValor_Cai.Text = FormatVal(!valor_fin)
        If !tipo_fin = "C" Then
            fr.optnCerd.Value = True
        Else
            fr.optnDeb.Value = True
        End If
        fr.optnCerd.Enabled = False
        fr.optnDeb.Enabled = False
        P_LocalizaOperacao !cod_con, fr.cmbxConta
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    If adc.Recordset.RecordCount = 0 Then
        cmndbtnExcl.Enabled = False
        Exit Sub
    End If
    If MsgBox("Confirma exclusão do Lançamento? ", vbYesNo + vbCritical) = vbYes Then
        strSQL = "Delete from tab_financeiro where seq_fin ='" & adc.Recordset!seq_fin & "'"
        Conn.Execute strSQL
        
    End If
End Sub

Private Sub cmndbtnImpr_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    P_CabRelTimbr "Movimento de Contas Financeiras", obj
    P_ImprX obj, "Período de " & txtbxDataIni.Text, 10
    P_ImprX obj, "até " & txtbxDataFin.Text, 40, True
    P_ImprX obj, "Conta:" & cmbxConta.Text, 100, False
    P_line1 obj
    With adc.Recordset
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        P_ImprX obj, "...", 2
        P_ImprX obj, "data", 15
        P_ImprX obj, "Valor", 35
        P_ImprX obj, "Tipo", 50
        P_ImprX obj, "Histórico", 60
        P_ImprX obj, "Conta", 140, False
        
        
        
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                If TypeName(obj) = "Printer" Then
                    P_RodapeVis obj
                    
                Else
                    P_AddPicVis obj
                End If
            End If
            P_ImprX obj, !seq_fin, 2
            P_ImprX obj, !data_fin, 15
            P_ImprX obj, FormatVal(!valor_fin), 35
            P_ImprX obj, !tipo_fin, 50
            P_ImprX obj, Left(!historico_fin, 35), 60
            P_ImprX obj, !nome_con, 140, False
            
            .MoveNext
        Loop
        
        P_line1 obj
        P_ImprX obj, "Crédito...:" & lblCred_din.caption, 2
        P_ImprX obj, "Débito...:" & lblDeb_Din.caption, 40
        P_ImprX obj, "Total...:" & lblTotal_Din.caption, 80, False
    End With
    
    
End Sub
Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanFin
    fr.Show 0
 End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnTrasf_Click()

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

Private Sub ado_Reposition()
    On Error Resume Next
    If adc.Recordset.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnExcl.Enabled = False
        Exit Sub
    Else
        If adc.Recordset!vinculado_fin = "N" Then
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
    cmbxTIpo.ListIndex = 0
    txtbxDataFin.Text = Date
    txtbxDataIni.Text = Date
    P_PreencheContas cmbxConta
    
    
    Set ado = New ADODB.Recordset
    
    
    
End Sub

Private Sub Form_Resize()
    On Error GoTo trata_erro
    dtgrd.width = Me.width - dtgrd.Left - 100
    dtgrd.height = pcBotoes.Top - dtgrd.Top - 100
    
    
    cmndbtnSair.Left = pcBotoes.width - cmndbtnSair.width - 100
    Exit Sub
trata_erro:
    
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
    If F_DataMaior(txtbxDataIni.Text, txtbxDataFin.Text) Then
        MsgBox "A data final deve ser maior que a inicial!", vbExclamation
        txtbxDataFin.SetFocus
        Exit Sub
    End If
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

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
