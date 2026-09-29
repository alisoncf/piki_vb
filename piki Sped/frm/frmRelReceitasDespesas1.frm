VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelReceitasDespesas1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Análise de Receitas e Despesas"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   Icon            =   "frmRelReceitasDespesas1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10425
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   10395
      TabIndex        =   19
      Top             =   6390
      Width           =   10425
      Begin VB.Label lblTotalD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   3600
         TabIndex        =   27
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Débito:"
         Height          =   195
         Left            =   3000
         TabIndex        =   26
         Top             =   90
         Width           =   510
      End
      Begin VB.Label lblTotalC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   840
         TabIndex        =   25
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   90
         Width           =   540
      End
      Begin VB.Label lblItens 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9360
         TabIndex        =   23
         Top             =   0
         Width           =   165
      End
      Begin VB.Label lblTOtal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6360
         TabIndex        =   22
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   5880
         TabIndex        =   21
         Top             =   90
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lançamentos:"
         Height          =   195
         Left            =   8280
         TabIndex        =   20
         Top             =   90
         Width           =   1005
      End
   End
   Begin VB.Frame frControle 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame1 
         Caption         =   "Tipo:"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2415
         Begin VB.ComboBox cmbxTipo 
            Height          =   315
            ItemData        =   "frmRelReceitasDespesas1.frx":0442
            Left            =   120
            List            =   "frmRelReceitasDespesas1.frx":044F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Análise:"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2415
         Begin VB.OptionButton optnSint 
            Caption         =   "S&intética"
            Height          =   195
            Left            =   1200
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optnDet 
            Caption         =   "De&talhada"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmRelReceitasDespesas1.frx":046E
         Height          =   5535
         Left            =   2640
         OleObjectBlob   =   "frmRelReceitasDespesas1.frx":0485
         TabIndex        =   18
         Top             =   240
         Width           =   7575
      End
      Begin VB.Data dtVendas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2415
         Begin VB.ComboBox cmbxContas 
            Height          =   315
            ItemData        =   "frmRelReceitasDespesas1.frx":16D2
            Left            =   120
            List            =   "frmRelReceitasDespesas1.frx":16E2
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtbxDAtaFIn 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   12
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Conta:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   465
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1320
            TabIndex        =   11
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D&e:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   2415
         Begin VB.CommandButton cmndbtnGraf 
            Caption         =   "&Gráfico"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelReceitasDespesas1.frx":170D
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1680
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnFechar 
            Cancel          =   -1  'True
            Caption         =   "&Sair"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelReceitasDespesas1.frx":1B4F
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnImprimir 
            Caption         =   "Im&primir"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelReceitasDespesas1.frx":2419
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnBuscar 
            Caption         =   "&Buscar"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelReceitasDespesas1.frx":2723
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmRelReceitasDespesas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxContas_Click()
    cmbxContas.Tag = F_BuscaUmCampo("Cod_Pla", "tab_planocontas", "nome_pla='" & _
        cmbxContas.Text & "'")
    If cmbxContas.Tag <> "" Then
        cmndbtnGraf.Enabled = False
    Else
        If optnDet.Value = False Then
            cmndbtnGraf.Enabled = True
        End If
    End If
    
End Sub

Private Sub cmbxContas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxContas.ListIndex = -1
    End If
End Sub

Sub BuscaSint()
    Dim strSQL As String
    strSQL = "select sum(tc.valor_cla) as total,tp.nome_pla,tp.tipo_pla "
    strSQL = strSQL & " from tab_class as tc, tab_planocontas as tp,"
    strSQL = strSQL & " tab_financeiro as tf "
    strSQL = strSQL & " where tc.cod_pla=tp.cod_pla "
    strSQL = strSQL & " and tc.seq_fin=tf.seq_fin "
    If txtbxDataFin.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin<=" & F_GravaData(txtbxDataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If cmbxTIpo.ListIndex = 2 Then
        strSQL = strSQL & " and tp.tipo_pla='D'"
    ElseIf cmbxTIpo.ListIndex = 1 Then
        strSQL = strSQL & " and tp.tipo_pla='C'"
    End If
    If cmbxContas.Tag <> "" Then
        strSQL = strSQL & " and tc.cod_pla='" & cmbxContas.Tag & "'"
    End If
    strSQL = strSQL & " group by tc.cod_pla "
    strSQL = strSQL & " order by tp.nome_pla"
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
    soma
End Sub
Sub BuscaDet()
    Dim strSQL As String
    strSQL = "select valor_cla as total,tf.cod_con,tf.data_fin,tf.Historico_fin,tc.cod_usu, tp.nome_pla,tp.tipo_pla "
    strSQL = strSQL & " from tab_class as tc, tab_planocontas as tp,"
    strSQL = strSQL & " tab_financeiro as tf "
    strSQL = strSQL & " where tc.cod_pla=tp.cod_pla "
    strSQL = strSQL & " and tc.seq_fin=tf.seq_fin "
    If txtbxDataFin.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin<=" & F_GravaData(txtbxDataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If cmbxTIpo.ListIndex = 2 Then
        strSQL = strSQL & " and tp.tipo_pla='D'"
    ElseIf cmbxTIpo.ListIndex = 1 Then
        strSQL = strSQL & " and tp.tipo_pla='C'"
    End If
    If cmbxContas.Tag <> "" Then
        strSQL = strSQL & " and tc.cod_pla='" & cmbxContas.Tag & "'"
    End If
    
    strSQL = strSQL & " order by tp.nome_pla,tf.data_fin"
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
    soma
End Sub

Sub soma()
    Dim cTotal As Currency
    Dim cTotalC As Currency
    Dim cTotalD As Currency
    With dtVendas.Recordset
            If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If !tipo_pla = "C" Then
                
                cTotalC = cTotalC + !total
            Else
                
                cTotalD = cTotalD + !total
                
            End If
            
            
            
            .MoveNext
        Loop
    End With
    cTotal = cTotalC - cTotalD
    
    lblTOtal.caption = FormatVal(cTotal)
    lblItens.caption = dtVendas.Recordset.RecordCount
    lblTotalC.caption = FormatVal(cTotalC)
    lblTotalD.caption = FormatVal(cTotalD)
End Sub
Private Sub Command1_Click()
    
End Sub

Private Sub cmndbtnBuscar_Click()
    
    If txtbxDataIni.Text = "" Then
        MsgBox "Informe a data Inicial!", vbExclamation
        txtbxDataIni.SetFocus
        Exit Sub
    End If
    If txtbxDataFin.Text = "" Then
        MsgBox "Informe a data Final!", vbExclamation
        txtbxDataFin.SetFocus
        Exit Sub
    End If
    
    P_Status "Aguarde..."
    If optnSint.Value Then
        BuscaSint
    Else
        BuscaDet
    End If
    P_Status ""
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmndbtnGraf_Click()
    Dim b As BlChart
    Dim iCount As Integer
    On Error GoTo trata_erro
    
    If dtVendas.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que ser gerado!", vbExclamation
        Exit Sub
    End If
    With dtVendas.Recordset
        .MoveFirst
        b.count = .RecordCount
        For iCount = 0 To .RecordCount - 1
            
            b.val(iCount).caption = !nome_pla
            b.val(iCount).Value = !total
            .MoveNext
        Next
    End With
    'P_Graph  b, 1, True, "Gráfico de Análise de " & _
        IIf(optnRec.Value, "Receitas", "Despesas") & " De: " & txtbxDataIni.Text & _
        " Até: " & txtbxDAtaFIn.Text, "Fonte: " & constSigla_sis & " " & constVersao_Sis
    
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Sub imprimeSint(obj As Object)
    Dim max As Integer
    max = 40
    P_CabRelTimbr "Análise Sintética de " & _
    IIf(optnDesp.Value, "Despesas", "Receitas"), obj
    P_ImprLinha obj, "-"
    Dim iCount As Integer
    If cmbxContas.Text = "" Then
        P_ImprX obj, "Conta: (Todas) ", 0, True
        
    Else
        P_ImprX obj, "Conta: " & cmbxContas.Text, 0, True
    End If
    P_ImprX obj, "De: " & txtbxDataIni.Text, 50
    P_ImprX obj, "Até: " & txtbxDataFin.Text, 80, False
    P_ImprLinha obj, "-"
    P_ImprX obj, "Conta", 0
    P_ImprX obj, "T", 85
    P_ImprX obj, "Valor", 120, False, True
    
    P_ImprLinha obj, "-"
    With dtVendas.Recordset
        .MoveFirst
        Do While Not .EOF
            
            If iCount = max Then
                iCount = 0
                P_RodapeVis obj, constNome_sis & constVersao_Sis, ConstEmpresaAbrevia
                P_ImprX obj, "Conta", 0
                P_ImprX obj, "T", 85
                P_ImprX obj, "Valor", 120, False, True
                P_ImprLinha obj, "-"
            End If
            
            P_ImprX obj, !nome_pla, 0
            P_ImprX obj, !tipo_pla, 85
            P_ImprX obj, FormatVal(!total), 120, False, True
            
            
            iCount = iCount + 1
            .MoveNext
        Loop
        P_ImprLinha obj
        obj.FontBold = True
        P_ImprX obj, "Total...: " & lblTOtal.caption, 120, False, True
        obj.FontBold = False
    End With
    
    P_RodapeVis obj, constNome_sis & constVersao_Sis, ConstEmpresaAbrevia
    
End Sub
Sub imprimeDet(obj As Object)
    Dim max As Integer
    max = 40
    P_CabRelTimbr "Análise Detalhada de " & _
    IIf(optnDesp.Value, "Despesas", "Receitas"), obj
    P_ImprLinha obj, "-"
    Dim iCount As Integer
    If cmbxContas.Text = "" Then
        P_ImprX obj, "Conta: (Todas) ", 0, True
        
    Else
        P_ImprX obj, "Conta: " & cmbxContas.Text, 0, True
    End If
    P_ImprX obj, "De: " & txtbxDataIni.Text, 50
    P_ImprX obj, "Até: " & txtbxDataFin.Text, 80, False
    P_ImprLinha obj, "-"
    P_ImprX obj, "Conta", 0
    P_ImprX obj, "T", 35
    P_ImprX obj, "Valor", 50, True, True
    P_ImprX obj, "Data", 55
    P_ImprX obj, "Histórico", 75
    P_ImprX obj, "C/C.", 165
    P_ImprX obj, "Usu.", 190, False
    P_ImprLinha obj, "-"
    With dtVendas.Recordset
        .MoveFirst
        Do While Not .EOF
            
            If iCount = max Then
                iCount = 0
                P_RodapeVis obj, constNome_sis & constVersao_Sis, ConstEmpresaAbrevia
                P_ImprX obj, "Conta", 0
                P_ImprX obj, "T", 35
                P_ImprX obj, "Valor", 50, True, True
                P_ImprX obj, "Data", 55
                P_ImprX obj, "Histórico", 75
                P_ImprX obj, "C/C.", 165
                P_ImprX obj, "Usu.", 190, False
                P_ImprLinha obj, "-"
            End If
            
            P_ImprX obj, !nome_pla, 0
            P_ImprX obj, !tipo_pla, 35
            P_ImprX obj, FormatVal(!total), 50, True, True
            P_ImprX obj, !data_fin, 55
            P_ImprX obj, Left(!historico_fin, 40), 75
            P_ImprX obj, F_BuscaUmCampo("nome_con", "tab_Contas", _
                "cod_con='" & !cod_Con & "'"), 165
            P_ImprX obj, !cod_usu, 190, False
            
            iCount = iCount + 1
            .MoveNext
        Loop
        P_ImprLinha obj
        obj.FontBold = True
        P_ImprX obj, "Total...: " & lblTOtal.caption, 50, False, True
        obj.FontBold = False
    End With
    
    P_RodapeVis obj, constNome_sis & constVersao_Sis, ConstEmpresaAbrevia
    
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_erro
    With dtVendas.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser impresso!", vbExclamation
            Exit Sub
        End If
        If optnDet.Value Then
            imprimeDet frmVisualiza.Pic
            frmVisualiza.Show 1
            If frmVisualiza.Tag <> "" Then
                imprimeDet Printer
                Printer.EndDoc
            End If
            Unload frmVisualiza
        End If
        If optnSint.Value Then
            imprimeSint frmVisualiza.Pic
            frmVisualiza.Show 1
            If frmVisualiza.Tag <> "" Then
                imprimeSint Printer
                Printer.EndDoc
            End If
            Unload frmVisualiza
        End If
    End With
Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    optnRec_Click
    optnDet_Click
    dtVendas.Connect = Banco.Connect
End Sub

Private Sub optnDesp_Click()
    P_PreenchePlanoContas cmbxContas, False
End Sub

Private Sub optnDesp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnDet_Click()
    Dim iCount As Integer
    For iCount = 3 To 6
        dbgrd.Columns(iCount).Visible = True
    Next
    cmndbtnGraf.Enabled = False
End Sub

Private Sub optnDet_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnRec_Click()
    P_PreenchePlanoContas cmbxContas, True
End Sub

Private Sub optnRec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnSint_Click()
    Dim iCount As Integer
    For iCount = 3 To 6
        dbgrd.Columns(iCount).Visible = False
    Next
    cmndbtnGraf.Enabled = True
End Sub

Private Sub optnSint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDataFin
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
