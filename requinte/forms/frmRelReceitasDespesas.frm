VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelReceitasDespesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Análise de Receitas e Despesas"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   Icon            =   "frmRelReceitasDespesas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10425
   Begin VB.Frame frControle 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame1 
         Caption         =   "Tipo:"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2415
         Begin VB.OptionButton optnDesp 
            Caption         =   "&Despesas"
            Height          =   195
            Left            =   1200
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optnRec 
            Caption         =   "&Receitas"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Análise:"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2415
         Begin VB.OptionButton optnSint 
            Caption         =   "S&intética"
            Height          =   195
            Left            =   1200
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optnDet 
            Caption         =   "De&talhada"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmRelReceitasDespesas.frx":0442
         Height          =   6135
         Left            =   2640
         OleObjectBlob   =   "frmRelReceitasDespesas.frx":0459
         TabIndex        =   24
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
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
         Begin VB.ComboBox cmbxContas 
            Height          =   315
            ItemData        =   "frmRelReceitasDespesas.frx":16A6
            Left            =   120
            List            =   "frmRelReceitasDespesas.frx":16B6
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtbxDAtaFIn 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   13
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Conta:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   465
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1320
            TabIndex        =   12
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D&e:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3975
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   2415
         Begin VB.CommandButton cmndbtnGraf 
            Caption         =   "&Gráfico"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelReceitasDespesas.frx":16E1
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1680
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnFechar 
            Cancel          =   -1  'True
            Caption         =   "&Sair"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelReceitasDespesas.frx":1B23
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnImprimir 
            Caption         =   "Im&primir"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelReceitasDespesas.frx":23ED
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnBuscar 
            Caption         =   "&Buscar"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelReceitasDespesas.frx":26F7
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   2175
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   705
            ScaleWidth      =   2145
            TabIndex        =   19
            Top             =   3120
            Width           =   2175
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
               Left            =   720
               TabIndex        =   23
               Top             =   277
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
               Left            =   720
               TabIndex        =   22
               Top             =   0
               Width           =   570
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gerado:"
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   83
               Width           =   570
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Itens:"
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   360
               Width           =   390
            End
         End
      End
   End
End
Attribute VB_Name = "frmRelReceitasDespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxContas_Click()
    cmbxContas.tag = F_BuscaUmCampo("Cod_Pla", "tab_planocontas", "nome_pla='" & _
        cmbxContas.Text & "'")
    If cmbxContas.tag <> "" Then
        cmndbtnGraf.Enabled = False
    Else
        If optnDet.Value = False Then
            cmndbtnGraf.Enabled = True
        End If
    End If
    
End Sub

Private Sub cmbxContas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If optnDesp.Value = True Then
        strSQL = strSQL & " and tp.tipo_pla='D'"
    Else
        strSQL = strSQL & " and tp.tipo_pla='C'"
    End If
    If cmbxContas.tag <> "" Then
        strSQL = strSQL & " and tc.cod_pla='" & cmbxContas.tag & "'"
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
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If optnDesp.Value = True Then
        strSQL = strSQL & " and tp.tipo_pla='D'"
    Else
        strSQL = strSQL & " and tp.tipo_pla='C'"
    End If
    If cmbxContas.tag <> "" Then
        strSQL = strSQL & " and tc.cod_pla='" & cmbxContas.tag & "'"
    End If
    
    strSQL = strSQL & " order by tp.nome_pla,tf.data_fin"
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
    soma
End Sub

Sub soma()
    Dim cTotal As Currency
    With dtVendas.Recordset
            If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            
            cTotal = cTotal + !total
            .MoveNext
        Loop
    End With
    lblTotal.caption = FormatVal(cTotal)
    lblItens.caption = dtVendas.Recordset.RecordCount
End Sub
Private Sub Command1_Click()
    
End Sub

Private Sub cmndbtnBuscar_Click()
    
    If txtbxDataIni.Text = "" Then
        MsgBox "Informe a data Inicial!", vbExclamation
        txtbxDataIni.SetFocus
        Exit Sub
    End If
    If txtbxdataFin.Text = "" Then
        MsgBox "Informe a data Final!", vbExclamation
        txtbxdataFin.SetFocus
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
    On Error GoTo trata_Erro
    
    If dtVendas.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que ser gerado!", vbExclamation
        Exit Sub
    End If
    With dtVendas.Recordset
        .MoveFirst
        b.Count = .RecordCount
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
trata_Erro:
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
    P_ImprX obj, "Até: " & txtbxdataFin.Text, 80, False
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
                P_RodapeVis obj, constNome_sis & constVersao_Sis, strNome_Emp
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
        P_ImprX obj, "Total...: " & lblTotal.caption, 120, False, True
        obj.FontBold = False
    End With
    
    P_RodapeVis obj, constNome_sis & constVersao_Sis, strNome_Emp
    
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
    P_ImprX obj, "Até: " & txtbxdataFin.Text, 80, False
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
                P_RodapeVis obj, constNome_sis & constVersao_Sis, strNome_Emp
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
                "cod_con='" & !cod_con & "'"), 165
            P_ImprX obj, !cod_usu, 190, False
            
            iCount = iCount + 1
            .MoveNext
        Loop
        P_ImprLinha obj
        obj.FontBold = True
        P_ImprX obj, "Total...: " & lblTotal.caption, 50, False, True
        obj.FontBold = False
    End With
    
    P_RodapeVis obj, constNome_sis & constVersao_Sis, strNome_Emp
    
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_Erro
    With dtVendas.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser impresso!", vbExclamation
            Exit Sub
        End If
        If optnDet.Value Then
            imprimeDet frmVisualiza.Pic
            frmVisualiza.Show 1
            If frmVisualiza.tag <> "" Then
                imprimeDet Printer
                Printer.EndDoc
            End If
            Unload frmVisualiza
        End If
        If optnSint.Value Then
            imprimeSint frmVisualiza.Pic
            frmVisualiza.Show 1
            If frmVisualiza.tag <> "" Then
                imprimeSint Printer
                Printer.EndDoc
            End If
            Unload frmVisualiza
        End If
    End With
Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    optnRec_Click
    optnDet_Click
  
End Sub

Private Sub optnDesp_Click()
    P_PreenchePlanoContas cmbxContas, False
End Sub

Private Sub optnDesp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnRec_Click()
    P_PreenchePlanoContas cmbxContas, True
End Sub

Private Sub optnRec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
