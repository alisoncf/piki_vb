VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelSaiCaixa 
   Caption         =   "Relatório de Saídas"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   Icon            =   "frmRelSaiCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   10140
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5655
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9975
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   7665
      TabIndex        =   17
      Top             =   5760
      Width           =   7695
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7080
         TabIndex        =   19
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   7920
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.Frame Frame2 
         Caption         =   "Busca:"
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1935
         Begin VB.OptionButton optnSeq 
            Caption         =   "Se&qüencial"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optnAgrupada 
            Caption         =   "&Agrupada"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmbxTurno 
         Height          =   315
         ItemData        =   "frmRelSaiCaixa.frx":08CA
         Left            =   120
         List            =   "frmRelSaiCaixa.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox cmbxCaixa 
         Height          =   315
         ItemData        =   "frmRelSaiCaixa.frx":08FA
         Left            =   120
         List            =   "frmRelSaiCaixa.frx":0907
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1665
         Width           =   1935
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cmbxOrderBY 
         Height          =   315
         ItemData        =   "frmRelSaiCaixa.frx":092A
         Left            =   120
         List            =   "frmRelSaiCaixa.frx":0937
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelSaiCaixa.frx":095A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelSaiCaixa.frx":0C64
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelSaiCaixa.frx":0F6E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Turno:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRelSaiCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ajustagrid()
    With msflxgrd
        If optnAgrupada.Value Then
            .Cols = 5
            .FixedCols = 0
            .rows = 1
            .Col = 0: .Text = "Código"
            .Col = 1: .Text = "Descrição"
            .Col = 2: .Text = "Un."
            .Col = 3: .Text = "Qtd."
            .Col = 4: .Text = "Total"
            .ColWidth(0) = 1100: .ColAlignment(0) = 0
            .ColWidth(1) = 4500: .ColAlignment(1) = 0
            .ColWidth(2) = 600: .ColAlignment(2) = 0
            .ColWidth(3) = 800:
            .ColWidth(4) = 800:
        Else
            .Cols = 8
            .FixedCols = 0
            .rows = 1
            .Col = 0: .Text = "Código"
            .Col = 1: .Text = "Descrição"
            .Col = 2: .Text = "Un."
            .Col = 3: .Text = "Qtd."
            .Col = 4: .Text = "Total"
            .Col = 5: .Text = "Unit."
            .Col = 6: .Text = "Venda"
            .Col = 7: .Text = "Hora"
            .ColWidth(0) = 1100: .ColAlignment(0) = 0
            .ColWidth(1) = 4500: .ColAlignment(1) = 0
            .ColWidth(2) = 600: .ColAlignment(2) = 0
            .ColWidth(3) = 800:
            .ColWidth(4) = 800:
            .ColWidth(5) = 800:
            .ColWidth(6) = 800: .ColAlignment(2) = 0
            .ColWidth(7) = 2000: .ColAlignment(2) = 0
        End If
    End With
End Sub

Private Sub cmbxCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxCaixa.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxOrderBY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxTurno.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub cmndbtnBuscar_Click()
    If optnAgrupada Then
        BuscaAgrupada
    Else
        BuscaAgrupada False
        
    End If
End Sub
Sub BuscaAgrupada(Optional bAgrup As Boolean = True)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    msflxgrd.rows = 1
    If pc.BackColor = vbYellow Then
        strSQL = "select tp.cod_pro,tp.nome_pro,tp.unidade_pro,"
        If bAgrup Then
            strSQL = strSQL & "sum(ti.qtd_Ite) as qtd,"
            strSQL = strSQL & "sum(final_ite) As Total,ti.unit_ite as unit "
            
        Else
            strSQL = strSQL & "ti.qtd_Ite as qtd,tp.pv_pro as unit,tv.num_ven as num,tv.data_ven as hora "
            
        End If
        strSQL = strSQL & " from tab_produtos as tp,"
        strSQL = strSQL & "  tab_itvendas as ti, tab_venfec as tc,tab_Vendas as tv  "
        strSQL = strSQL & " where tp.cod_pro=ti.cod_pro"
        
        strSQL = strSQL & " and tv.num_ven=ti.num_ven "
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxDataFin.Text)
        strSQL = strSQL & " and tc.num_ven=tv.num_ven"
        strSQL = strSQL & " and tv.situacao_ven=4 "
        If cmbxCaixa.Text <> "" Then
            strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
        End If
        If cmbxTurno.Text <> "" Then
            strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
        End If
        If bAgrup Then
            strSQL = strSQL & " group by cod_pro "
        End If
        If cmbxOrderBY.ListIndex = 1 Then
            strSQL = strSQL & " order by ti.cod_pro"
        ElseIf cmbxOrderBY.ListIndex = 2 Then
            strSQL = strSQL & " order by nome_pro"
        Else
            strSQL = strSQL & " order by total desc"
        End If
    Else
        
        strSQL = "select tp.cod_pro,tp.nome_pro,tp.unidade_pro,"
        If bAgrup Then
            strSQL = strSQL & " sum(ti.qtd_Its) as qtd,sum(qtd_its*unit_its) as total"
        Else
            strSQL = strSQL & " (ti.qtd_Its) as qtd,ti.unit_its as unit, "
            strSQL = strSQL & "tv.num_ven as num,tv.hora_sai as hora"
        End If
        strSQL = strSQL & " from tab_produtos as tp,"
        strSQL = strSQL & "  tab_itsai as ti, tab_saidas as tv  "
        strSQL = strSQL & " where tp.cod_pro=ti.cod_pro"
        
        strSQL = strSQL & " and tv.serie_sai <> 'CANC' "
        strSQL = strSQL & " and tv.emissao_sai>=" & F_GravaData(txtbxDataIni.Text)
        strSQL = strSQL & " and tv.emissao_Sai<=" & F_GravaData(txtbxDataFin.Text)
        strSQL = strSQL & " and tv.nota_sai=ti.nota_sai "
        strSQL = strSQL & " and tv.Serie_sai=ti.serie_sai "
        strSQL = strSQL & " and tv.cod_cai=ti.cod_cai "
        
        If cmbxCaixa.Text <> "" Then
            strSQL = strSQL & " and tv.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
        End If
        If cmbxTurno.Text <> "" Then
            
            strSQL = strSQL & " and tv.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
            
        End If
        strSQL = strSQL & " group by cod_pro "
        If cmbxOrderBY.ListIndex = 1 Then
            strSQL = strSQL & " order by ti.cod_pro"
        ElseIf cmbxOrderBY.ListIndex = 2 Then
            strSQL = strSQL & " order by nome_pro"
        Else
            strSQL = strSQL & " order by total desc"
        End If
    End If
        abreObjetoAdo rPesq, strSQL
        With rPesq
            Do While Not .EOF
                strSQL = FormatCodigo(!cod_pro) & Chr(9)
                strSQL = strSQL & !nome_pro & Chr(9) & !unidade_pro & Chr(9)
                strSQL = strSQL & FormatVal(!qtd) & Chr(9)
                strSQL = strSQL & FormatVal(!total) & Chr(9)
                cTotal = cTotal + !total
                If Not bAgrup Then
                    strSQL = strSQL & Chr(9) & FormatVal(!unit) & Chr(9)
                    strSQL = strSQL & !Num & Chr(9)
                    strSQL = strSQL & !hora & Chr(9)
                End If
                msflxgrd.AddItem strSQL
                .MoveNext
            Loop
            .Close
        End With
        
    
    
    
    ZebraGrid msflxgrd, 2, &HD3E7EB
    lblTotal.caption = "Total...........: " & FormatVal(cTotal)
End Sub

Private Sub cmndbtnImprimir_Click()
    msflxgrd.Visible = False
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
    msflxgrd.Visible = True
End Sub
Sub imprime(obj As Object)
    Dim iCount As Integer
    Dim iCol As Integer
    Dim sWith As Single
    P_CabRelTimbr "Relatório de Saídas de " & txtbxDataIni.Text & " até " & txtbxDataFin.Text, obj
    obj.fontsize = 12
    P_ImprX obj, "Caixa: " & cmbxCaixa, 0
    P_ImprX obj, "Turno: " & cmbxTurno, 60, False
    obj.fontsize = 10
    P_ImprLinha obj
    
        With msflxgrd
            For iCol = 0 To .Cols - 1
                .row = 0
                .Col = iCol
                
                P_ImprX obj, .Text, ScaleX(sWith, 1, 6)
                sWith = sWith + .CellWidth
            Next
            P_ImprX obj, "", 0, False
            sWith = 0
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_FimRel obj, "Continua..."
                For iCol = 0 To .Cols - 1
                    .row = 0
                    .Col = iCol
                
                    P_ImprX obj, .Text, ScaleX(sWith, 1, 6)
                    sWith = sWith + .CellWidth
                    Next
                    P_ImprX obj, "", 0, False
            End If
            sWith = 0
            For iCount = 1 To .rows - 1
                .row = iCount
                For iCol = 0 To .Cols - 1
                    .Col = iCol
                
                    P_ImprX obj, .Text, ScaleX(sWith, 1, 6)
                sWith = sWith + .CellWidth
            Next
            sWith = 0
            P_ImprX obj, "", 0, False
            Next
            P_ImprLinha obj
            P_ImprX obj, lblTotal.caption, 0, False
            P_FimRel obj, strNome_Emp
        End With
    
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        pc.BackColor = vbYellow
    End If
    If KeyCode = vbKeyF8 Then
        pc.BackColor = vbGreen
    End If
End Sub

Private Sub Form_Load()
    Me.width = 10220
    Me.height = 6660
    Me.Top = 10
    Me.Left = 10
    P_PreencheCaixas cmbxCaixa
    P_PreencheTurnos cmbxTurno
    txtbxDataFin.Text = Date
    txtbxDataIni.Text = Date
    pc.BackColor = vbGreen
    ajustagrid
End Sub

Private Sub optnAgrupada_Click()
    ajustagrid
    BuscaAgrupada
End Sub

Private Sub optnAgrupada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar.SetFocus
    End If
End Sub

Private Sub optnSeq_Click()
    ajustagrid
    BuscaAgrupada
End Sub

Private Sub optnSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar.SetFocus
    End If
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
