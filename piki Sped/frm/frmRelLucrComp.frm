VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelLucrComp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Lucratividade por Composição"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmndbtnGraph 
      Caption         =   "&Gráfico"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordenar Por:"
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
      Begin VB.OptionButton optnLiq 
         Caption         =   "Líquido"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optnNome 
         Caption         =   "Nome"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optnGer 
         Caption         =   "Gerado"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por"
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Buscar"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5415
      Left            =   2400
      TabIndex        =   15
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9551
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optnMarca 
         Caption         =   "Marca"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optnCat 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmRelLucrComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Dim rPesq As New ADODB.Recordset
    Dim sTab As String
    Dim sCampo As String
    Dim sNome As String
    Dim cFinal As Currency
    Dim cLiq As Currency
    Msflxgrd.rows = 1
    P_Status "Aguarde..."
    If optnMarca.Value Then
        sTab = "tab_marcas"
        sCampo = "cod_mar"
        sNome = "nome_mar"
        Msflxgrd.col = 0: Msflxgrd.Text = "Marca"
    Else
        Msflxgrd.col = 0: Msflxgrd.Text = "Categoria"
        sNome = "nome_cat"
        sTab = "tab_categorias"
        sCampo = "cod_Cat"
    End If
    strSQL = "    select tm." & sNome & " as nome, "
    strSQL = strSQL & " sum(final_ite) as final,sum(final_ite-cr_pro) as gl"
    strSQL = strSQL & " from " & sTab & " as tm," & TAB_ITVENDAS & " as ti,"
    strSQL = strSQL & " " & TAB_VENDAS & " As tv, " & TAB_PRODUTOS & " As tp"
    strSQL = strSQL & " Where tv.num_ven = ti.num_ven"
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro"
    strSQL = strSQL & " and tp." & sCampo & "=tm." & sCampo
    
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " group by tm." & sCampo
    If optnNome.Value Then
        strSQL = strSQL & " order by " & sNome
    ElseIf optnGer.Value Then
        strSQL = strSQL & " order by final desc "
    Else
        strSQL = strSQL & " order by gl desc "
    End If
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            Msflxgrd.AddItem !nome & Chr(9) & FormatVal(!Final) & Chr(9) & FormatVal(!gl)
            cFinal = cFinal + !Final
            cLiq = cLiq + !gl
            .MoveNext
        Loop
        .Close
        Msflxgrd.AddItem "Total...:" & Chr(9) & FormatVal(cFinal) & Chr(9) & FormatVal(cLiq)
    End With
    P_Status ""
    ZebraGrid Msflxgrd, 1
End Sub

Sub grafico(coluna As Integer)
    Dim b(21) As valChart
    Dim sTab As String
    Dim iCount As Integer
    Dim sNome As String
    Dim cTotal As Currency
    
    If optnCat.Value Then
        sTab = "Categorias"
    Else
        sTab = "Marcas"
    End If
    
    
    With Msflxgrd
        'b.count = msflxgrd.rows - 2
        For iCount = 1 To 20
            .Row = iCount
            .col = 0: b(iCount - 1).caption = .Text
            .col = coluna: b(iCount - 1).Value = CCur(.Text)
        Next
        For iCount = 21 To .rows - 2
            .Row = iCount
            col = coluna: cTotal = cTotal + CCur(.Text)
        Next
        b(20).caption = "Restante"
        b(20).Value = cTotal
    End With
    sNome = "Gráfico Comparativo de  " & sTab & " por " & IIf(coluna = 1, "Ganho Bruto", "Ganho Líquido")
    P_Graph b, 16, 1, True, sNome
End Sub

Private Sub cmndbtnGraph_Click()
    grafico 1
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    
    Dim iCount As Integer
    obj.FontName = "arial"
    P_CabRelTimbr "Relatório de Lucratividade por Composição", obj
    
    obj.fontsize = 11
    obj.FontBold = True
    
    P_ImprX obj, "De:" & txtbxDataIni.Text & " Até: " & txtbxdataFin.Text, 0, True, False, 4
    obj.Print ""
    P_Line obj, 10, obj.CurrentY, 170, obj.CurrentY, vbDot
    
    
    P_Line obj, 110, obj.CurrentY, 110, obj.CurrentY + 5, vbDot
    
    With Msflxgrd
        .Row = 0
    
        .col = 0: P_ImprX obj, .Text, 10, True, False, 4
        .col = 1:  P_ImprX obj, .Text, 140, True, True, 4
        
        .col = 2: P_ImprX obj, .Text, 170, False, True, 4
        P_Line obj, 10, obj.CurrentY, 170, obj.CurrentY, vbDot
        obj.FontBold = False
        For iCount = 1 To .rows - 1
            .Row = iCount
            .col = 0
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj, "Continua..."
            End If
            P_ImprX obj, .Text, 10, True, False, 4
            P_Line obj, 110, obj.CurrentY, 110, obj.CurrentY + 5, vbDot
            .col = 1
            P_ImprX obj, .Text, 140, True, True, 4
            .col = 2
            P_ImprX obj, .Text, 170, False, True, 4
            P_Line obj, 10, obj.CurrentY, 170, obj.CurrentY, vbDot
            
        Next
    End With
    P_RodapeVis obj, strNomeSis & " " & constVersao_Sis
    
    
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Me.Top = 10
    Me.Left = 10
    
    With Msflxgrd
        .FixedCols = 0
        .rows = 1
        .Cols = 3
        .col = 0: .Text = "Categoria"
        .col = 1: .Text = "Gerado"
        .col = 2: .Text = "Líquido"
        .ColWidth(0) = 2500
    End With
End Sub

Private Sub Label1_Click()
End Sub

Private Sub optnGer_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub optnLiq_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub optnNome_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
