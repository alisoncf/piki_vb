VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelPosVendaFaixa 
   Caption         =   "Informações de Pós-Venda"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   735
      Left            =   8040
      TabIndex        =   16
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   735
      Left            =   6720
      TabIndex        =   15
      Top             =   360
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grid_faixa 
      Height          =   2535
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4471
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Pesquisar"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtAte 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtDe 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grid_setores 
      Height          =   2415
      Left            =   240
      TabIndex        =   14
      Top             =   4920
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4260
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label lblFeminino 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   12
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Feminino:"
      Height          =   195
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label lblMasculino 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   10
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Masculino:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Clientes por Setores"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   1410
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Clientes por Faixa Etária"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblQtdCliCompradores 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   120
   End
   Begin VB.Label Label3 
      Caption         =   "Até:"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "De:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número de Clientes Compradores"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2355
   End
   Begin VB.Menu mnMenu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnCopy 
         Caption         =   "Copiar"
      End
   End
End
Attribute VB_Name = "frmRelPosVendaFaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoFaixa As ADODB.Recordset
Attribute adoFaixa.VB_VarHelpID = -1
Private Sub Faixa_Etaria()

    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim linha As Integer
    strSQL = "select    count(distinct tc.cod_cli) as qtd,"
    strSQL = strSQL & " case "
    strSQL = strSQL & " when ((to_days(curdate()) - to_days(datanasc_cli))/365) >70 then ('MAIS QUE 70 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <70) AND (((to_days(curdate()) - to_days(datanasc_cli))/365) >60)  then ('ENTRE 60 E 70 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <60) and (((to_days(curdate()) - to_days(datanasc_cli))/365) >50)  then ('ENTRE 50 E 60 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <50) and (((to_days(curdate()) - to_days(datanasc_cli))/365) >40)  then ('ENTRE 40 E 50 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <40) and (((to_days(curdate()) - to_days(datanasc_cli))/365) >30)  then ('ENTRE 30 E 40 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <30) and (((to_days(curdate()) - to_days(datanasc_cli))/365) >20)  then ('ENTRE 20 E 30 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <20) and (((to_days(curdate()) - to_days(datanasc_cli))/365) >0)  then ('MENOR QUE 20')"
    strSQL = strSQL & " else ('NULO')"
    strSQL = strSQL & " End"
    strSQL = strSQL & " as faixa"
    strSQL = strSQL & " from tab_clientes as tc inner join tab_vendas as tv on tv.cod_cli=tc.cod_cli "
    'strSQL = strSQL & " left join tab_acertos as ta on ta.cod_cli=tc.cod_cli "
    strSQL = strSQL & " where tc.cod_cli <> 0 "
    strSQL = strSQL & " and tv.tipo_operacao=1 "
    If txtDe.text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtDe.text)
    End If
    If txtAte.text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtAte.text)
    End If
    strSQL = strSQL & " group by faixa order by faixa "
    P_Status "Aguarde... pesquisando"
    
    abreObjetoAdo rPesq, strSQL
    grid_faixa.rows = 1
    
    With rPesq
        Do While Not .EOF
            grid_faixa.AddItem !faixa & Chr(9) & !qtd
            
            .MoveNext
        Loop
    End With
    
    
    
    
    
    
    
    strSQL = "select    count(*) as qtd,"
    strSQL = strSQL & " case "
    strSQL = strSQL & " when ((to_days(curdate()) - to_days(datanasc_cli))/365) >70 then ('MAIS QUE 70 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <70) AND (((to_days(curdate()) - to_days(datanasc_cli))/365) >60)  then ('ENTRE 60 E 70 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <60) and (((to_days(curdate()) - to_days(datanasc_cli))/365) >50)  then ('ENTRE 50 E 60 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <50) and (((to_days(curdate()) - to_days(datanasc_cli))/365) >40)  then ('ENTRE 40 E 50 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <40) and (((to_days(curdate()) - to_days(datanasc_cli))/365) >30)  then ('ENTRE 30 E 40 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <30) and (((to_days(curdate()) - to_days(datanasc_cli))/365) >20)  then ('ENTRE 20 E 30 ANOS')"
    strSQL = strSQL & " when (((to_days(curdate()) - to_days(datanasc_cli))/365) <20) and (((to_days(curdate()) - to_days(datanasc_cli))/365) >0)  then ('MENOR QUE 20')"
    strSQL = strSQL & " else ('NULO')"
    strSQL = strSQL & " End"
    strSQL = strSQL & " as faixa"
    strSQL = strSQL & " from tab_clientes as tc inner join tab_acertos as tv on tv.cod_cli=tc.cod_cli "
    'strSQL = strSQL & " left join tab_acertos as ta on ta.cod_cli=tc.cod_cli "
    strSQL = strSQL & " where tc.cod_cli <> 0 "
    If txtDe.text <> "" Then
        strSQL = strSQL & " and tv.data_ace>=" & F_GravaData(txtDe.text)
    End If
    If txtAte.text <> "" Then
        strSQL = strSQL & " and tv.data_ace<=" & F_GravaData(txtAte.text)
    End If

    strSQL = strSQL & " group by faixa "
    
    
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        Do While Not .EOF
            linha = F_InFlxgrd(grid_faixa, !faixa, 0)
            If linha > -1 Then
                grid_faixa.row = linha
                grid_faixa.col = 3
                grid_faixa.text = !qtd
                
            Else
                grid_faixa.AddItem Chr(9) & Chr(9) & Chr(9) & !faixa & Chr(9) & !qtd
            End If
            .MoveNext
        Loop
    End With
    
    'abreObjetoAdo adoFaixa, strSQL, gridFaixa
    'abreObjetoAdo rPesq, strSQL
    P_Status ""
    
    
    P_Total
End Sub
Private Sub buscaSetores()

    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim linha As Integer
    Dim i As Integer
    Dim cTotal As Currency
    Dim cPerc As Currency
    strSQL = "select    count( distinct tc.cod_cli) as qtd,ts.nome_set "
    
    strSQL = strSQL & " from tab_clientes as tc  inner join tab_vendas as tv on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " left join tab_setores as ts on ts.cod_set=tc.cod_set "
    strSQL = strSQL & " where tc.cod_cli <> 0 "
    If txtDe.text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtDe.text)
    End If
    If txtAte.text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtAte.text)
    End If
    strSQL = strSQL & " and tv.tipo_operacao=1"
    strSQL = strSQL & " group by nome_set order by nome_set "
    P_Status "Aguarde... pesquisando"
    
    abreObjetoAdo rPesq, strSQL
    grid_setores.rows = 1
    
    With rPesq
        Do While Not .EOF
            grid_setores.AddItem !nome_set & Chr(9) & !qtd
            cTotal = cTotal + !qtd
            
            .MoveNext
        Loop
        .Close
        grid_setores.AddItem "TOTAL" & Chr(9) & cTotal
    End With
    
    
    With grid_setores
        For i = 1 To .rows - 1
            .row = i
            .col = 1
            cPerc = CCur(.text) / cTotal * 100
            .col = 2
            .text = FormatVal(cPerc)
            
        Next
    End With
    
    P_Status ""
    
    
End Sub

Sub AjustaGrid()
    With grid_faixa
        .FixedCols = 0
        .rows = 1
        .Cols = 7
        .col = 0: .text = "Faixa"
        .ColWidth(0) = 2000
        .col = 1: .text = "Vendas"
        .col = 2: .text = "%"
        .col = 3: .text = "Acertos"
        .col = 4: .text = "%"
        .col = 5: .text = "Total"
        .col = 6: .text = "%"
    
    End With
    
    With grid_setores
        .FixedCols = 0
        .rows = 1
        .Cols = 3
        .col = 0: .text = "Setor"
        .ColWidth(0) = 4000
        .col = 1: .text = "Vendas"
        .col = 2: .text = "%"
        
    
    End With
    
    
End Sub

Private Sub cmdImprimir_Click()
    Dim a As String
    
    
    a = html.F_AbreHtml("pos")
    
    html.P_HtmlAbreH 1, 3, "Relatório de Informações de Clientes - Pós-Venda: de " & txtDe.text & " até " & txtAte.text, True
    html.P_HtmlAbreBody 1
    html.P_HtmlAbreH 1, 3, "Número de Clientes Compradores: " & lblQtdCliCompradores.caption
    html.P_HtmlAbreH 1, 3, "Masculino: " & lblMasculino.caption
    html.P_HtmlAbreH 1, 3, "Feminino: " & lblFeminino.caption
        
    
    
    html.P_HtmlFechaBody 1
    
    
    
    
    
    
    
    
End Sub

Private Sub cmdPesquisar_Click()
    geral
    Faixa_Etaria
    buscaSetores
End Sub
Sub geral()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim linha As Integer
    Dim cTotal As Currency
    Dim cM As Currency
    Dim cF As Currency
    Dim cOutros As Currency
    strSQL = "select count(distinct tc.cod_cli) as qtd,sexo_cli "
    strSQL = strSQL & " from tab_clientes as tc inner join tab_vendas as tv on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " where tc.cod_cli <> 0 "
    If txtDe.text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtDe.text)
    End If
    If txtAte.text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtAte.text)
    End If
    strSQL = strSQL & " and tipo_operacao=1"
    strSQL = strSQL & " group by sexo_cli"
    P_Status "Aguarde... pesquisando"
    
    abreObjetoAdo rPesq, strSQL
    lblFeminino.caption = "0"
    lblMasculino.caption = "0"
    With rPesq
        Do While Not .EOF
            If !sexo_cli = "M" Then
                cM = !qtd
            ElseIf !sexo_cli = "F" Then
                cF = !qtd
            
            Else
                cOutros = !qtd
            End If
            cTotal = cTotal + !qtd
            
            
            
            .MoveNext
        Loop
        .Close
        lblQtdCliCompradores.caption = cTotal
        If cTotal > 0 Then
            lblFeminino.caption = FormatVal(cF / cTotal * 100) & "%"
            lblMasculino.caption = FormatVal(cM / cTotal * 100) & "%"
        End If
        
    End With
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtDe.text = F_PrimDiaMes(Date)
    txtAte.text = Date
    Set adoFaixa = New ADODB.Recordset
    
    AjustaGrid
    
    
End Sub

Private Sub grid_faixa_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnMenu
    End If
End Sub

Private Sub grid_setores_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnMenu
    End If
End Sub

Private Sub mnMenu_Click()
    
    If Me.ActiveControl = grid_faixa Then
        P_MsflxToHtml "Faixa Etária", grid_faixa
    ElseIf Me.ActiveControl = grid_setores Then
        P_MsflxToHtml "Faixa Setores", grid_setores
    End If
    
End Sub

Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub
Sub P_Total()
    Dim i As Integer
    Dim cTotalV As Currency
    Dim cTotalA As Currency
    Dim a As Currency
    Dim b As Currency
    
    Dim cPerc As Currency
    With grid_faixa
        
        For i = 1 To .rows - 1
            .row = i
            .col = 1
            If .text = "" Then .text = "0"
            a = .text
            cTotalV = cTotalV + CCur(.text)
            .col = 3
            If .text = "" Then .text = "0"
            b = .text
            cTotalA = cTotalA + CCur(.text)
            .col = 5
            .text = a + b
            
        Next
        .AddItem "Total" & Chr(9) & cTotalV & Chr(9) & Chr(9) & cTotalA & Chr(9) & Chr(9) & (cTotalA + cTotalV)
        
        For i = 1 To .rows - 1
            .row = i
            .col = 1
            cPerc = CCur(.text) / cTotalV * 100
            .col = 2: .text = FormatVal(cPerc)
            
            .col = 3
            cPerc = CCur(.text) / cTotalA * 100
            .col = 4: .text = FormatVal(cPerc)
            
            .col = 5
            cPerc = CCur(.text) / (cTotalA + cTotalV) * 100
            .col = 6: .text = FormatVal(cPerc)
            
            
            
        Next
        
    End With
    
    ZebraGrid grid_faixa, 1
End Sub
Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub
