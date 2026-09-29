VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelPosVenda 
   Caption         =   "Informações de Pós-Venda"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
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
End
Attribute VB_Name = "frmRelPosVenda"
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
    strSQL = "select    count(*) as qtd,"
    strSQL = strSQL & " case "
    strSQL = strSQL & " when (datediff(curdate(),datanasc_cli)/365) >70 then ('MAIS QUE 70 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <70) AND ((datediff(curdate(),datanasc_cli)/365) >60)  then ('ENTRE 60 E 70 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <60) and ((datediff(curdate(),datanasc_cli)/365) >50)  then ('ENTRE 50 E 60 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <50) and ((datediff(curdate(),datanasc_cli)/365) >40)  then ('ENTRE 40 E 50 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <40) and ((datediff(curdate(),datanasc_cli)/365) >30)  then ('ENTRE 30 E 40 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <30) and ((datediff(curdate(),datanasc_cli)/365) >20)  then ('ENTRE 20 E 30 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <20) and ((datediff(curdate(),datanasc_cli)/365) >0)  then ('MENOR QUE 20')"
    strSQL = strSQL & " else ('NULO')"
    strSQL = strSQL & " End"
    strSQL = strSQL & " as faixa"
    strSQL = strSQL & " from tab_clientes as tc inner join tab_vendas as tv on tv.cod_cli=tc.cod_cli "
    'strSQL = strSQL & " left join tab_acertos as ta on ta.cod_cli=tc.cod_cli "
    strSQL = strSQL & " where tc.cod_cli <> 0 "
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtAte.Text)
    End If
    strSQL = strSQL & " group by faixa order by faixa "
    P_Status "Aguarde... pesquisando"
    
    abreObjetoAdo rPesq, strSQL
    grid_faixa.rows = 1
    
    With rPesq
        Do While Not .EOF
            grid_faixa.AddItem !faixa & Chr(9) & !QTD
            
            .MoveNext
        Loop
    End With
    
    
    
    
    
    
    
    strSQL = "select    count(*) as qtd,"
    strSQL = strSQL & " case "
    strSQL = strSQL & " when (datediff(curdate(),datanasc_cli)/365) >70 then ('MAIS QUE 70 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <70) AND ((datediff(curdate(),datanasc_cli)/365) >60)  then ('ENTRE 60 E 70 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <60) and ((datediff(curdate(),datanasc_cli)/365) >50)  then ('ENTRE 50 E 60 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <50) and ((datediff(curdate(),datanasc_cli)/365) >40)  then ('ENTRE 40 E 50 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <40) and ((datediff(curdate(),datanasc_cli)/365) >30)  then ('ENTRE 30 E 40 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <30) and ((datediff(curdate(),datanasc_cli)/365) >20)  then ('ENTRE 20 E 30 ANOS')"
    strSQL = strSQL & " when ((datediff(curdate(),datanasc_cli)/365) <20) and ((datediff(curdate(),datanasc_cli)/365) >0)  then ('MENOR QUE 20')"
    strSQL = strSQL & " else ('NULO')"
    strSQL = strSQL & " End"
    strSQL = strSQL & " as faixa"
    strSQL = strSQL & " from tab_clientes as tc inner join tab_acertos as tv on tv.cod_cli=tc.cod_cli "
    'strSQL = strSQL & " left join tab_acertos as ta on ta.cod_cli=tc.cod_cli "
    strSQL = strSQL & " where tc.cod_cli <> 0 "
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and tv.data_ace>=" & F_GravaData(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and tv.data_ace<=" & F_GravaData(txtAte.Text)
    End If

    strSQL = strSQL & " group by faixa "
    
    
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        Do While Not .EOF
            linha = F_InFlxgrd(grid_faixa, !faixa, 0)
            If linha > -1 Then
                grid_faixa.row = linha
                grid_faixa.Col = 3
                grid_faixa.Text = !QTD
                
            Else
                grid_faixa.AddItem Chr(9) & Chr(9) & Chr(9) & !faixa & Chr(9) & !QTD
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
    strSQL = "select    count(*) as qtd,ts.nome_set "
    
    strSQL = strSQL & " from tab_clientes as tc  inner join tab_vendas as tv on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " left join tab_setores as ts on ts.cod_set=tc.cod_set "
    strSQL = strSQL & " where tc.cod_cli <> 0 "
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtAte.Text)
    End If
    strSQL = strSQL & " group by nome_set order by nome_set "
    P_Status "Aguarde... pesquisando"
    
    abreObjetoAdo rPesq, strSQL
    grid_setores.rows = 1
    
    With rPesq
        Do While Not .EOF
            grid_setores.AddItem !nome_set & Chr(9) & !QTD
            cTotal = cTotal + !QTD
            
            .MoveNext
        Loop
        .Close
        grid_setores.AddItem "TOTAL" & Chr(9) & cTotal
    End With
    
    
    With grid_setores
        For i = 1 To .rows - 1
            .row = i
            .Col = 1
            cPerc = CCur(.Text) / cTotal * 100
            .Col = 2
            .Text = FormatVal(cPerc)
            
        Next
    End With
    
    P_Status ""
    
    
End Sub

Sub ajustaGrid()
    With grid_faixa
        .FixedCols = 0
        .rows = 1
        .Cols = 7
        .Col = 0: .Text = "Faixa"
        .ColWidth(0) = 2000
        .Col = 1: .Text = "Vendas"
        .Col = 2: .Text = "%"
        .Col = 3: .Text = "Acertos"
        .Col = 4: .Text = "%"
        .Col = 5: .Text = "Total"
        .Col = 6: .Text = "%"
    
    End With
    
    With grid_setores
        .FixedCols = 0
        .rows = 1
        .Cols = 3
        .Col = 0: .Text = "Setor"
        .ColWidth(0) = 4000
        .Col = 1: .Text = "Vendas"
        .Col = 2: .Text = "%"
        
    
    End With
    
    
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
    strSQL = "select count(*) as qtd,sexo_cli "
    strSQL = strSQL & " from tab_clientes as tc inner join tab_vendas as tv on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " where tc.cod_cli <> 0 "
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtAte.Text)
    End If
    strSQL = strSQL & " group by sexo_cli"
    P_Status "Aguarde... pesquisando"
    
    abreObjetoAdo rPesq, strSQL
    lblFeminino.caption = "0"
    lblMasculino.caption = "0"
    With rPesq
        Do While Not .EOF
            If !sexo_cli = "M" Then
                cM = !QTD
            ElseIf !sexo_cli = "F" Then
                cF = !QTD
            
            Else
                cOutros = !QTD
            End If
            cTotal = cTotal + !QTD
            
            
            
            .MoveNext
        Loop
        .Close
        lblQtdCliCompradores.caption = cTotal
        lblFeminino.caption = FormatVal(cF / cTotal * 100) & "%"
        lblMasculino.caption = FormatVal(cM / cTotal * 100) & "%"
        
    End With
End Sub
Private Sub Form_Load()
    txtDe.Text = F_PrimDiaMes(Date)
    txtAte.Text = Date
    Set adoFaixa = New ADODB.Recordset
    
    ajustaGrid
    
    
End Sub

Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
            .Col = 1
            If .Text = "" Then .Text = "0"
            a = .Text
            cTotalV = cTotalV + CCur(.Text)
            .Col = 3
            If .Text = "" Then .Text = "0"
            b = .Text
            cTotalA = cTotalA + CCur(.Text)
            .Col = 5
            .Text = a + b
            
        Next
        .AddItem "Total" & Chr(9) & cTotalV & Chr(9) & Chr(9) & cTotalA & Chr(9) & Chr(9) & (cTotalA + cTotalV)
        
        For i = 1 To .rows - 1
            .row = i
            .Col = 1
            cPerc = CCur(.Text) / cTotalV * 100
            .Col = 2: .Text = FormatVal(cPerc)
            
            .Col = 3
            cPerc = CCur(.Text) / cTotalA * 100
            .Col = 4: .Text = FormatVal(cPerc)
            
            .Col = 5
            cPerc = CCur(.Text) / (cTotalA + cTotalV) * 100
            .Col = 6: .Text = FormatVal(cPerc)
            
            
            
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub
