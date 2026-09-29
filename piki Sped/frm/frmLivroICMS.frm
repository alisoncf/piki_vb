VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLivIcms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Livro de ICMS"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   Icon            =   "frmLivroICMS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11085
   Begin VB.Frame frApur 
      Caption         =   "Apuração: (Créd- Déb)"
      Height          =   855
      Left            =   8400
      TabIndex        =   28
      Top             =   5520
      Width           =   2655
      Begin VB.Label lblApurImp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label labelimp 
         AutoSize        =   -1  'True
         Caption         =   "Valor Imposto:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label lblApurTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Valor Contábil:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Frame frSai 
      Caption         =   "Sub Totais de Saída:"
      Height          =   855
      Left            =   5040
      TabIndex        =   23
      Top             =   5520
      Width           =   2415
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Valor Contábil:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblTotal_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valor Imposto:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label lblImp_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   525
         Width           =   1095
      End
   End
   Begin VB.Frame frEntr 
      Caption         =   "Sub Totais de Entrada:"
      Height          =   855
      Left            =   2520
      TabIndex        =   18
      Top             =   5520
      Width           =   2415
      Begin VB.Label lblImp_Ent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   22
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valor Imposto:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label lblTotal_Ent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor Contábil:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1020
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrdLivro 
      Height          =   5295
      Left            =   2520
      TabIndex        =   14
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9340
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   192
      ForeColorFixed  =   -2147483634
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmndbtnConfig 
         Caption         =   "&Configurar Impressão"
         Height          =   615
         Left            =   120
         Picture         =   "frmLivroICMS.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnVisualizar 
         Caption         =   "&Visualizar"
         Height          =   615
         Left            =   120
         Picture         =   "frmLivroICMS.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   120
         Picture         =   "frmLivroICMS.frx":09CE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5400
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   120
         Picture         =   "frmLivroICMS.frx":0E10
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4680
         Width           =   2055
      End
      Begin VB.ComboBox cmbxOrderBy 
         Height          =   315
         ItemData        =   "frmLivroICMS.frx":119A
         Left            =   120
         List            =   "frmLivroICMS.frx":11AD
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnPesq 
         Caption         =   "&Localizar"
         Height          =   615
         Left            =   120
         Picture         =   "frmLivroICMS.frx":11E3
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtbxDataFin 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtbxData_Ini 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox cmbxLivro 
         Height          =   315
         ItemData        =   "frmLivroICMS.frx":156D
         Left            =   120
         List            =   "frmLivroICMS.frx":157A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Ordernar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Período de:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Livro:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Label lblie_ter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   9720
      TabIndex        =   17
      Top             =   6120
      Width           =   45
   End
   Begin VB.Label lblCNPJ_Ter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7800
      TabIndex        =   16
      Top             =   6120
      Width           =   45
   End
   Begin VB.Label lblNome_ter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   3120
      TabIndex        =   15
      Top             =   6120
      Width           =   45
   End
End
Attribute VB_Name = "frmLivIcms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bRow As Boolean
Dim cTotal_Ent As Currency
Dim cTotal_Sai As Currency
Dim cImp_ent As Currency
Dim cImp_Sai As Currency
Private Sub cmbxLivro_Click()
    bRow = False
    If cmbxLivro.ListIndex = 0 Or cmbxLivro.ListIndex = 1 Then
        With msflxgrdLivro
            .Rows = 1
            .Cols = 14
            .Col = 0
            .CellAlignment = 3
            .Text = "Entrada"
            
            .Col = 1
            .CellAlignment = 3
            .Text = "Tipo"
            
            .Col = 2
            .CellAlignment = 3
            .Text = "Série"
            
            .Col = 3
            .CellAlignment = 3
            .Text = "Número"
            .Col = 4
            .CellAlignment = 3
            .Text = "Emissão"
            .Col = 5
            .CellAlignment = 3
            .Text = "Terc."
            .Col = 6
            .CellAlignment = 3
            .Text = "UF"
            .Col = 7
            .CellAlignment = 3
            .Text = "Valor"
            .Col = 8
            .CellAlignment = 3
            .Text = "Oper."
            .Col = 9
            .CellAlignment = 3
            .Text = "Cód"
            .Col = 10
            .CellAlignment = 3
            .Text = "B.Cálculo"
            .Col = 11
            .CellAlignment = 3
            .Text = "Aliq"
            .Col = 12
            .CellAlignment = 3
            .Text = "Creditado"
            .Col = 13
            .CellAlignment = 3
            .Text = "Terceiro"
            .ColWidth(0) = 1000
            .ColAlignment(1) = 3
            .ColAlignment(2) = 3
            .ColWidth(1) = 800
            .ColWidth(2) = 800
            .ColWidth(3) = 800
            .ColWidth(4) = 1000
            .ColWidth(5) = 600
            .ColWidth(6) = 400
            .ColWidth(7) = 800
            .ColWidth(8) = 600
            .ColWidth(9) = 400
            .ColWidth(10) = 800
            .ColWidth(11) = 400
            .ColWidth(12) = 800
            .ColWidth(13) = 4000
            
        End With
    Else
            With msflxgrdLivro
            .Rows = 1
            .Cols = 6
            .Col = 0
            .CellAlignment = 3
            .Text = "Operação"
            .Col = 1
            .CellAlignment = 3
            .Text = "Valores"
            .Col = 2
            .CellAlignment = 3
            .Text = "B.Cálculo"
            .Col = 3
            .CellAlignment = 3
            .Text = "Creditado"
            .Col = 4
            .CellAlignment = 3
            .Text = "Is./N.Trib"
            .Col = 5
            .CellAlignment = 3
            .Text = "Outras"
            .ColWidth(0) = 800
            .ColWidth(1) = 1200
            .ColWidth(2) = 1200
            .ColWidth(3) = 1200
            .ColWidth(4) = 1200
            .ColWidth(5) = 1200
        End With
    End If
    bRow = True
End Sub

Private Sub cmndbtnConfig_Click()
    callobjP
End Sub

Sub P_Relat(obj As Object, Optional bPrinter As Boolean = True)
    Dim iCount As Integer
    Dim cTot As Currency
    Dim cImposto As Currency
    P_CabRelTimbr "Livro de ICMS: " & cmbxLivro.Text, obj
    obj.CurrentY = obj.CurrentY + 2
    P_ImprX obj, "Entrada", 5 '0
    P_ImprX obj, "Doc.", 22 '1
    P_ImprX obj, "Série", 30 '2
    P_ImprX obj, "Número", 40 '3
    P_ImprX obj, "Terc.", 60
    P_ImprX obj, "UF", 75
    P_ImprX obj, "Contábil", 95, True, True
    P_ImprX obj, "Operação", 105
    P_ImprX obj, "Cod", 125
    P_ImprX obj, "B.Cálc.", 155, True, True
    P_ImprX obj, "Vlr. Imposto", 175, False, True
    P_ImprLinha obj
    With msflxgrdLivro
        For iCount = 1 To msflxgrdLivro.Rows - 1
            .Row = iCount
            .Col = 0
            P_ImprX obj, .Text, 5  '0
            .Col = 1
            P_ImprX obj, .Text, 22 '1
            .Col = 2
            P_ImprX obj, .Text, 30 '2
            .Col = 3
            P_ImprX obj, .Text, 40 '3
            .Col = 5
            P_ImprX obj, .Text, 60
            .Col = 6
            
            P_ImprX obj, .Text, 75
            .Col = 7
            cTot = cTot + CCur(.Text)
            P_ImprX obj, .Text, 95, True, True
            .Col = 8
            P_ImprX obj, .Text, 105
            .Col = 9
            P_ImprX obj, .Text, 125
            .Col = 10
            P_ImprX obj, .Text, 155, True, True
            
            .Col = 11
            cImposto = cImposto + CCur(.Text)
            P_ImprX obj, .Text, 175, False, True
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                If bPrinter = True Then
                    P_Rodape "Continua", "Pág: " & Printer.Page
                Else
                    P_RodapeVis obj, "Continua", "Pág: XXX"
                End If
            End If
        Next
        P_ImprLinha obj
        obj.FontBold = True
        obj.FontUnderline = True
        P_ImprX obj, FormatVal(CStr(cTot)), 95, True, True
        P_ImprX obj, FormatVal(CStr(cImposto)), 175, False, True
        obj.FontBold = False
        obj.FontUnderline = False
        If bPrinter = True Then
            P_Rodape "", "Pág: " & Printer.Page
            Printer.EndDoc
        Else
            P_RodapeVis obj, "", "Pág: XXX"
        End If
        
    End With
    
    
    
End Sub

Sub P_RelatApur(obj As Object, Optional bPrinter As Boolean = True)
    Dim iCount As Integer
    Dim cTot As Currency
    Dim cImposto As Currency
    Dim iCol As Integer
    P_CabRelTimbr "Livro de ICMS: " & cmbxLivro.Text, obj
    obj.CurrentY = obj.CurrentY + 2
    
    P_ImprX obj, "Entradas...", 5, False '0
    P_ImprLinha obj
    
    P_ImprX obj, "Operação", 5 '1
    P_ImprX obj, "Valores", 45, True, True '2
    P_ImprX obj, "B.Cálculo", 60, True, True '2
    P_ImprX obj, "Creditado", 85, False, True '2
    P_ImprLinha obj
    With msflxgrdLivro
        
        For iCount = 1 To msflxgrdLivro.Rows - 1
            .Col = 0
            .Row = iCount
            If Left(.Text, 1) = "S" Then
                P_ImprLinha obj
                'iCol = iCount + 1
                'Exit For
            End If
            
            .Col = 0
            P_ImprX obj, .Text, 5  '0
            .Col = 1
            P_ImprX obj, .Text, 45, True, True '1
            .Col = 2
            P_ImprX obj, .Text, 60, True, True '2
            .Col = 3
            P_ImprX obj, .Text, 85, False, True '3
            
        Next
        P_ImprLinha obj
        obj.FontBold = True
        obj.FontUnderline = True

        P_ImprX obj, "Apuração Total...", 5
        obj.ForeColor = vbRed
        P_ImprX obj, "Valor Total...:" & FormatVal(CStr(cTotal_Sai - cTotal_Ent)), 80, True, True
        P_ImprX obj, "Valor do Imposto:" & (CStr(cImp_Sai - cImp_ent)), 130, False, True
        obj.FontBold = False
        obj.FontUnderline = False
        obj.ForeColor = vbBlack
        P_ImprLinha obj
        If bPrinter = True Then
            P_Rodape "", "Pág: " & Printer.Page
            Printer.EndDoc
        Else
            P_RodapeVis obj, "", "Pág: XXX"
        End If
        
    End With
    
    
    
End Sub

Private Sub cmndbtnImprimir_Click()
    If cmbxLivro.ListIndex <> 2 Then
        P_Relat Printer
    Else
        P_RelatApur Printer
    End If
End Sub

Private Sub cmndbtnPesq_Click()
    lblApurImp.caption = "0,00"
    lblApurTotal.caption = "0,00"
    lblTotal_Ent.caption = "0,00"
    lblTotal_Sai.caption = "0,00"
    lblImp_Ent.caption = "0,00"
    lblImp_Sai.caption = "0,00"
    msflxgrdLivro.Visible = False
    If cmbxLivro.ListIndex = 0 Then
        P_buscaReg "E"
    ElseIf cmbxLivro.ListIndex = 1 Then
        P_buscaReg "S"
    Else
        P_BuscaApur "E"
        P_BuscaApur "S"
        
    End If
    P_Resumo
    msflxgrdLivro.Visible = True
End Sub
Sub P_Resumo()
    With msflxgrdLivro
        If cmbxLivro.ListIndex = 0 Then
            frApur.Enabled = False
            frSai.Enabled = False
            frEntr.Enabled = True
            lblTotal_Ent.caption = FormatVal(CStr(cTotal_Ent))
            lblImp_Ent.caption = FormatVal(CStr(cImp_ent))
        ElseIf cmbxLivro.ListIndex = 1 Then
            frApur.Enabled = False
            frSai.Enabled = True
            frEntr.Enabled = False
            lblTotal_Sai.caption = FormatVal(CStr(cTotal_Sai))
            lblImp_Sai.caption = FormatVal(CStr(cImp_Sai))
        Else
            frApur.Enabled = True
            frSai.Enabled = True
            frEntr.Enabled = True
            lblApurTotal.caption = FormatVal(CStr(cTotal_Sai - cTotal_Ent))
            lblApurImp.caption = FormatVal(CStr(cImp_Sai - cImp_ent))
            
            lblTotal_Ent.caption = FormatVal(CStr(cTotal_Ent))
            lblImp_Ent.caption = FormatVal(CStr(cImp_ent))
            
            lblTotal_Sai.caption = FormatVal(CStr(cTotal_Sai))
            lblImp_Sai.caption = FormatVal(CStr(cImp_Sai))
            
            
        End If
    End With
End Sub
Sub P_BuscaApur(sTipo As String)
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sOrderBy As String
    Dim iCount As Integer
    If sTipo = "E" Then
        cTotal_Ent = 0
        cImp_ent = 0
    Else
        
        cTotal_Sai = 0
        cImp_Sai = 0
    End If
    
    strSQL = "Select sum(tn.Valor_not) as contabil,sum(bc_Not) as bc,sum(icms_not) icms, "
    strSQL = strSQL & " tn.operacao_not "
    strSQL = strSQL & " from tab_notasfiscais as tn where "
    strSQL = strSQL & " tn.cod_cli='" & strCOD_CLI & "' "
    strSQL = strSQL & " and tn.entrada_not>=" & F_PesqData(txtbxData_Ini.Text)
    strSQL = strSQL & " and tn.entrada_not<=" & F_PesqData(txtbxdataFin.Text)
    strSQL = strSQL & " and tipo_not='" & sTipo & "' "
    strSQL = strSQL & " group by tn.operacao_not "
    strSQL = strSQL & " order by tn.operacao_not "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        
        If sTipo = "E" Then
            msflxgrdLivro.Rows = 1
            msflxgrdLivro.AddItem "Entradas...  "
            
            
        Else
            
            msflxgrdLivro.AddItem "Saídas...  "
        End If
        
                msflxgrdLivro.Row = msflxgrdLivro.Rows - 1
                For iCount = 0 To msflxgrdLivro.Cols - 1
                    msflxgrdLivro.Col = iCount
                    msflxgrdLivro.CellBackColor = vbGreen
                Next
        Do While Not .EOF
            strSQL = !operacao_not & Chr(9) & FormatVal(!Contabil) & Chr(9)
            strSQL = strSQL & FormatVal(!bc) & Chr(9) & FormatVal(!Icms) & Chr(9)
            If sTipo = "E" Then
                cTotal_Ent = cTotal_Ent + !Contabil
                cImp_ent = cImp_ent + !Icms
            Else
                cTotal_Sai = cTotal_Sai + !Contabil
                cImp_Sai = cImp_Sai + !Icms
            End If
            
            msflxgrdLivro.AddItem strSQL
            .MoveNext
            msflxgrdLivro.Row = msflxgrdLivro.Rows - 1
            If msflxgrdLivro.Row Mod 2 = 0 Then
                For iCount = 0 To msflxgrdLivro.Cols - 1
                    msflxgrdLivro.Col = iCount
                    msflxgrdLivro.CellBackColor = vbYellow
                Next
            End If
        Loop
        .Close
        msflxgrdLivro.AddItem "Sub Totais..."
        msflxgrdLivro.Row = msflxgrdLivro.Rows - 1
        
        msflxgrdLivro.Col = 1
        If sTipo = "E" Then
            msflxgrdLivro.Text = FormatVal(CStr(cTotal_Ent))
        Else
            msflxgrdLivro.Text = FormatVal(CStr(cTotal_Sai))
        End If
        msflxgrdLivro.CellFontBold = True
        msflxgrdLivro.CellFontUnderline = True
        msflxgrdLivro.Col = 3
        msflxgrdLivro.CellFontBold = True
        msflxgrdLivro.CellFontUnderline = True
        If sTipo = "E" Then
            msflxgrdLivro.Text = FormatVal(CStr(cImp_ent))
        Else
            msflxgrdLivro.Text = FormatVal(CStr(cImp_Sai))
        End If
    End With
End Sub

    
Sub P_buscaReg(sTipo_Not As String)
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sOrder As String
    Dim iCount As Integer
    cTotal_Ent = 0
    cTotal_Sai = 0
    cImp_ent = 0
    cImp_Sai = 0
    bRow = False
    If cmbxOrderBy.ListIndex = 0 Then
        sOrder = "Num_Not,cod_Cli"
    ElseIf cmbxOrderBy.ListIndex = 1 Then
        sOrder = "emissao_Not,Num_Not,cod_Cli"
    ElseIf cmbxOrderBy.ListIndex = 2 Then
        sOrder = "entrada_not,Num_Not,cod_Cli"
    ElseIf cmbxOrderBy.ListIndex = 3 Then
        sOrder = "cod_Cli,num_Not"
    Else
        sOrder = "uf_cli,cod_Cli,num_not"
    End If
    
    strSQL = "Select tn.*,tt.nome_ter,tt.uf_ter from tab_notasfiscais as tn, "
    strSQL = strSQL & " tab_terceiros as tt where tt.cod_ter=tn.cod_ter and "
    strSQL = strSQL & " tn.cod_cli='" & strCOD_CLI & "' "
    strSQL = strSQL & " and tn.entrada_not>=" & F_PesqData(txtbxData_Ini.Text)
    strSQL = strSQL & " and tn.entrada_not<=" & F_PesqData(txtbxdataFin.Text)
    strSQL = strSQL & " and tipo_not='" & sTipo_Not & "' "
    strSQL = strSQL & " order by " & sOrder
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    msflxgrdLivro.Rows = 1
    With rPesq
        Do While Not .EOF
            strSQL = !Entrada_NOt & Chr(9) & F_TipoDoc(!tipodoc_Not, True) & Chr(9) & !Serie_Not & Chr(9) & !Num_Not & Chr(9)
            strSQL = strSQL & !Emissao_Not & Chr(9) & !Cod_Ter & Chr(9)
            strSQL = strSQL & F_BuscaUmCampo("UF_TER", "TAb_Terceiros", "Cod_Ter='" & _
                     !Cod_Ter & "'") & Chr(9)
            strSQL = strSQL & Format(!Valor_Not, "##,###,##0.00") & Chr(9)
            strSQL = strSQL & !operacao_not & Chr(9) & !codval_Not & Chr(9)
            strSQL = strSQL & Format(!BC_Not, "##,###,##0.00") & Chr(9)
            If !Icms_Not <> 0 Then
                strSQL = strSQL & FormatVal(CStr(!Valor_Not / !Icms_Not)) & Chr(9)
            Else
                strSQL = strSQL & "0,00" & Chr(9)
            End If
            strSQL = strSQL & FormatVal(!Icms_Not) & Chr(9)
            strSQL = strSQL & !Nome_Ter
            msflxgrdLivro.AddItem strSQL
            If sTipo_Not = "E" Then
                cTotal_Ent = cTotal_Ent + !Valor_Not
                cImp_ent = cImp_ent + !Icms_Not
            Else
                cTotal_Sai = cTotal_Sai + !Valor_Not
                cImp_Sai = cImp_Sai + !Icms_Not
            End If
            
            msflxgrdLivro.Row = msflxgrdLivro.Rows - 1
            If msflxgrdLivro.Row Mod 2 = 0 Then
                For iCount = 0 To msflxgrdLivro.Cols - 1
                    msflxgrdLivro.Col = iCount
                    msflxgrdLivro.CellBackColor = vbYellow
                Next
            End If
            .MoveNext
        Loop
        If .RecordCount = 0 Then
            MsgBox "Não foram encontrados registros nesta pesquisa!", vbExclamation, constNome_sis
        End If
        .Close
    End With
    bRow = True
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVisualizar_Click()
    If cmbxLivro.ListIndex <> 2 Then
        CallView
        P_Relat frmVisualiza.Pic, False
    Else
        CallView
        P_RelatApur frmVisualiza.Pic, False
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    Me.caption = Format(strCOD_CLI, "0###\-#") & "  " & _
        F_BuscaUmCampo("NOme_Cli", "Tab_Clientes", "cod_cli='" & strCOD_CLI & "'")
    cmbxOrderBy.ListIndex = 0
End Sub

Private Sub Label9_Click()

End Sub

Private Sub txtbxData_Ini_GotFocus()
    P_SelecTexto txtbxData_Ini
End Sub

Private Sub txtbxData_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxData_Ini_LostFocus()
    P_AjustaData txtbxData_Ini
End Sub

Private Sub txtbxdataFin_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxdataFin_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
            SendKeys Chr(9)
        End If
End Sub

Private Sub txtbxdataFin_LostFocus()
    
    P_AjustaData txtbxdataFin
    If F_DataMaior(txtbxData_Ini.Text, txtbxdataFin.Text) Then
        MsgBox "A Data Inicial não pode ser maior que a Data Final!", vbExclamation
    End If
End Sub
