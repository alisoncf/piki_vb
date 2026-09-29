VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmInventario 
   Caption         =   "Inventário"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   Icon            =   "frmInventario.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   12300
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Height          =   7290
      Left            =   0
      ScaleHeight     =   7230
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.CheckBox chkNeg 
         Caption         =   "Não mostrar Negativas(F) "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   3480
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnHtml 
         Caption         =   "Exportar para HTML"
         Height          =   615
         Left            =   120
         Picture         =   "frmInventario.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Emite Inventário"
         Top             =   6000
         Width           =   1815
      End
      Begin VB.ComboBox cmbxSt 
         Height          =   315
         ItemData        =   "frmInventario.frx":059E
         Left            =   120
         List            =   "frmInventario.frx":05AB
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2880
         Width           =   1785
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Gerar"
         Height          =   615
         Left            =   120
         Picture         =   "frmInventario.frx":05CA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "cria um novo inventário."
         Top             =   3825
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "&Cancelar Inventário"
         Height          =   615
         Left            =   120
         Picture         =   "frmInventario.frx":0E94
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "cria um novo inventário."
         Top             =   4545
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir p/ Conf."
         Height          =   615
         Left            =   120
         Picture         =   "frmInventario.frx":325B
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Emite Inventário"
         Top             =   5265
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   120
         Picture         =   "frmInventario.frx":33A5
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fecha formulário."
         Top             =   6720
         Width           =   1815
      End
      Begin VB.CheckBox chkEstZer 
         Caption         =   "Trazer estoque zerado"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.ComboBox cmbxCat 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   1785
      End
      Begin VB.ComboBox cmbxMarca 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2280
         Width           =   1785
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   735
      End
      Begin VB.ComboBox cmbxInvent 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cmbxOrderBy 
         BackColor       =   &H80000014&
         Height          =   315
         ItemData        =   "frmInventario.frx":34EF
         Left            =   120
         List            =   "frmInventario.frx":34FF
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação Tributária:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   2640
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblCat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1635
         TabIndex        =   29
         Tag             =   "!"
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label lblMar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1635
         TabIndex        =   28
         Tag             =   "!"
         Top             =   2040
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Inventário de:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Organizar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1005
      End
   End
   Begin VB.Data dtInvent 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox pcTotal 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12270
      TabIndex        =   17
      Top             =   7290
      Width           =   12300
      Begin VB.CommandButton cmndbtnEditaItem 
         Caption         =   "Editar Item"
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnExclItem 
         Caption         =   "Excluir Item"
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   975
      End
      Begin VB.Label lblTR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   10680
         TabIndex        =   40
         Top             =   0
         Width           =   105
      End
      Begin VB.Label lblTF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   9480
         TabIndex        =   39
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tributadas:"
         Height          =   195
         Left            =   8640
         TabIndex        =   38
         Top             =   0
         Width           =   795
      End
      Begin VB.Label lblRR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   7320
         TabIndex        =   37
         Top             =   0
         Width           =   105
      End
      Begin VB.Label lblRF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   6120
         TabIndex        =   36
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retidas:"
         Height          =   195
         Left            =   5400
         TabIndex        =   35
         Top             =   0
         Width           =   585
      End
      Begin VB.Label lblIR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   4080
         TabIndex        =   34
         Top             =   0
         Width           =   105
      End
      Begin VB.Label lblIF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   2880
         TabIndex        =   33
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Isentas:"
         Height          =   195
         Left            =   2280
         TabIndex        =   32
         Top             =   0
         Width           =   555
      End
      Begin VB.Label lblPecas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   2400
         TabIndex        =   27
         Top             =   390
         Width           =   105
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prod:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   1800
         TabIndex        =   26
         Top             =   390
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TE:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   3480
         TabIndex        =   25
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lblTotalInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   3840
         TabIndex        =   24
         Top             =   390
         Width           =   105
      End
      Begin VB.Label lbltotalItens 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   23
         Top             =   390
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Itens:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   390
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TR:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   5400
         TabIndex        =   21
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lblTotalFisc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   5760
         TabIndex        =   20
         Top             =   390
         Width           =   105
      End
      Begin VB.Label lblTV 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   7560
         TabIndex        =   19
         Top             =   390
         Width           =   105
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TV:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   7200
         TabIndex        =   18
         Top             =   390
         Width           =   315
      End
   End
   Begin MSDBGrid.DBGrid dbgrdInvent 
      Bindings        =   "frmInventario.frx":352C
      Height          =   5415
      Left            =   2160
      OleObjectBlob   =   "frmInventario.frx":3543
      TabIndex        =   16
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Sub P_Total()
    Dim cTotal As Currency
    Dim iItens As Single
    Dim ipecas As Single
    Dim cFisc As Currency
    Dim cNeg As Currency
    Dim cTV As Currency
    Dim cIF As Currency
    Dim cIR As Currency
    Dim cRF As Currency
    Dim cRR As Currency
    Dim cTF As Currency
    Dim cTR As Currency
    
    
    With dtInvent.Recordset
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            If !Estoque > 0 Then
                iItens = iItens + 1
                ipecas = ipecas + (!Estoque)
            End If
            cTotal = cTotal + (!total)
            If !total < 0 Then
                cNeg = !total + cNeg
            End If
            If !st_pro = "I" Then
                cIR = cIR + !total
                cIF = cIF + !cf_pro * !ef_pro
            ElseIf !st_pro = "R" Then
                cRR = cRR + !total
                cRF = cRF + !cf_pro * !ef_pro
            ElseIf !st_pro = "T" Then
                cTR = cTR + !total
                cTF = cTF + !cf_pro * !ef_pro
            End If
            cFisc = cFisc + !cf_pro * !ef_pro
            cTV = cTV + !tv
            P_Status "Somando...: " & Round((.AbsolutePosition + 1) / .RecordCount * 100) & "%"
            .MoveNext
        Loop
        .MoveFirst
        lblPecas.caption = Format(ipecas)
        lbltotalItens.caption = Format(iItens)
        lblTotalInv.caption = FormatVal(cTotal)
        lblTotalFisc.caption = FormatVal(cFisc)
        lblTV.caption = FormatVal(cTV)
        
        lblIR.caption = FormatVal(cIR)
        lblIF.caption = FormatVal(cIF)
        lblRR.caption = FormatVal(cRR)
        lblRF.caption = FormatVal(cRF)
        lblTR.caption = FormatVal(cTR)
        lblTF.caption = FormatVal(cTF)
        
    End With
'    MsgBox cNeg & " Negativo"
    P_Status ""
End Sub
Private Sub ListView1_Click()
        ListView1.ListItems.Add 1, , "teste"
End Sub
Sub P_BuscaInvent()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select distinct datahora_inv from tab_invent order by datahora_inv desc"
        Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        cmbxInvent.Clear
        Do While Not .EOF
            cmbxInvent.AddItem !DataHora_Inv
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub chkEstZer_Click()
    busca
End Sub

Private Sub cmbxCat_Click()
    lblCat.caption = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
    If cmbxInvent.ListIndex > -1 Then
        busca
    End If
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
End Sub

Private Sub cmbxInvent_Click()
    P_Status "Aguarde..."
    busca
    P_Status ""
End Sub

Function imprime(obj As Object) As Integer
    Dim sEnd As String
    Dim iCount As Integer
    Dim iPag As Integer
    Dim impr As impressao
    On Error GoTo trata_Erro
    If dtInvent.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que imprimir!", vbExclamation
        cmbxInvent.SetFocus
        Exit Function
    End If
    If MsgBox("Confirma emissão do Inventário " & cmbxInvent.Text & " ?", vbYesNo + vbQuestion) = vbNo Then
        Exit Function
    End If
    impr = callprinter(True, True)
    'Printer.PrintQuality = -2
    
    obj.ScaleMode = vbMillimeters
    'obj.Orientation = vbPRORLandscape
    Printer.FontName = "arial"
    Printer.fontsize = 10
    P_CentraImpressao obj, "INVENTÁRIO DE CONFERÊNCIA " & cmbxInvent.Text
    obj.FontName = "Arial"
    obj.fontsize = 9
    
    obj.FontBold = True
    P_ImprX obj, "Marca: " & cmbxMarca.Text, 0
    P_ImprX obj, "Categoria: " & cmbxCat.Text, 60, False
    P_ImprLinha obj
    P_ImprX obj, "Código", 0, True
    P_ImprX obj, "Marca", 18, True
    P_ImprX obj, "Produto", 28, True
    P_ImprX obj, "Ref.", 130
    P_ImprX obj, "Und", 145
    P_ImprX obj, "Estoque", 170, True, True
    
    P_ImprX obj, "Aferiçao", 175, False
    P_ImprLinha obj
    Printer.FontBold = False
    With dtInvent.Recordset
        .MoveFirst
        .Move impr.PagIni * 50
        
        Do While Not .EOF
            If iCount = 50 Then
                iCount = 0
                P_FimRel obj, constNome_sis & "-" & constVersao_Sis & " " & strNome_Emp, True, obj.Page + impr.PagIni
                Printer.FontBold = True
                P_ImprX obj, "Código", 0, True
                P_ImprX obj, "Marca", 18, True
                P_ImprX obj, "Produto", 28, True
                P_ImprX obj, "Ref.", 130
                P_ImprX obj, "Und", 145
                P_ImprX obj, "Estoque", 170, True, True
                P_ImprX obj, "Aferiçao", 175, False
                P_ImprLinha obj
                Printer.FontBold = False
            End If
            Printer.FontBold = False
            P_ImprX obj, FormatCodigo(!cod_pro), 0, True
            P_ImprX obj, F_BuscaUmCampo("Cod_mar", "tab_produtos", "cod_Pro='" & !cod_pro & "'"), 18, True
            P_ImprX obj, !nome_pro, 28, True
            P_ImprX obj, !codfab_pro, 130
            P_ImprX obj, !unidade_pro, 145
            P_ImprX obj, FormatVal(!estoque_pro - reservado_pro), 170, True, True
            P_ImprX obj, "________", 175, False
            iCount = iCount + 1
            
            .MoveNext
        Loop
    End With
    P_FimRel obj, constNome_sis & "-" & constVersao_Sis & " " & strNome_Emp, True, obj.Page + impr.PagIni & " de " & obj.Page + impr.PagIni
    Printer.EndDoc
    Exit Function
trata_Erro:
    MsgBox "Não foi possível concluir a impressão!" & Chr(13) & Err.Description, vbCritical
    
End Function
Function imprimeEst(obj As Object) As Integer
    Dim sEnd As String
    Dim iCount As Integer
    Dim iPag As Integer
    Dim impr As impressao
    On Error GoTo trata_Erro
    If dtInvent.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que imprimir!", vbExclamation
        cmbxInvent.SetFocus
        Exit Function
    End If
    If MsgBox("Confirma emissão do Inventário " & cmbxInvent.Text & " ?", vbYesNo + vbQuestion) = vbNo Then
        Exit Function
    End If
    impr = callprinter(True, True)
    Printer.PrintQuality = -2
    
    obj.ScaleMode = vbMillimeters
    'obj.Orientation = vbPRORLandscape
    Printer.FontName = "arial"
    Printer.fontsize = 10
    P_CentraImpressao obj, "INVENTÁRIO DE ESTOQUE " & cmbxInvent.Text
    obj.FontName = "Arial"
    obj.fontsize = 9
    
    obj.FontBold = True
    P_ImprX obj, "Marca: " & cmbxMarca.Text, 0
    P_ImprX obj, "Categoria: " & cmbxCat.Text, 60, False
    P_ImprLinha obj
    P_ImprX obj, "Código", 0, True
    P_ImprX obj, "Produto", 18, True
    P_ImprX obj, "Un.", 100, True
    P_ImprX obj, "E.E", 125, True, True
    P_ImprX obj, "C.E", 140, True, True
    P_ImprX obj, "T.E", 155, True, True
    P_ImprX obj, "E.R", 170, True, True
    P_ImprX obj, "C.R", 185, True, True
    P_ImprX obj, "T.R", 200, False, True
    
    P_ImprLinha obj
    Printer.FontBold = False
    With dtInvent.Recordset
        .MoveFirst
        .Move impr.PagIni * 50
        
        Do While Not .EOF
            If iCount = 50 Then
                iCount = 0
                P_FimRel obj, constNome_sis & "-" & constVersao_Sis & " " & strNome_Emp, True, obj.Page + impr.PagIni
                Printer.FontBold = True
                P_ImprX obj, "Código", 0, True
                P_ImprX obj, "Produto", 18, True
                P_ImprX obj, "Un.", 100, True
                P_ImprX obj, "E.E", 125, True, True
                P_ImprX obj, "C.E", 140, True, True
                P_ImprX obj, "T.E", 155, True, True
                P_ImprX obj, "E.R", 170, True, True
                P_ImprX obj, "C.R", 185, True, True
                P_ImprX obj, "T.R", 200, False, True
                P_ImprLinha obj
                Printer.FontBold = False
            End If
            Printer.FontBold = False
            P_ImprX obj, FormatCodigo(!cod_pro), 0, True
            
            P_ImprX obj, !nome_pro, 18, True
            P_ImprX obj, !unidade_pro, 100
            P_ImprX obj, FormatVal(!ef_pro), 125, True, True
            P_ImprX obj, FormatVal(!cf_pro), 140, True, True
            P_ImprX obj, FormatVal(!ef_pro) * !cf_pro, 155, True, True
            P_ImprX obj, FormatVal(!estoque_pro), 170, True, True
            P_ImprX obj, FormatVal(!custo_pro), 185, True, True
            P_ImprX obj, FormatVal(!custo_pro * !estoque_pro), 200, False, True
            iCount = iCount + 1
            
            .MoveNext
        Loop
    End With
    P_FimRel obj, constNome_sis & "-" & constVersao_Sis & " " & strNome_Emp, True, obj.Page + impr.PagIni & " de " & obj.Page + impr.PagIni
    Printer.EndDoc
    Exit Function
trata_Erro:
    MsgBox "Não foi possível concluir a impressão!" & Chr(13) & Err.Description, vbCritical
    
End Function

Private Sub cmndbtnLocalizar_Click()
    

    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "Select TI.*,TI.ESTOQUE_PRO as estoque,((TI.ESTOQUE_PRO)*TI.CUSTO_PRO) AS TOTAL,TM.st_pro,"
    strSQL = strSQL & "((TI.ef_pro)*TI.cf_pro) as totalf,preco_pro * estoque_pro as tv, "
    strSQL = strSQL & " TM.codfab_pro,TM.endereco_pro,TM.unidade_pro from tab_Invent as TI, TAB_produtos AS TM"
    
    strSQL = strSQL & " where TI.COD_pro = TM.COD_pro  and "
    If chkEstZer.Value = 0 Then
        strSQL = strSQL & "TI.ESTOQUE_PRO > 0 and "
    End If
    If chkNeg.Value = 1 Then
        strSQL = strSQL & "TI.ef_PRO >= 0 and "
    End If
    
    strSQL = strSQL & " DATAHORA_INV ='" & Format(cmbxInvent.Text, "YYYY-MM-DD-HH:MM:SS") & "' "
    If cmbxCat.Text <> "" Then
        strSQL = strSQL & " and TM.cod_cat='" & lblCat.caption & "' "
    End If
    If cmbxMarca.Text <> "" Then
        strSQL = strSQL & " and TM.cod_mar='" & lblMar.caption & "' "
    End If
    If cmbxSt.Text <> "" Then
        strSQL = strSQL & " and st_pro='" & Left(cmbxSt.Text, 1) & "' "
    End If
    If cmbxOrderBy.ListIndex = 0 Then
        strSQL = strSQL & " ORDER BY TM.cod_pro"
    ElseIf cmbxOrderBy.ListIndex = 1 Then
        strSQL = strSQL & " ORDER BY NOME_pro"
     ElseIf cmbxOrderBy.ListIndex = 2 Then
        strSQL = strSQL & " ORDER BY Endereco_pro,NOME_pro"
    ElseIf cmbxOrderBy.ListIndex = 3 Then
        strSQL = strSQL & "Order by codfab_pro "
    End If
    dtInvent.RecordSource = strSQL
    dtInvent.Refresh
    P_Total
End Sub

Private Sub cmbxMarca_Click()
    lblMar.caption = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxMarca.Text & "'")
    If cmbxInvent.ListIndex > -1 Then
        busca
    End If
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnAnalise_Click()
    Dim iCount As Integer
    
    
    iCount = imprimeEst(Printer)
End Sub

Private Sub cmbxSt_Click()
    busca
End Sub

Private Sub cmndbtnEditaItem_Click()
    Dim strSQL As String
    On Error Resume Next
    Dim i As Integer
    frmEdiItemInvent.lblInvent.caption = cmbxInvent.Text
    frmEdiItemInvent.lblNomePro.caption = dtInvent.Recordset!nome_pro
    frmEdiItemInvent.lblNomePro.Tag = dtInvent.Recordset!cod_pro
    frmEdiItemInvent.txtbxCusto.Text = dtInvent.Recordset!cf_pro
    frmEdiItemInvent.txtbxQtd.Text = dtInvent.Recordset!ef_pro
    
    frmEdiItemInvent.Show 1
    
    busca
End Sub

Private Sub cmndbtnExcl_Click()
    On Error GoTo trata_Erro
    Dim strSQL As String
    If strExclVenda_usu <> "S" Then
        MsgBox "Você não tem permissão para esta Operação!", vbCritical
        Exit Sub
    End If
    With dtInvent.Recordset
        If MsgBox("Deseja excluir o inventário:" & !DataHora_Inv & "?", vbYesNo + vbCritical) = vbYes Then
            strSQL = "delete from tab_invent where datahora_inv=" & F_GravaDataTempo(!DataHora_Inv) & " "
            Banco.Execute strSQL
            P_BuscaInvent
            busca
        End If
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnExclItem_Click()
    Dim strSQL As String
    On Error Resume Next
    If MsgBox("Confirma exclusão do item " & dtInvent.Recordset!nome_pro & "?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = "delete from tab_invent where cod_pro='" & dtInvent.Recordset!cod_pro & "' "
        strSQL = strSQL & " and datahora_inv=" & F_GravaDataTempo(dtInvent.Recordset!DataHora_Inv)
        Banco.Execute strSQL
    End If
    busca
End Sub

Private Sub cmndbtnHtml_Click()
    P_DbgrdToHtml "POSIÇÃO DOS PRODUTOS EM " & cmbxInvent.Text, dbgrdInvent, dtInvent
End Sub

Private Sub cmndbtnImprimir_Click()
    Dim iCount As Integer

    
    iCount = imprime(Printer)
    On Error Resume Next
    'frmVisualiza.Show 1
End Sub

Private Sub cmndbtnNovo_Click()
     Dim sDataHora As String
     Dim strSQL As String
    sDataHora = CStr(Now)
    If F_BuscaUmCampo("count(datahora_inv)", "tab_invent", "datahora_Inv='" & _
                       Format(sDataHora, "yyyy-mm-dd hh:mm:dd") & "'") > "0" Then
        MsgBox "Já Existe um Inventário com esta data e hora!", vbCritical
        Exit Sub
    ElseIf F_BuscaUmCampo("count(datahora_inv)", "tab_invent", "left(datahora_Inv,10)='" & _
                       Format(sDataHora, "yyyy-mm-dd") & "'") > "0" Then
        
        If MsgBox("Já Existe um Inventário com a data de hoje, deseja cancelá-lo?!", vbCritical + vbYesNo) = vbYes Then
            strSQL = "delete from tab_invent where left(datahora_inv,10)=" & F_GravaData(Date)
            P_Status "Excluindo..."
            Banco.Execute strSQL
            P_Status ""
        Else
            MsgBox "Não é possível continuar!", vbExclamation
            Exit Sub
        End If
        
    
        
    End If
    If MsgBox("Confirma geração do Inventário de " & sDataHora & " ?", vbYesNo + vbQuestion) = vbYes Then
            P_Status "Gerando Inventários..."
            P_GravaInvent sDataHora
            P_BuscaInvent
            MsgBox "Inventário Gerado com Sucesso!", vbInformation
            P_Status
        End If
End Sub
Sub P_GravaInvent(sDataHora As String)
    Dim strSQL As String
    strSQL = "Insert Into tab_invent(DataHora_Inv,Cod_Pro,Estoque_pro,"
    strSQL = strSQL & "Reservado_Pro,Unidade_Pro,Custo_Pro,"
    strSQL = strSQL & "Preco_Pro,Nome_Pro,ef_pro,cf_pro) Select '" & Format(sDataHora, "YYYY-MM-DD-HH:MM:SS") & "',"
    strSQL = strSQL & "TM.COD_Pro, TM.ER_Pro,0,"
    strSQL = strSQL & "TM.UNIDADE_pro,TM.CR_Pro,"
    strSQL = strSQL & "TM.PV_Pro,TM.NOME_Pro,ef_pro,cf_pro FROM TAB_produtos AS TM  "
    
    Banco.Execute strSQL
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    P_BuscaInvent
End Sub


Private Sub dbgrdInvent_DblClick()
    cmndbtnEditaItem_Click
End Sub

Private Sub dbgrdInvent_GotFocus()
    dbgrdInvent.MarqueeStyle = 3
End Sub

Private Sub dbgrdInvent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEditaItem_Click
    End If
    If KeyCode = vbKeyDelete Then
        cmndbtnExclItem_Click
    End If
    
End Sub

Private Sub dbgrdInvent_LostFocus()
    dbgrdInvent.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    
    resizeForm Me
    
    cmbxOrderBy.ListIndex = 0
    P_BuscaInvent
    bActive = True
End Sub

Private Sub Form_Load()
    bActive = False
    P_CentraForm Me
    dtInvent.Connect = Banco.Connect
    P_PreencheCateg cmbxCat
    P_PreencheMarca cmbxMarca
    With dbgrdInvent
        If Left(strNomeResp_Emp, 10) = strNome_Usu Or strNome_Usu = "ALISON" Then
            .Columns(8).Visible = True
            .Columns(9).Visible = True
            .Columns(10).Visible = True
            .Columns(8).DataField = "ef_pro"
            .Columns(9).DataField = "cf_pro"
            .Columns(10).DataField = "totalf"
            .Columns(8).caption = "Est.F"
            .Columns(9).caption = "Custo F"
            .Columns(10).caption = "Total F"
        Else
            .Columns(8).Visible = False
            .Columns(9).Visible = False
            .Columns(10).Visible = False
        End If
    End With

End Sub

Private Sub Form_Resize()
    dbgrdInvent.Height = pcTotal.Top - dbgrdInvent.Top - 100
    dbgrdInvent.Width = Me.Width - dbgrdInvent.Left - 200
End Sub

