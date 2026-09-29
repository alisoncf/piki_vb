VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelGiro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Giro Periódico"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11385
   Icon            =   "frmRelGiro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   11325
      TabIndex        =   25
      Top             =   6375
      Width           =   11385
      Begin VB.CommandButton cmndbtnExpor 
         Caption         =   "&Exportar"
         Height          =   685
         Left            =   3840
         Picture         =   "frmRelGiro.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   685
         Left            =   9360
         Picture         =   "frmRelGiro.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   685
         Left            =   5760
         Picture         =   "frmRelGiro.frx":0709
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   685
         Left            =   0
         Picture         =   "frmRelGiro.frx":0B16
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico"
         Enabled         =   0   'False
         Height          =   685
         Left            =   1920
         Picture         =   "frmRelGiro.frx":0EA0
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11355
      TabIndex        =   16
      Top             =   6000
      Width           =   11385
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Últ. Ent:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4920
         TabIndex        =   24
         Top             =   45
         Width           =   570
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Itens:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3720
         TabIndex        =   23
         Top             =   45
         Width           =   390
      End
      Begin VB.Label lblUltEnt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 Qdt: 000"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5640
         TabIndex        =   22
         Top             =   30
         Width           =   1545
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gerado:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   21
         Top             =   45
         Width           =   570
      End
      Begin VB.Label lblTotalGerado 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.000,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   600
         TabIndex        =   20
         Top             =   30
         Width           =   630
      End
      Begin VB.Label lblItens 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4200
         TabIndex        =   19
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblLiq 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.000,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2520
         TabIndex        =   18
         Top             =   30
         Width           =   630
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Líqüido:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         TabIndex        =   17
         Top             =   45
         Width           =   585
      End
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.ComboBox cmbxSub 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2760
         Width           =   1905
      End
      Begin VB.ComboBox cmbxSec 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2160
         Width           =   1905
      End
      Begin VB.ComboBox cmbxPrincipio 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1560
         Width           =   1905
      End
      Begin VB.Frame frGroup 
         Caption         =   "Agrupamento:"
         Height          =   1335
         Left            =   55
         TabIndex        =   11
         Top             =   3960
         Width           =   2055
         Begin VB.OptionButton optnSec 
            Caption         =   "por S&eção"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optnPro 
            Caption         =   "por &Produto."
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optnPri 
            Caption         =   "por P&rincípio"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optnSub 
            Caption         =   "por &S&ubseção"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbxOrderBY 
         Height          =   315
         ItemData        =   "frmRelGiro.frx":0FEA
         Left            =   120
         List            =   "frmRelGiro.frx":0FFD
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   945
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S&ubseção:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seção"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Principio"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelGiro.frx":103C
      Height          =   5535
      Left            =   2280
      OleObjectBlob   =   "frmRelGiro.frx":104D
      TabIndex        =   15
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmRelGiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmbxCat_Click()
    cmbxCat.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.text & "'")
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
End Sub

Private Sub cmbxMarca_Click()
    cmbxMarca.tag = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxMarca.text & "'")
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub

Private Sub cmbxOrderBY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxSec_Click()
    cmbxSec.tag = F_BuscaUmCampo("cod_sec", "tab_secoes", "nome_sec='" & cmbxSec.text & "'")
    P_PreencheSubSec cmbxSub, cmbxSec.tag
    busca
End Sub

Private Sub cmndbtnBuscar_Click()
    P_Status "Buscando..."
    busca
    soma
    P_Status ""
End Sub

Private Sub cmndbtnGraf_Click()
    ReDim b(max) As valChart
    ReDim row2(max) As valChart
    Dim iCount As Integer
    Dim cPrim As Currency
    Dim cSoma As Currency
    'Dim max As Integer
    Dim strSQL As String
    With adc.Recordset
        .MoveFirst
        If .RecordCount < 15 Then
            max = .RecordCount - 1
        Else
            max = 15
        End If
        ReDim b(max + 1) As valChart
        ReDim row2(max + 1) As valChart
        .MoveFirst
        For iCount = 0 To max
    
            b(iCount).caption = !nome & " "
            If cmbxOrderBY.ListIndex = 2 Then
                b(iCount).Value = !gerado
                cPrim = cPrim + !gerado
            ElseIf cmbxOrderBY.ListIndex = 3 Then
                b(iCount).Value = !liquido
                cPrim = cPrim + !liquido
            ElseIf cmbxOrderBY.ListIndex = 4 Then
                b(iCount).Value = !soma
                cPrim = cPrim + !soma
            End If
            .MoveNext
        Next
        row2(0).caption = "15 Primeiros"
        row2(0).Value = cPrim
        
        Do While Not .EOF
            If cmbxOrderBY.ListIndex = 2 Then
                cSoma = cSoma + !gerado
            ElseIf cmbxOrderBY.ListIndex = 3 Then
                cSoma = cSoma + !liquido
            ElseIf cmbxOrderBY.ListIndex = 4 Then
                cSoma = cSoma + !soma
            End If
            
            .MoveNext
        Loop
        
        ' b.val(max).caption = "Restante"
       ' b.val(max).Value = cSoma
       ' row2.val(1).caption = "Restante"
       ' row2.val(1).Value = cSoma
    End With
    If cmbxOrderBY.ListIndex = 2 Then
    
    ElseIf cmbxOrderBY.ListIndex = 3 Then
    
    ElseIf cmbxOrderBY.ListIndex = 4 Then
    
    End If
    P_Graph b, Int(max), 1, True, "Gráfico de Giro de Produtos(" & cmbxOrderBY.text & ")", "Giro Periódico", True
    
    
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_Erro
    If adc.Recordset.RecordCount <= 0 Then
        MsgBox "Não há o quer ser impresso!", vbExclamation
        Exit Sub
    End If
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag = "" Then
        Unload frmVisualiza
        Exit Sub
    End If
    'If MsgBox("Confirma Impressão do Relatório?", vbYesNo + vbQuestion) = vbYes Then
        imprime Printer
    'End If
    Exit Sub
trata_Erro:
    P_MostraErro Me.Name
End Sub
Sub imprime(obj As Object)
    Dim iItens As Integer
    Dim iPag As Integer
    Dim strSQL As String
    Dim cTotal As Currency
    Dim cQPage As Integer
    cQPage = 60
    P_CabRelTimbr "Relatório de Giro Periódico " & de & txtbxDataIni.text & _
        " Até " & txtbxDataFin.text, obj
    obj.FontName = "arial"
    obj.fontsize = 9
    P_ImprX obj, "Cód.:", 0
    P_ImprX obj, "Descrição", 20
    P_ImprX obj, "UN", 120
    P_ImprX obj, "QTD", 140, True, True
    P_ImprX obj, "Ger.", 160, True, True
    P_ImprX obj, "Líquido", 175, True, True
    P_ImprX obj, "Estoque", 192, True, True
    P_ImprX obj, "Expr(%)", 205, False, True
    
    P_ImprLinha obj
    
    With adc.Recordset
        
        .MoveFirst
        strSQL = CStr(.RecordCount / cQPage)
        If InStr(1, strSQL, ",") <> 0 Then
            iPag = Int(strSQL) + 1
        End If
        
        Do While Not .EOF
            iItens = iItens + 1
            If iItens = cQPage Then
                If TypeName(obj) = "Printer" Then
                    P_Rodape "Continua...", True, "Pág:" & IIf(TypeName(obj) = "Printer", obj.Page, "%") & " de " & iPag
                
                End If
                If obj.Name = "Pic" Then
                    frmVisualiza.imglst.ListImages.Add , , frmVisualiza.Pic.image
                    frmVisualiza.Pic.Cls
                   
                    frmVisualiza.Pic.Picture = Nothing
                    obj.CurrentY = yIni
                End If
                iItens = 0
                P_ImprX obj, "Cód.:", 0
                P_ImprX obj, "Descrição", 20
                P_ImprX obj, "UN", 120
                P_ImprX obj, "QTD", 140, True, True
                P_ImprX obj, "Ger.", 160, True, True
                P_ImprX obj, "Líquido", 175, True, True
                P_ImprX obj, "Estoque", 192, True, True
                P_ImprX obj, "Expr(%)", 205, False, True
                P_ImprLinha obj
            End If
            obj.FontName = "arial"
            obj.fontsize = 9
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome, 20
            P_ImprX obj, !unidade_pro, 120
            P_ImprX obj, FormatVal(!soma), 140, True, True
            P_ImprX obj, FormatVal(!gerado), 160, True, True
            P_ImprX obj, FormatVal(!liquido), 175, True, True
            P_ImprX obj, FormatVal(!er_pro), 192, True, True
            P_ImprX obj, FormatVal(!gerado * 100 / CCur(lblTotalGerado.caption)) & "%", 205, False, True
            cTotal = cTotal + !gerado * 100 / CCur(lblTotalGerado.caption)
           
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "Total...: " & lblTotalGerado.caption, 80, True
        'P_ImprX obj, lblTotal.caption, 130, False
        On Error Resume Next
        P_RodapeVis obj, strNome_Emp, "Pág: " & obj.Page & " de " & obj.Page
        If obj.Name = "Pic" Then
                    frmVisualiza.imglst.ListImages.Add , , frmVisualiza.Pic.image
                    frmVisualiza.Pic.Cls
                   
                    frmVisualiza.Pic.Picture = Nothing
                    obj.CurrentY = yIni
                End If
    End With
    
    
    obj.EndDoc
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub dt_Reposition()
    If bActive = True Then Exit Sub
    If adc.Recordset.RecordCount > 0 Then
        If cmbxOrderBY.ListIndex >= 2 Then
            cmndbtnGraf.Enabled = True
        Else
            cmndbtnGraf.Enabled = False
        End If
    Else
        cmndbtnGraf.Enabled = False
    End If
    BuscaEnt
    bActive = True
End Sub
Sub BuscaEnt()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error Resume Next
    strSQL = "select ti.*,tc.entrada_not from tab_cmp as tc, tab_itcmp as ti "
    strSQL = strSQL & " where tc.num_not=ti.num_not "
    strSQL = strSQL & " and tc.serie_not=ti.serie_not "
    strSQL = strSQL & " and tc.cod_for=ti.cod_for "
    strSQL = strSQL & " and ti.cod_pro='" & adc.Recordset!cod_pro & "' "
    strSQL = strSQL & " order by entrada_not desc "
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        lblUltEnt.caption = rPesq!entrada_not
        lblUltEnt.caption = lblUltEnt.caption & "-Qdt:" & FormatVal(rPesq!qtd_ite)
    Else
        lblUltEnt.caption = "Ult. Entr: Indisponível"
    End If
    rPesq.Close
End Sub

Private Sub Form_Activate()
    bative = True
End Sub

Private Sub Form_Load()
    adc.ConnectionString = Conn.ConnectionString
    
    cmbxOrderBY.ListIndex = 0
    P_CentraForm Me
    P_PreenchePrincipios cmbxPrincipio
    P_PreencheSecao cmbxSec
    
    txtbxDataFin.text = Date
    txtbxDataIni.text = Date
    
    
End Sub
Sub soma()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim cTotalGer As Currency
    Dim iCountPro As Single
    bActive = True
    With adc.Recordset
    
        Do While Not .EOF
           ' cTotal = cTotal + !Total
            cTotalGer = cTotalGer + !gerado
            cTotal = cTotal + !liquido
            iCountPro = iCountPro + !soma
            .MoveNext
            
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    
    
    lblTotalGerado.caption = FormatVal(cTotalGer)
    lblLiq.caption = FormatVal(cTotal)
    lblItens.caption = iCountPro
    
    bActive = False
    
End Sub
Sub busca()
    Dim strSQL As String
    
    
    If optnPro.Value = True Then
        strSQL = "select distinct tp.nome_pro as nome,tp.cod_pro as cod,"
    ElseIf optnPri.Value = True Then
        strSQL = "select distinct tpr.nome_pri as nome,tpr.cod_pri as cod,"
    ElseIf optnSec.Value = True Then
        strSQL = "select distinct  ts.nome_sec as nome,ts.cod_sec as cod,"
    ElseIf optnSub.Value = True Then
        strSQL = "select distinct tu.nome_sub as nome,tu.cod_sub as cod,"
    End If
    
    strSQL = strSQL & "sum(ti.qtd_ite) as soma,sum(er_pro) as er,sum(ef_pro) as ef, "
    strSQL = strSQL & " sum(ti.final_ite) as gerado, "
    strSQL = strSQL & " sum((ti.final_ite/ti.qtd_ite-tp.cr_pro)*ti.qtd_ite) as liquido,tp.* "
    strSQL = strSQL & " from tab_vendas as tv, tab_itvendas as ti,"
    strSQL = strSQL & " tab_produtos as tp,tab_estoque as te,"
    strSQL = strSQL & " tab_principios as tpr,tab_subsec as tu,tab_secoes as ts "
    strSQL = strSQL & " where  tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tp.cod_pro=te.cod_pro "
    strSQL = strSQL & " and tu.cod_sub=tp.cod_sub "
    strSQL = strSQL & " and ts.cod_sec=tu.cod_sec "
    strSQL = strSQL & " and tp.cod_pri=tpr.cod_pri "
        
    strSQL = strSQL & " and ti.num_ven=tv.num_ven and tv.situacao_ven = 4 "
    If cmbxPrincipio.ListIndex > -1 Then
        strSQL = strSQL & " and nome_pri='" & cmbxCat.text & "' "
    End If
    If cmbxSec.ListIndex > -1 Then
        strSQL = strSQL & " and nome_sec='" & cmbxSec.text & "' "
    End If
    If cmbxSub.ListIndex > -1 Then
        strSQL = strSQL & " and nome_sub='" & cmbxSub.text & "' "
    End If
    strSQL = strSQL & " and tv.dataefet_ven >=" & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and tv.dataefet_ven <=" & F_GravaData(txtbxDataFin.text)
    strSQL = strSQL & " and tipo_operacao=1"
    If optnPro.Value = True Then
        strSQL = strSQL & " group by ti.cod_pro "
    End If
    If optnPri.Value = True Then
        strSQL = strSQL & " group by tp.cod_pri,tp.cod_pro  "
    End If
    If optnSec.Value = True Then
        strSQL = strSQL & " group by ts.cod_sec,tp.cod_pro  "
    End If
    If optnSub.Value = True Then
        strSQL = strSQL & " group by tp.cod_sub,tp.cod_pro "
    End If
    strSQL = strSQL & " order by "
    If cmbxOrderBY.ListIndex = 0 Then
        strSQL = strSQL & "nome"
    ElseIf cmbxOrderBY.ListIndex = 1 Then
        strSQL = strSQL & "tp.cod_Pro "
    ElseIf cmbxOrderBY.ListIndex = 2 Then
        strSQL = strSQL & "gerado desc "
    ElseIf cmbxOrderBY.ListIndex = 3 Then
        strSQL = strSQL & "liquido desc"
    Else
        strSQL = strSQL & "soma desc"
    End If
    dt.RecordSource = strSQL
    dt.Refresh
End Sub

Private Sub optnCat_Click()
    frGroup.caption = "Agrupar por Categoria"
End Sub

Private Sub optnMar_Click()
    frGroup.caption = "Agrupar por Marca"
End Sub

Private Sub optnPro_Click()
    frGroup.caption = "Agrupar por Produto"
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


