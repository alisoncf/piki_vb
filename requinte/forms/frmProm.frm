VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmProm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtos em Promoção"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "frmProm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9270
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   4320
      TabIndex        =   11
      Top             =   5280
      Width           =   2415
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "&Retirar"
         Height          =   855
         Left            =   1250
         Picture         =   "frmProm.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Alterar"
         Height          =   855
         Left            =   120
         Picture         =   "frmProm.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frPro 
      Caption         =   "&Adicionar Produto:"
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   5280
      Width           =   4215
      Begin VB.OptionButton optnR 
         Caption         =   "&R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnOk 
         Caption         =   "O&k"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         Picture         =   "frmProm.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbxPreco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Text            =   "0,00"
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optnN 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optnC 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optnBar 
         Caption         =   "&B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxCodBarra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblCusto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   20
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo:"
         Height          =   195
         Left            =   1800
         TabIndex        =   19
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblPreco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   18
         Top             =   900
         Width           =   480
      End
      Begin VB.Label label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pr&omoção:"
         Height          =   195
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         Caption         =   "&Cod. Barra:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Tag             =   "B"
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Data dtPro 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   6840
      TabIndex        =   14
      Top             =   5280
      Width           =   2415
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   855
         Left            =   120
         Picture         =   "frmProm.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   855
         Left            =   1250
         Picture         =   "frmProm.frx":096A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmProm.frx":0AB4
      Height          =   5295
      Left            =   0
      OleObjectBlob   =   "frmProm.frx":0AC8
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmProm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    With dtPro.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser editado!", vbExclamation
            Exit Sub
        End If
        txtbxCodBarra.Text = !nome_pro
        txtbxCodBarra.Tag = !cod_pro
        frPro.caption = !nome_pro
        lblCusto.caption = FormatVal(!c2_pro)
        lblPreco.caption = FormatVal(!p2_pro)
        txtbxPreco.Text = FormatVal(!p3_pro)
        txtbxPreco.SetFocus
    End With
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    With dtPro.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser excluído!", vbExclamation
            Exit Sub
        End If
        If MsgBox("Confirma exclusão do produto da Promoção?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "update tab_composicao set prom_pro='N' where cod_pro='"
            strSQL = strSQL & !cod_pro & "' "
            banco.Execute strSQL
            busca
        End If
    End With
End Sub

Private Sub cmndbtnImpr_Click()
    With dtPro.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser impresso!", vbExclamation
            Exit Sub
        End If
    End With
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    Dim iCount As Integer
    P_CabRelTimbr "Produtos em Promoção em " & Date, obj
    P_ImprLinha obj
    P_ImprX obj, "Código", 0
    P_ImprX obj, "Produto", 20
    P_ImprX obj, "Estoque", 150, True, True
    P_ImprX obj, "Pr.Venda", 170, True, True
    P_ImprX obj, "Promocional", 190, False, True
    P_ImprLinha obj
    With dtPro.Recordset
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj, constNome_sis
                P_ImprX obj, "Código", 0
                P_ImprX obj, "Produto", 20
                P_ImprX obj, "Estoque", 150, True, True
                P_ImprX obj, "Pr.Venda", 170, True, True
                P_ImprX obj, "Promocional", 190, False, True
                P_ImprLinha obj
            End If
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome_pro, 20
            P_ImprX obj, !estoque, 150, True, True
            P_ImprX obj, FormatVal(!p2_pro), 170, True, True
            P_ImprX obj, FormatVal(!p3_pro), 190, False, True
            .MoveNext
        Loop
        P_RodapeVis obj, constNome_sis, "(F)"
    End With
End Sub
Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    If txtbxCodBarra.Tag = "" Then
        MsgBox "Escolha primeiro o produto!", vbExclamation
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
    If txtbxPreco.Text = "0,00" Then
        MsgBox "Informe um preço promocional!", vbExclamation
        txtbxPreco.SetFocus
        Exit Sub
    End If
    If MsgBox("Deseja realmente adicionar o Produto na promoção?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = "update tab_composicao set p3_pro=" & F_GravaDec(txtbxPreco.Text)
        strSQL = strSQL & ",prom_pro='S' "
        strSQL = strSQL & " where cod_pro='" & txtbxCodBarra.Tag & "' "
        banco.Execute strSQL
        busca
        txtbxCodBarra.SetFocus
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dtPro.Connect = banco.Connect
    busca
    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tp.nome_pro,tp.cod_pro,tp.codbarra_pro,tc.*,(tc.e2_pro-tc.r_pro) as estoque,"
    strSQL = strSQL & "tm.nome_mar,tt.nome_cat from tab_produtos as tp,"
    strSQL = strSQL & " tab_composicao as tc, tab_marcas as tm, "
    strSQL = strSQL & " tab_categorias as tt where tp.cod_pro=tc.cod_pro "
    strSQL = strSQL & " and tp.cod_cat=tt.cod_cat and tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " and tc.prom_pro ='S' order by nome_pro "
    dtPro.RecordSource = strSQL
    dtPro.Refresh
End Sub

Private Sub optnBar_Click()
    lblExpr.caption = "Cód &Barra:"
End Sub

Private Sub optnC_Click()
    lblExpr.caption = "Có&digo:"
End Sub

Private Sub optnN_Click()
    lblExpr.caption = "&Nome:"
End Sub

Sub buscaProd()
    Dim p As Produto
    
    
    If optnBar.Value Then
        p = F_BuscaProdSaida(txtbxCodBarra.Text, "B")
        txtbxCodBarra.Text = p.codbarra_pro
    ElseIf optnC.Value Then
        p = F_BuscaProdSaida(txtbxCodBarra.Text, "C")
        txtbxCodBarra.Text = p.cod_pro
    ElseIf optnR.Value Then
        p = F_BuscaProdSaida(txtbxCodBarra.Text, "R")
        txtbxCodBarra.Text = p.cod_pro
    Else
        p = F_BuscaProdSaida(txtbxCodBarra.Text, "N")
        txtbxCodBarra.Text = p.nome_pro
    End If
    If p.cod_pro = "" Then
        txtbxCodBarra.Tag = ""
        Exit Sub
    End If
    txtbxPreco.Tag = p.p2_pro
    txtbxPreco.Text = FormatVal(p.p3_pro)
    frPro.caption = p.nome_pro
    lblCusto.caption = FormatVal(p.c2_pro)
    lblPreco.caption = FormatVal(p.p2_pro)
    txtbxCodBarra.Tag = p.cod_pro
    txtbxPreco.SetFocus
    
End Sub

Private Sub txtbxCodBarra_GotFocus()
    P_SelecTexto txtbxCodBarra
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        buscaProd
    End If
End Sub

Private Sub txtbxPreco_GotFocus()
    P_SelecTexto txtbxPreco
End Sub

Private Sub txtbxPreco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPreco_LostFocus()
    P_ValorLostFocus txtbxPreco
End Sub
