VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
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
      TabIndex        =   10
      Top             =   5280
      Width           =   2415
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "&Retirar"
         Height          =   855
         Left            =   1250
         Picture         =   "frmProm.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Alterar"
         Height          =   855
         Left            =   120
         Picture         =   "frmProm.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frPro 
      Caption         =   "&Adicionar Produto:"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
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
         TabIndex        =   6
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
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbxPreco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   8
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxCodBarra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   2
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
         TabIndex        =   19
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo:"
         Height          =   195
         Left            =   1800
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   900
         Width           =   480
      End
      Begin VB.Label label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pr&omoção:"
         Height          =   195
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         Caption         =   "&Cod. Barra:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Tag             =   "B"
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   6840
      TabIndex        =   13
      Top             =   5280
      Width           =   2415
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   855
         Left            =   120
         Picture         =   "frmProm.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   1200
      Top             =   5760
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmProm.frx":0AB4
      Height          =   5175
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
        txtbxCodBarra.tag = !cod_pro
        frPro.caption = !nome_pro
        lblCusto.caption = FormatVal(!cr_pro)
        lblPreco.caption = FormatVal(!pv_pro)
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
            conn.Execute strSQL
            Busca
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
    If frmVisualiza.tag <> "" Then
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
                P_RodapeVis obj, strNomeSis
                P_ImprX obj, "Código", 0
                P_ImprX obj, "Produto", 20
                P_ImprX obj, "Estoque", 150, True, True
                P_ImprX obj, "Pr.Venda", 170, True, True
                P_ImprX obj, "Promocional", 190, False, True
                P_ImprLinha obj
            End If
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome_pro, 20
            P_ImprX obj, !Estoque, 150, True, True
            P_ImprX obj, FormatVal(!pv_pro), 170, True, True
            P_ImprX obj, FormatVal(!p3_pro), 190, False, True
            .MoveNext
        Loop
        P_RodapeVis obj, strNomeSis, "(F)"
    End With
End Sub
Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    If txtbxCodBarra.tag = "" Then
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
        strSQL = strSQL & " where cod_pro='" & txtbxCodBarra.tag & "' "
        conn.Execute strSQL
        Busca
        txtbxCodBarra.SetFocus
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dtPro.ConnectionString = strConnect
    Busca
    
End Sub
Sub Busca()
    Dim strSQL As String
    strSQL = "select tp.nome_pro,tp.cod_pro,tp.codbarra_pro,tc.*,(tc.er_pro-tc.r_pro) as estoque,"
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
        p = F_BuscaProdSaida(txtbxCodBarra.Text)
        txtbxCodBarra.Text = p.codbarra_pro
    ElseIf optnC.Value Then
        p = F_BuscaProdSaida(txtbxCodBarra.Text)
        txtbxCodBarra.Text = p.cod_pro
    ElseIf optnR.Value Then
        p = F_BuscaProdSaida(txtbxCodBarra.Text)
        txtbxCodBarra.Text = p.cod_pro
    Else
        p = F_BuscaProdSaida(txtbxCodBarra.Text)
        txtbxCodBarra.Text = p.nome_pro
    End If
    If p.cod_pro = "" Then
        txtbxCodBarra.tag = ""
        Exit Sub
    End If
    txtbxPreco.tag = p.pv_pro
    txtbxPreco.Text = FormatVal(p.p3_pro)
    frPro.caption = p.nome_pro
    lblCusto.caption = FormatVal(p.cr_pro)
    lblPreco.caption = FormatVal(p.pv_pro)
    txtbxCodBarra.tag = p.cod_pro
    txtbxPreco.SetFocus
    
End Sub

Private Sub txtbxCodBarra_GotFocus()
    ganhaFoco txtbxCodBarra
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        buscaProd
    End If
End Sub

Private Sub txtbxPreco_GotFocus()
    ganhaFoco txtbxPreco
End Sub

Private Sub txtbxPreco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxPreco_LostFocus()
    PerdeFoco txtbxPreco
End Sub
