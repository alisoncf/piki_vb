VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelAtvCliProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Atividade de Clientes por Produto"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmRelAtvCliProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9855
   Begin VB.Frame Frame3 
      Caption         =   "Opções de Busca:"
      Height          =   735
      Left            =   4440
      TabIndex        =   24
      Top             =   840
      Width           =   5295
      Begin VB.OptionButton optnAna 
         Caption         =   "&Analítico"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optnSint 
         Caption         =   "&Sintético"
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   2280
      TabIndex        =   5
      Top             =   0
      Width           =   1935
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmRelAtvCliProd.frx":014A
         Left            =   120
         List            =   "frmRelAtvCliProd.frx":015A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   9705
      TabIndex        =   16
      Top             =   6120
      Width           =   9735
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblQtd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   600
         TabIndex        =   21
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   1560
         TabIndex        =   20
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   2160
         TabIndex        =   19
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final"
         Height          =   195
         Left            =   3000
         TabIndex        =   18
         Top             =   120
         Width           =   330
      End
      Begin VB.Label lblFinal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   3720
         TabIndex        =   17
         Top             =   120
         Width           =   390
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelAtvCliProd.frx":0190
      Height          =   4335
      Left            =   0
      OleObjectBlob   =   "frmRelAtvCliProd.frx":01A1
      TabIndex        =   23
      Top             =   1680
      Width           =   9735
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   8400
      Picture         =   "frmRelAtvCliProd.frx":13D2
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   100
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   735
      Left            =   7080
      Picture         =   "frmRelAtvCliProd.frx":1C9C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   100
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnGraf 
      Caption         =   "&Gráfico"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5760
      Picture         =   "frmRelAtvCliProd.frx":1FA6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   100
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   4440
      Picture         =   "frmRelAtvCliProd.frx":23E8
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   100
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   4215
      Begin VB.TextBox txtbxNomeCli 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3735
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   285
         Left            =   3840
         Picture         =   "frmRelAtvCliProd.frx":26F2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmRelAtvCliProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxOrder_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub cmndbtnBuscar_Click()
    If optnAna.Value Then
        ana
    Else
        sint
    End If
    soma

End Sub
Sub ana()
    Dim strSQL As String
    
    strSQL = "select ti.qtd_ite,tp.nome_pro,tp.cod_pro,tq.p2_pro,"
    strSQL = strSQL & "ti.final_ite,(ti.qtd_ite*ti.unit_ite) as total,tv.*   "
    
    strSQL = strSQL & " from tab_vendas as tv, tab_venfec as tf, tab_clientes as tc,"
    strSQL = strSQL & " tab_itvendas as ti, tab_produtos as tp, tab_composicao as tq   "
    strSQL = strSQL & " where tc.cod_cli=tv.cod_cli and tv.num_ven=tf.num_ven "
    strSQL = strSQL & " and tq.cod_pro=tp.cod_pro"
    strSQL = strSQL & " and ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        
    End If
    strSQL = strSQL & " and situacao_ven='4' "
    strSQL = strSQL & " and tc.cod_cli='" & txtbxNomeCli.tag & "' "
    
    cmndbtnGraf.Enabled = True
    If cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by nome_pro "
        cmndbtnGraf.Enabled = False
    ElseIf cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by total desc "
    ElseIf cmbxOrder.ListIndex = 2 Then
        strSQL = strSQL & " order by final_ite desc "
    ElseIf cmbxOrder.ListIndex = 3 Then
        strSQL = strSQL & " order by qtd_ite desc "
    End If
    dt.RecordSource = strSQL
    dt.Refresh
End Sub
Sub sint()
    Dim strSQL As String
    
    strSQL = "select sum(ti.qtd_ite) as qtd,tp.*,"
    strSQL = strSQL & "sum(ti.final_ite) as final,sum(ti.qtd_ite*ti.unit_ite) as total   "
    
    strSQL = strSQL & " from tab_vendas as tv, tab_venfec as tf, tab_clientes as tc,"
    strSQL = strSQL & " tab_itvendas as ti, tab_produtos as tp "
    strSQL = strSQL & " where tc.cod_cli=tv.cod_cli and tv.num_ven=tf.num_ven "
    
    strSQL = strSQL & " and ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        
    End If
    strSQL = strSQL & " and situacao_ven='4' "
    strSQL = strSQL & " and tc.cod_cli='" & txtbxNomeCli.tag & "' "
        
    strSQL = strSQL & " group by tp.cod_pro  "
    
    
    cmndbtnGraf.Enabled = True
    If cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by nome_pro "
        cmndbtnGraf.Enabled = False
    ElseIf cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by total desc "
    ElseIf cmbxOrder.ListIndex = 2 Then
        strSQL = strSQL & " order by final desc "
    ElseIf cmbxOrder.ListIndex = 3 Then
        strSQL = strSQL & " order by qtd desc "
    End If
    dt.RecordSource = strSQL
    dt.Refresh
End Sub
Sub soma()
    Dim cFinal As Currency
    Dim cTotal As Currency
    Dim cQtd As Currency
    On Error GoTo trata_erro
    With dt.Recordset
        Do While Not .EOF
            If optnSint.Value Then
                cFinal = cFinal + !final
                cTotal = cTotal + !total
                cQtd = cQtd + !qtd
            Else
                cFinal = cFinal + !final_ite
                cTotal = cTotal + !total
                cQtd = cQtd + !qtd_ite
            End If
            .MoveNext
        Loop
        
        lblTotal.caption = FormatVal(cTotal)
        lblQtd.caption = cQtd
        lblFinal.caption = FormatVal(cFinal)
        
        If .RecordCount > 0 Then .MoveFirst
    End With
    Exit Sub
trata_erro:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnGraf_Click()
    On Error GoTo trata_erro
    If cmbxOrder.ListIndex = 1 Then
        grafico dt.Recordset!final
    
    ElseIf cmbxOrder.ListIndex = 2 Then
        grafico dt.Recordset!avista
    ElseIf cmbxOrder.ListIndex = 3 Then
        grafico dt.Recordset!aprazo
    ElseIf cmbxOrder.ListIndex = 4 Then
        grafico dt.Recordset!finan
    ElseIf cmbxOrder.ListIndex = 5 Then
        grafico dt.Recordset!desconto
    ElseIf cmbxOrder.ListIndex = 6 Then
        grafico dt.Recordset!qtd
    End If
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub grafico(f As Field)
    Dim b As BlChart
    Dim iCount As Integer
    Dim cRest As Currency
    Dim iRows As Integer
    
        With dt.Recordset
            .MoveFirst
            iRows = .RecordCount
            If iRows > 9 Then
                iRows = 9
                b.Count = 10 + 1
            Else
                b.Count = iRows
            End If
            
            For iCount = 0 To iRows - 1
                
                b.val(iCount).caption = !nome_pro
                b.val(iCount).Value = f
                 
                .MoveNext
            Next
            Do While Not .EOF
                cRest = cRest + f
                
                .MoveNext
            Loop
            If cRest > 0 Then
                b.val(iRows + 1).caption = "RESTANTE"
                b.val(iRows + 1).Value = cRest
            End If
        End With
        
    
    'P_Graph b, 1, True
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
    P_CabRelTimbr Me.caption & " de " & txtbxDataIni.Text & " até  " & txtbxdataFin.Text, obj
    P_ImprX obj, "Cód", 0
    P_ImprX obj, "Cliente", 15
    P_ImprX obj, "Final", 115, True, True
    P_ImprX obj, "À Vista", 135, True, True
    P_ImprX obj, "A Prazo", 155, True, True
    P_ImprX obj, "Financeira", 175, True, True
    P_ImprX obj, "Desconto", 195, True, True
    P_ImprX obj, "Qtd", 196, False
    P_ImprLinha obj
    With dt.Recordset
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj, constNome_sis, "Continua..."
                P_ImprX obj, "Cód", 0
                P_ImprX obj, "Cliente", 15
                P_ImprX obj, "Final", 115, True, True
                P_ImprX obj, "À Vista", 135, True, True
                P_ImprX obj, "A Prazo", 155, True, True
                P_ImprX obj, "Financeira", 175, True, True
                P_ImprX obj, "Desconto", 195, True, True
                P_ImprX obj, "Qtd", 196, False
                P_ImprLinha obj
            End If
            P_ImprX obj, txtbxNomeCli.tag, 0
            P_ImprX obj, txtbxNomeCli.Text, 15
            P_ImprX obj, FormatVal(!final), 115, True, True
            P_ImprX obj, FormatVal(!avista), 135, True, True
            P_ImprX obj, FormatVal(!aprazo), 155, True, True
            P_ImprX obj, FormatVal(!finan), 175, True, True
            P_ImprX obj, FormatVal(!desconto), 195, True, True
            P_ImprX obj, !qtd, 198, False
            
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "Totais...", 0
        P_ImprX obj, lblFinal.caption, 115, True, True
        P_ImprX obj, lblAv.caption, 135, True, True
        P_ImprX obj, lblAp.caption, 155, True, True
        P_ImprX obj, lblFinan.caption, 175, True, True
        P_ImprX obj, lblDesc.caption, 195, True, True
        P_ImprX obj, lblQtd.caption, 198, False
        
        P_ImprLinha obj, "_"
        P_RodapeVis obj
    End With
    
    
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    dt.Connect = Banco.Connect
    P_CentraForm Me
    cmbxOrder.ListIndex = 0
End Sub

Private Sub optnAna_Click()
    With dbgrd
        .Columns(5).Visible = True
        .Columns(6).Visible = True
        .Columns(0).DataField = "Cod_pro"
        .Columns(1).DataField = "nome_pro"
        .Columns(2).DataField = "qtd_ite"
        .Columns(3).DataField = "total"
        .Columns(4).DataField = "final_ite"
        
        .Columns(5).DataField = "dataefet_ven"
        .Columns(6).DataField = "num_ven"
        .Columns(5).caption = "Data"
        .Columns(6).caption = "Venda"
    End With
End Sub

Private Sub optnSint_Click()
    With dbgrd
        .Columns(5).Visible = False
        .Columns(6).Visible = False
        .Columns(0).DataField = "Cod_pro"
        .Columns(1).DataField = "nome_pro"
        .Columns(2).DataField = "qtd"
        .Columns(3).DataField = "total"
        .Columns(4).DataField = "final"
    End With
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

