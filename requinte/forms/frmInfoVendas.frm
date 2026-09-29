VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelInfoVendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Comissão"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   Icon            =   "frmInfoVendas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11400
   Begin VB.Frame Frame2 
      Caption         =   "Opções"
      Height          =   735
      Left            =   2640
      TabIndex        =   20
      Top             =   0
      Width           =   8775
      Begin VB.OptionButton optnVendedor 
         Caption         =   "Vendedor"
         Height          =   435
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optnTipoPag 
         Caption         =   "Tipo de Pagamento"
         Height          =   435
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optnVendaFormaP 
         Caption         =   "Vendas por Forma de Pagamento"
         Height          =   435
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optnVendaCargo 
         Caption         =   "Vendas por Cargo"
         Height          =   435
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
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
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "&Detalhes"
         Height          =   735
         Left            =   120
         Picture         =   "frmInfoVendas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3960
         Width           =   2175
      End
      Begin VB.ComboBox cmbxFunc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmInfoVendas.frx":052C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmInfoVendas.frx":08B6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmInfoVendas.frx":0CC3
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4800
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   825
         ScaleWidth      =   2145
         TabIndex        =   11
         Top             =   5640
         Width           =   2175
         Begin VB.Label lblComissao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   17
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comiss:"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   300
            Width           =   540
         End
         Begin VB.Label lblVendas 
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
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   840
            TabIndex        =   15
            Top             =   -15
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendas:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   45
            Width           =   585
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Devol:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   540
            Width           =   465
         End
         Begin VB.Label lblDev 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   12
            Top             =   480
            Width           =   480
         End
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico de Vendas"
         Height          =   735
         Left            =   120
         Picture         =   "frmInfoVendas.frx":0DB8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmInfoVendas.frx":0F02
      Height          =   5295
      Left            =   2640
      OleObjectBlob   =   "frmInfoVendas.frx":0F13
      TabIndex        =   18
      Top             =   960
      Width           =   8775
   End
End
Attribute VB_Name = "frmRelInfoVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxFunc_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxFunc.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    If optnVendaCargo.Value Then
        strSQL = "select tc.nome_car,sum(valor_ven) as Total  from tab_vendas as tv,"
        strSQL = strSQL & "tab_cargos As tc, tab_funcionarios As tf"

        strSQL = strSQL & " where "
        strSQL = strSQL & " tf.cod_fun = tv.abrevenda_usu "
        strSQL = strSQL & " and "
        strSQL = strSQL & " tf.cod_Car = tc.cod_Car"
    
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        End If
    
        strSQL = strSQL & " and situacao_ven=4"
        strSQL = strSQL & " group by tc.cod_car "
    ElseIf optnVendaFormaP Then
        strSQL = "select tf.nome_for,sum(valor_ven) as Total  from tab_vendas as tv,"
        strSQL = strSQL & "tab_formaspag As tf"

        strSQL = strSQL & " where "
        strSQL = strSQL & " tf.cod_for = tv.cod_for "
        
    
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        End If
    
        strSQL = strSQL & " and situacao_ven=4"
        strSQL = strSQL & " group by tf.cod_for "
    ElseIf optnTipoPag Then
        strSQL = "select tf.tipo_cai,sum(valor_ven) as Total  from tab_vendas as tv,"
        strSQL = strSQL & "tab_formaspag As tf"

        strSQL = strSQL & " where "
        strSQL = strSQL & " tf.cod_for = tv.cod_for "
        
    
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        End If
    
        strSQL = strSQL & " and situacao_ven=4"
        strSQL = strSQL & " group by tf.tipo_cai"
    ElseIf optnVendedor Then
        strSQL = "select tf.guerra_fun,sum(valor_ven) as Total  from tab_vendas as tv,"
        strSQL = strSQL & "tab_cargos As tc, tab_funcionarios As tf"

        strSQL = strSQL & " where "
        strSQL = strSQL & " tf.cod_fun = tv.abrevenda_usu "
        strSQL = strSQL & " and "
        strSQL = strSQL & " tf.cod_Car = tc.cod_Car"
    
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        End If
    
        strSQL = strSQL & " and situacao_ven=4"
        strSQL = strSQL & " group by tf.cod_fun"
    
    End If
    dt.RecordSource = strSQL
    dt.Refresh
    dbgrd.ClearFields
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    With dt.Recordset
        Do While Not .EOF
            cTotal = cTotal + !total
            .MoveNext
        Loop
        lblVendas.caption = FormatVal(cTotal)
        dt.Refresh
    End With
    
End Sub

Private Sub cmndbtnImprimir_Click()
    imprHTML
    Exit Sub
    imprime frmVisualiza.Pic
    
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprHTML()
    Dim sNomeArq As String
    Dim sCodFun As String
    Dim cTotal As Currency
    Dim cTotalComiss As Currency
    Dim cTotalTotalComiss As Currency
    Dim cTotalTotal As Currency
    sNomeArq = F_AbreHtml("rel_comiss")
    P_HtmlAbreHead
    P_HtmlAbreTitle 1, "Relatório de Comissões", True
    P_HtmlFechaHead
    P_HtmlAbreBody
    P_HtmlAbreTable 1, 1
    
    With dt.Recordset
        .MoveFirst
        Do While Not .EOF
            If sCodFun <> !cod_fun Then
                If .AbsolutePosition > 0 Then
                    P_HtmlAbreLinha
                        P_HtmlAbreColuna 1, "Total"
                        P_HtmlAbreColuna 1, FormatVal(cTotal)
                        P_HtmlAbreColuna 1, FormatVal(cTotalComiss)
                        P_HtmlAbreColuna 1, "."
                    P_HtmlFechaLinha
                End If
                
                sCodFun = !cod_fun
                cTotal = 0
                cTotalComiss = 0
                P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, !guerra_fun, , True, , "4"
                P_HtmlFechaLinha
                P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, "Tipo Pag."
                    P_HtmlAbreColuna 1, "Vlr Vendido."
                    P_HtmlAbreColuna 1, "Comissão"
                    P_HtmlAbreColuna 1, "%"
                P_HtmlFechaLinha
            End If
            P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, !tipo_cai
                    P_HtmlAbreColuna 1, FormatVal(!total)
                    P_HtmlAbreColuna 1, FormatVal(!comiss)
                    
                    P_HtmlAbreColuna 1, !comiss_car
            P_HtmlFechaLinha
            
            cTotal = cTotal + !total
            cTotalComiss = cTotalComiss + !comiss
            cTotalTotal = cTotalTotal + !total
            cTotalTotalComiss = cTotalTotalComiss + !comiss
            .MoveNext
        Loop
        
    End With
    
    P_HtmlAbreLinha
        P_HtmlAbreColuna 1, "Soma", , True, , "4"
    P_HtmlFechaLinha
    P_HtmlAbreLinha
        P_HtmlAbreColuna 1, "Geral"
        P_HtmlAbreColuna 1, FormatVal(cTotalTotal)
        P_HtmlAbreColuna 1, FormatVal(cTotalTotalComiss)
        P_HtmlAbreColuna 1, "."
    P_HtmlFechaLinha
    
    
    P_HtmlFechaTable 1
    P_HtmlFechaBody
    P_FechaHtml , True, sNomeArq
    
    
    
    
End Sub
    
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub
Sub imprime(obj As Object)
    P_CabRelTimbr "Relatório de Comissões de " & txtbxDataIni.Text & " até " & txtbxdataFin.Text, obj
    Dim sCodFun As String
    P_ImprLinha obj, "-"
    
    
    
    With dt.Recordset
        .MoveFirst
        Do While Not .EOF
            If !cod_fun <> sCodFun Then
                sCodFun = !cod_fun
                P_ImprX obj, "Nome", 0
                P_ImprLinha obj, "-"
                P_ImprX obj, "TipoCai", 10
                P_ImprX obj, "Vendido", 70, True, True
                P_ImprX obj, "Comissão", 100, True, True
                P_ImprX obj, "Perc.", 115, False, True
                P_ImprLinha obj, "-"
            End If
            
            P_ImprX obj, !guerra_fun, 10
            P_ImprX obj, FormatVal(!total), 70, True, True
            P_ImprX obj, FormatVal(!comiss), 100, True, True
            P_ImprX obj, FormatVal(!total / CCur(lblVendas.caption) * 100), 115, False, True

            .MoveNext
        Loop
        P_RodapeVis obj, constNome_sis
    End With
    
    
End Sub
Private Sub Form_Load()
    dt.Connect = Banco.Connect
    P_CentraForm Me
    P_PreencheFunc cmbxFunc, 1
    
    
    If strVisrelfin_usu <> "S" Then
        P_LocalizaOperacao strNome_Usu, cmbxFunc
        
    End If
    
End Sub

Private Sub optnTipoPag_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub optnVendaCargo_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub optnVendaFormaP_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub optnVendedor_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
