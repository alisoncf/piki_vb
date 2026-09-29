VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmListaVales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vales de Funcionários"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico de Vendas"
         Height          =   735
         Left            =   120
         Picture         =   "frmListaVales.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3120
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
         TabIndex        =   8
         Top             =   5640
         Width           =   2175
         Begin VB.Label lblVendas 
            Alignment       =   2  'Center
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
            TabIndex        =   9
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmListaVales.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmListaVales.frx":023F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmListaVales.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbxFunc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Detalhes"
         Height          =   735
         Left            =   120
         Picture         =   "frmListaVales.frx":09D6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   13
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   870
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
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmListaVales.frx":0AC0
      Height          =   6615
      Left            =   2520
      OleObjectBlob   =   "frmListaVales.frx":0AD1
      TabIndex        =   14
      Top             =   120
      Width           =   9375
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   8280
      TabIndex        =   15
      Top             =   6720
      Width           =   3615
      Begin VB.OptionButton optnGeral 
         Caption         =   "Geral"
         Height          =   195
         Left            =   1680
         TabIndex        =   17
         Top             =   120
         Width           =   1695
      End
      Begin VB.OptionButton optnDet 
         Caption         =   "Detalhado"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmListaVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxFunc_Click()
    cmbxFunc.tag = F_BuscaUmCampo("cod_fun", "tab_funcionarios", "guerra_fun='" & cmbxFunc.Text & "'")
    
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
    
    If cmbxFunc.Text <> strNome_Usu Then
        If strVisrelest_usu <> "S" Then
            MsgBox "Você não pode visualizar uma comissão de outra pessoa!", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
        strSQL = "select distinct tca.nome_car,tc.tipo_cai,sum(tc.valor_cai) as total,tc.data_cai,tc.historico_cai,"
        strSQL = strSQL & " tf.cod_fun,tf.guerra_fun,tf.nome_fun from tab_caixa as tc,tab_cargos as tca, "
        strSQL = strSQL & " tab_valescai as tv, tab_funcionarios as tf where tc.seq_cai=tv.seq_cai "
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and tc.data_cai>=" & F_PesqData(txtbxDataIni.Text)
        End If
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and tc.data_cai<=" & F_PesqData(txtbxdataFin.Text)
        End If
        
        
        
        
        strSQL = strSQL & " and tca.cod_car=tf.cod_car"
        strSQL = strSQL & " and tf.cod_fun=tv.cod_fun "
        If cmbxFunc.Text <> "" Then
            strSQL = strSQL & " and tv.cod_fun='" & cmbxFunc.tag & "' "
        End If
        
        
        
        If optnDet.Value Then
            strSQL = strSQL & " group by tf.cod_fun order by tf.guerra_fun,data_cai "
            dbgrd.Columns(3).Visible = True
            dbgrd.Columns(4).Visible = True
        Else
            strSQL = strSQL & " group by tf.cod_fun order by guerra_fun"
            dbgrd.Columns(3).Visible = False
            dbgrd.Columns(4).Visible = False
            
        End If
    
    
    dt.RecordSource = strSQL
    dt.Refresh
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
    P_DbgrdToHtml "Lista de Vales", dbgrd, dt, "De:" & txtbxDataIni.Text & " Até: " & txtbxdataFin.Text, "2", "Funcionário: " & cmbxFunc.Text
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
    P_HtmlAbreTable 1, 1, 700
    
    P_HtmlAbreLinha
    P_HtmlAbreColuna 1, "Relatório de " & IIf(optnVendas.Value, "Comissões", "Diferenças de Trocas") & ". Período de " & txtbxDataIni.Text & " Até " & txtbxdataFin.Text, , True, , "4"
    P_HtmlFechaLinha
    
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
                    P_HtmlAbreColuna 1, !nome_fun & " (Cargo: " & !nome_car & ")", , True, , "4"
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
    P_PreencheVendedor cmbxFunc, 1
    
    
    If strVisrelfin_usu <> "S" Then
        P_LocalizaOperacao strNome_Usu, cmbxFunc
        cmbxFunc.Enabled = False
    End If
    
End Sub

Private Sub Option1_Click()
    dbgrd.caption = "Diferenças de Troca no Período"
    cmndbtnBuscar_Click
End Sub

Private Sub optnVendas_Click()
    dbgrd.caption = "Comissões de Vendas no Período"
    cmndbtnBuscar_Click
End Sub

Private Sub txtbxdatafin_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxdatafin_LostFocus()
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

