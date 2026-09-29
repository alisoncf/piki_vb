VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEfetivaVen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Efetivação de Vendas"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9885
   Begin VB.Frame Frame1 
      Caption         =   "&Vendas a Efetivar:"
      Height          =   5655
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9855
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   4440
         Top             =   1200
      End
      Begin VB.Data dtVendas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   1260
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEfetivaVen.frx":0000
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "frmEfetivaVen.frx":0017
         TabIndex        =   7
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   9855
      Begin VB.CommandButton cmndbtnReabrir 
         Caption         =   "&Reabrir"
         Height          =   660
         Left            =   3960
         Picture         =   "frmEfetivaVen.frx":1420
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Nota Fiscal"
         Enabled         =   0   'False
         Height          =   660
         Left            =   2040
         Picture         =   "frmEfetivaVen.frx":156A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   660
         Left            =   5880
         Picture         =   "frmEfetivaVen.frx":1AF4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   660
         Left            =   7800
         Picture         =   "frmEfetivaVen.frx":1E7E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnEfetivar 
         Caption         =   "&Efetivar"
         Height          =   660
         Left            =   120
         Picture         =   "frmEfetivaVen.frx":1FC8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmEfetivaVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub busca()
     Dim strSQL As String
    
    
    
    
    
    strSQL = "select tc.cod_cli,tc.nome_cli,tu.guerra_fun,tv.num_ven,tv.data_ven,tv.situacao_ven,tv.valor_ven, "
    strSQL = strSQL & " tf.*,(tf.desconto_ven*tv.valor_ven)/100 as desconto"
    strSQL = strSQL & " from tab_clientes as tc, "
    strSQL = strSQL & " tab_vendas as tv,tab_venfec as tf, tab_funcionarios as tu "
    strSQL = strSQL & " where tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tv.situacao_ven='2' "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    strSQL = strSQL & " and tv.cod_fun=tu.cod_fun "
    If UCase(App.Title) = "PIKIBOX_CX" Then
        strSQL = strSQL & " and tf.cod_cai='" & strCod_Cai & "' "
    End If
    strSQL = strSQL & " order by tf.num_ven"
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
End Sub
    
Private Sub cmndbtnEfetivar_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    If dbgrd.Row = -1 Then
        MsgBox "selecione a Venda!", vbInformation
        dbgrd.SetFocus
        Exit Sub
    End If
    With dtVendas.Recordset
        If MsgBox("Deseja Realmente Efetivar a Venda?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
            
        End If
    End With

    strSQL = "select * from tab_venfec where num_ven='" & dtVendas.Recordset!num_ven & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
            frmFecharVenda.operacao = "venda"
            frmFecharVenda.lblvlrVenda.caption = FormatVal(!valorfinal_ven)
            frmFecharVenda.lblDesconto.caption = FormatVal(!desconto_ven)
            If !aprazo_ven > 0 Then
                frmFecharVenda.frAp.Enabled = True
                frmFecharVenda.frForm.Enabled = True
                frmFecharVenda.frVal.Enabled = True
                If Not IsNull(!cod_tip) Then
                    frmFecharVenda.cmbxFormas_Pag.Tag = !cod_tip
                End If
            End If
            If !tipo_ven = "C" Then
                frmFecharVenda.optnCheque.Value = True
            Else
                frmFecharVenda.otpnCart.Value = True
            End If
            frmFecharVenda.lblAprazo.caption = FormatVal(!aprazo_ven)
            frmFecharVenda.lblAprazo.Tag = FormatVal(!aprazo_ven)
            frmFecharVenda.lblAvista.caption = FormatVal(!avista_ven)
            frmFecharVenda.lblFinan.caption = FormatVal(!finan_ven)
            
            frmFecharVenda.lblvlrVenda.caption = FormatVal(!valorfinal_ven + (!valorfinal_ven * !desconto_ven / 100))
            frmFecharVenda.lblTotal.caption = FormatVal(!valorfinal_ven)
            frmFecharVenda.lblNome_cli.caption = dtVendas.Recordset!nome_Cli
            frmFecharVenda.lblNome_cli.Tag = dtVendas.Recordset!cod_cli
            frmFecharVenda.Tag = !num_ven
            frmFecharVenda.caption = "Fechar Venda: " & !num_ven & " de " & dtVendas.Recordset!nome_Cli
            
            frmFecharVenda.Show 1
            
            dtVendas.Refresh
            rPesq.Close
         
        End With
End Sub

Private Sub cmndbtnImpr_Click()
        P_Status "Gerando impressão... "
    P_ImprimeVenda frmVisualiza.Pic, dtVendas.Recordset!num_ven, False
    P_Status ""
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
        
        P_ImprimeVenda Printer, dtVendas.Recordset!num_ven, True
               
    End If
    Unload frmVisualiza

End Sub

Private Sub cmndbtnReabrir_Click()
    Dim strSQL As String
    If dbgrd.Row = -1 Then
        MsgBox "selecione a Venda!", vbInformation
        dbgrd.SetFocus
        Exit Sub
    End If
    With dtVendas.Recordset
        If MsgBox("Deseja Realmente Reabrir a Venda?", vbYesNo + vbQuestion) = vbYes Then
            
            strSQL = strSQL & " update tab_vendas set situacao_ven='1'"
            strSQL = strSQL & " where num_ven='" & !num_ven & "' "
            
            
            Banco.Execute strSQL
            strSQL = "delete from tab_venfec where num_ven='" & !num_ven & "' "
            Banco.Execute strSQL
            busca
        End If
    End With
End Sub

Private Sub cmndbtnSair_Click()
        Unload Me
End Sub

Private Sub dtVendas_Reposition()
    If dtVendas.Recordset.RecordCount = 0 Then
        cmndbtnCancelar.Enabled = False
        cmndbtnEfetivar.Enabled = False
        cmndbtnImpr.Enabled = False
        cmndbtnReabrir.Enabled = False
    Else
        cmndbtnCancelar.Enabled = True
        cmndbtnEfetivar.Enabled = True
        cmndbtnImpr.Enabled = True
        cmndbtnReabrir.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    
    busca
    Me.caption = "Vendas a Efetivar  - Usuário: " & strNome_Usu
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dtVendas.Connect = Banco.Connect
End Sub

Private Sub Timer1_Timer()
    
    Dim i As Integer
    On Error Resume Next
    i = dtVendas.Recordset.AbsolutePosition
    busca
    dtVendas.Recordset.Move i
End Sub
