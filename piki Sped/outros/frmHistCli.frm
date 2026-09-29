VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmHistCli 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Histórico de Produtos Adquiridos pelo Cliente"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   735
      Left            =   5520
      Picture         =   "frmHistCli.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   7080
      Picture         =   "frmHistCli.frx":040D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Data dtVendas 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   4095
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmHistCli.frx":0557
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmHistCli.frx":056E
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem:"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   2775
      Begin VB.OptionButton Option1 
         Caption         =   "&Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optnData 
         Caption         =   "&Data"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmHistCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    Dim iPag As Integer
    P_CabRelTimbr "Histórico de Cliente", obj
    P_ImprX obj, "Cliente: " & Me.caption, 5, False
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth, obj.CurrentY
    
    P_ImprX obj, "Data", 2
    P_ImprX obj, "Venda", 20
    P_ImprX obj, "Ref.", 35
    P_ImprX obj, "Produto", 60
    P_ImprX obj, "Qtd.", 150, True, True
    P_ImprX obj, "Unitário", 170, True, True
    P_ImprX obj, "Usuário", 175, False
    P_line1 obj
    With dtVendas.Recordset
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If obj.CurrentY > obj.ScaleHeight - 30 Then
                 
                P_RodapeVis obj, "", (iPag + 1)
                iPag = iPag + 1
                
                P_AddPicVis obj
                
                P_ImprX obj, "Data", 2
                P_ImprX obj, "Venda", 20
                P_ImprX obj, "Ref.", 35
                P_ImprX obj, "Produto", 60
                P_ImprX obj, "Qtd.", 150, True, True
                P_ImprX obj, "Unitário", 170, True, True
                P_ImprX obj, "Usuário", 175, False
                P_line1 obj
                
            End If
            P_ImprX obj, !dataefet_ven, 2
            P_ImprX obj, !seqefet_ven, 20
            P_ImprX obj, !codfab_pro, 35
            P_ImprX obj, StrConv(!nome_pro, vbProperCase), 60
            P_ImprX obj, !qtd_ite, 150, True, True
            P_ImprX obj, !unit_ite, 170, True, True
            P_ImprX obj, !guerra_fun, 175, False
            
            .MoveNext
        Loop
    End With
    
    
    P_RodapeVis obj
    P_AddPicVis obj
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    dtVendas.Connect = Banco.Connect
    P_CentraForm Me
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tf.guerra_fun,tc.cod_cli,tc.nome_cli,data_ven,dataefet_ven,"
    strSQL = strSQL & " tp.*, ti.unit_ite,ti.qtd_ite "
    strSQL = strSQL & " from tab_vendas as tv, tab_itvendas as ti,"
    strSQL = strSQL & " tab_produtos as tp,tab_clientes as tc,tab_funcionarios as tf  "
    strSQL = strSQL & " where ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and tc.cod_Cli=tv.cod_cli "
    strSQL = strSQL & " and tc.cod_cli='" & Me.tag & "' "
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tf.cod_fun=tv.fechavenda_usu "
    If optnData.Value Then
        strSQL = strSQL & " order by dataefet_ven,nome_pro "
    Else
        strSQL = strSQL & " order by nome_pro,Ref_pro,dataefet_ven "
    End If
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
    
    With dtVendas.Recordset
        dbgrd.caption = .RecordCount & " Registros encontrados"
        If .RecordCount > 0 Then
            Me.caption = !cod_cli & "-" & !nome_Cli
        End If
    End With
    End Sub

Private Sub Option1_Click()
    busca
End Sub

Private Sub optnData_Click()
    busca
End Sub
