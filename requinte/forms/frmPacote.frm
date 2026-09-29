VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPacote 
   Caption         =   "Pacote"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14205
   Icon            =   "frmPacote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   14205
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   8160
      Left            =   5910
      ScaleHeight     =   8130
      ScaleWidth      =   8265
      TabIndex        =   7
      Top             =   0
      Width           =   8295
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   6255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   11033
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
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
         Left            =   6360
         TabIndex        =   15
         Top             =   120
         Width           =   690
      End
      Begin VB.Label lblCaixa 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   6360
         TabIndex        =   14
         Top             =   480
         Width           =   75
      End
      Begin VB.Label lblVendedor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   4320
         TabIndex        =   13
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
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
         Left            =   4320
         TabIndex        =   12
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label lblDataVen 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   120
         Width           =   570
      End
      Begin VB.Label lblNumVen 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.Data dtVendas 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   14175
      TabIndex        =   0
      Top             =   8160
      Width           =   14205
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar"
         Height          =   825
         Left            =   6240
         Picture         =   "frmPacote.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnFinalizarVenda 
         Caption         =   "&F7 Entregar"
         Height          =   825
         Left            =   4200
         Picture         =   "frmPacote.frx":0E14
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   825
         Left            =   2160
         Picture         =   "frmPacote.frx":119E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar (Esc)"
         Height          =   825
         Left            =   12120
         Picture         =   "frmPacote.frx":12E8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   825
         Left            =   120
         Picture         =   "frmPacote.frx":1453
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Timer tmr 
      Interval        =   2000
      Left            =   2640
      Top             =   4200
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmPacote.frx":17DD
      Height          =   8055
      Left            =   0
      OleObjectBlob   =   "frmPacote.frx":17F4
      TabIndex        =   6
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmPacote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblnum_ven_Click()
End Sub

Private Sub cmndbtnFinalizarVenda_Click()
    On Error GoTo trata_Erro
     
    
    With dtVendas.Recordset
    
        
            If F_BaixaPac(dtVendas.Recordset!seq_pac) Then
                MsgBox "Pacote entregue com sucesso!", vbInformation
                busca
            End If
        
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnImpr_Click()
    On Error Resume Next
    ImprVenda dtVendas.Recordset!num_ven
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dtVendas_Reposition()
    On Error Resume Next
    With dtVendas.Recordset
        If .RecordCount > 0 Then
            lblNumVen.caption = !num_ven
            lblDataVen.caption = !data_ven
            lblVendedor.caption = !guerra_fun
            lblCaixa.caption = F_BuscaUmCampo("Guerra_fun", "tab_funcionarios", "cod_fun='" & !fechavenda_usu & "'")
        
            BuscaItVendas (!num_ven)
            
        End If
        
    End With
    
End Sub
Sub BuscaItVendas(num_ven As String)
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    msflxgrd.rows = 1
    strSQL = "select ti.* from tab_itvendas as ti where num_ven='" & num_ven & "' "
    abreObjetoAdo rPesq, strSQL
    strSQL = ""
    With rPesq
        Do While Not .EOF
            strSQL = .AbsolutePosition + 1 & Chr(9)
           strSQL = strSQL & (!cod_pro) & Chr(9)
           strSQL = strSQL & !desc_ite & Chr(9)
           strSQL = strSQL & !qtd_ite & Chr(9)
           strSQL = strSQL & FormatVal(!unit_ite) & Chr(9)
           strSQL = strSQL & FormatVal(!final_ite)
           msflxgrd.AddItem strSQL
            .MoveNext
        Loop
        
        .Close
    End With
    ZebraGrid msflxgrd, 1
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnFinalizarVenda_Click
    End If
End Sub

Private Sub Form_Load()
   ' dtVendas.Connect = Banco.Connect
    Me.top = 10
    Me.Left = 10
    With msflxgrd
        
        msflxgrd.Cols = 6
        .rows = 1
        .col = 0: .Text = "..."
        .col = 1: .Text = "Código"
        .col = 2: .Text = "Descrição"
        .col = 3: .Text = "Qtd"
        .col = 4: .Text = "Unit"
        .col = 5: .Text = "Final"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 3500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        
        
    End With
    
End Sub

Private Sub tmr_Timer()
     busca
End Sub
Sub busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    strSQL = "select tc.nome_cli,fechavenda_usu,tv.valor_ven,data_ven,tv.nomecli_ven,tp.*,tf.guerra_fun from tab_vendas as tv,tab_clientes as tc,tab_pacotes as tp,"
    strSQL = strSQL & "tab_funcionarios as tf where tf.cod_fun=tv.abrevenda_usu "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    strSQL = strSQL & " and tp.num_ven=tv.num_ven "
    strSQL = strSQL & " and status_pac='1' "
    strSQL = strSQL & " order by data_pac desc"
    dtVendas.RecordSource = strSQL
        
    RefreshDataPosition dtVendas
    
    
    Exit Sub
trata_Erro:
    dtVendas.Refresh
    
End Sub

