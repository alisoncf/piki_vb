VERSION 5.00
Begin VB.Form frmFecharMes 
   Caption         =   "Fechar Mês"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1665
      ScaleWidth      =   6345
      TabIndex        =   12
      Top             =   1200
      Width           =   6375
      Begin VB.CommandButton cmndbtnFechar 
         Caption         =   "Fechar Mês"
         Height          =   495
         Left            =   4440
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxImpEst 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtbxImpFed 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Text            =   "0,00"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impostos Estaduais:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impostos Federais:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   615
      Left            =   7800
      TabIndex        =   8
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnConectar 
      Caption         =   "Conectar"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Text            =   "dw"
      Top             =   600
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   9555
      TabIndex        =   4
      Top             =   3360
      Width           =   9615
      Begin VB.Label lblDetalhe 
         AutoSize        =   -1  'True
         Caption         =   "..."
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
         TabIndex        =   11
         Top             =   600
         Width           =   195
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Aguardando Comando"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Andamento do Processo:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1785
      End
   End
   Begin VB.TextBox txtbxAno 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cmbxMes 
      Height          =   315
      ItemData        =   "frmFecharMes.frx":0000
      Left            =   120
      List            =   "frmFecharMes.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banco de Dados"
      Height          =   195
      Left            =   4680
      TabIndex        =   5
      Top             =   360
      Width           =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Ano:"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mês:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   345
   End
End
Attribute VB_Name = "frmFecharMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bDw As New ADODB.Connection
Private Sub cmndbtnConectar_Click()
    abreBancoDw
End Sub

Private Sub cmndbtnFechar_Click()
    abreBancoDw
    Dim strSQL As String
    Dim rAdo As New ADODB.Recordset
    Dim rPesq As New ADODB.Recordset
    Dim sValorSaida As String
    Dim saida As String
    Dim saida2 As String
    Dim valorVen  As String
    Dim cmp1 As String
    Dim cmp2 As String
    Dim customax As String
    Dim customin As String
    Dim r2 As New ADODB.Recordset
    Dim qtd_ite As String
    Dim Entrada As String
    
    Dim primDia As String
    Dim ultDia As String
    
    
    
    primDia = F_PrimDiaMes("01/" & cmbxMes.ListIndex + 1 & "/" & txtbxAno.Text)
    ultDia = F_UltDiaMes("01/" & cmbxMes.ListIndex + 1 & "/" & txtbxAno.Text)
    
    
    P_Status "Processando..."
    lblStatus.caption = "Movimento Mensal"
    lblDetalhe.caption = "Buscando Movimento Mensal"
    strSQL = "select count(*) as total from movmensal where mes=" & cmbxMes.ListIndex + 1 & " and ano=" & txtbxAno.Text
    rAdo.Open strSQL, bDw
    
    If rAdo!total > 0 Then
        MsgBox "Já existe movimento mensal para esse mês e esse ano!", vbExclamation
        rAdo.Close
        Exit Sub
        
    End If
    
    saida = F_BuscaUmCampo("sum(valor_cup)", "cupons", "left(data_cup,10)>=" & F_GravaData(primDia) & _
        " and left(data_cup,10)<=" & F_GravaData(ultDia) & " and cancelado_cup='N'")
        
    cmp1 = F_BuscaUmCampo("sum(valor_not)", "tab_compras", "emissao_not>=" & F_GravaData(primDia) & _
        " and emissao_not<=" & F_GravaData(ultDia))
        
    cmp2 = F_BuscaUmCampo("sum(valor_not)", "tab_cmp", "emissao_not>=" & F_GravaData(primDia) & _
        " and emissao_not<=" & F_GravaData(ultDia))
        
        
    strSQL = "select sum(avista_ven) as avista, sum(aprazo_ven) as aprazo, sum(finan_ven) as finan from  " & TAB_VENFEC & "  as tf,"
    strSQL = strSQL & "" & TAB_VENDAS & " as tv where tv.num_ven=tf.num_ven "
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(primDia)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(ultDia)
    strSQL = strSQL & " and situacao_ven=4"
    rPesq.Open strSQL, conn_client
    
    With rPesq
        Do While Not .EOF
            DW.P_GravaMovMensal cmbxMes.ListIndex + 1, txtbxAno.Text, !avista, !aprazo, !finan, txtbxImpFed.Text, _
            txtbxImpEst.Text, saida, !avista + !aprazo + !finan, cmp1, cmp2
            .MoveNext
        Loop
    End With
    
    
    
    strSQL = "select tp.nome_pro,tp.cod_pro, tp.er_pro,tp.ef_pro  from " & TAB_PRODUTOS & " as tp, " & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti "
    strSQL = strSQL & " where tv.num_ven=ti.num_ven "
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
    trSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(primDia)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(ultDia)
    strSQL = strSQL & " and tv.situacao_ven=4 "
    
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "select max(custo_ite) as maximo, min(custo_ite) as minimo,sum(qtd_ite) as total "
            strSQL = strsqç & " from tab_cmp as tc, tab_itcmp as ti"
            strSQL = strSQL & " where tc.seq_not=ti.seq_not "
            strSQL = strSQL & " and emissao_not>=" & F_GravaData(F_PrimDiaMes)
            strSQL = strSQL & " and emissao_not<=" & F_GravaData(F_UltDiaMes)
            strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " group by cod_pro "
            r.Open strSQL
            If r2.RecordCount > 0 Then
                customax = r2!maximo
                customin = r2!minimo
                qtd_ite = !total
            Else
                customax = 0
                customin = 0
                qtd_ite = 0
            End If
            
            strSQL = "select sum(qtd_ite) as total tab_compras as tc, tab_itcompras as ti"
            strSQL = strSQL & " where tc.seq_not=ti.seq_not "
            strSQL = strSQL & " and emissao_not>=" & F_GravaData(F_PrimDiaMes)
            strSQL = strSQL & " and emissao_not<=" & F_GravaData(F_UltDiaMes)
            strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
            r.Open strSQL
            If r2.RecordCount > 0 Then
                Entrada = !total
            End If
            
            saida = F_BuscaUmCampo("sum(qtd_ite)", "cupons", "left(data_cup,10)>=" & F_GravaData(primDia) & _
        " and left(data_cup,10)<=" & F_GravaData(ultDia) & " and cancelado_cup='N' and cod_pro='" & !cod_pro & "'")
        
            saida2 = F_BuscaUmCampo("sum(qtd_ite)", "" & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti", "dataefet_ven>=" & F_GravaData(primDia) & _
        " and dataefet_ven<=" & F_GravaData(ultDia) & " and situacao_ven='4' and ti.seq_ven=tv.seq_ven and cod_pro='" & !cod_pro & "'")
            valorVen = F_BuscaUmCampo("sum(valor_ven)", "" & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti", "dataefet_ven>=" & F_GravaData(primDia) & _
        " and dataefet_ven<=" & F_GravaData(ultDia) & " and situacao_ven='4' and ti.seq_ven=tv.seq_ven and cod_pro='" & !cod_pro & "'")
            
            
            DW.P_GravaProdutoPreco !cod_pro, !nome_pro, customax, customin, Entrada, saida, qtd_ite, saida2, !er_pro, !ef_pro, cmbxMes.ListIndex + 1, txtbxAno.Text, !pv_pro
            DW.P_GravaProdutoVenda !cod_pro, !nome_pro, cmbxMes.ListIndex + 1, txtbxAno.Text, saida2, valorVen, saida
            .MoveNext
        Loop
    End With
    
    
    
    
    
    rAdo.Close
    
    P_Status ""
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    bDw.Close
    Unload Me
End Sub

Private Sub Form_Load()
    txtbxAno.Text = Year(Date)
End Sub
Sub abreBancoDw()
    On Error Resume Next
    bDw.Close
    Dim strCon As String
    
    On Error GoTo trata_erro
    
    strCon = "Driver={MySQL ODBC 3.51 Driver};" & _
    "Server=" & "localhost" & ";" & _
    "Port=3306;" & _
    "Database=" & "dw" & ";" & _
    "Uid=root;" & _
    "Pwd=;" & _
    "Option=1 2 8 32 2048 16384;" '131072
    
    'DBEngine.DefaultType = dbUseODBC
    
    
    
    bDw.Open strCon
Exit Sub
trata_erro:
    MostraErro
End Sub
