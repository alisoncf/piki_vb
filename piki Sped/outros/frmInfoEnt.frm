VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmInfoEnt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informações de Compra"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optnGer 
      Caption         =   "Usar &Filtro"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.OptionButton optnFiscal 
      Caption         =   "Informações &Gerais"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.Data dtProd 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   5280
      Picture         =   "frmInfoEnt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   7320
      Picture         =   "frmInfoEnt.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid dbgrdProd 
      Bindings        =   "frmInfoEnt.frx":0294
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "frmInfoEnt.frx":02A9
      TabIndex        =   3
      Top             =   600
      Width           =   9135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   3600
      Width           =   405
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3120
      TabIndex        =   6
      Top             =   3525
      Width           =   570
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Produto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2460
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9240
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmInfoEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim sTabCab As String
    Dim sTabIte As String
    If optnFiscal.Value Then
        sTabCab = "tab_Compras"
        sTabIte = "tab_itcompras"
    Else
        sTabCab = "tab_Cmp"
        sTabIte = "tab_itcmp"
    End If
    
    strSQL = "select tp.nome_pro,tf.nome_for,status_not,data_not,num_not,ti.*, "
    strSQL = strSQL & " ti.unit_ite*ti.qtd_ite as total,ti.qtd_ite*ti.qtd_emb as qtd_total "
    strSQL = strSQL & " from " & sTabCab & " as tc, " & sTabIte & " as ti,"
    strSQL = strSQL & " tab_fornecedores as tf,tab_produtos as tp "
    strSQL = strSQL & " where ti.seq_not=tc.seq_not and ti.cod_pro='" & Me.Tag & "' "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tf.cod_for=tc.cod_for "
    strSQL = strSQL & " group by ti.Seq_not,ti.cod_pro "
    strSQL = strSQL & " order by data_not desc"
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    With dtProd.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        Me.caption = Me.Tag & " - " & !nome_pro & " [Informações de Entradas]  "
        lblNome.caption = !nome_pro
        
    End With
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    With dtProd.Recordset
        Do While Not .EOF
            cTotal = cTotal + (!qtd_ite * !qtd_emb)
            .MoveNext
        Loop
        lblTOtal.caption = FormatVal(cTotal)
    End With
End Sub
Private Sub Form_Load()
    P_CentraForm Me
    dtProd.Connect = Banco.Connect
    
End Sub

Private Sub optnFiscal_Click()
    busca
End Sub

Private Sub optnGer_Click()
    busca
End Sub
