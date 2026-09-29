VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frmInfSai 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informações de Saída"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   7320
      Picture         =   "frmInfSai.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   5280
      Picture         =   "frmInfSai.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Data dtProd 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton optnFiscal 
      Caption         =   "Informações &Gerais"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton optnGer 
      Caption         =   "Usar &Filtro"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid dbgrdProd 
      Bindings        =   "frmInfSai.frx":0294
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "frmInfSai.frx":02A9
      TabIndex        =   4
      Top             =   480
      Width           =   9135
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   195
      Left            =   3600
      TabIndex        =   9
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr Gerado: :"
      Height          =   195
      Left            =   2640
      TabIndex        =   8
      Top             =   3720
      Width           =   885
   End
   Begin VB.Label lblQtd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   3480
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd Total:"
      Height          =   195
      Left            =   2640
      TabIndex        =   6
      Top             =   3480
      Width           =   705
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9240
      Y1              =   360
      Y2              =   360
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
      TabIndex        =   5
      Top             =   0
      Width           =   2460
   End
End
Attribute VB_Name = "frmInfSai"
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
    If optnFiscal.Value = False Then
        dbgrdProd.Columns(0).caption = "NF.:"
        dbgrdProd.Columns(0).DataField = "nota_sai"
        strSQL = "select tv.data_ven,ts.nota_sai,tv.cod_usu,tv.nomecli_ven,"
        strSQL = strSQL & "ti.unit_ite as unit,ti.qtd_ite as qtd,ti.unit_ite*ti.qtd_ite as total"
        strSQL = strSQL & " from tab_vendas as tv, tab_itvendas as ti "
        strSQL = strSQL & " where ti.num_ven=tv.num_ven "
        strSQL = strSQL & " and ti.cod_pro='" & Me.Tag & "' "
        strSQL = strSQL & " order by data_ven "
    Else
        strSQL = "select tv.data_ven,tv.num_ven,tv.cod_usu,tv.nomecli_ven,"
        strSQL = strSQL & "ti.unit_ite as unit,ti.qtd_ite as qtd,ti.unit_ite*ti.qtd_ite as total,tp.nome_pro "
        strSQL = strSQL & " from tab_vendas as tv, tab_itvendas as ti,tab_produtos as tp  "
        
        strSQL = strSQL & " where ti.num_ven=tv.num_ven "
        strSQL = strSQL & " and ti.cod_pro='" & Me.Tag & "' "
        strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
        strSQL = strSQL & " and situacao_ven=4 "
        strSQL = strSQL & " order by data_ven "
    End If
    

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
    Dim cQtd As Currency
    With dtProd.Recordset
        Do While Not .EOF
            cTotal = cTotal + !Total
            cQtd = cQtd + !qtd
            .MoveNext
        Loop
    End With
    lblQtd.caption = cQtd
    lblTotal.caption = FormatVal(cTotal)
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

