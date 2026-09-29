VERSION 5.00
Begin VB.Form frmBaixarEstoque 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBaixar 
      Caption         =   "Baixar"
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtDe 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Baixar de:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   705
   End
End
Attribute VB_Name = "frmBaixarEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBaixar_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select cod_pro, sum(ti.qtd_its) as total"
    strSQL = strSQL & " from tab_saidas as ts, tab_itsai as ti "
    strSQL = strSQL & " where ti.nota_sai=ts.nota_sai "
    strSQL = strSQL & " and ts.emissao_sai>=" & F_GravaData(txtDe.Text)
    strSQL = strSQL & " group by cod_pro order by cod_pro "
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_produtos set ef_pro=ef_pro-" & F_GravaDec(!total)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Conn.Execute strSQL
            
            .MoveNext
        Loop
    End With
    
    
    
    
    
End Sub

Private Sub Form_Load()
    txtDe.Text = Date
End Sub
