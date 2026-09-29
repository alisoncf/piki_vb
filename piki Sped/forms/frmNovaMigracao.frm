VERSION 5.00
Begin VB.Form frmNovaMigracao 
   Caption         =   "Form1"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVendas 
      Caption         =   "Ajustar id_cliente nas vendas"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmNovaMigracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVendas_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select id,cod_cli from clientes order by id "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update vendas set id_cliente =" & !id
            strSQL = strSQL & " where cod_cli=" & !cod_cli
            strSQL = strSQL & " and cod_emp='01'"
            Conn.Execute strSQL
            
            cmdVendas.caption = .AbsolutePosition & "/" & .RecordCount
            cmdVendas.Refresh
            .MoveNext
            
        Loop
    End With
    
    
End Sub
