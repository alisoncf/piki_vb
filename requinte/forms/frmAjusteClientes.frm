VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAjusteClientes 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CPFs"
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Carnês"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   4920
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAcertos 
      Caption         =   "Acertos"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   3840
      Width           =   2415
   End
End
Attribute VB_Name = "frmAjusteClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAcertos_Click()
    ajustagrid
    
    acertossemCli
    
End Sub
Sub acertossemCli()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rItens As New ADODB.Recordset
    strSQL = "select distinct cod_ter,seq_doc from tab_doc where cod_ter > 120418 and tipo_ter = 'C' order by cod_ter"
    abreObjetoAdo rPesq, strSQL
    
    
    
    With rPesq
        Do While Not .EOF
            
            strSQL = F_BuscaUmCampo("cod_cli", "tab_clientes", "cod_cli='" & !cod_ter & "'")
            If strSQL = "" Then
                grid.AddItem "" & Chr(9) & !cod_ter
                strSQL = Clientes.F_GravaCli("CADASTRO PERDIDO - CORRIGIR", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", !cod_ter)
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    
    strSQL = "select distinct cod_cli,seq_ace from tab_doc where cod_cli > 120418 order by cod_cli desc "
    abreObjetoAdo rPesq, strSQL
    
    
End Sub
    
Sub ajustagrid()
    With grid
        .FixedCols = 0
        .rows = 1
        .Cols = 4
        .Col = (0): .Text = "Acerto"
        .Col = (1): .Text = "Cliente"
        
    End With
End Sub

Private Sub Command1_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select cod_cli from tab_clientes as tc inner join tab_doc as td on td.cod_cli=tc.cod_cli"
    strSQL = strSQL & " inner join tab_itdoc as ti on td.seq_doc=ti.seq_doc where tipo_ter='C' "
    strSQL = strSQL & " and ti.vencini_itd <='2018-12-31' and reman_itd>0"
    
    abreObjetoAdo rPesq, strSQL
    
End Sub

Private Sub Command2_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim r2 As New ADODB.Recordset
    
    strSQL = "select cod_cli from tab_clientes where cpf_cli=''"
    abreObjetoAdo rPesq, strSQL
    strSQL = "select cod_cli from tab_clientes where cpf_cli=''"
    abreObjetoAdo rPesq, strSQL
    
End Sub
