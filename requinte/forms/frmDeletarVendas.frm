VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeletarVendas 
   Caption         =   "Deletar Vendas Abertas"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNenhuma 
      Caption         =   "Nenhuma"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdTodas 
      Caption         =   "Todas"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Excluir Selecionadas"
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   735
      Left            =   10800
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   6375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11245
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Pesquisar"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Data:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmDeletarVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub todas(valor As Boolean)
    Dim i As Integer
    With grid
        For i = 1 To .rows - 1
            .row = i
            .col = 0
            If valor = True Then
                .Text = "S"
            Else
                .Text = ""
            End If
        Next
    End With
End Sub
Private Sub cmdDel_Click()
    With grid
        If strExclVenda_usu <> "S" Then
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
        For i = 1 To .rows - 1
            .row = i
            
            .col = 0
            If .Text = "S" Then
                .col = 1
                strSQL = "delete from tab_itvendas where num_ven='" & .Text & "' "
                Conn.Execute strSQL
                strSQL = "delete from tab_vendas where num_ven='" & .Text & "' "
                Conn.Execute strSQL
                
                
                P_Status "Deletando... " & .row
            End If
            
            
        Next
        busca
        P_Status
    End With
End Sub

Private Sub cmdNenhuma_Click()
    todas False
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdTodas_Click()
    todas True
End Sub

Private Sub Form_Load()
    txtData.Text = DateAdd("D", -1, Date)
        
        AjustaGrid
    
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    grid.rows = 1
    strSQL = "select tv.*, count(ti.cod_pro) as itens "
    strSQL = strSQL & " from tab_vendas as tv left join tab_itvendas as ti "
    strSQL = strSQL & " on ti.num_ven=tv.num_ven"
    strSQL = strSQL & " where tv.situacao_ven='1' "
    strSQL = strSQL & " and tv.data_ven<=" & F_GravaDataTempo(Date & "23:59:00")
    strSQL = strSQL & " group by tv.num_ven "
    strSQL = strSQL & " order by num_ven desc "
    
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        Do While Not .EOF
            strSQL = "" & Chr(9) & !num_ven & Chr(9) & FormatVal(!valor_ven) & Chr(9)
            strSQL = strSQL & !data_ven & Chr(9)
            strSQL = strSQL & !abrevenda_usu & Chr(9)
            strSQL = strSQL & !itens & Chr(9)
            strSQL = strSQL & !nomecli_ven
            
            grid.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub


Sub AjustaGrid()
    With grid
        .rows = 1
        .FixedCols = 0
        .Cols = 7
        .col = 0: .Text = "*"
        .col = 1: .Text = "Venda"
        .col = 2: .Text = "Valor"
        .col = 3: .Text = "Data"
        .col = 4: .Text = "Vendedor"
        .col = 5: .Text = "Produtos"
        .col = 6: .Text = "Cliente"
        .ColWidth(0) = 500
        .ColWidth(3) = 2500
        .ColWidth(6) = 2500
        
        
    End With
End Sub

Private Sub grid_DblClick()
    With grid
        .col = 0
        If .Text = "" Then
            .Text = "S"
        Else
            .Text = ""
        End If
    End With
End Sub
