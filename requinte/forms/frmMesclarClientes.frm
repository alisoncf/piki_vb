VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMesclarClientes 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Mesclar Clientes"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   12060
      TabIndex        =   7
      Top             =   7515
      Width           =   12120
      Begin VB.CommandButton cmndbtnAjustar 
         Caption         =   "Mesclar Selecionados"
         Height          =   735
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   9960
         TabIndex        =   9
         Top             =   0
         Width           =   2055
      End
      Begin VB.TextBox txtbxCliPri 
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Listados:"
         Height          =   195
         Left            =   5880
         TabIndex        =   15
         Top             =   0
         Width           =   630
      End
      Begin VB.Label lblListados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   6720
         TabIndex        =   14
         Top             =   0
         Width           =   120
      End
      Begin VB.Label lblSelec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   3360
         TabIndex        =   13
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Selecionados:"
         Height          =   195
         Left            =   2280
         TabIndex        =   12
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Principal:"
         Height          =   195
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar:"
      Height          =   615
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton optnGroup 
         Caption         =   "Iguais Agrupados"
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optnNomes 
         Caption         =   "Todos os Nomes"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9128
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtbxExpr 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expressão:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmMesclarClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnAjustar_Click()
    
    If strExclVenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente excluir os clientes selecionados?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    
    
    If optnNomes.Value Then
        If txtbxCliPri.text = "" Then
            MsgBox "Escolha, ou informe o código do item que será preservado!", vbExclamation
            Exit Sub
        End If
        If CCur(lblSelec.caption) > 1 Then
                     
                         
            mescla
        End If
    End If
End Sub
Sub mescla()
    Dim iCount As Integer
    Dim strSQL As String
    Dim sCod As String
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 0
            If .text = "S" Then
                .col = 1
                sCod = .text
                If sCod <> txtbxCliPri.text Then
                    strSQL = "update tab_vendas set cod_cli='" & txtbxCliPri.text & "' "
                    strSQL = strSQL & " where cod_cli='" & sCod & "'"
                    Conn.Execute strSQL
                
                
                    strSQL = "update tab_saidas set cod_cli='" & txtbxCliPri.text & "' "
                    strSQL = strSQL & " where cod_cli='" & sCod & "'"
                    Conn.Execute strSQL
                    strSQL = "update tab_acertos set cod_cli='" & txtbxCliPri.text & "' "
                    strSQL = strSQL & " where cod_cli='" & sCod & "'"
                    Conn.Execute strSQL
                    strSQL = "update tab_doc set cod_ter='" & txtbxCliPri.text & "' "
                    strSQL = strSQL & " where cod_ter='" & sCod & "' and tipo_ter='C'"
                    Conn.Execute strSQL
                
                    strSQL = "delete from tab_telefones where tipo_ter='C' and cod_cli='" & sCod & "'"
                    Conn.Execute strSQL
                    strSQL = "delete from tab_clientes where cod_cli='" & sCod & "'"
                    Conn.Execute strSQL
                    strSQL = "delete from tab_clientetrab where cod_cli='" & sCod & "'"
                    Conn.Execute strSQL
                End If
            End If
        Next
        lblSelec.caption = "0"
        MsgBox "Clientes mesclados!", vbInformation
        busca
    End With
End Sub
Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AjustaGrid
    resizeForm Me
    
End Sub
Sub AjustaGrid()
    
    With msflxgrd
        If optnNomes.Value Then
            .Cols = 8
            .rows = 1
            .col = 0: .text = "S": .ColWidth(0) = 200
            .col = 1: .text = "Código": .ColWidth(1) = 800
            .col = 2: .text = "Nome": .ColWidth(2) = 4500
            .col = 3: .text = "CPF": .ColWidth(3) = 2500
            .col = 4: .text = "Endereço": .ColWidth(4) = 3500
            .col = 5: .text = "Bairro Antigo": .ColWidth(5) = 2500
            .col = 6: .text = "Cidade Antiga": .ColWidth(6) = 2500
            .col = 7: .text = "Data": .ColWidth(7) = 2500
        Else
            .Cols = 9
            .rows = 1
            .col = 0: .text = "S": .ColWidth(0) = 200
            .col = 1: .text = "Código": .ColWidth(1) = 800
            .col = 2: .text = "Nome": .ColWidth(2) = 4500
            .col = 3: .text = "CPF": .ColWidth(3) = 2500
            .col = 4: .text = "Endereço": .ColWidth(4) = 3500
            .col = 5: .text = "Bairro Antigo": .ColWidth(5) = 2500
            .col = 6: .text = "Cidade Antiga": .ColWidth(6) = 2500
            .col = 7: .text = "Data": .ColWidth(7) = 1000
            .col = 8: .text = "Qtd": .ColWidth(8) = 1000
        End If
    End With
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    P_AjustaTexto txtbxExpr
    txtbxCliPri.text = ""
    msflxgrd.rows = 1
    If optnNomes.Value Then
        strSQL = "select data_cli,cod_cli,nome_cli,cpf_cli,endereco_cli,bairro_cli,cidade_cli from tab_clientes where (nome_cli like '" & txtbxExpr.text & "%' or "
        strSQL = strSQL & " cpf_cli='" & F_FormataCPF(txtbxExpr.text) & "')"
        strSQL = strSQL & " order by nome_cli, cpf_cli,data_cli"
    Else
        strSQL = "select data_cli,cod_cli,nome_cli,cpf_cli,endereco_cli,bairro_cli,cidade_cli,count(*) as qtd from tab_clientes as tc where (nome_cli like '" & txtbxExpr.text & "%' or "
        strSQL = strSQL & " cpf_cli='" & F_FormataCPF(txtbxExpr.text) & "') group by nome_cli "
        strSQL = strSQL & " having qtd > 1 "
        strSQL = strSQL & " order by nome_cli, cpf_cli,data_cli "
        
    End If
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "" & Chr(9) & !cod_cli & Chr(9)
            strSQL = strSQL & !nome_Cli & Chr(9)
            strSQL = strSQL & !cpf_cli & Chr(9)
            strSQL = strSQL & !endereco_cli & Chr(9)
            strSQL = strSQL & !bairro_cli & Chr(9)
            strSQL = strSQL & !cidade_cli & Chr(9) & !data_cli
            If Not optnNomes.Value Then
                strSQL = strSQL & Chr(9) & !qtd
            End If
            P_Status "Buscando"
            msflxgrd.AddItem strSQL
            
            .MoveNext
        Loop
        lblListados.caption = .RecordCount
        .Close
    End With
    P_Status ""
End Sub

Private Sub Form_Resize()
    msflxgrd.width = pcBot.top - msflxgrd.top - 100
    msflxgrd.width = Me.width - msflxgrd.Left - 100
End Sub

Private Sub msflxgrd_DblClick()
    With msflxgrd
        .col = 0
        If .text = "" Then
            lblSelec.caption = CCur(lblSelec.caption) + 1
            .text = "S"
            .col = 1
            txtbxCliPri.text = .text
            
            
        Else
            lblSelec.caption = CCur(lblSelec.caption) - 1
            txtbxCliPri.text = ""
            .text = ""
        End If
    End With
End Sub

Private Sub msflxgrd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then
'        msflxgrd.Col = 0: msflxgrd.Text = ""
'        msflxgrd_DblClick
'        msflxgrd.Col = 1
'        txtbxCliPri.Text = msflxgrd.Text
'    End If
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        cmndbtnBuscar_Click
    End If
End Sub
