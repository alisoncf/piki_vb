VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelClienteSpc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Clientes no SPC"
   ClientHeight    =   8430
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   13305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   7215
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   12726
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.Frame Frame2 
         Caption         =   "Ordem"
         Height          =   615
         Left            =   2040
         TabIndex        =   6
         Top             =   120
         Width           =   2895
         Begin VB.OptionButton optnData 
            Caption         =   "Data"
            Height          =   195
            Left            =   1920
            TabIndex        =   9
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optnDoc 
            Caption         =   "Doc"
            Height          =   195
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optAlfab 
            Caption         =   "Código"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   120
         Picture         =   "frmRelClienteSpc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   8040
         Picture         =   "frmRelClienteSpc.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   11520
         Picture         =   "frmRelClienteSpc.frx":0797
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEditar 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   9840
         Picture         =   "frmRelClienteSpc.frx":088C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Menu mnAcao 
      Caption         =   "Ação"
      Begin VB.Menu mnBuscar 
         Caption         =   "Buscar"
      End
      Begin VB.Menu mnImpr 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnEdit 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "frmRelClienteSpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxOpcao_Click()
    busca
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim r2 As Recordset
    Dim sMsg As String
    Dim sCli As String
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sCPF As String
    Dim r3 As Recordset
    msflxgrd.rows = 1
    On Error GoTo trata_Erro
    
    strSQL = "select * from tab_clientespc where baixa_spc='0000-00-00' "
    
    If optnDoc.Value Then
        strSQL = strSQL & " order by doc_spc,data_spc desc"
    ElseIf optnData.Value Then
        strSQL = strSQL & " order by data_spc desc,cod_Cli "
    Else
        strSQL = strSQL & " order by cod_cli, data_spc desc "
    End If
    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        
        Do While Not .EOF
            If !tipo_ter = "C" Then
                sTab = "tab_clientes"
                sCod = "cod_cli"
                sNome = "nome_cli"
                sCPF = "cpf_cli"
            ElseIf !tipo_ter = "R" Then
                sTab = "tab_representantes"
                sCod = "cod_rep"
                sNome = "nome_rep"
                sCPF = "cpf_Rep"
            End If
            
            If !tipo_doc = "C" Then
                strSQL = "select "
            Else
                strSQL = "select ti.*,tc." & sCod & " as cod, "
                strSQL = strSQL & "tc." & sNome & " as nome, tc." & sCPF & " as cpf "
                
                strSQL = strSQL & " from tab_itdoc as ti,tab_doc as td, "
                strSQL = strSQL & sTab & " as tc "
                strSQL = strSQL & " where tc." & sCod & "=td.cod_ter "
                strSQL = strSQL & " and td.seq_doc=ti.seq_doc "
                strSQL = strSQL & " and td.cod_ter='" & !cod_cli & "' "
                strSQL = strSQL & " and ti.seq_itd='" & !doc_spc & "' "
                
                Set r2 = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
                sMsg = !seq_spc & Chr(9)
                sMsg = sMsg & !data_spc & Chr(9)
                sMsg = sMsg & !tipo_ter & Chr(9)
                
                
                    If r2.RecordCount = 0 Then
                        
                        If !tipo_ter = "C" Then
                            strSQL = "select cod_cli as cod,nome_cli as nome, cpf_cli as cpf "
                            strSQL = strSQL & " from tab_clientes where cod_cli='" & !cod_cli & "' "
                            Set r3 = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
                        Else
                            strSQL = "select cod_rep as cod,nome_rep as nome, cpf_rep as cpf "
                            strSQL = strSQL & " from tab_representantes where cod_rep='" & !cod_cli & "' "
                            Set r3 = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
                        End If
                        If r3.RecordCount > 0 Then
                            sMsg = sMsg & r3!cod & Chr(9)
                            sMsg = sMsg & r3!nome & Chr(9)
                            sMsg = sMsg & r3!CPF & Chr(9)
                            sMsg = sMsg & FormatVal(!valor_spc) & Chr(9)
                            sMsg = sMsg & "Não encontrado" & Chr(9)
                            sMsg = sMsg & "Não encontrado" & Chr(9)
                            sMsg = sMsg & "Não encontrado" & Chr(9)
                        End If
                    Else
                        sMsg = sMsg & r2!cod & Chr(9)
                        sMsg = sMsg & r2!nome & Chr(9)
                        sMsg = sMsg & r2!CPF & Chr(9)
                        sMsg = sMsg & FormatVal(!valor_spc) & Chr(9)
                        sMsg = sMsg & r2!seq_itd & Chr(9)
                        sMsg = sMsg & r2!venc_itd & Chr(9)
                        sMsg = sMsg & FormatVal(r2!reman_itd) & Chr(9)

                    End If
                    msflxgrd.AddItem sMsg
            End If
            
            strSQL = "select sum(ti.reman_itd) from tab_itdoc where    "
            
            .MoveNext
        Loop
        .Close
    End With
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub soma()
    Dim cTotal As Currency
    With dt.Recordset
        Do While Not .EOF
            cTotal = cTotal + !total
            
            .MoveNext
        Loop
        lblVendas.caption = FormatVal(cTotal)
    End With
End Sub

Private Sub cmndbtnBuscaRep_Click()
    buscaRep
End Sub

Private Sub cmndbtnGraf_Click()

End Sub

Private Sub cmndbtnEditar_Click()
    msflxgrd.Col = 3
    P_ClienteSpc msflxgrd.Text
End Sub

Private Sub cmndbtnImprimir_Click()
    P_MsflxToHtml "Clientes no SPC", msflxgrd
    
    
    
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'dt.Connect = Banco.Connect
    
    
    ajustaGrid
End Sub
Sub ajustaGrid()
    With msflxgrd
        .Cols = 10
        
        .FixedCols = 0
        .rows = 1
        .Col = 0: .Text = "..."
        .Col = 1: .Text = "Data"
        .Col = 2: .Text = "T"
        .Col = 3: .Text = "Cód."
        .Col = 4: .Text = "Nome"
        .Col = 5: .Text = "CPF"
        .Col = 6: .Text = "Valor"
        .Col = 7: .Text = "Doc"
        .Col = 8: .Text = "Vencimento"
        .Col = 9: .Text = "Reman"
        
        .ColWidth(4) = 3500
        .ColWidth(5) = 2000
        .ColWidth(2) = 400
        .ColWidth(3) = 800
    End With
End Sub

Private Sub txtbxRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaRep
    End If
End Sub

Sub buscaRep()
    Dim r As representante.representante
    r = F_BuscaRep(txtbxRepresentante.Text)
    txtbxRepresentante.Text = r.nome_rep
    txtbxRepresentante.tag = r.cod_Rep
    busca
End Sub

Private Sub mnBuscar_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub mnEdit_Click()
    cmndbtnEditar_Click
End Sub

Private Sub mnImpr_Click()
    cmndbtnImprimir_Click
End Sub

Private Sub mnSair_Click()
    cmndbtnSair_Click
End Sub

Private Sub msflxgrd_DblClick()
    cmndbtnEditar_Click
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        msflxgrd_DblClick
    End If
End Sub

Private Sub msflxgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnAcao
    End If
End Sub

Private Sub optAlfab_Click()
    busca
End Sub

Private Sub optnData_Click()
    busca
End Sub

Private Sub optnDoc_Click()
    busca
End Sub
