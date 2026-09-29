VERSION 5.00
Begin VB.Form frmMigracaoSMP 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "prod"
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Esp"
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "div"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Migrar"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmMigracaoSMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim str As String
    Dim strSQL
    Dim iCount As Integer
    Dim strAux As String
    Dim strPalavra As String
    Dim x As Integer
    Open App.Path & "\tabela.txt" For Input As #1
        Line Input #1, str
        Do While Not EOF(1)
            
            Line Input #1, str
            strSQL = "insert into tab_migra(nome_pro,nome_div,nome_esp,nome_gen,scusto_pro,spreco_pro)values("
            For iCount = 1 To Len(str)
                
                
                strAux = Mid(str, iCount, 1)
                If iCount = 1 Then
                    strAux = "'" & strAux
                End If
                If strAux = Chr(9) Then
                    
                    If iCount = Len(str) - 1 Then
                        strSQL = strSQL & "'"
                    Else
                        strSQL = strSQL & "',' "
                    End If
                    
                    
                ElseIf strAux = ";" Then
                    strSQL = strSQL & ");"
                Else
                    
                    strSQL = strSQL & strAux
                End If
                
                
            Next
            'On Error Resume Next
            On Error Resume Next
            banco.Execute UCase(strSQL) & "') "
            str = ""
            strSQL = ""
            x = x + 1
            P_Status x & " Registros"
        Loop
    Close #1
End Sub

Private Sub Command2_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim iCod As Integer
    Dim iCod2 As Integer
    Dim rPesq2 As Recordset
    
    banco.Execute "delete from tab_generos"
    banco.Execute "delete from tab_divisoes"
    strSQL = "select distinct nome_gen from tab_migra "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    iCod = 1
    iCod2 = 1
    With rPesq
        Do While Not .EOF
            strSQL = "insert into tab_generos(cod_gen,nome_gen)values("
            strSQL = strSQL & "'" & Format(iCod, "00#") & "',"
            strSQL = strSQL & "'" & !nome_gen & "')"
            banco.Execute strSQL
            strSQL = "select distinct nome_div from tab_migra where nome_gen='" & !nome_gen & "' "
            Set rPesq2 = banco.OpenRecordset(strSQL, dbOpenSnapshot)
            With rPesq2
                Do While Not .EOF
                    'MsgBox !nome_div
                    If F_BuscaUmCampo("COD_DIV", "TAB_DIVisoes", "Nome_Div='" & !nome_div & "'") = "" Then
                        strSQL = "insert into tab_divisoes(cod_gen,nome_div,cod_div)values("
                        strSQL = strSQL & "'" & Format(iCod, "00#") & "',"
                        strSQL = strSQL & "'" & !nome_div & "',"
                        strSQL = strSQL & "'" & Format(iCod2, "00#") & "')"
                        banco.Execute strSQL
                        iCod2 = iCod2 + 1
                    Else
                       ' MsgBox "Não "
                    End If
                    .MoveNext
                    
                Loop
            End With
            iCod = iCod + 1
            
            .MoveNext
            
        Loop
        .Close
    End With
End Sub

Private Sub Command3_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim iCod As Integer
    
    
    
    
    banco.Execute "delete from tab_especies"
    strSQL = "select distinct nome_esp from tab_migra "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    iCod = 1
    
    With rPesq
        Do While Not .EOF
            strSQL = "insert into tab_especies(cod_esp,nome_esp)values("
            strSQL = strSQL & "'" & Format(iCod, "00#") & "',"
            strSQL = strSQL & "'" & !nome_esp & "')"
            banco.Execute strSQL
            
            iCod = iCod + 1
            
            .MoveNext
            
        Loop
        .Close
    End With
End Sub
 
Private Sub Command4_Click()
    P_grava
    
    
    
End Sub
Sub P_grava()
    Dim strSQL As String
    Dim sCod As String
    Dim rPesq As Recordset
    Dim iCount As Integer
    banco.Execute "delete from tab_produtos"
    strSQL = "select distinct tm.*,tg.cod_gen,td.cod_div,te.cod_esp "
    strSQL = strSQL & "from tab_migra as tm, tab_divisoes as td, tab_generos as tg,"
    strSQL = strSQL & "tab_especies as te where tm.nome_gen = tg.nome_gen "
    strSQL = strSQL & " and tm.nome_esp = te.nome_esp  and tm.nome_div = td.nome_div "
    strSQL = strSQL & " order by cod_esp"
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            sCod = F_BuscaUmCampo(" max(cod_pro)", _
            "tab_produtos", " cod_esp='" & !cod_esp & "'")
            
            If sCod = "" Then
                sCod = "001"
            Else
                sCod = Mid(sCod, 4, 3)
                sCod = Format(Int(sCod) + 1, "00#")
            End If
            sCod = F_GeraDig(sCod)
            sCod = !cod_esp & sCod
            If Len(sCod) < 7 Then
                MsgBox "menor"
            End If
            strSQL = "insert into tab_produtos(cod_pro,nome_pro,unidade_pro,"
            strSQL = strSQL & "cod_gen,cod_div,cod_esp,descricao_pro,codfab_pro,"
            strSQL = strSQL & "custo_pro,preco_pro,estoque_pro"
            strSQL = strSQL & ")Values("
            strSQL = strSQL & "'" & sCod & "' ,"
            strSQL = strSQL & "'" & !nome_pro & "',"
            strSQL = strSQL & "'PC',"
            strSQL = strSQL & "'" & !cod_gen & "',"
            strSQL = strSQL & "'" & !cod_div & "', "
            strSQL = strSQL & "'" & !cod_esp & "', "
            strSQL = strSQL & "'',"
            strSQL = strSQL & "'',"
            strSQL = strSQL & f_GravaDec(!scusto_pro) & ","
            strSQL = strSQL & f_GravaDec(!sPreco_pro) & ","
            strSQL = strSQL & "0)"
            banco.Execute strSQL
            .MoveNext
            iCount = iCount + 1
            P_Status .AbsolutePosition / .RecordCount * 100 & " Por cento..."
            mdiFrm.StatusBar1.Refresh
        Loop
        .Close
    End With
    
    
    
    
    ' MsgBox "Produto Gravado com Sucesso", vbInformation
    
    
End Sub

