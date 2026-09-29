VERSION 5.00
Begin VB.Form frmMigraAbcFarma 
   Caption         =   "Migração ABCFarma"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Perfumaria"
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Migra"
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmMigraAbcFarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As database
Dim p As database
Private Sub Command1_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sLista As String
    Dim sEtico As String
    
    strSQL = "delete from tab_principios where cod_pri>1"
    Banco.Execute strSQL
    strSQL = "delete from tab_produtos "
    Banco.Execute strSQL
    strSQL = "select * from tabela "
    Set rPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            If !med_negpos = "SIM" Then
                sLista = "+"
            ElseIf !med_negpos = "NAO" Then
                sLista = "-"
            End If
            If med_neutro = "SIM" Then
                sLista = "N"
            ElseIf !med_neutro = "OUT" Then
                sLista = "O"
            End If
        
            strSQL = Produtos.F_GravaProduto(F_AjustaTexto(!med_des & " " & !med_apr), Format(CDec(!med_barra), "000000000000#"), IIf(Format(!med_gene) = "GEN", "006", "001"), _
                F_Pri(Format(!med_princi)), !med_abc, !med_pla17, !med_pco17, _
                "UN", F_AjustaTexto(!med_apr), !med_pla17, F_Lab(F_AjustaTexto(!lab_nom)), "000", sLista, "", "", "", "", "")
            Me.caption = .AbsolutePosition + 1 & " de " & .RecordCount
            .MoveNext
        Loop
        .Close
    End With
    
End Sub

Function F_sub()

End Function
Function F_Lab(nome As String) As String
    Dim strSQL As String
    strSQL = F_BuscaUmCampo("cod_fab", "tab_fabricantes", "nome_fab='" & nome & "'")
    If strSQL = "" Then
        strSQL = fabricante.F_GravaFabricante(nome, "", "", "", "", "", "", "")
        F_Lab = strSQL
    Else
        F_Lab = strSQL
    End If
End Function
Function F_Pri(nome As String) As String
    Dim strSQL As String
    strSQL = F_BuscaUmCampo("cod_pri", "tab_principios", "nome_pri='" & nome & "'")
    If strSQL = "" Then
        strSQL = principio.F_GravaPri(nome)
        F_Pri = strSQL
    Else
        F_Pri = strSQL
    End If
End Function

Private Sub Command2_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sCat As String
    Dim sCodSec As String
    sCodSec = "002"
    strSQL = "delete from tab_produtos where cod_sub in ('002')"
    Banco.Execute strSQL
    
    
    
    
    
    
    For iCount = 1 To 2
        If iCount = 1 Then
            strSQL = "select distinct tp.*,tf.descricao as desc_cla,tl.fantasia from produto as tp, classificacaoperfumaria as tf, "
            strSQL = strSQL & "laboratorio as tl where tp.class_perf=tf.codigo and tl.codigo=tp.codlab "
        Else
            strSQL = "select distinct tp.*,'GERAL' as desc_cla, tl.fantasia from produto as tp, "
            strSQL = strSQL & "laboratorio as tl where  tl.codigo=tp.codlab "
            strSQL = strSQL & " and tp.class_perf is null and tp.class_tera is null  "
        End If
        
        
        Set rPesq = p.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
            Do While Not .EOF
                sCat = F_BuscaUmCampo("cod_sub", "tab_subsec", "cod_sec='" & sCodSec & "' and nome_sub='" & !desc_cla & "'")
                If sCat = "" Then
                    sCat = subsecAO.F_GravaSub(sCodSec, !desc_cla, "0", "0")
            
                End If
                strSQL = Produtos.F_GravaProduto(F_AjustaTexto(!Descricao), !Codigo, sCat, "00001", !codigosc, _
                    !Preco * 0.7, !Preco, "UN", "", !Preco * 0.7, F_Lab(!fantasia), "000", _
                      !lista, "", "0", "", "", "")
                Me.caption = .AbsolutePosition + 1 & " de " & .RecordCount
                .MoveNext
            Loop
            .Close
        End With
    Next
End Sub

Private Sub Form_Load()
    If Not F_AbreBancoODBC(b, "abcfarma") Then
    End If
    If Not F_AbreBancoODBC(p, "perfum") Then
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    b.Close
End Sub
