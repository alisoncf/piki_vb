VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmResumoContasPagar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumo de Documentos a Pagar"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C6DADD&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   10275
      TabIndex        =   9
      Top             =   6600
      Width           =   10305
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total em Cheques:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblTotalCH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total em Dupl:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   13
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lblTotalItd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   12
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Geral:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6720
         TabIndex        =   11
         Top             =   0
         Width           =   1065
      End
      Begin VB.Label lblGTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7920
         TabIndex        =   10
         Top             =   0
         Width           =   420
      End
   End
   Begin VB.Frame frControle 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin VB.ComboBox cmbxConta 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   10
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   5295
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   9340
         _Version        =   393216
         SelectionMode   =   1
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   4920
         Picture         =   "frmResumoContasPagar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   6600
         Picture         =   "frmResumoContasPagar.frx":00F9
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   8280
         Picture         =   "frmResumoContasPagar.frx":0506
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Corrente:"
         Height          =   195
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&De:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmResumoContasPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cTotalCH As Currency
Dim cTotalItd As Currency

Sub soma()
    Dim cTotal As Currency
    With dtNot.Recordset
        Do While Not .EOF
                        
            cTotal = cTotal + !Valor_che
            
            .MoveNext
        Loop
    End With
    lblTotal.caption = FormatVal(cTotal)
    
End Sub

Private Sub cmbxOrder_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxTipo_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub


Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    Dim sCampo As String
   ' On Error GoTo trata_erro
    cTotalCH = 0
    cTotalItd = 0
    buscaCH
    buscaITD
    lblTotalCH.caption = FormatVal(cTotalCH)
    lblTotalItd.caption = FormatVal(cTotalItd)
    lblGTotal.caption = FormatVal(cTotalCH + cTotalItd)
  
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub buscaCH()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCodCon As String
    Dim cTotal As Currency
    strSQL = "select  tf.nome_for,tc.nome_con,td.historico_doc,td.valor_Doc,td.parc_doc,td.num_doc,td.data_doc, td.valor_doc,"
    strSQL = strSQL & "ti.* from " & TAB_FORNECEDORES & " as tf,  " & TAB_DOC & "  as td,tab_cheques as ti,  "
    strSQL = strSQL & "tab_contas as tc where tc.cod_con=ti.cod_con"
    strSQL = strSQL & " and tf.cod_for=td.cod_ter and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.tipo_doc='P' "
    strSQL = strSQL & " and ti.dep_che is null "
    If txtbxDAtaFIn.Text <> "" Then
        strSQL = strSQL & " and ti.venc_che < " & F_GravaData(txtbxDAtaFIn.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and ti.venc_che > " & F_GravaData(txtbxDataIni.Text)
    End If
    
    strSQL = strSQL & " order by num_con,venc_che,nome_for"
    rPesq.Open strSQL, conn_client
    
    msflxgrd.rows = 1
    With rPesq
        Do While Not .EOF
            If !cod_con <> sCodCon Then
                If msflxgrd.rows > 1 Then
                    msflxgrd.AddItem ">>" & Chr(9) & "Total...:" & Chr(9) & FormatVal(cTotal)
                    msflxgrd.Row = msflxgrd.rows - 1
                    ColoreLinhaGrid msflxgrd, msflxgrd.Row, , True, vbRed
                End If
                msflxgrd.AddItem ">>>" & Chr(9) & !cod_con & Chr(9)
                
                msflxgrd.col = 5
                msflxgrd.Row = msflxgrd.rows - 1
                msflxgrd.Text = !nome_con
                ColoreLinhaGrid msflxgrd, msflxgrd.Row, 13880000, True, vbBlue
                
                sCodCon = !cod_con
                cTotal = 0
            End If
            strSQL = !num_doc & Chr(9)
            strSQL = strSQL & !data_doc & Chr(9)
            strSQL = strSQL & !Venc_che & Chr(9)
            strSQL = strSQL & FormatVal(!Valor_che) & Chr(9)
            strSQL = strSQL & !num_che & Chr(9)
            strSQL = strSQL & !nome_for & Chr(9)
            strSQL = strSQL & !historico_doc
            msflxgrd.AddItem strSQL
            cTotal = cTotal + !Valor_che
            cTotalCH = cTotalCH + !Valor_che
            .MoveNext
        Loop
        If cTotal > 0 Then
            msflxgrd.AddItem ">>" & Chr(9) & "Total...:" & Chr(9) & FormatVal(cTotal)
            msflxgrd.Row = msflxgrd.rows - 1
            ColoreLinhaGrid msflxgrd, msflxgrd.Row, , True, vbRed
        End If
        .Close
    End With
End Sub
Sub buscaITD()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCodCon As String
    Dim cTotal As Currency
    strSQL = "select  tf.nome_for,td.historico_doc,td.valor_Doc,td.parc_doc,td.num_doc,td.data_doc, td.valor_doc,"
    strSQL = strSQL & "ti.* from " & TAB_FORNECEDORES & " as tf,  " & TAB_DOC & " 1 as td, " & TAB_ITDOC & " 1 as ti  "
    strSQL = strSQL & "where "
    strSQL = strSQL & " tf.cod_for=td.cod_ter and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.tipo_doc='P' "
    strSQL = strSQL & " and reman_itd > 0 "
    If txtbxDAtaFIn.Text <> "" Then
        strSQL = strSQL & " and ti.venc_itd < " & F_GravaData(txtbxDAtaFIn.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and ti.venc_itd > " & F_GravaData(txtbxDataIni.Text)
    End If
    
    strSQL = strSQL & " order by venc_itd,nome_for"
    rPesq.Open strSQL, conn_client
    
    msflxgrd.AddItem "Duplicatas"
    msflxgrd.Row = msflxgrd.rows - 1
    ColoreLinhaGrid msflxgrd, msflxgrd.Row, vbYellow, True
    With rPesq
        Do While Not .EOF
            
            strSQL = !num_doc & Chr(9)
            strSQL = strSQL & !data_doc & Chr(9)
            strSQL = strSQL & !venc_itd & Chr(9)
            strSQL = strSQL & FormatVal(!reman_itd) & Chr(9)
            strSQL = strSQL & FormatVal(!valor_itd) & Chr(9)
            strSQL = strSQL & !nome_for & Chr(9)
            strSQL = strSQL & !historico_doc
            msflxgrd.AddItem strSQL
            cTotal = cTotal + !reman_itd
            cTotalItd = cTotalItd + !reman_itd
            .MoveNext
        Loop
        If cTotal > 0 Then
            msflxgrd.AddItem ">>" & Chr(9) & "Total...:" & Chr(9) & FormatVal(cTotal)
            msflxgrd.Row = msflxgrd.rows - 1
            ColoreLinhaGrid msflxgrd, msflxgrd.Row, , True, vbRed
        End If
        strSQL = ">" & Chr(9) & "Total CH.:" & Chr(9) & FormatVal(cTotalCH) & Chr(9)
        strSQL = strSQL & "Total It.:" & Chr(9) & FormatVal(cTotalItd) & Chr(9)
        strSQL = strSQL & "T.Geral:" & Chr(9) & FormatVal(cTotalItd + cTotalCH)
        msflxgrd.AddItem strSQL
        msflxgrd.Row = msflxgrd.rows - 1
        ColoreLinhaGrid msflxgrd, msflxgrd.Row, vbCyan, True
        .Close
    End With
End Sub

Private Sub cmndbtnLocaliza_Click()
    
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.Show 0
End Sub

Sub imprime(obj As Object)
    P_CabRelTimbr "Resumo de Documentos a Pagar de: " & F_EsticaTexto(txtbxDataIni.Text, 10) & " até: " & F_EsticaTexto(txtbxDAtaFIn.Text, 10), obj
    
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth - 0, obj.CurrentY
    impr obj
    
                
                
                
    
    
End Sub
Sub impr(obj As Object, Optional linh As Integer = 1)
    With msflxgrd
        On Error GoTo trata_Erro
        .Row = 0
        .col = 0: P_ImprX obj, .Text, 2
        .col = 1: P_ImprX obj, .Text, 20
        .col = 2: P_ImprX obj, .Text, 40
        .col = 3: P_ImprX obj, .Text, 75, True, True
        .col = 4: P_ImprX obj, .Text, 80
        .col = 5: P_ImprX obj, .Text, 100
        .col = 6: P_ImprX obj, .Text, 150, False
        P_Line obj, 0, obj.CurrentY, obj.ScaleWidth - 0, obj.CurrentY
        
        For iCount = 1 To .rows - 2
            
            
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                If TypeName(obj) <> "Printer" Then
                    
                    frmVisualiza.imglst.ListImages.Add , , frmVisualiza.Pic.image
                    frmVisualiza.Pic.Cls
                   
                    frmVisualiza.Pic.Picture = Nothing
                    obj.CurrentY = yIni
                
                End If
                P_RodapeVis obj, Me.caption, "Continua..."
                .Row = 0
                .col = 0: P_ImprX obj, .Text, 2
                .col = 1: P_ImprX obj, .Text, 20
                .col = 2: P_ImprX obj, .Text, 40
                .col = 3: P_ImprX obj, .Text, 75, True, True
                .col = 4: P_ImprX obj, .Text, 80
                .col = 5: P_ImprX obj, .Text, 100
                .col = 6: P_ImprX obj, .Text, 150, False
            End If
            msflxgrd.Row = iCount
            msflxgrd.col = 0
            If .Text = ">>>" Then
                obj.FontBold = True
                P_Line obj, 0, obj.CurrentY, obj.ScaleWidth - 0, obj.CurrentY, vbDot
                .col = 1
                P_ImprX obj, .Text, 2
                .col = 5
                P_ImprX obj, .Text, 20, False
                P_Line obj, 0, obj.CurrentY, obj.ScaleWidth - 0, obj.CurrentY
            ElseIf .Text = ">>" Then
                obj.FontBold = True
                P_Line obj, 0, obj.CurrentY, obj.ScaleWidth - 0, obj.CurrentY, vbDot
                .col = 1
                P_ImprX obj, .Text, 20
                .col = 2
                P_ImprX obj, .Text, 40, False
                P_Line obj, 0, obj.CurrentY, obj.ScaleWidth - 0, obj.CurrentY

            Else
                obj.FontBold = False
                .col = 0: P_ImprX obj, .Text, 2
                .col = 1: P_ImprX obj, .Text, 20
                .col = 2: P_ImprX obj, .Text, 40
                .col = 3: P_ImprX obj, .Text, 75, True, True
                .col = 4: P_ImprX obj, .Text, 80
                .col = 5: P_ImprX obj, StrConv(Left(.Text, 30), vbProperCase), 100
                .col = 6: P_ImprX obj, StrConv(Left(.Text, 30), vbProperCase), 150, False
            End If
        Next
             
    End With
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth, obj.CurrentY
    obj.FontBold = True
    P_ImprX obj, "Total em Cheques: " & lblTotalCH.caption, 5
    P_ImprX obj, "Total em Duplicatas: " & lblTotalItd.caption, 60
    P_ImprX obj, "Total Geral: " & lblGTotal.caption, 200, False, True
    obj.FontBold = True
        P_RodapeVis obj, Me.caption, "Continua..."
        If TypeName(obj) <> "Printer" Then
                    
                    frmVisualiza.imglst.ListImages.Add , , frmVisualiza.Pic.image
                    frmVisualiza.Pic.Cls
                   
                    frmVisualiza.Pic.Picture = Nothing
                    obj.CurrentY = yIni
                
                End If
trata_Erro:
                
        
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
         imprime Printer
    End If
    Unload frmVisualiza
End Sub

Private Sub cmndbtnSair_Click()
        Unload Me
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.BorderStyle = 1
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.BorderStyle = 0
End Sub

Private Sub cmndbtnTerc_Click()
    
    If cmbxBuscar.ListIndex <> 1 Then
        cmndbtnBuscar_Click
        Exit Sub
    End If
    If cmbxTIpo.ListIndex = 0 Then
        frmSelectCli.txtbxNOme.tag = "C"
    Else
        frmSelectCli.txtbxNOme.tag = "F"
    End If
    frmSelectCli.txtbxNOme.Text = txtbxExpr.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxExpr.Text = frmSelectCli.adoR!nome
    txtbxExpr.tag = frmSelectCli.adoR!cod
    Unload frmSelectCli
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dbgrdChe_GotFocus()
    dbgrdChe.MarqueeStyle = 3
End Sub

Private Sub dbgrdChe_LostFocus()
    dbgrdChe.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    Dim strSQL As String
    With msflxgrd
        .FixedCols = 0
        .Cols = 7
        .rows = 1
        
        .col = 0: .Text = "Documento"
        .col = 1: .Text = "Data"
        .col = 2: .Text = "Vencimento"
        .col = 3: .Text = "Valor"
        .col = 4: .Text = "Cheque"
        .col = 5: .Text = "Fornecedor"
        .col = 6: .Text = "Histórico"
        .ColWidth(5) = 2500
        .ColWidth(6) = 3500
        .ColAlignment(0) = 0
        
    End With
    
 
    'P_PreencheContas cmbxConta
    P_CentraForm Me
 
        
End Sub

Private Sub optnCli_Click()
    frExpr.Enabled = True
End Sub

Private Sub optnTodas_Click()
    frExpr.Enabled = False
 
    
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_Change()
    'cmndbtnTerc_Click
End Sub








