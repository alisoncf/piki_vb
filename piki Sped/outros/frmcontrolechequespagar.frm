VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmControleChequesPagar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controle de Cheques"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   15120
   Icon            =   "frmcontrolechequespagar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.Data dtChe 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox pcBusca 
      Align           =   1  'Align Top
      BackColor       =   &H00B0B0B0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   15060
      TabIndex        =   7
      Top             =   855
      Width           =   15120
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Height          =   645
         Left            =   14280
         Picture         =   "frmcontrolechequespagar.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmcontrolechequespagar.frx":04D4
         Left            =   120
         List            =   "frmcontrolechequespagar.frx":04E1
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxConta 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   31
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnConta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         Picture         =   "frmcontrolechequespagar.frx":0516
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxNumDoc 
         Height          =   285
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxNumChe 
         Height          =   285
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Height          =   285
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Height          =   285
         Left            =   9000
         MaxLength       =   10
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12240
         Picture         =   "frmcontrolechequespagar.frx":08A0
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   10080
         MaxLength       =   50
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmcontrolechequespagar.frx":0C2A
         Left            =   12600
         List            =   "frmcontrolechequespagar.frx":0C3A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbxSituacao 
         Height          =   315
         ItemData        =   "frmcontrolechequespagar.frx":0C66
         Left            =   2160
         List            =   "frmcontrolechequespagar.frx":0C76
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "Conta:"
         Height          =   195
         Left            =   5760
         TabIndex        =   32
         Tag             =   "!"
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "Num Doc:"
         Height          =   195
         Left            =   4680
         TabIndex        =   29
         Tag             =   "!"
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblLocal 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "Num Che:"
         Height          =   195
         Left            =   3600
         TabIndex        =   12
         Tag             =   "!"
         Top             =   0
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "&De:"
         Height          =   195
         Left            =   7920
         TabIndex        =   14
         Tag             =   "!"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "&Até:"
         Height          =   195
         Left            =   9000
         TabIndex        =   16
         Tag             =   "!"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "Fornecedor:"
         Height          =   195
         Left            =   10080
         TabIndex        =   18
         Tag             =   "!"
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "&Ordenar Por:"
         Height          =   195
         Left            =   12600
         TabIndex        =   8
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   2160
         TabIndex        =   10
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.PictureBox pcBotes 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   15090
      TabIndex        =   0
      Top             =   0
      Width           =   15120
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmcontrolechequespagar.frx":0CA7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmcontrolechequespagar.frx":0DF1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmcontrolechequespagar.frx":117B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   9240
         Picture         =   "frmcontrolechequespagar.frx":12C5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnDev 
         Caption         =   "&Cheque Devolvido"
         Height          =   735
         Left            =   5160
         Picture         =   "frmcontrolechequespagar.frx":1813
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   6840
         Picture         =   "frmcontrolechequespagar.frx":195D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Width           =   1575
      End
   End
   Begin VB.PictureBox pCdados 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   10905
      TabIndex        =   21
      Top             =   7800
      Width           =   10935
      Begin VB.Label lblDev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblQtd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9240
         TabIndex        =   25
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados:"
         Height          =   195
         Left            =   8520
         TabIndex        =   24
         Top             =   60
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   60
         Width           =   405
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   22
         Top             =   0
         Width           =   480
      End
   End
   Begin MSDBGrid.DBGrid dbgrdChePag 
      Bindings        =   "frmcontrolechequespagar.frx":1D6A
      Height          =   5415
      Left            =   0
      OleObjectBlob   =   "frmcontrolechequespagar.frx":1D7E
      TabIndex        =   27
      Top             =   1680
      Width           =   10575
   End
End
Attribute VB_Name = "frmControleChequesPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxBuscar_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    If cmbxTipo.ListIndex = 0 Then
        dbgrdChe.Visible = True
        dbgrdChePag.Visible = False
    Else
        dbgrdChe.Visible = False
        dbgrdChePag.Visible = True
    End If
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBaixar_Click()
    Dim sSEQ As String
    Dim iCount As Integer
    Dim c As cheque
    Dim sChe As String
    Dim sEsp As String
    Dim cValor As Currency
    Dim strSQL As String
    Dim sCod_Con  As String
    Dim sData As String
    With dtChe.Recordset
    
        If .RecordCount = 0 Then
            MsgBox "Não há o que baixar!", vbExclamation
            Exit Sub
        End If
        
      
            sCod_Con = !cod_con
            sData = Date
      
        If sCod_Con = "" Then
            Exit Sub
        End If
        If MsgBox("Confirma baixa do cheque " & !num_che & " ?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        If sCod_Con = "" Then
            Exit Sub
        End If
        
        
        sChe = F_GravaFinanceiro("D", Date, CStr(!Valor_che), _
                 "Baixa Cheque " & !num_che & "-" & !nome, sCod_Con, "", "000", !seq_che, "", "000")
                
        If sChe <> "" Then
            'P_Class !Valor_che, Left(cmbxTipo.Text, 1), sChe
            strSQL = "update tab_cheques set dep_che = " & F_GravaData(Date) & ","
            strSQL = strSQL & " status_Che=2 "
            strSQL = strSQL & " where seq_che = '" & !seq_che & "' "
            banco.Execute strSQL
            
            strSQL = "insert into tab_chequedep(seq_che,data_dep,cod_con)values("
            strSQL = strSQL & "'" & !seq_che & "',"
            strSQL = strSQL & F_GravaData(sData) & ","
            strSQL = strSQL & "'" & sCod_Con & "')"
            banco.Execute strSQL
            
            MsgBox "Cheque baixado!", vbInformation
            dtChe.Refresh
            
        End If
        
    End With
            
End Sub

Private Sub cmndbtnBuscar_Click()
    
    Dim strSQL As String
    Dim sOrderBy As String
    On Error GoTo trata_Erro
    
    
    If cmbxOrder.ListIndex = 0 Then
        sOrderBy = "venc_che"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderBy = "Nome"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderBy = "num_doc"
    Else
        sOrderBy = "Valor_Che"
    End If
    strSQL = Cheques.F_selectChPag(Left(cmbxTipoTer.Text, 1), txtbxExpr.Tag, txtbxDataIni.Text, txtbxdataFin.Text, _
        txtbxNumChe.Text, "", Left(cmbxSituacao.Text, 1), txtbxConta.Tag, "", "", txtbxNumDoc.Text, sOrderBy)
    dtChe.RecordSource = strSQL
    dtChe.Refresh
    
    soma
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnLocaliza_Click()
    frmBusca_Ter.txtbxCli.Text = UCase(txtbxExpr.Text)
    On Error Resume Next
    frmBusca_Ter.Show 1
    txtbxExpr.Text = frmBusca_Ter.dtCli.Recordset!Nome_Ter
    txtbxExpr.Tag = frmBusca_Ter.dtCli.Recordset!cod_ter
    Unload frmBusca_Ter
End Sub

Private Sub cmndbtnConta_Click()
    buscaConta
End Sub
Sub buscaConta()
    Dim c As ContaCorrente
    c = F_BuscaContaCorrente(txtbxConta.Text)
    If c.cod_con <> "" Then
        txtbxConta.Text = c.nome_con
        txtbxConta.Tag = c.cod_con
    End If
End Sub
Private Sub cmndbtnDev_Click()
    On Error GoTo trata_Erro
    With dtChe.Recordset
        If F_DevolveChe(!seq_che) Then
            dtChe.Refresh
        End If
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub buscaTer()
    
        buscaCLi
    
End Sub
Sub buscaCLi()
    Dim c As tCliente
    c = Clientes.F_BuscaCliente("D", txtbxExpr.Text)
    
    If c.cod_cli <> "" Then
        txtbxExpr.Text = c.nome_Cli
        txtbxExpr.Tag = c.cod_cli
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    Dim c As cheque
    Dim cp As chePag
    Dim strSQL As String
    With dtChe.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser editado!", vbExclamation
            Exit Sub
        End If
        
            cp = F_DadosChePag(!Venc_che, !num_che, _
                 !cod_con, CStr(!Valor_che))
            If cp.num_che = "" Then
                MsgBox "Não foi informado um número para o cheque!", vbExclamation
                Exit Sub
            ElseIf cp.Valor_che = 0 Then
                MsgBox "Não foi informado um valor válido para o cheque!  ", vbExclamation
                Exit Sub
            End If
            If MsgBox("Confirma alteração do cheque?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "update tab_cheques set num_che = '" & cp.num_che & "', "
                
                
                strSQL = strSQL & "cod_con='" & cp.cod_con & "',"
                
                strSQL = strSQL & "venc_che=" & F_GravaData(cp.Venc_che) & ","
                strSQL = strSQL & "valor_che=" & F_GravaDec(cp.Valor_che)
                strSQL = strSQL & " where seq_che = '" & !seq_che & "' "
                banco.Execute strSQL
                MsgBox "cheque alterado!", vbInformation
                RefreshDataPosition dtChe
            End If
        
    
    End With
End Sub

Private Sub cmndbtnImpr_Click()
    If dtChe.Recordset.RecordCount <= 0 Then
        Exit Sub
    End If
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    
    Dim strSQL As String
    Printer.Orientation = vbPRORLandscape
    If cmbxTipo.ListIndex = 0 Then
        strSQL = "A Receber"
    End If
    If cmbxSituacao.ListIndex = 2 Then
        strSQL = strSQL & " Devolvidos "
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " de " & txtbxDataIni.Text
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " até " & txtbxDataIni.Text
    End If
    P_CabRelTimbr "Relatório de Cheques " & strSQL, obj
    
    With dtChe.Recordset
        obj.fontsize = 8
        obj.FontBold = True
        P_line1 obj
        P_ImprX obj, "Cliente", 1
        P_ImprX obj, "Titular", 60
        P_ImprX obj, "CPF", 120
        P_ImprX obj, "Número", 150
        P_ImprX obj, "Conta", 170
        P_ImprX obj, "Ag.", 185
        P_ImprX obj, "Valor", 210
        P_ImprX obj, "Bco", 220

        P_ImprX obj, "Data", 235
        P_ImprX obj, "Venc.", 250
        P_ImprX obj, "Status", 265, False
        P_line1 obj
        obj.FontBold = False
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 25 Then
                P_RodapeVis obj, "Continua..."
                P_AddPicVis obj
                obj.fontsize = 8
                obj.FontBold = True
                
                obj.Print
                P_line1 obj
                P_ImprX obj, "Cliente", 1
                P_ImprX obj, "Titular", 60
                P_ImprX obj, "CPF", 120
                P_ImprX obj, "Número", 150
                P_ImprX obj, "Conta", 170
                P_ImprX obj, "Ag.", 185
                P_ImprX obj, "Bco", 200
                P_ImprX obj, "Data", 215
                P_ImprX obj, "Venc.", 240
                P_ImprX obj, "Status", 255, False
                P_line1 obj
                obj.FontBold = False
            End If
            obj.fontsize = 8
            P_ImprX obj, !nome, 1
            P_ImprX obj, !nome_che, 60
            P_ImprX obj, !cpf_che, 120
            P_ImprX obj, !num_che, 150
            P_ImprX obj, !cco_che, 170
            P_ImprX obj, !age_che, 185
            P_ImprX obj, FormatVal(!Valor_che), 217, True, True
            P_ImprX obj, !cod_bco, 220
            P_ImprX obj, !Data_Doc, 235
            
            P_ImprX obj, !Venc_che, 250
            P_ImprX obj, !situacao, 265, False
            P_line1 obj, vbDot
            .MoveNext
        Loop
        P_line1 obj
        obj.FontBold = True
        P_ImprX obj, "Cheques listados: " & lblQtd.caption, 10
        P_ImprX obj, "Valor Total: " & lblTotal.caption, obj.ScaleWidth - 10, False, True
        obj.FontBold = False
        
        
        P_AddPicVis obj
    End With
    
    
    
End Sub
Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.Show 0
    
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
    buscaTer
End Sub

Private Sub dtNot_Reposition()
    Dim strSQL As String
    With dtNot.Recordset
        If .RecordCount = 0 Then Exit Sub
        strSQL = "select * from tab_cheques where seq_doc='" & !seq_doc & "' "
        dtChe.RecordSource = strSQL
        dtChe.Refresh
        If dtChe.Recordset.RecordCount > 0 Then
            dbgrdChe.caption = "Cheques referentes ao documento " & !num_doc
            dbgrdChe.Visible = True
        Else
            dbgrdChe.Visible = False
        End If
        If .RecordCount > 0 Then
            If !reman_itd = 0 Then
                cmndbtnEdit.Enabled = False
                cmndbtnBaixar.Enabled = False
            Else
                cmndbtnEdit.Enabled = True
                cmndbtnBaixar.Enabled = True
            End If
            Exit Sub
        
        End If
        
    End With
End Sub


Private Sub dbgrdChe_DblClick()
    If cmndbtnEdit.Enabled Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub dbgrdChe_GotFocus()
    dbgrdChe.MarqueeStyle = 3
End Sub

Private Sub dbgrdChe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrdChe_DblClick
    End If
End Sub

Private Sub dbgrdChe_LostFocus()
    dbgrdChe.MarqueeStyle = 0
End Sub

Private Sub dbgrdChePag_GotFocus()
    'dbgrdChe.MarqueeStyle = 3
End Sub

Private Sub dbgrdChePag_LostFocus()
'    dbgrdChe.MarqueeStyle = 0
End Sub

Sub soma()
    Dim cTotal As Currency
    
    On Error GoTo trata_Erro
    
    With dtChe.Recordset
        .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !Valor_che
            .MoveNext
        Loop
    End With
    lblTotal.caption = FormatVal(cTotal)
    Exit Sub
trata_Erro:
    lblTotal.caption = "0,00"
End Sub

Private Sub dtChe_Reposition()
    On Error Resume Next
    If dtChe.Recordset!status_che = 2 Then
        cmndbtnEdit.Enabled = False
        cmndbtnBaixar.Enabled = False
        
    
    Else
        cmndbtnEdit.Enabled = True
        cmndbtnBaixar.Enabled = True
    End If
    If dtChe.Recordset!status_che = "3" Then
        devChe
    Else
        lblDev.Visible = False
    End If
    
End Sub
Sub devChe()
    Dim strSQL As String
    Dim rPesq As Recordset
    lblDev.caption = ""
    With dtChe.Recordset
        strSQL = "select * from tab_chequedev where seq_che='" & !seq_che & "' "
        Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        
    End With
    With rPesq
        If .RecordCount > 0 Then
            lblDev.caption = !motivo_Dev & " em " & !data_dev
        End If
        .Close
    End With
    lblDev.Visible = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    Dim strSQL As String
  '  On Error GoTo trata_Erro
    resizeForm Me
    txtbxDataIni.Text = "01" & Mid(Format(Date, "dd/mm/yyyy"), 3, 8)
    strSQL = DateAdd("m", 1, txtbxDataIni.Text)
    txtbxdataFin.Text = DateAdd("D", -1, strSQL)
   
    
    dtChe.Connect = banco.Connect
    
    

    
    cmbxOrder.ListIndex = 0
    cmbxTipoTer.ListIndex = 0
    'cmndbtnBuscar_Click
    Exit Sub
trata_Erro:
    
End Sub

Private Sub optnCli_Click()
    frExpr.Enabled = True
End Sub

Private Sub optnTodas_Click()
    frExpr.Enabled = False
 
    
End Sub

Private Sub Form_Resize()
    posicionaControles
End Sub
Sub posicionaControles()
    On Error Resume Next
    pcBotes.Width = Me.Width - 300
    pcBusca.Width = pcBotes.Width
    dbgrdChe.Width = pcBotes.Width
    dbgrdChePag.Width = pcBotes.Width
    
    pCdados.Top = Me.Height - pCdados.Height - 400
    dbgrdChe.Height = pCdados.Top - dbgrdChe.Top - 50
    dbgrdChePag.Height = dbgrdChe.Height
    pCdados.Width = pcBotes.Width
    cmndbtnSair.Left = pcBotes.Width - cmndbtnSair.Width - 100
End Sub

Private Sub txbtxLocalche_Change()

End Sub

Private Sub txtbtxLocalche_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaLoc
    End If
End Sub
Sub buscaLoc()
    Dim l As localche
    l = f_BuscaLocalChe(txtbxLocal.Text)
    If l.cod_loc <> "" Then
        txtbxLocal.Text = l.nome_loc
        txtbxLocal.Tag = l.cod_loc
        cmndbtnBuscar_Click
    End If
End Sub
Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_Change()
    cmndbtnBuscar_Click
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxExpr_LostFocus
        cmndbtnTerc_Click
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    If cmbxBuscar.ListIndex = 0 Then
        txtbxExpr.Text = F_SoNumero(txtbxExpr.Text)
    Else
        P_AjustaTexto txtbxExpr
    End If
End Sub


Private Sub txtbxLocal_GotFocus()
    'P_SelecTexto txtbxLocal
End Sub

Private Sub txtbxLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaLoc
    End If
End Sub
