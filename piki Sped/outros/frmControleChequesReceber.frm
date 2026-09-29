VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmControleChequesReceber 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controle de Cheques"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14580
   Icon            =   "frmControleChequesReceber.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   Begin VB.Data dtChe 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgrdChe 
      Bindings        =   "frmControleChequesReceber.frx":014A
      Height          =   5175
      Left            =   0
      OleObjectBlob   =   "frmControleChequesReceber.frx":015E
      TabIndex        =   45
      Top             =   2400
      Width           =   11055
   End
   Begin VB.PictureBox pcBusca 
      Align           =   1  'Align Top
      BackColor       =   &H00B0B0B0&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   14520
      TabIndex        =   7
      Top             =   855
      Width           =   14580
      Begin VB.Frame Frame1 
         Caption         =   "Data:"
         Height          =   495
         Left            =   8400
         TabIndex        =   17
         Top             =   0
         Width           =   2055
         Begin VB.OptionButton optnDoc 
            Caption         =   "Doc"
            Height          =   195
            Left            =   840
            TabIndex        =   19
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optnVenc 
            Caption         =   "Venc."
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Height          =   885
         Left            =   12720
         Picture         =   "frmControleChequesReceber.frx":228A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxTitular 
         Height          =   285
         Left            =   10560
         MaxLength       =   50
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtbxCpf 
         Height          =   285
         Left            =   10560
         MaxLength       =   50
         TabIndex        =   31
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtbxSeqChe 
         Height          =   285
         Left            =   9480
         MaxLength       =   50
         TabIndex        =   29
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtbxSeqDoc 
         Height          =   285
         Left            =   8400
         MaxLength       =   50
         TabIndex        =   27
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtbxNumDoc 
         Height          =   285
         Left            =   7320
         MaxLength       =   50
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtbxNumChe 
         Height          =   285
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   23
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmControleChequesReceber.frx":2614
         Left            =   120
         List            =   "frmControleChequesReceber.frx":2621
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxLocal 
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   36
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtbxDataIni 
         Height          =   285
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Height          =   285
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         Picture         =   "frmControleChequesReceber.frx":2653
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   11
         Top             =   240
         Width           =   3495
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmControleChequesReceber.frx":29DD
         Left            =   4440
         List            =   "frmControleChequesReceber.frx":29ED
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cmbxSituacao 
         Height          =   315
         ItemData        =   "frmControleChequesReceber.frx":2A19
         Left            =   120
         List            =   "frmControleChequesReceber.frx":2A29
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "Titular Cheque:"
         Height          =   195
         Left            =   10560
         TabIndex        =   20
         Tag             =   "!"
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "CPF Cheque"
         Height          =   195
         Left            =   10560
         TabIndex        =   30
         Tag             =   "!"
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "Seq.Che"
         Height          =   195
         Left            =   9480
         TabIndex        =   28
         Tag             =   "!"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "Seq.Doc"
         Height          =   195
         Left            =   8400
         TabIndex        =   26
         Tag             =   "!"
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "Nº Doc"
         Height          =   195
         Left            =   7320
         TabIndex        =   24
         Tag             =   "!"
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "Nº Cheque"
         Height          =   195
         Left            =   6240
         TabIndex        =   22
         Tag             =   "!"
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblLocal 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "&Local do Cheque:"
         Height          =   195
         Left            =   2280
         TabIndex        =   35
         Tag             =   "!"
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "&De:"
         Height          =   195
         Left            =   6240
         TabIndex        =   13
         Tag             =   "!"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "&Até:"
         Height          =   195
         Left            =   7320
         TabIndex        =   15
         Tag             =   "!"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "&Terceiro"
         Height          =   195
         Left            =   2280
         TabIndex        =   10
         Tag             =   "!"
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         Caption         =   "&Ordenar Por:"
         Height          =   195
         Left            =   4440
         TabIndex        =   37
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   600
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
      ScaleWidth      =   14550
      TabIndex        =   0
      Top             =   0
      Width           =   14580
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmControleChequesReceber.frx":2A5A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmControleChequesReceber.frx":2BA4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmControleChequesReceber.frx":2F2E
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
         Picture         =   "frmControleChequesReceber.frx":3078
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnDev 
         Caption         =   "&Cheque Devolvido"
         Height          =   735
         Left            =   5160
         Picture         =   "frmControleChequesReceber.frx":35C6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   6840
         Picture         =   "frmControleChequesReceber.frx":3710
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
      TabIndex        =   39
      Top             =   8640
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados:"
         Height          =   195
         Left            =   8520
         TabIndex        =   42
         Top             =   60
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   4560
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   0
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmControleChequesReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbxBuscar_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    
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
        
        If !tipo_che = "R" Then
            frmContaCorrente.Show 1
            frmContaCorrente.caption = "Baixa de Cheque"
            sCod_Con = Left(frmContaCorrente.Tag, 3)
            sData = frmContaCorrente.txtbxData.Text
            
            Unload frmContaCorrente
        Else
            sCod_Con = !cod_con
            sData = Date
        End If
        If sCod_Con = "" Then
            Exit Sub
        End If
        If MsgBox("Confirma baixa do cheque " & !num_che & " ?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        If sCod_Con = "" Then
            Exit Sub
        End If
        
        
        sChe = F_GravaFinanceiro("R", Date, CStr(!Valor_che), _
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
    Dim sTab As String
    
    On Error GoTo trata_Erro
    
    
    
    strSQL = F_SelectChRec(Left(cmbxTipoTer.Text, 1), txtbxExpr.Tag, txtbxDataIni.Text, txtbxdataFin.Text, _
        txtbxNumChe.Text, "", cmbxSituacao.Text, txtbxCpf.Text, txtbxTitular.Text, txtbxSeqChe.Text, txtbxSeqDoc.Text, _
        txtbxNumDoc.Text, , , txtbxLocal.Text, optnVenc.Value)
    
    dtChe.RecordSource = strSQL
    dtChe.Refresh
    
    
'
'    Exit Sub
'    If cmbxOrder.ListIndex = 0 Then
'        sOrderby = "venc_che"
'    ElseIf cmbxOrder.ListIndex = 1 Then
'        sOrderby = "Nome"
'    ElseIf cmbxOrder.ListIndex = 2 Then
'        sOrderby = "num_doc"
'    Else
'        sOrderby = "Valor_Che"
'    End If
'
'    If tipo_ter = "C" Then
'        sTab = "tab_clientes"
'        sCod = "cod_cli"
'        sNome = "Nome_cli"
'        sCpf = "Cpf_cli"
'    ElseIf tipo_ter = "R" Then
'        sTab = "tab_representantes"
'        sCod = "cod_rep"
'        sNome = "Nome_rep"
'        sCpf = "Cpf_Rep"
'    ElseIf tipo_ter = "U" Then 'funcionário
'        sTab = "tab_funcionarios"
'        sCod = "cod_fun"
'        sNome = "Nome_fun"
'        sCpf = "cpf_fun"
'    End If
'
'        strSQL = "select tc.nome_cli as nome, tl.nome_loc,tx.*, td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc, "
'        strSQL = strSQL & "if(status_che='1','Aberto',if(status_che='2','Baixado','DEVOLVIDO')) as situacao "
'        strSQL = strSQL & " from tab_clientes as tc, tab_doc as td,tab_cheques as tx,tab_localche as tl  "
'        strSQL = strSQL & " where tc.cod_cli=td.cod_ter and td.seq_doc=tx.seq_doc "
'        strSQL = strSQL & " and tl.cod_loc=tx.cod_loc "
'        strSQL = strSQL & " and tx.tipo_che='R' "
'
'
'        If cmbxSituacao.ListIndex > 0 Then
'            strSQL = strSQL & " and status_che='" & Left(cmbxSituacao.Text, 1) & "' "
'        End If
'        If txtbxdataFin.Text <> "" Then
'            strSQL = strSQL & " and tx.venc_che < " & F_GravaData(txtbxdataFin.Text)
'        End If
'        If txtbxDataIni.Text <> "" Then
'            strSQL = strSQL & " and tx.venc_che > " & F_GravaData(txtbxDataIni.Text)
'        End If
'
'        If txtbxExpr.tag <> "" Then
'            strSQL = strSQL & " and td.cod_ter='" & txtbxExpr.tag & "'"
'        End If
'        If txtbxCpf.Text <> "" Then
'            strSQL = strSQL & " and cpf_che ='" & F_FormataCPF(txtbxCpf.Text) & "' "
'        End If
'        If txtbxTitular.Text <> "" Then
'            strSQL = strSQL & " and nome_che like '" & txtbxTitular.Text & "%'"
'
'        End If
'        If txtbxNumChe.Text <> "" Then
'            strSQL = strSQL & " and num_che='" & txtbxNumChe.Text & "' "
'        End If
'        If txtbxNumDoc.Text <> "" Then
'            strSQL = strSQL & " and num_doc='" & txtbxNumDoc.Text & "' "
'        End If
'        If txtbxSeqChe.Text <> "" Then
'            strSQL = strSQL & " and tx.seq_che='" & txtbxSeqChe.Text & "' "
'        End If
'        If txtbxSeqDoc.Text <> "" Then
'            strSQL = strSQL & " and tx.seq_doc='" & txtbxSeqDoc.Text & "' "
'        End If
'
'        strSQL = strSQL & " order by " & sOrderby
'
'    dtChe.Connect = Banco.Connect
'    dtChe.RecordSource = strSQL
'    dtChe.Refresh
'    If dtChe.Recordset.RecordCount = 0 Then
''        MsgBox "Não foram encontrados registros para serem listados!" & Chr(13) & _
'                "Verifique sua Pesquisa!", vbExclamation, constNome_sis
'
'
'
'    End If
    lblQtd.caption = dtChe.Recordset.RecordCount
    soma
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
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

Private Sub cmndbtnEdit_Click()
    Dim c As cheque
    Dim cp As chePag
    Dim strSQL As String
    With dtChe.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser editado!", vbExclamation
            Exit Sub
        End If
        
            c = F_DadosCheque(True, !Venc_che, !nome_che, _
                !num_che, !cco_che, !age_che, !cod_bco, CStr(!Valor_che), Format(!cpf_che), !cod_loc)
            If c.num_che = "" Then
'                MsgBox "Não foi informado um número para o cheque!", vbExclamation
                Exit Sub
            ElseIf c.Valor_che = 0 Then
                MsgBox "Não foi informado um valor válido para o cheque!  ", vbExclamation
                Exit Sub
            End If
            If MsgBox("Confirma alteração do cheque?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "update tab_cheques set num_che = '" & c.num_che & "', "
                strSQL = strSQL & "age_che='" & c.age_che & "',"
                strSQL = strSQL & "cco_che='" & c.cco_che & "',"
                strSQL = strSQL & "cod_bco='" & c.cod_bco & "',"
                strSQL = strSQL & "nome_che='" & c.nome_che & "',"
                strSQL = strSQL & "cod_loc='" & c.cod_loc & "',"
                strSQL = strSQL & "venc_che=" & F_GravaData(c.Venc_che) & ","
                strSQL = strSQL & "valor_che=" & F_GravaDec(c.Valor_che) & ","
                strSQL = strSQL & "cpf_che='" & c.cpf_che & "'"
                
                strSQL = strSQL & " where seq_che = '" & !seq_che & "' "
                banco.Execute strSQL
                MsgBox "cheque alterado!", vbInformation
                RefreshDataPosition dtChe
            End If
        
    End With
End Sub

Private Sub cmndbtnImpr_Click()
    
    P_DbgrdToHtml "Listagem de Cheques a Receber ", dbgrdChe, dtChe, _
        txtbxDataIni.Text & " A " & txtbxdataFin.Text
    
    
    Exit Sub
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
    
        strSQL = "A Receber"
    
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
            P_ImprX obj, !Status, 265, False
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


Private Sub cmndbtnTerc_Click()
   On Error GoTo trata_Erro
    buscaTer
    
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub buscaTer()
    
        buscaCLi
    
End Sub
Sub buscaCLi()
    Dim c As tCliente
    c = Clientes.F_BuscaCliente("C", txtbxExpr.Text)
    
    If c.cod_cli <> "" Then
        txtbxExpr.Text = c.nome_Cli
        txtbxExpr.Tag = c.cod_cli
    End If
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
    dbgrdChe.MarqueeStyle = 3
End Sub

Private Sub dbgrdChePag_LostFocus()
    dbgrdChe.MarqueeStyle = 0
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
    On Error GoTo trata_Erro
    resizeForm Me
    txtbxDataIni.Text = "01" & Mid(Format(Date, "dd/mm/yyyy"), 3, 8)
    strSQL = DateAdd("m", 1, txtbxDataIni.Text)
    txtbxdataFin.Text = DateAdd("D", -1, strSQL)
    cmbxTipoTer.ListIndex = 0
    cmbxOrder.ListIndex = 0
    
    dtChe.Connect = banco.Connect
    
    
 '   P_CentraForm Me
'    strSQL = "select tn.*,tt.* from tab_compras as tn, tab_fornecedores as tt "
'    strSQL = strSQL & " where tt.cod_for=tn.cod_for "
'
'    strSQL = strSQL & " order by num_not "
'    dtNot.RecordSource = strSQL
'    dtNot.Refresh
    
    cmbxOrder.ListIndex = 0
    
    'cmndbtnBuscar_Click
    Exit Sub
trata_Erro:
    
End Sub

Private Sub optnCli_Click()
    
End Sub

Private Sub optnTodas_Click()
    
 
    
End Sub

Private Sub Form_Resize()
    posicionaControles
End Sub
Sub posicionaControles()
    On Error Resume Next
    pcBotes.Width = Me.Width - 300
    pcBusca.Width = pcBotes.Width
    dbgrdChe.Width = pcBotes.Width
    'dbgrdChePag.Width = pcBotes.Width
    
    pCdados.Top = Me.Height - pCdados.Height - 400
    dbgrdChe.Height = pCdados.Top - dbgrdChe.Top - 50
    'dbgrdChePag.Height = dbgrdChe.Height
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


Private Sub txtbxCpf_KeyDown(KeyCode As Integer, Shift As Integer)
    cmndbtnBuscar_Click
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
    'cmndbtnBuscar_Click
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'txtbxExpr_LostFocus
        cmndbtnTerc_Click
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxLocal_Change()
    txtbxLocal.Tag = ""
End Sub

Private Sub txtbxLocal_GotFocus()
    P_SelecTexto txtbxLocal
End Sub

Private Sub txtbxLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaLoc
    End If
End Sub

Private Sub txtbxTitular_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub
