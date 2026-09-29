VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmControleChequesPagar 
   Caption         =   "Controle de Cheques a Pagar"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15120
   Icon            =   "frmControleChequesPagar.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15120
   Begin VB.PictureBox pcBusca 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   15090
      TabIndex        =   7
      Top             =   0
      Width           =   15120
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Height          =   645
         Left            =   14280
         Picture         =   "frmControleChequesPagar.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   0
         Width           =   1095
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmControleChequesPagar.frx":04D4
         Left            =   120
         List            =   "frmControleChequesPagar.frx":04E1
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxConta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnConta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         Picture         =   "frmControleChequesPagar.frx":0516
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxNumChe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9000
         MaxLength       =   10
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10080
         MaxLength       =   50
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmControleChequesPagar.frx":08A0
         Left            =   12600
         List            =   "frmControleChequesPagar.frx":08B0
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbxSituacao 
         Height          =   315
         ItemData        =   "frmControleChequesPagar.frx":08DC
         Left            =   2160
         List            =   "frmControleChequesPagar.frx":08EC
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12240
         Picture         =   "frmControleChequesPagar.frx":091D
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Conta:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5760
         TabIndex        =   31
         Tag             =   "!"
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Num Doc:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4680
         TabIndex        =   28
         Tag             =   "!"
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblLocal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Num Che:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   12
         Tag             =   "!"
         Top             =   0
         Width           =   705
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&De:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7920
         TabIndex        =   14
         Tag             =   "!"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Até:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9000
         TabIndex        =   16
         Tag             =   "!"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Fornecedor:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10080
         TabIndex        =   18
         Tag             =   "!"
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Ordenar Por:"
         ForeColor       =   &H80000008&
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
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   15090
      TabIndex        =   0
      Top             =   8325
      Width           =   15120
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmControleChequesPagar.frx":0CA7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmControleChequesPagar.frx":0DF1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmControleChequesPagar.frx":117B
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
         Picture         =   "frmControleChequesPagar.frx":12C5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnDev 
         Caption         =   "&Cheque Devolvido"
         Height          =   735
         Left            =   5160
         Picture         =   "frmControleChequesPagar.frx":1813
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   6840
         Picture         =   "frmControleChequesPagar.frx":195D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Width           =   1575
      End
   End
   Begin VB.PictureBox pCdados 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   15090
      TabIndex        =   21
      Top             =   7830
      Width           =   15120
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
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frmControleChequesPagar.frx":1D6A
      Height          =   6735
      Left            =   0
      TabIndex        =   35
      Top             =   960
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   11880
      _Version        =   393216
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "status"
         Caption         =   "Status"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nome"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "num_che"
         Caption         =   "Número"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "valor_che"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "venc_che"
         Caption         =   "Vencimento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "nome_con"
         Caption         =   "Conta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "num_doc"
         Caption         =   "Documento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2819,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1874,835
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1214,929
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmControleChequesPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmbxBuscar_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    If adp!status_che = 2 Then
        cmndbtnEdit.Enabled = False
        cmndbtnBaixar.Enabled = False
        
    
    Else
        cmndbtnEdit.Enabled = True
        cmndbtnBaixar.Enabled = True
    End If
    If ado!status_che = "3" Then
        devChe
    Else
        lblDev.Visible = False
    End If
    
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBaixar_Click()
    Dim sSeq As String
    Dim iCount As Integer
    Dim c As cheque
    Dim sChe As String
    Dim sEsp As String
    Dim cValor As Currency
    Dim strSQL As String
    Dim sCod_Con  As String
    Dim sData As String
    With ado
    
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
                 "Baixa Cheque " & !num_che & "-" & !nome, sCod_Con, "", "000")
                
        If sChe <> "" Then
            'P_Class !Valor_che, Left(cmbxTipo.Text, 1), sChe
            strSQL = "update tab_cheques set dep_che = " & F_GravaData(Date) & ","
            strSQL = strSQL & " status_Che=2 "
            strSQL = strSQL & " where seq_che = '" & !seq_che & "' "
            Conn.Execute strSQL
            
            strSQL = "insert into tab_chequedep(seq_che,data_dep,cod_con)values("
            strSQL = strSQL & "'" & !seq_che & "',"
            strSQL = strSQL & F_GravaData(sData) & ","
            strSQL = strSQL & "'" & sCod_Con & "')"
            Conn.Execute strSQL
            
            MsgBox "Cheque baixado!", vbInformation
            
            ado.Requery
            
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
    strSQL = Cheques.F_selectChPag(Left(cmbxTipoTer.Text, 1), txtbxExpr.tag, txtbxDataIni.Text, txtbxdataFin.Text, _
        txtbxNumChe.Text, "", Left(cmbxSituacao.Text, 1), txtbxConta.tag, "", "", txtbxNumDoc.Text, sOrderBy)
    abreObjetoAdo ado, strSQL, grid
    
    soma
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub



Private Sub cmndbtnConta_Click()
    buscaConta
End Sub
Sub buscaConta()
    Dim c As ContaCorrente
    c = F_BuscaContaCorrente(txtbxConta.Text)
    If c.cod_con <> "" Then
        txtbxConta.Text = c.nome_con
        txtbxConta.tag = c.cod_con
    End If
End Sub
Private Sub cmndbtnDev_Click()
    On Error GoTo trata_Erro
    With ado
        If F_DevolveChe(!seq_che) Then
            ado.Requery
        End If
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub buscaTer()
    If cmbxTipoTer.ListIndex = 1 Then
        buscaRep
    ElseIf cmbxTipoTer.ListIndex = 2 Then
        buscaFun
    Else
        buscaCli
    End If
End Sub
Sub buscaCli()
    Dim c As cliente
    c = Clientes.F_BuscaCliente(txtbxExpr.Text)
    
    If c.cod_cli <> "" Then
        txtbxExpr.Text = c.nome_Cli
        txtbxExpr.tag = c.cod_cli
    End If
End Sub
Sub buscaRep()
    Dim c As representante.representante
    c = F_BuscaRep(txtbxExpr.Text)
    
    If c.cod_Rep <> "" Then
        txtbxExpr.Text = c.nome_rep
        txtbxExpr.tag = c.cod_Rep
    End If
End Sub
Sub buscaFun()
    Dim c As funcionario.funcionario
    c = funcionario.F_SelectFuncionario(txtbxExpr.Text)
    
    If c.cod_fun <> "" Then
        txtbxExpr.Text = c.nome_fun
        txtbxExpr.tag = c.cod_fun
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    Dim c As cheque
    Dim cp As chePag
    Dim strSQL As String
    With ado
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
                Conn.Execute strSQL
                MsgBox "cheque alterado!", vbInformation
                RefreshAdoPosition ado
            End If
        
    
    End With
End Sub

Private Sub cmndbtnImpr_Click()
    If ado.RecordCount <= 0 Then
        Exit Sub
    End If
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
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
    
    With ado
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
            P_ImprX obj, !data_doc, 235
            
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
    
    With ado
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

Sub devChe()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    lblDev.caption = ""
    With ado
        strSQL = "select * from tab_chequedev where seq_che='" & !seq_che & "' "
        abreObjetoAdo rPesq, strSQL
        
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
   
    
    Set ado = New ADODB.Recordset
    
    

    
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
    pcBotes.width = Me.width - 300
    pcBusca.width = pcBotes.width
    dbgrdChe.width = pcBotes.width
    dbgrdChePag.width = pcBotes.width
    
    pCdados.top = Me.height - pCdados.height - 400
    dbgrdChe.height = pCdados.top - dbgrdChe.top - 50
    dbgrdChePag.height = dbgrdChe.height
    pCdados.width = pcBotes.width
    cmndbtnSair.Left = pcBotes.width - cmndbtnSair.width - 100
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
        txtbxLocal.tag = l.cod_loc
        cmndbtnBuscar_Click
    End If
End Sub
Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        
        Sendkeys Chr(9)
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
