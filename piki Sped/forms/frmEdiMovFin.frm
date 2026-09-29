VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiMovFin 
   Caption         =   "Movimento de Financeiras"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   720
   ClientWidth     =   13785
   Icon            =   "frmEdiMovFin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   13785
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   13755
      TabIndex        =   19
      Top             =   6570
      Width           =   13785
      Begin VB.CheckBox chkAberto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apenas Parc. em Aberto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Total 0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2940
         TabIndex        =   21
         Top             =   0
         Width           =   1140
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   13755
      TabIndex        =   0
      Top             =   7065
      Width           =   13785
      Begin VB.CommandButton cmdEmLote 
         Caption         =   "Em Lote"
         Height          =   615
         Left            =   10080
         Picture         =   "frmEdiMovFin.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1260
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdComprovante 
         Caption         =   "Comprovante"
         Height          =   615
         Left            =   14040
         Picture         =   "frmEdiMovFin.frx":08FF
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1260
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   595
         Left            =   11400
         Picture         =   "frmEdiMovFin.frx":0C89
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtFinanceira 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   15360
         Picture         =   "frmEdiMovFin.frx":10B8
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1260
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   7440
         Picture         =   "frmEdiMovFin.frx":1302
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1260
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   12720
         Picture         =   "frmEdiMovFin.frx":144C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1260
      End
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   615
         Left            =   8760
         Picture         =   "frmEdiMovFin.frx":17D6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1260
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   615
         Left            =   6120
         Picture         =   "frmEdiMovFin.frx":1B60
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1260
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Valor"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4920
         TabIndex        =   9
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Doc./Aut."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3960
         TabIndex        =   7
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Financeira:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2280
         TabIndex        =   5
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Até:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&De:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
   End
   Begin MSDataGridLib.DataGrid dtgrid 
      Height          =   6375
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   11245
      _Version        =   393216
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
      Caption         =   "Movimento de Financeiras no Período"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "num_ven"
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
      BeginProperty Column01 
         DataField       =   "valor_fin"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "parcela_fin"
         Caption         =   "Parcela"
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
         DataField       =   "vencimento_fin"
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
      BeginProperty Column04 
         DataField       =   "operacao_fin"
         Caption         =   "Operação"
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
         DataField       =   "autorizacao_fin"
         Caption         =   "Autorização"
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
         DataField       =   "data_fin"
         Caption         =   "Data"
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
      BeginProperty Column07 
         DataField       =   "hora_fin"
         Caption         =   "Hora"
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
      BeginProperty Column08 
         DataField       =   "nome_fin"
         Caption         =   "Financeira"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   3179,906
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Opções"
      Begin VB.Menu mnPesq 
         Caption         =   "Pesquisar"
      End
      Begin VB.Menu mnReimp 
         Caption         =   "Reimprimir Comprovante"
      End
   End
End
Attribute VB_Name = "frmEdiMovFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim colAutorizacao As Integer
Dim colParcela As Integer
Dim colValor As Integer

Private Sub chkAberto_Click()
    If chkAberto.Value = 1 Then
        cmndbtnBaixar.Enabled = True
    Else
        cmndbtnBaixar.Enabled = False
    End If
End Sub


Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmdComprovante_Click()
    On Error GoTo trata_Erro
    ImprimeComprovanteNaoFiscalTEF ado!msg
    Exit Sub
trata_Erro:
MsgBox Err.Description

    
End Sub

Private Sub cmdEmLote_Click()
    frmBaixaFinanceirasLote.Show 1
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV dtgrid, ado
End Sub

Private Sub cmndbtnBaixar_Click()
    Dim iCount As Integer
    Dim strSQL As String
    Dim sNumVen As String
    Dim sValor As String
    Dim sSeq_Fin As String
    Dim parcela As String
    
    If strEditarFin_Usu <> "S" And strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    If IsLoad("frmListaDoc") Then
        If frmListaDoc.tag = Me.Name Then
            If MsgBox("Confirma baixa de todas as parcelas listadas no valor de " _
                & frmListaDoc.lblTotal.caption & "?", vbYesNo + vbQuestion) = vbYes Then
                If frmListaDoc.txtbxConta.tag = "" Then
                    MsgBox "Você deve informar uma conta corrente para o crédito do valor!", vbInformation
                    frmListaDoc.txtbxConta.SetFocus
                    Exit Sub
                    
                End If
                
                sSeq_Fin = F_GravaFinanceiro("C", Date, frmListaDoc.lblTotal.caption, _
                "Compensação Financeira " & ado!nome_fin, frmListaDoc.txtbxConta.tag, _
                 "N", "001")
                 
                 
                 If sSeq_Fin = "" Then
                    MsgBox "O movimento financeiro não foi gravado, tente novamente!", vbExclamation
                    Exit Sub
                 End If
                 
                
                With frmListaDoc.msflxgrd
                    For iCount = 0 To .rows - 1
                        .row = iCount
                        .col = 0
                        sNumVen = .Text
                        .col = 2
                        sValor = .Text
                         .col = 3
                        parcela = .Text
                        strSQL = "update " & TAB_VENFIN & "    set reman_fin=reman_fin-" & F_GravaDec(sValor)
                        
                        
                        strSQL = strSQL & " where num_ven='" & sNumVen & "' "
                        
                        strSQL = strSQL & " and parcela_fin=" & parcela
                        
                        conn.Execute strSQL
                        strSQL = F_GravaBaixaVenFin(sNumVen, sValor, "DIN", sSeq_Fin)
                        
                    Next
                    
                    
                    
                    
                Unload frmListaDoc
                ado.Requery
                End With
            End If
        End If
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    dtgrid.Visible = False
    strSQL = "select tf.nome_fin,tvf.* from tab_financeiras as tf, "
    strSQL = strSQL & TAB_VENFIN & "   as tvf where tvf.cod_fin=tf.cod_fin "
    
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tvf.data_fin>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tvf.data_fin<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If chkAberto.Value = 1 Then
        strSQL = strSQL & " and tvf.reman_fin>0 "
    End If
    If txtDoc.Text <> "" Then
        strSQL = strSQL & " and (tvf.num_ven='" & txtDoc.Text & "' or tvf.autorizacao_fin='" & txtDoc.Text & "')"
    End If
    If txtFinanceira.tag <> "" Then
        strSQL = strSQL & " and tvf.cod_fin='" & txtFinanceira.tag & "' "
    End If
    If txtValor.Text <> "" Then
        strSQL = strSQL & " and tvf.valor_fin=" & F_GravaDec(txtValor.Text)
    End If
    
        strSQL = strSQL & " order by  data_fin desc, hora_fin desc,nome_fin"
    
    abreObjetoAdo ado, strSQL, dtgrid
    If ado.RecordCount > 0 Then
        cmndbtnBaixar.Enabled = True
        cmndbtnEdit.Enabled = True
        cmndbtnImpr.Enabled = True
    Else
        cmndbtnBaixar.Enabled = False
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
    End If
    soma
    dtgrid.Visible = True
End Sub
Sub soma()
    Dim cTotal As Currency
    With ado
        Do While Not .EOF
            cTotal = cTotal + !reman_fin
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    lblTotal.caption = "Total " & FormatVal(cTotal)
End Sub

Private Sub cmndbtnEdit_Click()
    On Error Resume Next
    EditaVenFin ado!num_ven
    ado.Requery
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dtgrid_DblClick()
    With ado
    Dim i As Integer
    Dim r As Integer
        If !reman_fin = 0 Then
            Exit Sub
        End If
        If IsLoad("frmListaDoc") Then
            If frmListaDoc.tag <> Me.Name Then
                MsgBox "A lista está sendo usada por outro formulário!", vbExclamation
                Exit Sub
            End If
            i = F_InFlxgrd(frmListaDoc.msflxgrd, !num_ven, 0)
            If i = -1 Then
                
                frmListaDoc.msflxgrd.AddItem !num_ven & Chr(9) & !reman_fin & Chr(9) & !reman_fin & Chr(9) & !parcela_fin
            
            Else
                frmListaDoc.msflxgrd.row = i
                frmListaDoc.msflxgrd.col = 3
                If frmListaDoc.msflxgrd.Text <> !parcela_fin Then
                    frmListaDoc.msflxgrd.AddItem !num_ven & Chr(9) & !reman_fin & Chr(9) & !reman_fin & Chr(9) & !parcela_fin
                End If
            End If
            
            
            frmListaDoc.soma
        Else
            frmListaDoc.tag = Me.Name
            frmListaDoc.caption = "Documentos de Financeiras"
            frmListaDoc.msflxgrd.Cols = 4
            frmListaDoc.msflxgrd.rows = 0
            frmListaDoc.msflxgrd.AddItem !num_ven & Chr(9) & !reman_fin & Chr(9) & !reman_fin & Chr(9) & !parcela_fin
            frmListaDoc.Show 0, Me
            frmListaDoc.soma
            
        End If
    
    End With
End Sub


Private Sub dtgrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrid_DblClick
    End If
End Sub

Private Sub dtgrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        dtgrid_DblClick
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    txtbxdataFin.Text = F_UltDiaMes
    txtbxDataIni.Text = F_PrimDiaMes
    resizeForm Me
End Sub

Private Sub Form_Resize()
    PosicionaGrade dtgrid, Me, pcBot
    posicionaADireita cmndbtnSair, pcBot
End Sub


Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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

Private Sub txtDoc_GotFocus()
    P_SelecTexto txtDoc
End Sub

Private Sub txtDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtFinanceira_Change()
    txtFinanceira.tag = ""
End Sub

Private Sub txtFinanceira_GotFocus()
    P_SelecTexto txtFinanceira
End Sub

Private Sub txtFinanceira_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFin
    End If
End Sub
Sub buscaFin()
    Dim f As Financeira.Financeira
    f = F_BuscaFinanceira(txtFinanceira.Text)
    If f.cod_fin <> "" Then
        txtFinanceira.Text = f.nome_fin
        txtFinanceira.tag = f.cod_fin
        busca
    End If
    
End Sub

Private Sub txtValor_GotFocus()
    P_SelecTexto txtValor
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
