VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelBaixaItd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação de Recebimentos"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13905
   Icon            =   "frmRelBaixaItd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   13905
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   13875
      TabIndex        =   15
      Top             =   0
      Width           =   13905
      Begin VB.TextBox txtbxTipoCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtbxContaCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbxEnd 
         Height          =   315
         ItemData        =   "frmRelBaixaItd.frx":2CFA
         Left            =   6480
         List            =   "frmRelBaixaItd.frx":2D07
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmRelBaixaItd.frx":2D35
         Left            =   8280
         List            =   "frmRelBaixaItd.frx":2D3F
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmRelBaixaItd.frx":2D5B
         Left            =   10680
         List            =   "frmRelBaixaItd.frx":2D65
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtbxGrupoCli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         MaxLength       =   60
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtbxSeqDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         MaxLength       =   60
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxSeqItd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         MaxLength       =   60
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxTerceiro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         MaxLength       =   10
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   570
         Left            =   10680
         Picture         =   "frmRelBaixaItd.frx":2D7F
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnCob 
         Height          =   285
         Left            =   1800
         Picture         =   "frmRelBaixaItd.frx":2E78
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnContaCob 
         Height          =   285
         Left            =   4200
         Picture         =   "frmRelBaixaItd.frx":3202
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnGrupo 
         Height          =   285
         Left            =   6000
         Picture         =   "frmRelBaixaItd.frx":358C
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnTerc 
         Height          =   285
         Left            =   4200
         Picture         =   "frmRelBaixaItd.frx":3916
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblTipoCob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cobrança:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Cobrança (Baixada):"
         Height          =   195
         Left            =   2160
         TabIndex        =   43
         Top             =   0
         Width           =   1905
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listar Endereço:"
         Height          =   195
         Left            =   6480
         TabIndex        =   42
         Top             =   0
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         Height          =   195
         Left            =   8280
         TabIndex        =   41
         Top             =   0
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar Por:"
         Height          =   195
         Left            =   10680
         TabIndex        =   40
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Cientes:"
         Height          =   195
         Left            =   4560
         TabIndex        =   39
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         Height          =   195
         Left            =   6960
         TabIndex        =   38
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. Doc:"
         Height          =   195
         Left            =   4560
         TabIndex        =   36
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. Parcela:"
         Height          =   195
         Left            =   5760
         TabIndex        =   35
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   8280
         TabIndex        =   34
         Tag             =   "!"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   9480
         TabIndex        =   33
         Tag             =   "!"
         Top             =   600
         Width           =   285
      End
   End
   Begin VB.Data dtNot 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FEF4E4&
      ForeColor       =   &H00FEF4E4&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   13875
      TabIndex        =   0
      Top             =   7455
      Width           =   13905
      Begin VB.CommandButton cmndbtnCancelaBaixa 
         Caption         =   "&Cancelar Baixa"
         Height          =   685
         Left            =   6840
         Picture         =   "frmRelBaixaItd.frx":3CA0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnFechar 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   685
         Left            =   11400
         Picture         =   "frmRelBaixaItd.frx":402A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   685
         Left            =   9120
         Picture         =   "frmRelBaixaItd.frx":4195
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   2175
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2040
         ScaleHeight     =   855
         ScaleWidth      =   2175
         TabIndex        =   7
         Top             =   120
         Width           =   2175
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FEF8ED&
            Height          =   195
            Left            =   855
            TabIndex        =   9
            Tag             =   "!"
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Recebido:"
            ForeColor       =   &H00FEF8ED&
            Height          =   195
            Left            =   480
            TabIndex        =   8
            Tag             =   "!"
            Top             =   120
            Width           =   1140
         End
      End
      Begin VB.Label lblJuros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   4560
         TabIndex        =   6
         Tag             =   "!"
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   5520
         TabIndex        =   5
         Tag             =   "!"
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblValDoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   915
         TabIndex        =   4
         Tag             =   "!"
         Top             =   450
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total das Parcelas:"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   390
         TabIndex        =   3
         Tag             =   "!"
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Juros:"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   4440
         TabIndex        =   2
         Tag             =   "!"
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descontos:"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   5400
         TabIndex        =   1
         Tag             =   "!"
         Top             =   120
         Width           =   810
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   1695
      Left            =   0
      TabIndex        =   10
      Top             =   5760
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   2990
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelBaixaItd.frx":45A2
      Height          =   4215
      Left            =   120
      OleObjectBlob   =   "frmRelBaixaItd.frx":45B6
      TabIndex        =   11
      Top             =   1440
      Width           =   13695
   End
End
Attribute VB_Name = "frmRelBaixaItd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rPesq As Recordset

Private Sub cmndbrnTransf_Click()
    If txtbxTransfCob.tag = "" Then
        MsgBox "Escolha a conta que deseja Transferir", vbExclamation
        txtbxTransfCob.SetFocus
        Exit Sub
        
    End If
    If MsgBox("Confirma Transferência dos " & lblSelect.caption & " para " & txtbxTransfCob.Text & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    transfere
End Sub
Sub transfere()
    Dim strSQL As String
    Dim iCount As Integer
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 1
            If .Text = "S" Then
                
                strSQL = "update tab_itdoc set cod_cob='" & txtbxTransfCob.tag & "' "
                .Col = 3
                strSQL = strSQL & "where seq_itd='" & .Text & "' "
                Banco.Execute strSQL
            End If
        Next
        
    End With
    busca
    
    
End Sub
Private Sub cmndbtnBuscaCobTransf_Click()
        Dim c As ContaCob
    c = F_BuscaContaCob(txtbxTransfCob.Text)
    txtbxTransfCob.Text = c.nome_cob
    txtbxTransfCob.tag = c.cod_cob

End Sub

Private Sub cmndbtnBuscar_Click()
    busca
    
    
End Sub
Sub busca()
    Dim strSQL As String
    Dim s As setor
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sTipoTer As String
    Dim sEnd As String
    
    
    sOrderby = "venc_Itd"
    
    
    
    
    strSQL = "select tcb.nome_cob,"
    strSQL = strSQL & ",td.historico_doc,td.parcelas_doc, tf.valor_fin,tf.hora_fin,"
    strSQL = strSQL & "td.valor_Doc,td.num_doc,td.data_doc, td.valor_doc,"
    
    strSQL = strSQL & " ti.* "
    
    strSQL = strSQL & "from tab_doc as td,tab_baixaitd2 as tb, "
    strSQL = strSQL & "tab_itdoc as ti,tab_contacob as tcb,tab_Financeiro as tf  "
    
    strSQL = strSQL & "where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and ti.seq_itd=tb.seq_itd "
    strSQL = strSQL & " and tb.seq_fin=tf.seq_fin "
    strSQL = strSQL & " and ti.cod_cob=tcb.cod_cob "
    strSQL = strSQL & "and td.tipo_doc='R' "
    If cmbxTipoTer.ListIndex > 0 Then
        strSQL = strSQL & " and td.tipo_ter='" & Left(cmbxTipoTer.Text, 1) & "' "
    End If
    
    If txtbxGrupoCli.tag <> "" And cmbxTipoTer.ListIndex = 1 Then
        
        strSQL = strSQL & " and tc.cod_gru='" & txtbxGrupoCli.tag & "' "
        
    End If
    
    
    If txtbxContaCob.tag <> "" Then
        strSQL = strSQL & " and ti.cod_cob='" & txtbxContaCob.tag & "' "
    End If
    If txtbxDataFin.Text <> "" Then
        strSQL = strSQL & " and ti.venc_itd <= " & F_GravaData(txtbxDataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and ti.venc_itd >= " & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxTipoCob.tag <> "" Then
        strSQL = strSQL & " and ti.cod_tip='" & txtbxTipoCob.tag & "' "
    End If
    If txtbxNumDoc.Text <> "" Then
        strSQL = strSQL & " and num_doc='" & txtbxNumDoc.Text & "' "
    End If
    If txtbxSeqDoc.Text <> "" Then
        strSQL = strSQL & " and td.seq_doc='" & txtbxSeqDoc.Text & "' "
    End If
    If txtbxSeqItd.Text <> "" Then
        strSQL = strSQL & " and ti.seq_itd='" & txtbxSeqDoc.Text & "' "
    End If
    If txtbxTerceiro.tag <> "" Then
        strSQL = strSQL & " and td.cod_ter='" & txtbxTerceiro.tag & "' "
    End If
    
    strSQL = strSQL & " order by " & sOrderby
    

    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = .AbsolutePosition + 1 & Chr(9)
            strSQL = strSQL & Chr(9)
            strSQL = strSQL & !num_doc & Chr(9)
            strSQL = strSQL & !seq_itd & Chr(9)
            strSQL = strSQL & !parcela_itd & "/" & !parcelas_doc & Chr(9)
            strSQL = strSQL & !venc_itd & Chr(9)
            strSQL = strSQL & FormatVal(!valor_itd) & Chr(9)
            strSQL = strSQL & FormatVal(!reman_itd) & Chr(9)
            strSQL = strSQL & !Codigo & Chr(9)
            strSQL = strSQL & !Nome & Chr(9)
            '
            
            If cmbxEnd.ListIndex = 0 Then
                If !cobranca_cli = "2" Then
                    s = F_RetornaSetorTrab(!Codigo)
                    strSQL = strSQL & s.Cidade & Chr(9)
                    strSQL = strSQL & s.Nome & Chr(9)
                    strSQL = strSQL & !Endereco & Chr(9)
                    strSQL = strSQL & "T" & Chr(9)
                Else
                    strSQL = strSQL & !nome_cid & Chr(9)
                    strSQL = strSQL & !nome_set & Chr(9)
                    strSQL = strSQL & !Endereco & Chr(9)
                    strSQL = strSQL & "R" & Chr(9)
                End If
            ElseIf cmbxEnd.ListIndex = 1 Then
                strSQL = strSQL & !Endereco & Chr(9)
                strSQL = strSQL & !nome_set & Chr(9)
                strSQL = strSQL & !Endereco & Chr(9)
                strSQL = strSQL & "R" & Chr(9)
            Else
                
                    s = F_RetornaSetorTrab(!Codigo)
                    strSQL = strSQL & s.Cidade & Chr(9)
                    strSQL = strSQL & s.Nome & Chr(9)
                    strSQL = strSQL & !Endereco & Chr(9)
                    strSQL = strSQL & "T" & Chr(9)
                
            End If
            
            strSQL = strSQL & F_BuscaTel(!Codigo, sTipoTer) & Chr(9)
            strSQL = strSQL & !cod_tip & Chr(9)
            strSQL = strSQL & !nome_cob
            msflxgrd.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    lblListados.caption = msflxgrd.rows - 1
    'ZebraGrid  msflxgrd, 1
    
End Sub

Private Sub cmndbtnCancelaBaixa_Click()
    On Error GoTo trata_Erro
    Dim strSQL As String
    If strExcldoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirma cancelamento da baixa?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    With dtNot.Recordset
        strSQL = "delete from tab_baixaitd where seq_cab='" & !seq_cab & "'"
        Banco.Execute strSQL
        
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub

Private Sub cmndbtnCob_Click()
    Dim t As tipocob
    t = F_BuscaTipoCob(txtbxTipoCob.Text)
    txtbxTipoCob.Text = t.nome_tip
    txtbxTipoCob.tag = t.cod_tip
    
    
End Sub

Private Sub cmndbtnContaCob_Click()
    Dim c As ContaCob
    c = F_BuscaContaCob(txtbxContaCob.Text)
    txtbxContaCob.Text = c.nome_cob
    txtbxContaCob.tag = c.cod_cob
    
    
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub



Private Sub cmndbtnGrupo_Click()
    Dim g As grupoCli
    g = F_BuscaGrupoCli()
    
    txtbxGrupoCli.Text = g.nome_gru
    txtbxGrupoCli.tag = g.cod_gru
End Sub

Private Sub cmndbtnImprimir_Click()
    Dim strSQL As String
    strSQL = "Tipo:[" & txtbxTipoCob.Text & "], "
    strSQL = strSQL & "Conta:[" & txtbxContaCob.Text & "], "
    strSQL = strSQL & "Grupo:[" & txtbxGrupoCli.Text & "]," & Chr(9)
    strSQL = strSQL & "Endereço:[" & cmbxEnd.Text & "], " & Chr(9)
    strSQL = strSQL & "Tipo Terc.:[" & cmbxTipoTer.Text & "],"
    strSQL = strSQL & "De:[" & txtbxDataIni.Text & "], "
    strSQL = strSQL & "Até:[" & txtbxDataFin.Text & "]"
    
    'P_MsflxToHtml "Listagem de Cobrança", msflxgrd, strSQL
End Sub


Private Sub Form_Load()
    cmbxTipoTer.ListIndex = 0
    cmbxEnd.ListIndex = 0
    resizeForm Me
    txtbxDataFin.Text = F_UltDiaMes(Date)
    txtbxDataIni.Text = F_PrimDiaMes(Date)
    ajustaGrid
End Sub

Private Sub Form_Resize()
'    msflxgrd.Width = Me.Width - msflxgrd.Left - 200
'    msflxgrd.Height = pcBot.Top - msflxgrd.Top - 100
 '   pcBotoes.Left = pcBot.Width - pcBotoes.Width - 100
    
End Sub
Sub ajustaGrid()
    With msflxgrd
        .rows = 1
        .Cols = 17
        .FixedCols = 2
        .Col = 0
        
        .Col = 0: .Text = "..."
        .Col = 1: .Text = "S"
        .Col = 2: .Text = "Doc"
        .Col = 3: .Text = "Itdoc"
        .Col = 4: .Text = "Parcela"
        .Col = 5: .Text = "Vencimento"
        .Col = 6: .Text = "Valor"
        .Col = 7: .Text = "Restam"
        .Col = 8: .Text = "Código"
        .Col = 9: .Text = "Cliente"
        .Col = 10: .Text = "Cidade"
        .Col = 11: .Text = "Setor"
        .Col = 11: .Text = "Setor"
        .Col = 12: .Text = "Endereço"
        .Col = 13: .Text = "T"
        .Col = 14: .Text = "Fones"
        .Col = 15: .Text = "Cob."
        .Col = 16: .Text = "Conta Cob."
        
        .ColWidth(0) = 0
        .ColWidth(1) = 300
        .ColWidth(2) = 700
        .ColWidth(3) = 700
        .ColWidth(4) = 600
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 600
        .ColWidth(9) = 2600
        .ColWidth(10) = 1000
        .ColWidth(11) = 2600
        .ColWidth(12) = 2600
        .ColWidth(13) = 500
        .ColWidth(14) = 2600
        .ColWidth(15) = 300
        .ColWidth(16) = 1300
        .ColAlignment(14) = 1
        
    End With
End Sub

Private Sub msflxgrd_DblClick()
    
    With msflxgrd
        .Col = 1
        If .Text = "" Then
            
            'ZebraLinha msflxgrd, msflxgrd.Row
            .Col = 1
            .CellBackColor = vbBlue
            .CellForeColor = vbBlue
            .Text = "S"
            lblSelect.caption = Int(lblSelect.caption) + 1
        Else
            .Text = ""
            'ZebraLinha msflxgrd, msflxgrd.Row, vbWhite
            .CellBackColor = vbWhite
            lblSelect.caption = Int(lblSelect.caption) - 1
        End If
        
    End With
End Sub

Private Sub txtbxContaCob_Change()
    txtbxContaCob.tag = ""
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDataFin
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

Private Sub txtbxGrupoCli_Change()
    txtbxGrupoCli.tag = ""
End Sub

Private Sub cmndbtnTerc_Click()
    
    
        If cmbxTipoTer.ListIndex = 0 Then
            buscaCli
        Else
            buscaRep
        End If
    
    
    cmndbtnBuscar_Click
End Sub
Sub buscaCli()
    Dim c As Cliente
    c = F_BuscaCliente(txtbxTerceiro.Text)
        If c.cod_cli <> "" Then
            txtbxTerceiro.Text = c.nome_Cli
            txtbxTerceiro.tag = c.cod_cli
        End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
        f = F_SelectForn(txtbxTerceiro.Text)
        If f.cod_for <> "" Then
            txtbxTerceiro.Text = f.nome_for
            txtbxTerceiro.tag = f.cod_for
        End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxTerceiro.Text)
        If r.cod_rep <> "" Then
            txtbxTerceiro.Text = r.nome_rep
            txtbxTerceiro.tag = r.cod_rep
        End If
    
End Sub

Private Sub txtbxNumDoc_GotFocus()
    P_SelecTexto txtbxNumDoc
End Sub

Private Sub txtbxNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNumDoc_LostFocus()
    P_AjustaTexto txtbxNumDoc
End Sub

Private Sub txtbxSeqDoc_GotFocus()
    P_SelecTexto txtbxSeqDoc
End Sub

Private Sub txtbxSeqDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSeqDoc_LostFocus()
    P_AjustaNota txtbxSeqDoc, 7
End Sub

Private Sub txtbxSeqItd_GotFocus()
    P_SelecTexto txtbxSeqItd
End Sub

Private Sub txtbxSeqItd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSeqItd_LostFocus()
    P_AjustaNota txtbxSeqItd, 7
End Sub

Private Sub txtbxTerceiro_Change()
    txtbxContaCob.tag = c.cod_cob
End Sub

Private Sub txtbxTipoCob_Change()
    txtbxTipoCob.Text = ""
End Sub

Private Sub txtbxTipoCob_LostFocus()
    If txtbxTipoCob.Text = "" Then txtbxTipoCob.tag = ""
End Sub

