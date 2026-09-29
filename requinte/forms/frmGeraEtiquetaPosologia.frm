VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGeraEtiquetasPosologia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Geração de Etiquetas"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9480
   Icon            =   "frmGeraEtiquetaPosologia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   6600
      Picture         =   "frmGeraEtiquetaPosologia.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "&Sair"
      Height          =   615
      Left            =   7920
      Picture         =   "frmGeraEtiquetaPosologia.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   7575
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2520
         ScaleHeight     =   825
         ScaleWidth      =   2385
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   7305
         TabIndex        =   9
         Top             =   240
         Width           =   7335
         Begin VB.Label lblDescEti 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Etiqueta"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   0
            Width           =   585
         End
      End
      Begin VB.TextBox txtbxIni 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Text            =   "1"
         Top             =   3120
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4260
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.CommandButton cmndbtnEdit 
         Height          =   375
         Left            =   7080
         Picture         =   "frmGeraEtiquetaPosologia.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnExcl 
         Height          =   375
         Left            =   6600
         Picture         =   "frmGeraEtiquetaPosologia.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnLimpa 
         Height          =   375
         Left            =   6120
         Picture         =   "frmGeraEtiquetaPosologia.frx":0672
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   600
         Picture         =   "frmGeraEtiquetaPosologia.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salvar Modelo"
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   120
         Picture         =   "frmGeraEtiquetaPosologia.frx":0906
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Abrir um modelo pre-definido"
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar em:"
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         Top             =   3120
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Adicionar Posologia"
      Height          =   1335
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton cmndbtnAdd 
         Caption         =   "&Adicionar"
         Height          =   495
         Left            =   7200
         Picture         =   "frmGeraEtiquetaPosologia.frx":0A50
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6000
         MaxLength       =   40
         TabIndex        =   33
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   675
         Left            =   120
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label lblTotalEti 
         AutoSize        =   -1  'True
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
         Left            =   6360
         TabIndex        =   40
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Total de Etiquetas:"
         Height          =   195
         Left            =   4920
         TabIndex        =   39
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
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
         Left            =   7680
         TabIndex        =   38
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Páginas:"
         Height          =   195
         Left            =   6960
         TabIndex        =   37
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&qtd:"
         Height          =   195
         Left            =   6000
         TabIndex        =   36
         Top             =   240
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3495
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1455
      Begin VB.TextBox txtbxLarg 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   21
         Text            =   "45"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtbxAlt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   20
         Text            =   "12,7"
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtbxMargEsq 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   19
         Text            =   "14"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtbxMarS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   18
         Text            =   "13"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtbxQtdCol 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   17
         Text            =   "4"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtbxDistH 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   16
         Text            =   "3"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtbxDistV 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   15
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtbxQtdLinha 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   14
         Text            =   "20"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtbxCopias 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   13
         Text            =   "1"
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Largura:"
         Height          =   195
         Left            =   195
         TabIndex        =   30
         Top             =   2835
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Altura:"
         Height          =   195
         Left            =   330
         TabIndex        =   29
         Top             =   2475
         Width           =   450
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Marg.Esq:"
         Height          =   195
         Left            =   60
         TabIndex        =   28
         Top             =   2108
         Width           =   720
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Marg.Sup:"
         Height          =   195
         Left            =   45
         TabIndex        =   27
         Top             =   1755
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Qtd Col:"
         Height          =   195
         Left            =   210
         TabIndex        =   26
         Top             =   308
         Width           =   570
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dist.Hor.:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1028
         Width           =   660
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dist.Vert:"
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   1388
         Width           =   645
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd Linha:"
         Height          =   195
         Left            =   45
         TabIndex        =   23
         Top             =   668
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cópias:"
         Height          =   195
         Left            =   255
         TabIndex        =   22
         Top             =   3195
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmGeraEtiquetasPosologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmbxBusca_Click()
    cmbxBusca.tag = Left(cmbxBusca.Text, 1)
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxCopia_Click()
    cmbxCopia.tag = F_BuscaUmCampo("tipo_cop", "tab_copias", "cod_cop='" & cmbxCopia.Text & "'")
End Sub

Private Sub cmndbtnAdd_Click()
    Dim strSQL As String
    If txtbxExpr.tag = "" Then
'        MsgBox "Escolha primeiramente o Produto!", vbExclamation
'        txtbxExpr.SetFocus
 '       Exit Sub
    End If
    If CCur(txtbxQtd.Text) = 0 Then
        MsgBox "Informe uma quantidade válida!", vbExclamation
        txtbxQtd.SetFocus
        Exit Sub
    End If
    If cmndbtnAdd.tag = "" Then
       strSQL = txtbxQtd.Text & Chr(9)
       
       strSQL = strSQL & txtbxExpr.Text & Chr(9)
       
       
       msflxGrd.AddItem strSQL
       lblTotalEti.caption = Int(lblTotalEti.caption) + Int(txtbxQtd.Text)
       Inf
    Else
        With msflxGrd
            .Row = Int(cmndbtnAdd.tag)
            .col = 0: .Text = txtbxQtd.Text
            .col = 1: .Text = FormatCodigo(txtbxExpr.tag)
            .col = 2: .Text = txtbxExpr.Text
            .col = 3: .Text = lblCodFab.caption
            .col = 4: .Text = lblPreco.caption
            cmndbtnAdd.tag = ""
        End With
    End If
    txtbxExpr.SetFocus
End Sub
Sub Inf()
    On Error Resume Next
    Dim iTot As Single
    Dim iAtual As Single
    iAtual = CSng(lblTotalEti.caption)
    iTot = Int(txtbxQtdCol.Text) * Int(txtbxQtdLinha.Text)
    If iAtual = 0 Then
        lblPaginas.caption = "0"
        Exit Sub
    End If
    If iAtual <= iTot Then
        lblPaginas.caption = 1
        Exit Sub
    End If
    If iTot Mod iAtual = 0 Then
        lblPaginas.caption = iTot / iAtual
    Else
        lblPaginas.caption = Int(iAtual / iTot) + 1
    End If
End Sub
Private Sub cmndbtnConf_Click()
    'On Error Resume Next
    If optnEstoq.Value = False And cmbxBusca.ListIndex = -1 Then
        MsgBox "A buscar se encontra incompleta!", vbExclamation
        Exit Sub
        cmbxBusca.SetFocus
    End If
    If MsgBox("Confirma busca?", vbYesNo + vbQuestion) Then
        If optnEnt.Value Then
            BuscaEntrada
        Else
            BuscaAvancada
        End If
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    If msflxGrd.rows = 1 Then Exit Sub
    With msflxGrd
        .col = 0:  txtbxQtd.Text = .Text
        .col = 1: txtbxQtd.tag = F_SoNumero(.Text)
        .col = 2: lblCodFab.caption = .Text
        .col = 3: txtbxExpr.Text = .Text
        .col = 4: lblMarca.caption = .Text
        .col = 5: lblEnd.caption = .Text
        .col = 6: lblPreco.caption = FormatVal(.Text)
        cmndbtnAdd.tag = msflxGrd.Row
        txtbxExpr.SetFocus
    End With
End Sub

Private Sub cmndbtnExcl_Click()
    Dim iAux As Integer
    If msflxGrd.rows = 1 Then Exit Sub
    
    If MsgBox("Confirma exclusão da linha?", vbYesNo + vbQuestion) = vbYes Then
        msflxGrd.col = 0
        iAux = msflxGrd.Text
        
        If msflxGrd.rows = 2 Then
            msflxGrd.Row = 1
        Else
            msflxGrd.RemoveItem msflxGrd.Row
        End If
        lblTotalEti.caption = Int(lblTotalEti.caption) - iAux
    End If
End Sub

Private Sub cmndbtnImprimir_Click()
    ImprimeEti frmVisualiza.Pic, CSng(txtbxAlt.Text), CSng(txtbxLarg.Text), CSng(txtbxDistH.Text), _
            CSng(txtbxDistV.Text), CSng(txtbxQtdCol.Text)
    frmVisualiza.Show 1
    
    If frmVisualiza.tag <> "" Then
        ImprimeEti Printer, CSng(txtbxAlt.Text), CSng(txtbxLarg.Text), CSng(txtbxDistH.Text), _
            CSng(txtbxDistV.Text), CSng(txtbxQtdCol.Text)
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub

Sub ImprimeEti(obj As Object, Optional sAlt As Single = 20, Optional sLarg As Single = 89, _
     Optional sDistH As Single = 2, Optional sDistV As Single = 2, Optional IQtdFolha As Integer = 2)
    Dim iCount As Integer
    Dim iQtd As Integer
    Dim iLinha As Integer
    Dim iEti As Integer
    Dim sMar As String
    Dim sNome As String
    Dim sCod As String
    Dim sRef As String
    Dim sEnd As String
    Dim sPre As String
    Dim sPos As Single
    Dim bCont As Boolean
    Dim sYIni As Single
    Dim sInicioX As Single
    Dim sInicioY As Single
    Dim sAuxY As Single
    Dim iImprimiu As Integer
    Dim iPage As Integer
    Dim sCentro As Single, sQtdLinha As Single
    Dim sUltY As Single, iAux As Integer, iPos As Integer, iPrimEsp As Integer
    Dim iPulaLin As Integer
    Dim iPulaCol As Integer
    Dim iRest As Integer
    Dim c As New Collection
    Set c = Nothing
    Dim jCount As Integer
 '   iPulaLin = Int(txtbxIni.Text) / Int(txtbxQtdCol.Text)
        
   
        
    
'        iPulaCol = (Int(txtbxIni.Text) Mod Int(txtbxQtdCol.Text)) - 1
        msflxGrd.AddItem CStr(Int(txtbxIni.Text) - 1), 1
    
        
    
        obj.ScaleMode = vbMillimeters
'        On Error GoTo trata_erro
        
        If TypeName(obj) = "Printer" Then
            sQtdLinha = CSng(txtbxQtdLinha.Text)
            'obj.PaperSize = vbPRPSUser
            'obj.ScaleHeight = sAlt * (sQtdLinha) + sDistV * (sQtdLinha) + CSng(txtbxMarS.Text)
            'obj.ScaleWidth = sLarg * (IQtdFolha) + sDistH * (IQtdFolha) + CSng(txtbxMargEsq.Text)
        End If
        sInicioX = Int(txtbxMargEsq.Text)
        sInicioY = Int(txtbxMarS.Text)
        iImprimiu = iPulaLin
        With msflxGrd
        
        obj.FontName = "arial narrow"
        obj.fontsize = 6
            iPos = iPulaCol
            For iLinha = 1 To .rows - 1
                .Row = iLinha
                .col = 0
                iEti = CInt(.Text) 'quantidade de etiquetas
                
                    For iAux = 1 To iEti
                        
                        .col = 1: sNome = .Text
                        
                        
                        sPos = sLarg * (iPos) + (sDistH * (iPos)) + sInicioX
                        
                        sAuxY = sInicioY + ((sAlt) * (iImprimiu)) + (sDistV * (iImprimiu))
                        
                        
                        
                        obj.CurrentY = sAuxY
                        obj.fontsize = 12
                        P_ImprQuebra sNome, obj, sPos, sLarg + sPos
                        
                        
                        
                        
                        
                        
                        iPos = iPos + 1
                        
                        If iPos >= IQtdFolha Then
                            iPos = 0
                            iImprimiu = iImprimiu + 1
                            If iImprimiu >= Int(txtbxQtdLinha.Text) Then
                                iImprimiu = CInt(txtbxIni.Text) - 1
                                If TypeName(obj) = "Printer" Then
                                    obj.NewPage
                                End If
                            End If
                        End If
                    
                Next
            
            Next
        End With
    msflxGrd.RemoveItem 1
Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnLimpa_Click()
    If msflxGrd.rows = 1 Then Exit Sub
    If MsgBox("Confirma exclusão de todas as linhas?", vbYesNo + vbQuestion) = vbYes Then
        msflxGrd.rows = 1
        lblTotalEti.caption = "0"
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub Command1_Click()
   ' Dim e As etiqueta
  '  e = F_BuscaEti
    If e.seq_eti <> "" Then
        lblDescEti.caption = e.desc
        txtbxAlt.Text = e.altura
        txtbxDistH.Text = e.distH
        txtbxDistV.Text = e.distV
        txtbxLarg.Text = e.largura
        txtbxMargEsq.Text = e.mEsq
        txtbxMarS.Text = e.mSup
        txtbxQtdCol.Text = e.qtdCol
        txtbxQtdLinha.Text = e.qtdLinha
    End If
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    If bActive Then Exit Sub
    
    If Me.tag <> "" Then
'        cmbxComp.Text = Me.tag
        BuscaEntrada
    End If
    bActive = True
End Sub

Private Sub Form_Load()
    bActive = False
    P_CentraForm Me
 '   cmbxBusca.ListIndex = 0
    msflxGrd.Cols = 2
    msflxGrd.FixedCols = 0
    msflxGrd.FixedRows = 1
    msflxGrd.rows = 1
    msflxGrd.col = 0
    msflxGrd.Text = "Qtd"
    msflxGrd.col = 1
    msflxGrd.Text = "Descrição"
    
    msflxGrd.ColWidth(0) = 200
    msflxGrd.ColWidth(1) = 4000
    
    
    
End Sub




Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If msflxGrd.rows = 2 Then
            msflxGrd.rows = 1
        ElseIf msflxGrd.rows > 2 Then
            msflxGrd.RemoveItem msflxGrd.Row
        End If
    End If
End Sub

Private Sub Option4_Click()

End Sub

Private Sub optnEstoque_Click()
    cmbxComp.Enabled = False
End Sub

Private Sub optnEstoque_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optn_Click()
    cmbxComp.Clear
End Sub

Private Sub optnGrupo_Click()
    'P_PreencheCateg cmbxComp
    cmbxComp.Enabled = True
End Sub

Private Sub optnGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnMarca_Click()
    'P_PreencheMarca cmbxComp
    cmbxComp.Enabled = True
End Sub

Private Sub optnMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnSub_Click()
    'P_PreencheSubCat cmbxComp
    cmbxComp.Enabled = True
End Sub

Private Sub optnSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
        
    End If
End Sub
Sub buscaProd()
    
    Dim p As ProdutoSaida
    p = F_BuscaProdSaida(txtbxExpr.Text)
    If p.cod_pro <> "" Then
        txtbxExpr.Text = p.nome_pro
        txtbxExpr.tag = p.cod_pro
        lblCodFab.caption = p.ref_pro
        lblPreco.caption = FormatVal(p.pv_pro)
        
    End If
    
    
End Sub
Sub BuscaAvancada()
    Dim strSQL As String
    Dim rPesq As Recordset
'    On Error GoTo trata_erro
    P_Status "Buscando produtos..."
    strSQL = "select tp.nome_pro,tm.nome_mar,tp.cod_pro,tp.codfab_pro,tp.local_pro,tp.p2_pro,tc.e2_pro "
    strSQL = strSQL & " from tab_produtos as tp, tab_composicao as tc, "
    strSQL = strSQL & " tab_categorias as tg, tab_marcas as tm "
    strSQL = strSQL & " where tp.cod_pro=tc.cod_pro and tg.cod_cat=tp.cod_cat "
    strSQL = strSQL & "  and tm.cod_mar=tp.cod_mar "
    If optnGrupo.Value Then
        strSQL = strSQL & " and tg.nome_cat='" & cmbxComp.Text & "' "
    
    ElseIf optnMarca.Value Then
        strSQL = strSQL & " and tm.nome_mar='" & cmbxComp.Text & "' "
    End If
    strSQL = strSQL & " and e2_pro>0 "
    strSQL = strSQL & " order by nome_pro "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        msflxGrd.rows = 1
        msflxGrd.ColAlignment(2) = 1
        lblTotalEti.caption = "0"
        Do While Not .EOF
            strSQL = !e2_pro & Chr(9)
            strSQL = strSQL & FormatCodigo(!cod_pro) & Chr(9)
            strSQL = strSQL & Format(!codfab_pro) & Chr(9)
            strSQL = strSQL & !nome_pro & Chr(9)
            strSQL = strSQL & !nome_mar & Chr(9)
            strSQL = strSQL & Format(!local_pro) & Chr(9)
            strSQL = strSQL & FormatVal(!p2_pro)
            msflxGrd.AddItem strSQL
            lblTotalEti.caption = Int(lblTotalEti.caption) + !e2_pro
            .MoveNext
        Loop
        Inf
        .Close
        
    End With
    P_Status ""
Exit Sub
trata_erro:
    On Error Resume Next
    P_Status ""
    MsgBox Err.Description, vbCritical
End Sub
Sub BuscaEntrada()
    Dim strSQL As String
    Dim rPesq As Recordset
'    On Error GoTo trata_erro
    P_Status "Buscando produtos..."
    strSQL = "select qtd_ite,desc_ite,descricao_pos,ti.num_ven  from tab_vendas as tv, tab_itvendas as ti, tab_posologia as tp "
    
    strSQL = strSQL & " where tv.num_ven=ti.num_ven "
    strSQL = strSQL & " and ti.num_ven=tp.num_ven "
    strSQL = strSQL & " and ti.seq_ite=tp.seq_ite "
    strSQL = strSQL & " and tv.num_ven='" & Me.tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        msflxGrd.rows = 1
        'msflxgrd.ColAlignment(2) = 1
        lblTotalEti.caption = "0"
        Do While Not .EOF
            strSQL = !qtd_ite & Chr(9)
            strSQL = strSQL & !descricao_pos
            msflxGrd.AddItem strSQL
            lblTotalEti.caption = Int(lblTotalEti.caption) + !qtd_ite
            .MoveNext
        Loop
        Inf
        .Close
        
    End With
    P_Status ""
Exit Sub
trata_erro:
    On Error Resume Next
    P_Status ""
    MsgBox Err.Description, vbCritical
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
End Sub

Private Sub txtbxIni_GotFocus()
    P_SelecTexto txtbxIni
End Sub

Private Sub txtbxIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIni_LostFocus()
    P_ValorLostFocus txtbxIni, "#"
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd, "###,##0"
End Sub
