VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmImagens 
   Caption         =   "Imagens"
   ClientHeight    =   8805
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14850
   Icon            =   "frmImagens.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   14850
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   14820
      TabIndex        =   28
      Top             =   0
      Width           =   14850
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   29
         Top             =   0
         Width           =   12975
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Produto:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   1545
      End
   End
   Begin VB.TextBox txtCaminhoImg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "C:\"
      Top             =   8400
      Width           =   10335
   End
   Begin VB.PictureBox pcNavegaImgs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4200
      ScaleHeight     =   345
      ScaleWidth      =   10305
      TabIndex        =   11
      Top             =   600
      Width           =   10335
      Begin VB.CommandButton cmdEsq 
         Caption         =   "<"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   ">"
         Height          =   375
         Left            =   9360
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
      Begin VB.Label lblCapImg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Imagem 0 de 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox pcQuadro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   4200
      ScaleHeight     =   7785
      ScaleWidth      =   10305
      TabIndex        =   3
      Top             =   960
      Width           =   10335
      Begin VB.PictureBox pcImg 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1785
         ScaleWidth      =   3825
         TabIndex        =   10
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.PictureBox pcLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   0
      ScaleHeight     =   8625
      ScaleWidth      =   4065
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      Begin VB.PictureBox pcAjustesImagem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   0
         ScaleHeight     =   1305
         ScaleWidth      =   4185
         TabIndex        =   19
         Top             =   6600
         Width           =   4215
         Begin VB.CommandButton cmdSave 
            Caption         =   "Salvar Cópia"
            Height          =   495
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdMoveCima 
            Height          =   495
            Left            =   3360
            Picture         =   "frmImagens.frx":1084A
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   120
            Width           =   615
         End
         Begin VB.CommandButton cmdMoveBaixo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            Picture         =   "frmImagens.frx":10D31
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   120
            Width           =   615
         End
         Begin VB.CommandButton cmdMoveEsq 
            Height          =   495
            Left            =   2640
            Picture         =   "frmImagens.frx":11219
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   720
            Width           =   615
         End
         Begin VB.CommandButton cmdMoveDir 
            Height          =   495
            Left            =   3360
            Picture         =   "frmImagens.frx":116FD
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   720
            Width           =   615
         End
         Begin VB.CommandButton cmdAjuste 
            Caption         =   "Ajustar Imagem"
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdMenos 
            Height          =   495
            Left            =   1920
            Picture         =   "frmImagens.frx":11BC1
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   120
            Width           =   615
         End
         Begin VB.CommandButton cmdMais 
            Height          =   495
            Left            =   1920
            Picture         =   "frmImagens.frx":11FC8
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   2640
         Picture         =   "frmImagens.frx":1255D
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   8040
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid gradeImg 
         Height          =   3255
         Left            =   0
         TabIndex        =   15
         Top             =   1680
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         Caption         =   "Imagens"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "id"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "caminho"
            Caption         =   "Caminho"
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
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3030,236
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdObj 
         Caption         =   "Nova Imagem"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imagens Relacionadas"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   3090
      End
      Begin VB.Label lblMarca 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   435
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   3195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblCodPro 
         AutoSize        =   -1  'True
         Caption         =   "..."
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   645
      End
      Begin VB.Label lblCodFab 
         AutoSize        =   -1  'True
         Caption         =   "..."
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref:"
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
         Left            =   1680
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Menu mnOp 
      Caption         =   "Opções"
      Begin VB.Menu mnCopia 
         Caption         =   "Salvar uma cópia"
      End
      Begin VB.Menu mnAjuste 
         Caption         =   "Auto Ajustar Imagem"
      End
      Begin VB.Menu mnEsq 
         Caption         =   "Mover para Esquerda"
      End
      Begin VB.Menu mnDir 
         Caption         =   "Mover para Direita"
      End
      Begin VB.Menu mnCima 
         Caption         =   "Mover pra cima"
      End
      Begin VB.Menu mnBaixo 
         Caption         =   "Mover pra Baixo"
      End
      Begin VB.Menu div 
         Caption         =   "-"
      End
      Begin VB.Menu mnSair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmImagens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xPoint As Single
Dim yPoint As Single
Dim xSolta As Single
Dim ySolta As Single
Dim prop As Single
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If ado.AbsolutePosition > -1 Then
        lblCapImg.caption = "Imagem " & ado.AbsolutePosition & " de " & ado.RecordCount
    End If
End Sub

Private Sub cmdAjuste_Click()
    AjusteImagem
End Sub
Sub AjusteImagem()
    If pcImg.image = 0 Then
    
    Else
        resizeImage pcImg, pcQuadro
    End If
End Sub
Sub resizeImage(img As Object, obj As PictureBox)
    Dim i As Integer
    Dim h As Single
    Dim w As Single
    Dim p As Single
    Dim horizontal As Single
    Dim proporcao As Single
    pcImg.Left = 0
    pcImg.Top = 0
    h = img.height
    w = img.width
    p = h / 10
    
    
    If w > obj.ScaleWidth Then
    
        For i = w To obj.width Step -p
            horizontal = i
            
            
        Next
        proporcao = i / w
     ElseIf h > obj.ScaleHeight Then
        For i = h To obj.height Step -p
            horizontal = i
            
            
        Next
        proporcao = i / w
    End If
    
    MostraImagemMiniatura txtCaminhoImg.Text, img.width * proporcao, img.height * proporcao
    prop = proporcao
End Sub

Private Sub cmdDir_Click()
    If ado.AbsolutePosition < ado.RecordCount Then
        ado.Move 1
        gradeImg_Click
    End If
End Sub

Private Sub cmdEsq_Click()
    If ado.AbsolutePosition > 1 Then
        ado.Move -1
        gradeImg_Click
    End If
End Sub

Private Sub cmdExcluir_Click()
    If MsgBox("Confirma a exclusão do registro?", vbYesNo + vbCritical) = vbYes Then
        ExcluirImagem ado!id
        BuscaImagens
    End If
End Sub

Private Sub cmdMais_Click()
    prop = prop + 0.05
    
    MostraImagemMiniatura txtCaminhoImg.Text, pcImg.ScaleWidth * prop, pcImg.ScaleHeight * prop
End Sub

Private Sub cmdMenos_Click()
    prop = prop - 0.05
    
    MostraImagemMiniatura txtCaminhoImg.Text, pcImg.ScaleWidth * prop, pcImg.ScaleHeight * prop
End Sub

Private Sub cmdMoveBaixo_Click()
    pcImg.Top = pcImg.Top + 100
End Sub

Private Sub cmdMoveCima_Click()
    pcImg.Top = pcImg.Top - 100
End Sub

Private Sub cmdMoveDir_Click()
    pcImg.Left = pcImg.Left + 100
End Sub

Private Sub cmdMoveEsq_Click()
    pcImg.Left = pcImg.Left - 100
End Sub

Private Sub cmdObj_Click()
    Dim caminho As String
    caminho = BuscaNomeArquivoExternoAbrir("*.jpg;*.jpeg")
    MostraImagemMiniatura caminho, 0, 0
    imagem.AdicionarImagem txtProduto.tag, caminho
    BuscaImagens
End Sub
Sub BuscaImagens()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "select * from imagem_produto where cod_pro='" & txtProduto.tag & "' "
    
    abreObjetoAdo ado, strSQL, gradeImg
    If ado.RecordCount > 0 Then
        gradeImg_Click
    End If
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub



Private Sub cmdSave_Click()
    On Error GoTo trata_erro
    Dim file As String
    file = BuscaNomeArquivoExternoSalvar(".jpg", RetornaNomeArquivoResumido(ado!caminho))
    
    If file = "" Then
        Exit Sub
    End If
    FileCopy txtCaminhoImg.Text, file
    
    Shell "explorer " & file
Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    Carrega
    
    bActive = True
End Sub
Sub Carrega()
    
    
    If Me.tag <> "" Then
        buscaProd
    End If
End Sub
Private Sub Form_Load()
    
    
    bActive = False
    Set ado = New ADODB.Recordset
    prop = 1
    resizeForm Me
End Sub

Private Sub Form_Resize()
   ' On Error Resume Next
    pcQuadro.width = Me.width - pcQuadro.Left - 300
    pcNavegaImgs.width = pcQuadro.width
    txtCaminhoImg.width = pcQuadro.width
    posicionaADireita cmdDir, pcNavegaImgs
    pcLeft.height = Me.height - pcLeft.Top - 800
    pcQuadro.height = pcLeft.height - 500
    txtCaminhoImg.Top = pcQuadro.height + pcQuadro.Top - txtCaminhoImg.height
    txtCaminhoImg.Left = pcQuadro.Left
    
    cmndbtnSair.Top = txtCaminhoImg.Top - 500
    
End Sub

Private Sub gradeImg_Click()
    MostraImagemMiniatura ado!caminho
End Sub

Private Sub mnAjuste_Click()
    cmdAjuste_Click
End Sub

Private Sub mnBaixo_Click()
    cmdMoveBaixo_Click
End Sub

Private Sub mnCima_Click()
    cmdMoveCima_Click
End Sub

Private Sub mnCopia_Click()
    cmdSave_Click
End Sub

Private Sub mnDir_Click()
        cmdMoveDir_Click
End Sub

Private Sub mnEsq_Click()
    cmdMoveEsq_Click
End Sub

Private Sub mnSair_Click()
    cmndbtnSair_Click
End Sub

Private Sub pcImg_DblClick()
    On Error Resume Next
    Shell "explorer " & txtCaminhoImg.Text
End Sub

Private Sub pcimg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        yPoint = y
        xPoint = x
    End If
    If Button = 2 Then
        Me.PopupMenu mnOp
    End If
End Sub

Private Sub pcimg_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        xSolta = (x - xPoint)
        ySolta = (y - yPoint)
        
        pcImg.Left = pcImg.Left + xSolta
        pcImg.Top = pcImg.Top + ySolta
        
    End If
End Sub

Sub MostraImagemMiniatura(img As String, Optional w As Single, Optional h As Single)
    Dim s As String
    On Error Resume Next
    pcImg.ScaleMode = 1
    
    pcImg.picture = Nothing
    pcImg.AutoSize = True

    img = lstImagens.Text
    If w > 0 And h > 0 Then
         pcImg.PaintPicture LoadPicture(img), 0, 0, w, h
    Else
        prop = 1
        pcImg.Left = 0
        pcImg.Top = 0
        pcImg.picture = LoadPicture(img)
    End If
  
    pcImg.Refresh
    
    txtCaminhoImg.Text = img
    
   ' Set pcMiniatura.Picture = LoadPicture(s)
    
End Sub

Private Sub txtProduto_Change()
    txtProduto.tag = ""
End Sub

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
    End If
End Sub
Sub buscaProd()
    Dim p As Produto
    
    p = Produtos.F_BuscaProdSaida(txtProduto.Text)
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
        lblCodPro.caption = p.cod_pro
        lblMarca.caption = p.nome_mar
        lblCodFab.caption = p.codfab_pro
        BuscaImagens
    End If
    
End Sub
