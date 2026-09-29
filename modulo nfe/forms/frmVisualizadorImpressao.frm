VERSION 5.00
Begin VB.Form frmVisualizadorImpressao 
   Caption         =   "Visualizar Impressão"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   Icon            =   "frmVisualizadorImpressao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsc 
      Height          =   255
      LargeChange     =   1000
      Left            =   120
      SmallChange     =   100
      TabIndex        =   5
      Top             =   7200
      Width           =   11055
   End
   Begin VB.VScrollBar vsc 
      Height          =   6975
      LargeChange     =   1000
      Left            =   11280
      SmallChange     =   100
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox pcFundo 
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6915
      ScaleWidth      =   10995
      TabIndex        =   2
      Top             =   120
      Width           =   11055
      Begin VB.PictureBox pcImpressao 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   0
         ScaleHeight     =   1575
         ScaleWidth      =   1335
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   7500
      Width           =   11625
      Begin VB.PictureBox pcBotoes 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   8040
         ScaleHeight     =   855
         ScaleWidth      =   3255
         TabIndex        =   11
         Top             =   0
         Width           =   3255
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "Imprimir"
            Height          =   615
            Left            =   120
            Picture         =   "frmVisualizadorImpressao.frx":1CFA
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "Cancelar"
            Height          =   615
            Left            =   1680
            Picture         =   "frmVisualizadorImpressao.frx":21E8
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.TextBox txtPagina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Text            =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
      Begin VB.Label lblImpressora 
         Caption         =   "Impressora Selecionada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impressora"
         Height          =   195
         Left            =   1800
         TabIndex        =   8
         Top             =   0
         Width           =   765
      End
      Begin VB.Label lblNPags 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1260
         TabIndex        =   7
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de"
         Height          =   240
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmVisualizadorImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nPag As Integer
Public iPrinter As Integer
Private Sub cmdCancelar_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmdImprimir_Click()
    Me.tag = iPrinter
    Me.Hide
End Sub

Private Sub Form_Activate()
    If iPrinter = -1 Then
        iPrinter = BuscaImpressoraPadrao
    End If
    
    lblImpressora.caption = BuscaNomeImpressoraPorIndice(iPrinter)
    mudaPagina
    resizeForm Me
End Sub

Private Sub Form_Load()
    nPag = 0
    iPrinter = -1
    pcImpressao.ScaleMode = Printer.ScaleMode
    pcImpressao.width = Printer.width
    pcImpressao.Height = Printer.Height
    hsc.max = pcImpressao.width / 2
    vsc.max = (pcImpressao.Height / 2) + 3000
    
End Sub

Private Sub Form_Resize()
    PosicionaGrade pcFundo, Me, pcBot
    
    posicionaADireita vsc, Me
    vsc.Left = vsc.Left - vsc.width
    hsc.Top = pcFundo.Top + pcFundo.Height - 100
    hsc.width = pcFundo.width - 200
    vsc.Height = pcFundo.Height - 100
    posicionaADireita pcBotoes, Me
End Sub

Private Sub hsc_Change()
    pcImpressao.Left = 0 - hsc.Value
End Sub

Private Sub lblImpressora_Click()
    Dim i As Integer
    i = EscolheImpressora
    
    If i > -1 Then
        iPrinter = i
        lblImpressora.caption = BuscaNomeImpressoraPorIndice(iPrinter)
        
    End If
    
End Sub

Private Sub pcImpressao_Paint()
'    MsgBox "pintou"
End Sub

Private Sub txtPagina_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mudaPagina
    End If
    
End Sub

Sub mudaPagina()
On Error Resume Next
    If Int(txtPagina.Text) > Int(lblNPags.caption) Then
        
        MsgBox "Página não existe!", vbExclamation
        txtPagina.Text = nPag
        txtPagina.SetFocus
        Exit Sub
    End If
    pcImpressao.Picture = LoadPicture(App.Path & "\pagina" & txtPagina.Text & ".bmp")
    lblNPags.caption = nPag
End Sub

Private Sub txtZoom_Change()
    
End Sub


Private Sub vsc_Change()
    pcImpressao.Top = 0 - vsc.Value
End Sub


Public Sub addNewPage()
    nPag = nPag + 1
    SavePicture pcImpressao.image, App.Path & "\pagina" & nPag & ".bmp"
    pcImpressao.Cls
    
End Sub

