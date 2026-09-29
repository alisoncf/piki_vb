VERSION 5.00
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVisualiza 
   Caption         =   "Visualização de Impressão:"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   Icon            =   "frmVisualiza.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcBotoes 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   10020
      ScaleHeight     =   8145
      ScaleWidth      =   1425
      TabIndex        =   5
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton cmdImpressora 
         Caption         =   "Impressora"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtbxCopias 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Text            =   "1"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.PictureBox pcBarcode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   825
         ScaleWidth      =   2385
         TabIndex        =   14
         Top             =   5520
         Visible         =   0   'False
         Width           =   2415
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   0
            ScaleHeight     =   825
            ScaleWidth      =   2385
            TabIndex        =   17
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
            Begin BARCODEXLib.BarcodeX brcdx 
               Height          =   855
               Left            =   120
               Top             =   0
               Visible         =   0   'False
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   1508
               _StockProps     =   13
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   ""
               Title           =   ""
            End
         End
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   720
         Left            =   120
         Picture         =   "frmVisualiza.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4320
         Width           =   1215
      End
      Begin VB.ComboBox cmbxZoom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmVisualiza.frx":0B14
         Left            =   120
         List            =   "frmVisualiza.frx":0B2D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   1260
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Imprimir"
         Default         =   -1  'True
         Height          =   720
         Left            =   120
         Picture         =   "frmVisualiza.frx":0B5A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox cmbxPag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmVisualiza.frx":0EE4
         Left            =   120
         List            =   "frmVisualiza.frx":0EFD
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   1260
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir Esta Página"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblCOpias 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cópias"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Zoom:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Página:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.Frame frContainer 
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9735
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   7575
         Left            =   120
         ScaleHeight     =   7545
         ScaleWidth      =   9585
         TabIndex        =   13
         Top             =   120
         Width           =   9615
         Begin MSComctlLib.ImageList imglst 
            Left            =   8760
            Top             =   6360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
      End
   End
   Begin VB.HScrollBar Hscrc 
      Height          =   250
      LargeChange     =   500
      Left            =   0
      SmallChange     =   100
      TabIndex        =   3
      Top             =   7800
      Width           =   9840
   End
   Begin VB.VScrollBar vscr 
      Height          =   7800
      LargeChange     =   500
      Left            =   9840
      SmallChange     =   100
      TabIndex        =   2
      Top             =   0
      Width           =   250
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox Pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   7680
         Left            =   120
         ScaleHeight     =   7620
         ScaleWidth      =   6375
         TabIndex        =   1
         Top             =   -1800
         Width           =   6435
      End
   End
End
Attribute VB_Name = "frmVisualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxPag_Click()
    pic.Cls
    pic.picture = Nothing
    pic.PaintPicture imglst.ListImages(cmbxPag.ListIndex + 1).picture, 0, 0
    Picture1.picture = pic.image
End Sub

Private Sub cmbxZoom_Click()
    Dim w As Single
    Dim h As Single
    On Error GoTo trata_erro
    If cmbxZoom.ListCount - 1 Then
        'Image1.Stretch = False
        'Image1.Width = Image1.Width * (Int(Left(cmbxZoom.Text, 2)) / 100)
        'Image1.Height = Image1.Height * (Int(Left(cmbxZoom.Text, 2)) / 100)
        'Image1.Stretch = True
        w = Picture1.width * (Int(Left(cmbxZoom.Text, 2)) / 100)
        h = Picture1.height * (Int(Left(cmbxZoom.Text, 2)) / 100)
        Picture1.Cls
        Picture1.PaintPicture pic.image, 0, 0, w, h
        Picture1.AutoSize = False
        Picture1.height = h
        Picture1.width = w
    Else
        'Image1.Stretch = False
    End If
    Exit Sub
trata_erro:
End Sub

Private Sub cmdImpressora_Click()
    Set Printer = Printers(BuscaImpressoraPadrao)
End Sub

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Command3_Click()
    If cmbxPag.ListIndex > 0 Then
        Printer.PaintPicture imglst.ListImages(cmbxPag.ListIndex + 1).picture, 0, 0
    Else
        Printer.PaintPicture pic.image, 0, 0
        
    End If
    Printer.EndDoc
End Sub

Private Sub Command4_Click()
'    If MsgBox("Confirma Impressão?", vbYesNo + vbQuestion) = vbYes Then
        'Printer.PaintPicture Pic.Image, 0, 0
        'Printer.EndDoc
        Me.tag = "X"
        Me.Hide
 '   End If
End Sub

Private Sub Form_Activate()
    Dim iCount As Integer
    cmbxZoom.ListIndex = cmbxZoom.ListCount - 1
    'Image1.Picture = Pic.Image
    
    'SavePicture Pic.Image, App.Path & "\impressao.tmp"
    Picture1.width = pic.width
    Picture1.height = pic.height
    
    Picture1.PaintPicture pic.image, 0, 0
    
    P_Status ""
    MudaAparencia Me
    
    If imglst.ListImages.Count > 0 Then
        cmbxPag.Clear
        For iCount = 1 To imglst.ListImages.Count
            cmbxPag.AddItem iCount
        Next
        cmbxPag.ListIndex = 0
    Else
        cmbxPag.Enabled = False
    End If
    
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Printer.ScaleMode = 1
    pic.ScaleMode = 1
    
    
    pic.height = Printer.ScaleHeight
    pic.width = Printer.ScaleWidth
    P_CentraForm Me
    Hscrc.max = pic.width / 2
    vscr.max = (pic.height / 2) + 1200
    Frame1.width = pic.width + 50
    pic.AutoRedraw = True
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    pcBotoes.Left = Me.width - pcBotoes.width - 300
    
    vscr.Left = pcBotoes.Left - 300
    frContainer.width = vscr.Left - 50
    Hscrc.Top = Me.height - 800
    frContainer.height = Hscrc.Top - 50
    Hscrc.width = vscr.Left
    vscr.height = Hscrc.Top
    pcBotoes.height = Hscrc.Top
End Sub

Private Sub Hscrc_Change()
    Picture1.Left = 0 - Hscrc.Value
End Sub


Private Sub vscr_Change()
    Picture1.Top = 0 - vscr.Value
End Sub
    
Private Sub Pic_Paint()
    pic.Refresh
End Sub


