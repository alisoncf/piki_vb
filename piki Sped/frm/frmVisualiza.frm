VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
      Height          =   8055
      Left            =   10080
      ScaleHeight     =   7995
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   0
      Width           =   1335
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   600
         Left            =   120
         Picture         =   "frmVisualiza.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2640
         Width           =   1095
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
         ItemData        =   "frmVisualiza.frx":0A14
         Left            =   120
         List            =   "frmVisualiza.frx":0A2D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   1140
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Imprimir"
         Height          =   600
         Left            =   120
         Picture         =   "frmVisualiza.frx":0A5A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
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
         ItemData        =   "frmVisualiza.frx":0DE4
         Left            =   120
         List            =   "frmVisualiza.frx":0DFD
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   1140
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir Esta Página"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Zoom:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
      End
      Begin MSComctlLib.ImageList imglst 
         Left            =   2040
         Top             =   3960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
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
    Pic.Cls
    Pic.Picture = Nothing
    Pic.PaintPicture imglst.ListImages(cmbxPag.ListIndex + 1).Picture, 0, 0
    Picture1.Picture = Pic.image
End Sub

Private Sub cmbxZoom_Click()
    Dim w As Single
    Dim H As Single
    On Error GoTo trata_erro
    If cmbxZoom.ListCount - 1 Then
        'Image1.Stretch = False
        'Image1.Width = Image1.Width * (Int(Left(cmbxZoom.Text, 2)) / 100)
        'Image1.Height = Image1.Height * (Int(Left(cmbxZoom.Text, 2)) / 100)
        'Image1.Stretch = True
        w = Picture1.Width * (Int(Left(cmbxZoom.Text, 2)) / 100)
        H = Picture1.Height * (Int(Left(cmbxZoom.Text, 2)) / 100)
        Picture1.Cls
        Picture1.PaintPicture Pic.image, 0, 0, w, H
        Picture1.AutoSize = False
        Picture1.Height = H
        Picture1.Width = w
    Else
        'Image1.Stretch = False
    End If
    Exit Sub
trata_erro:
End Sub

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Command3_Click()
    If cmbxPag.ListIndex > 0 Then
        Printer.PaintPicture imglst.ListImages(cmbxPag.ListIndex + 1), 0, 0
    Else
        Printer.PaintPicture Pic.image, 0, 0
        
    End If
    Printer.EndDoc
End Sub

Private Sub Command4_Click()
    If MsgBox("Confirma Impressão?", vbYesNo + vbQuestion) = vbYes Then
        'Printer.PaintPicture Pic.Image, 0, 0
        'Printer.EndDoc
        Me.tag = "X"
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    Dim iCount As Integer
    cmbxZoom.ListIndex = cmbxZoom.ListCount - 1
    'Image1.Picture = Pic.Image
    
    'SavePicture Pic.Image, App.Path & "\impressao.tmp"
    Picture1.Width = Pic.Width
    Picture1.Height = Pic.Height
    
    Picture1.PaintPicture Pic.image, 0, 0
    
    P_Status ""
    'MudaAparencia Me
    
    If imglst.ListImages.count > 0 Then
        cmbxPag.Clear
        For iCount = 1 To imglst.ListImages.count
            cmbxPag.AddItem iCount
        Next
        cmbxPag.ListIndex = 0
    Else
        cmbxPag.Enabled = False
    End If
    
End Sub

Private Sub Form_Load()
    
    
    Printer.ScaleMode = 1
    Pic.ScaleMode = 1
    
    
    Pic.Height = Printer.ScaleHeight
    Pic.Width = Printer.ScaleWidth
    P_CentraForm Me
    Hscrc.max = Pic.Width / 2
    vscr.max = (Pic.Height / 2) + 1200
    Frame1.Width = Pic.Width + 50
    Pic.AutoRedraw = True
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    pcBotoes.Left = Me.Width - pcBotoes.Width - 300
    
    vscr.Left = pcBotoes.Left - 300
    frContainer.Width = vscr.Left - 50
    Hscrc.Top = Me.Height - 800
    frContainer.Height = Hscrc.Top - 50
    Hscrc.Width = vscr.Left
    vscr.Height = Hscrc.Top
    pcBotoes.Height = Hscrc.Top
End Sub

Private Sub Hscrc_Change()
    Picture1.Left = 0 - Hscrc.Value
End Sub


Private Sub vscr_Change()
    Picture1.Top = 0 - vscr.Value
End Sub
    
Private Sub Pic_Paint()
    Pic.Refresh
End Sub

