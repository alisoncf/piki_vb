VERSION 5.00
Object = "{70383A57-8640-4150-AB17-F7D6588F8CF9}#4.0#0"; "Cfx4032ce.ocx"
Begin VB.Form frmGraf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gráfico"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   Icon            =   "frmGraf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbxGraf 
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
      ItemData        =   "frmGraf.frx":0442
      Left            =   120
      List            =   "frmGraf.frx":0464
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   7560
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   600
      Left            =   9600
      Picture         =   "frmGraf.frx":04D7
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7380
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
      ItemData        =   "frmGraf.frx":0861
      Left            =   3720
      List            =   "frmGraf.frx":087D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Imprimir"
      Height          =   600
      Left            =   7440
      Picture         =   "frmGraf.frx":08A2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7380
      Width           =   1935
   End
   Begin VB.Frame frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10575
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   11895
         Left            =   0
         ScaleHeight     =   11865
         ScaleWidth      =   11865
         TabIndex        =   3
         Top             =   120
         Width           =   11895
         Begin ChartfxLibCtl.ChartFX ChrtFX 
            Height          =   8775
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   10575
            _cx             =   18653
            _cy             =   15478
            Build           =   20
            TypeMask        =   44564485
            LeftGap         =   29
            RightGap        =   9
            TopGap          =   9
            BottomGap       =   23
            DblClk          =   2
            RigClk          =   0
            Volume          =   50
            Axis(0).Decimals=   0
            Axis(0).TickMark=   -32767
            Axis(2).Min     =   0
            Axis(2).Max     =   100
            RGBBk           =   -2147483643
         End
      End
   End
   Begin VB.VScrollBar vscr 
      Height          =   6960
      LargeChange     =   500
      Left            =   10560
      SmallChange     =   100
      TabIndex        =   1
      Top             =   120
      Width           =   250
   End
   Begin VB.HScrollBar Hscrc 
      Height          =   250
      LargeChange     =   500
      Left            =   0
      SmallChange     =   100
      TabIndex        =   0
      Top             =   7080
      Width           =   10560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Gráfico:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   7320
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Zoom:"
      Height          =   195
      Left            =   3720
      TabIndex        =   8
      Top             =   7320
      Width           =   450
   End
End
Attribute VB_Name = "frmGraf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub graf()
    With mschrt
'        .chartType = 1
 '       .ColumnCount = 5
  '      .RowCount = 1
         .Ser
    End With
End Sub


Private Sub Combo1_Change()
    
End Sub

Private Sub Combo1_Click()
    
End Sub

Private Sub ChrtFX_LButtonDblClk(ByVal X As Integer, ByVal Y As Integer, ByVal nSerie As Integer, ByVal nPoint As Long, nRes As Integer)
    ChrtFX.ShowTips = True
End Sub

Private Sub cmbxGraf_Click()
    If cmbxGraf.ListIndex = 0 Then
        ChrtFX.Gallery = 5
    End If
    If cmbxGraf.ListIndex = 1 Then
        ChrtFX.Gallery = 7
    End If
    If cmbxGraf.ListIndex = 2 Then
        ChrtFX.Gallery = 6
    End If
    If cmbxGraf.ListIndex = 3 Then
        ChrtFX.Gallery = 19
    End If
    If cmbxGraf.ListIndex = 4 Then
        ChrtFX.Gallery = 1
        
    End If
    If cmbxGraf.ListIndex = 5 Then
        ChrtFX.Gallery = 13
        
    End If
    If cmbxGraf.ListIndex = 6 Then
        ChrtFX.Gallery = 20
    End If
    If cmbxGraf.ListIndex = 7 Then
        ChrtFX.Gallery = 3
    End If
    If cmbxGraf.ListIndex = 8 Then
        ChrtFX.Gallery = 8
    End If
    If cmbxGraf.ListIndex = 9 Then
        ChrtFX.Gallery = 10
    End If
End Sub

Private Sub cmbxZoom_Click()
    
    ChrtFX.Volume = Int(cmbxZoom.Text)
    ChrtFX.Refresh
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    pic.Height = ChrtFX.Height
End Sub

Private Sub Form_Load()
    Hscrc.Max = pic.Width / 2
    vscr.Max = pic.Height / 2
End Sub
Private Sub Hscrc_Change()
    pic.Left = 0 - Hscrc.Value
End Sub

Private Sub vscr_Change()
    pic.Top = 0 - vscr.Value
End Sub
