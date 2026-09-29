VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditaFoto 
   Caption         =   "Edição de Figuras"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   2160
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Image img 
         Height          =   1695
         Left            =   0
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Frame frAss 
      Caption         =   "Assina&tura Por Extenso:"
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   855
         Left            =   7200
         TabIndex        =   15
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnMenos 
         Caption         =   "Menos"
         Height          =   495
         Left            =   7200
         TabIndex        =   13
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnMais 
         Cancel          =   -1  'True
         Caption         =   "Mais"
         Height          =   495
         Left            =   7200
         TabIndex        =   12
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   255
         Left            =   7080
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   4815
         LargeChange     =   500
         Left            =   6600
         SmallChange     =   100
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   500
         Left            =   120
         SmallChange     =   100
         TabIndex        =   2
         Top             =   5040
         Width           =   6495
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5055
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6735
         Begin VB.PictureBox PicDisplay 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5040
            Left            =   0
            ScaleHeight     =   336
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   439
            TabIndex        =   8
            Top             =   0
            Width           =   6585
            Begin VB.PictureBox PicSelect 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1095
               Left            =   240
               MouseIcon       =   "frmEditaFoto.frx":0000
               MousePointer    =   5  'Size
               ScaleHeight     =   73
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   97
               TabIndex        =   10
               Top             =   240
               Visible         =   0   'False
               Width           =   1455
               Begin VB.Shape PicSelectShape 
                  BorderStyle     =   3  'Dot
                  DrawMode        =   6  'Mask Pen Not
                  Height          =   615
                  Left            =   0
                  Top             =   0
                  Width           =   735
               End
            End
            Begin VB.PictureBox PicTemp 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1095
               Left            =   2520
               MouseIcon       =   "frmEditaFoto.frx":030A
               ScaleHeight     =   73
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   97
               TabIndex        =   9
               Top             =   1440
               Visible         =   0   'False
               Width           =   1455
            End
            Begin MSComDlg.CommonDialog cmndlg 
               Left            =   0
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Line Line1 
               BorderStyle     =   3  'Dot
               Visible         =   0   'False
               X1              =   56
               X2              =   152
               Y1              =   128
               Y2              =   128
            End
            Begin VB.Shape ShapeSelect 
               BorderStyle     =   3  'Dot
               DrawMode        =   6  'Mask Pen Not
               Height          =   615
               Left            =   2160
               Top             =   360
               Visible         =   0   'False
               Width           =   735
            End
         End
      End
      Begin VB.CommandButton cmndbtnGravaFoto 
         Caption         =   "&Apagar"
         Height          =   345
         Left            =   7080
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Capturar"
         Height          =   345
         Left            =   7080
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   345
         Left            =   7080
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Fonte"
         Height          =   345
         Left            =   7080
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditaFoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive


'This demo shows in the simplest possible code, the basics of
'image editing. It is intended for the novice but others may
'find it useful in optimising or simplifying their code.

'When first starting coding graphics apps, one of the functions
'that can cause problems is making selections and controlling
'input to those selections. I hope this demo makes those tasks
'clearer. Whilst it is in no way a full blown editor, it may provide
'a basis from which to learn.

'If you want an example of a more advanced image editor I
'suggest you look at "Bobo Image Workshop". You can download
'this from Planet Source Code
'address :http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=21482&lngWId=1



Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Dim sX As Long, sY As Long, CurColor As Long, CurTool As Long
Dim DontMouseUp As Boolean, DontDraw As Boolean

Private Sub cmndbtnBuscar_Click()
    cmndlg.Filter = "Pictures (*.bmp;*.jpg)|*.bmp;*.jpg"
    Set fs = CreateObject("Scripting.FileSystemObject")
    

    cmndlg.ShowOpen
    PicDisplay.Picture = LoadPicture(cmndlg.FileName)
End Sub

Private Sub cmndbtnMais_Click()
    img.Stretch = True
    img.Picture = PicDisplay.Picture
    
    img.Height = img.Height * 1.75
    img.Width = img.Width * 1.75
    img.Stretch = True
    PicDisplay.PaintPicture img.Picture, 0, 0
    
    PicDisplay.Refresh
End Sub

Private Sub Command3_Click()
    Dim z As Long
    For z = 0 To mnuTools.count - 1 'uncheck all menu items
        mnuTools(z).Checked = False
    Next
    mnuTools(Index).Checked = True 'check clicked menu
    CurTool = Index 'current tool
    Select Case Index 'provide a cursor to indicate tool
        Case 0, 4 'selection,line
            PicDisplay.MousePointer = 99
            PicDisplay.MouseIcon = PicTemp.MouseIcon
        Case 3 'flood fill
            PicDisplay.MousePointer = 99
            PicDisplay.MouseIcon = PicSelect.MouseIcon
        Case Else 'just the default arrow
            PicDisplay.MousePointer = 0
    End Select
    
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command4_Click()

    Dim strSQL As String
    
   

    con.Open "banco.Connect"
    tabela.Open "select * from tab_fotos", con, adOpenDynamic, adLockOptimistic
    tabela.AddNew
    tabela!cod_cli = Me.Tag
    tabela!data_fir = Date
    tabela!cod_usu = "000"
    GravaBLOB cmndlg.FileName, tabela!assinatura_fir
    tabela.Update
    con.Close



End Sub

Private Sub Form_Load()
    CurColor = vbRed
    CurTool = 2 'pen
    mnuTools_Click 0
End Sub
Private Sub Form_Resize()
    On Error Resume Next 'center displayed image
    PicDisplay.Left = (Me.ScaleWidth - PicDisplay.Width) / 2
    PicDisplay.Top = (Me.ScaleHeight - PicDisplay.Height) / 2
End Sub
Private Sub mnuColor_Click(Index As Integer)
    Dim z As Long, col As Long
    For z = 0 To mnuColor.count - 1 'uncheck all menu items
        If z <> 5 Then mnuColor(z).Checked = False
    Next
    mnuColor(Index).Checked = True 'check clicked menu
    Select Case Index 'change current color accordingly
        Case 0: CurColor = vbRed
        Case 1: CurColor = vbGreen
        Case 2: CurColor = vbBlue
        Case 3: CurColor = vbBlack
        Case 4: CurColor = vbCyan
        Case 6 'select a color using Windows color picker
            col = ShowColor
            If col <> -1 Then CurColor = col
    End Select
End Sub
Private Sub mnuDrawWidth_Click(Index As Integer)
    Dim z As Long
    For z = 1 To mnuDrawWidth.count 'uncheck all menu items
        mnuDrawWidth(z).Checked = False
    Next
    mnuDrawWidth(Index).Checked = True 'check clicked menu
    PicDisplay.DrawWidth = Index 'set draw width
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Select Case Index
        Case 0 'cut
            Clipboard.Clear
            Clipboard.SetData PicSelect.Picture
            PicSelect.Visible = False
        Case 1 'copy
            Clipboard.Clear
            Clipboard.SetData PicSelect.Picture
        Case 2 'paste
            PicSelect.Picture = LoadPicture()
            PicSelect.AutoSize = True
            PicSelect.Picture = Clipboard.GetData
            PicSelect.AutoSize = False
            PicSelect.Left = 0
            PicSelect.Top = 0
            PicSelect.Visible = True
        Case 3 'delete
            PicSelect.Visible = False
        Case 5 'select all
            PicSelect.Left = 0
            PicSelect.Top = 0
            PicSelect.Width = PicDisplay.Width
            PicSelect.Height = PicDisplay.Height
            PicSelect.Picture = PicDisplay.Image
            PicDisplay.Picture = LoadPicture()
            PicSelect.Visible = True
        Case 6 'select none
            BitBlt PicDisplay.hdc, PicSelect.Left, PicSelect.Top, PicSelect.Width, PicSelect.Height, PicSelect.hdc, 0, 0, vbSrcCopy
            PicDisplay.Picture = PicDisplay.Image
            PicSelect.Visible = False
    End Select
End Sub

Private Sub mnuEditBase_Click()
    'mnuEdit(0).Enabled = PicSelect.Visible 'cut enabled
    'mnuEdit(1).Enabled = PicSelect.Visible 'copy enabled
    'mnuEdit(2).Enabled = Clipboard.GetFormat(vbCFBitmap) 'paste enabled
    'mnuEdit(3).Enabled = PicSelect.Visible 'delete enabled
    'mnuEdit(6).Enabled = PicSelect.Visible 'select none enabled
End Sub

Private Sub mnuFileNew_Click()
    frmNew.Show vbModal, Me 'see frmNew code
    Form_Resize
End Sub

Private Sub mnuFileOpen_Click()
    DontDraw = True 'stops accidental drawing when the dialog closes
    With cmndlg
        .Filter = "Image files |*.bmp;*.jpg;*.gif;*.ico;*.cur"
        'ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        PicDisplay.Picture = LoadPicture(.FileName) 'open an image file
    End With
End Sub
Private Sub mnuFileSaveAs_Click()
    With cmndlg
        .Filter = "Bitmap files (*.bmp)|*.bmp"
        'ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        SavePicture PicDisplay.Image, .FileName 'save an image file as bitmap
    End With
End Sub
Private Sub mnuTools_Click(Index As Integer)
    
    CurTool = 0 'current tool
    Select Case Index 'provide a cursor to indicate tool
        Case 0, 4 'selection,line
            PicDisplay.MousePointer = 99
            PicDisplay.MouseIcon = PicTemp.MouseIcon
        Case 3 'flood fill
            PicDisplay.MousePointer = 99
            PicDisplay.MouseIcon = PicSelect.MouseIcon
        Case Else 'just the default arrow
            PicDisplay.MousePointer = 0
    End Select
    
    
End Sub


Private Sub HScroll1_Change()
    PicDisplay.Left = 100 - HScroll1.Value
End Sub

Private Sub VScroll1_Change()
    PicDisplay.Top = 100 - VScroll1.Value
End Sub

Private Sub PicDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sX = X 'remember the start position of editing
    sY = Y
    If Button = 1 Then
        DontDraw = False
        If PicSelect.Visible Then 'a selection is in progress so paste in situ
            BitBlt PicDisplay.hdc, PicSelect.Left, PicSelect.Top, PicSelect.Width, PicSelect.Height, PicSelect.hdc, 0, 0, vbSrcCopy
            PicDisplay.Picture = PicDisplay.Image
            PicSelect.Visible = False
            DontMouseUp = True 'flag to tell MouseUp sub not to respond
            Exit Sub
        End If
        Select Case CurTool
            Case 0
                ShapeSelect.Left = X 'display selection area
                ShapeSelect.Top = Y
                ShapeSelect.Width = 0
                ShapeSelect.Height = 0
                ShapeSelect.Visible = True
            Case 4
                Line1.x1 = sX 'display line co-ordinates
                Line1.X2 = sX
                Line1.y1 = sY
                Line1.y2 = sY
                Line1.Visible = True
        End Select
    End If
End Sub
Private Sub PicDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And DontDraw = False Then
        Select Case CurTool
            Case 0 'selection
                ShapeSelect.Left = IIf(X > sX, sX, X)
                ShapeSelect.Top = IIf(Y > sY, sY, Y)
                ShapeSelect.Width = IIf(X > sX, X - sX, sX - X)
                ShapeSelect.Height = IIf(Y > sY, Y - sY, sY - Y)
            Case 1, 2 'pen, eraser
                PicDisplay.Line (sX, sY)-(X, Y), IIf(CurTool = 1, CurColor, vbWhite)
                sX = X
                sY = Y
            Case 4 'line
                Line1.X2 = X
                Line1.y2 = Y
        End Select
    End If
End Sub
Private Sub PicDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sX = -1
    If Button = 1 Then
        Select Case CurTool
            Case 0
                If DontMouseUp Then
                    DontMouseUp = False 'reset flag
                    Exit Sub 'user clicked to remove selection so bail out here
                End If
                ShapeSelect.Visible = False 'hide selection shape
                PicSelect.Left = ShapeSelect.Left 'position PicSelect to match selection
                PicSelect.Top = ShapeSelect.Top
                PicSelect.Width = ShapeSelect.Width
                PicSelect.Height = ShapeSelect.Height
                PicTemp.Width = ShapeSelect.Width 'used to blank out behind PicSelect
                PicTemp.Height = ShapeSelect.Height
                PicDisplay.Picture = PicDisplay.Image 'set the picture property before using Bitblt
                'fill PicSelect with selection
                BitBlt PicSelect.hdc, 0, 0, PicSelect.Width, PicSelect.Height, PicDisplay.hdc, ShapeSelect.Left, ShapeSelect.Top, vbSrcCopy
                'clear image behind PicSelect
                BitBlt PicDisplay.hdc, PicSelect.Left, PicSelect.Top, PicSelect.Width, PicSelect.Height, PicTemp.hdc, 0, 0, vbSrcCopy
                PicDisplay.Picture = PicDisplay.Image 'set the picture property
                PicSelect.Picture = PicSelect.Image 'set the picture property
                PicSelect.Visible = True 'show PicSelect
            Case 3 'Flood Fill
                PicDisplay.FillStyle = 0
                PicDisplay.FillColor = CurColor
                ExtFloodFill PicDisplay.hdc, X, Y, GetPixel(PicDisplay.hdc, X, Y), 1
                PicDisplay.FillStyle = 1
                PicDisplay.Refresh
            Case 4 'Line
                PicDisplay.Line (Line1.x1, Line1.y1)-(X, Y), CurColor
                Line1.Visible = False
        End Select
    End If
End Sub
Private Sub PicDisplay_Resize()
    On Error Resume Next
    Dim W As Long, H As Long
    'center image
    W = PicDisplay.Width * Screen.TwipsPerPixelX + 480
    H = PicDisplay.Height * Screen.TwipsPerPixelY + 1020
    'adjust form size to fit
    'If Me.Width < W Then Me.Width = W
    'If Me.Height < H Then Me.Height = H
    Form_Resize
End Sub

Private Sub PicSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'enable dragging of selection
    ReleaseCapture
    SendMessage PicSelect.hwnd, &HA1, 2, 0&
    
End Sub

Private Sub PicSelect_Resize()
    On Error Resume Next
    'show a dotted edge to selection
    PicSelectShape.Width = PicSelect.Width
    PicSelectShape.Height = PicSelect.Height
End Sub


