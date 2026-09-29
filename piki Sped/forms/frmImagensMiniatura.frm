VERSION 5.00
Begin VB.Form frmImagensMiniatura 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Imagem"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   93.927
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   105.569
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcNext 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5760
      Picture         =   "frmImagensMiniatura.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pcQuadro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5265
      ScaleWidth      =   5865
      TabIndex        =   1
      Top             =   0
      Width           =   5895
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
         TabIndex        =   2
         Top             =   0
         Width           =   3855
         Begin VB.PictureBox pcClose 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   4
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.TextBox txtCaminhoImg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   7920
      Width           =   10335
   End
End
Attribute VB_Name = "frmImagensMiniatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40



Dim xPoint As Single
Dim yPoint As Single
Dim xSolta As Single
Dim ySolta As Single
Dim prop As Single

Dim bActive As Boolean
Dim listaImagens As New Collection
Dim Indice As Integer
Dim Quantidade As Integer



Private Sub ajusteTopo()
    ' Mantém tamanho e posição, e mostra o formulário no topo
    SetWindowPos Me.hwnd, HWND_TOPMOST, _
                 Me.Left / Screen.TwipsPerPixelX, _
                 Me.Top / Screen.TwipsPerPixelY, _
                 Me.width / Screen.TwipsPerPixelX, _
                 Me.height / Screen.TwipsPerPixelY, _
                 SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
                 
End Sub

Private Sub cmdAjuste_Click()
    AjusteImagem
End Sub
Private Sub Form_Load()
     bActive = False
    
    prop = 1
   
    
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
    
    MostraImagemMiniatura listaImagens(1), img.width * proporcao, img.height * proporcao
    prop = proporcao
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    If bActive = True Then Exit Sub
    Carrega
    ajusteTopo
    bActive = True
End Sub
Sub Carrega()
    Set listaImagens = F_SeparaTexto(Me.tag, ",", False)
    Quantidade = listaImagens.Count
    
    If Me.tag <> "" Then
        MostraImagemMiniatura listaImagens(1)
    End If
End Sub


Private Sub Form_Resize()
   ' On Error Resume Next
    pcQuadro.width = Me.width - pcQuadro.Left - 300
'        pcNavegaImgs.width = pcQuadro.width
    ' txtCaminhoImg.width = pcQuadro.width
    
   
End Sub

Private Sub gradeImg_Click()
    MostraImagemMiniatura Me.tag
End Sub



Private Sub mnSair_Click()
    cmndbtnSair_Click
End Sub

Private Sub pcClose_Click()
    Unload Me
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
    
    pcImg.Picture = Nothing
    pcImg.AutoSize = True

    
    If w > 0 And h > 0 Then
         pcImg.PaintPicture LoadPicture(img), 0, 0, w, h
    Else
        prop = 1
        pcImg.Left = 0
        pcImg.Top = 0
        pcImg.Picture = LoadPicture(img)
    End If
  
    pcImg.Refresh
    
    txtCaminhoImg.Text = img
    
   ' Set pcMiniatura.Picture = LoadPicture(s)
    
End Sub





