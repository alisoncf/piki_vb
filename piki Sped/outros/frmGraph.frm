VERSION 5.00
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "graphs32.ocx"
Begin VB.Form frmGraph 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Visualização do Gráficos"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   Icon            =   "frmGraph.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   240
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6135
      Left            =   11160
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox pcGr 
      Height          =   6135
      Left            =   240
      ScaleHeight     =   6075
      ScaleWidth      =   10875
      TabIndex        =   2
      Top             =   360
      Width           =   10935
      Begin GraphsLib.Graph grph 
         Height          =   6135
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   10575
         _Version        =   393216
         _ExtentX        =   18653
         _ExtentY        =   10821
         _StockProps     =   96
         BorderStyle     =   1
         Background      =   "15~-1~-1~-1~-1~-1~-1"
         DataLabels      =   1
         FontSize        =   "175~150~100~100"
         GraphStyle      =   4
         GraphType       =   1
         LeftTitleStyle  =   1
         OverlayColor    =   "0"
         OverlayPattern  =   "3"
         OverlaySymbol   =   "0"
         PropertyPages   =   1
         ToolStat        =   "0~0~0~1~0~0~0~0~0~0~0~0~0~0~1~0~0~0~0~0"
         SmartLabels     =   1
         EnhancedFontBehavior=   "1~1~0~0"
         FontSizeEx      =   "12~8~8~8"
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   9000
      Picture         =   "frmGraph.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   6600
      Picture         =   "frmGraph.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   2295
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub imprime(obj As Object)
    grph.DrawMode = graphCopy
    obj.ScaleMode = 3
    P_CabRelTimbr grph.GraphTitle & " " & Date, obj
    obj.PaintPicture Clipboard.GetData, 0, 15
    
    
    
    P_FimRel obj, constNome_sis & "-" & constVersao_Sis, , strFantasia_Emp
    
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
'    frmMenu.Visible = False
End Sub

Private Sub Form_Load()
    Me.Height = 7395
    Me.Width = 11475
    P_CentraForm Me
End Sub

Private Sub Form_LostFocus()
    frmMenu.Visible = True
End Sub

Private Sub Form_Resize()
    
    pcGr.Width = Me.Width - 100
    
    
    P_CentraObejct grph, Me, False
    
    
    
    cmndbtnImprimir.Top = Me.Height - 1100
    cmndbtnSair.Top = Me.Height - 1100
    
    
    
    cmndbtnSair.Left = grph.Left + grph.Width - cmndbtnSair.Width
    cmndbtnImprimir.Left = grph.Left
    pcGr.Height = cmndbtnSair.Top - 400
    
    grph.Height = cmndbtnSair.Top - 500
    
    '.Top = Me.Height - 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    frmMenu.Visible = True
End Sub
