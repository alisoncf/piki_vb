VERSION 5.00
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Begin VB.Form frmGeraBarCode 
   BackColor       =   &H80000018&
   Caption         =   "Geração de Etiquetas"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   Icon            =   "frmGeraBarCode.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   5130
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin BARCODEXLib.BarcodeX picture1 
         Height          =   735
         Left            =   1560
         Top             =   3240
         Visible         =   0   'False
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   1296
         _StockProps     =   13
         BackColor       =   16777215
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
      Begin VB.Frame Frame2 
         Caption         =   "&Estiqueta:"
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   4560
         Width           =   4575
         Begin VB.TextBox txtbxDist 
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
            Left            =   3120
            TabIndex        =   20
            Text            =   "50"
            Top             =   240
            Width           =   735
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
            Left            =   2280
            TabIndex        =   18
            Text            =   "5"
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cmbxTam 
            Height          =   315
            ItemData        =   "frmGeraBarCode.frx":030A
            Left            =   120
            List            =   "frmGeraBarCode.frx":0311
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtbxQtd 
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
            Left            =   3960
            TabIndex        =   14
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Dist.(mm)"
            Height          =   195
            Left            =   3120
            TabIndex        =   21
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Qtd Linha"
            Height          =   195
            Left            =   2280
            TabIndex        =   19
            Top             =   0
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cópias:"
            Height          =   195
            Left            =   3960
            TabIndex        =   13
            Top             =   0
            Width           =   525
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Códigos de Barra:"
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1815
         Begin VB.ListBox lstBarCode 
            Appearance      =   0  'Flat
            Height          =   810
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   1320
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   144
         TabIndex        =   17
         Top             =   3000
         Width           =   2160
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   2880
         Picture         =   "frmGeraBarCode.frx":0330
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5400
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Gerar Etiqueta"
         Height          =   735
         Left            =   3120
         Picture         =   "frmGeraBarCode.frx":047A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtbxDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox txtbxPreco 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   120
         Picture         =   "frmGeraBarCode.frx":05C4
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         BorderWidth     =   2
         Height          =   1575
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Figura gerada (Copiada para o Clipboard)"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de &Barra:"
         Height          =   195
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descrição do Produto:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1590
      End
      Begin VB.Label Label6 
         Caption         =   "Preço:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmGeraBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Printer.PaintPicture Picture1.Picture, 0, 0
    Printer.EndDoc
End Sub


Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub Command2_Click()
    Dim iCount As Integer
    Dim iPula As Integer
    Dim i2 As Integer
    Dim hei As Single
    Dim wid As Single
    Dim x As Single
    Dim y As Single
    For iCount = 1 To CInt(txtbxQtd.Text)
        x = 200: y = 0
        For i2 = 1 To CInt(txtbxQtdLinha.Text)
            iPula = iPula + 1
            Printer.PaintPicture Picture2.image, x, y, Picture2.Width * 0.8, Picture2.Height * 0.75
            x = x + 1900
            If iPula = 5 Then
                y = y + Picture2.Height + 150
                x = 200
            End If
            
        Next
    Next
    Printer.EndDoc
End Sub
Sub imprime(obj As Object)
    obj.PaintPicture Picture2.image, 0, 0
    
End Sub
Private Sub Command3_Click()
    ImprimeEti frmVisualiza.Pic
    frmVisualiza.Show 1
End Sub
Sub ImprimeEti(obj As Object)
    Dim iCount As Integer
    'For iCount = 1 To 30
        obj.PaintPicture Picture1.Picture, 0, 0
        
    'Next
End Sub

Private Sub Command4_Click()
    Dim sTexto As String
    Dim iCount As Integer
    Dim sLinha1 As String
    Dim sLinha2 As String
    On Error Resume Next
    Picture1.caption = Text1.Text
    Picture2.Cls
    Picture2.FontName = "times new roman"
    Picture2.Picture = Nothing
    Picture2.Refresh
    Picture2.ScaleMode = 3
    Picture2.Height = 1480
        
    Picture2.PaintPicture Picture1.Picture, 0, 25
    Picture2.FontBold = False
    FontStuff 10, 10, 90, Picture2, "R$", 110, 85, "Arial"
    FontStuff 10, 13, 90, Picture2, txtbxPreco.Text, 121, 85, "Arial black"
    Picture2.fontsize = 7
    sTexto = QuebraLinha(txtbxDesc.Text, 27)
    If Len(sTexto) <= 28 Then
        sLinha1 = sTexto
        sLinha2 = ""
    Else
    
        For iCount = 1 To Len(sTexto)
        
            If Mid(sTexto, iCount, 1) = Chr(13) Then
                sLinha1 = Left(sTexto, iCount)
                sLinha2 = Mid(sTexto, iCount + 1, Len(sTexto))
            End If
        Next
    End If
    P_PrintXY Picture2, sLinha1, 1, 0, False, 2
    
    P_PrintXY Picture2, sLinha2, 1, 2, False, 2
    ' FontStuff 10, 10, 0, Picture2, Text1, 12, 19, "arial"
    Picture2.FontBold = True
   ' P_PrintXY Picture2, Text1.Text, 1, 19, False, 2
    Picture2.FontBold = False
    Picture2.fontsize = 6
    Picture2.FontUnderline = True
    Picture2.FontName = "arial"
    P_PrintXY Picture2, strFantasia_Emp, 1, 20, False, 2
    Picture2.FontUnderline = False
    'FontStuff 10, 8, 0, Picture2, ConstEmpresaAbrevia, 2, 23, "Times New Roman"
   
End Sub

Private Sub Command5_Click()
    Image1.Height = Picture1.Height
    Image1.Width = Picture1.Width
    If Image1.Stretch = False Then
        Image1.Stretch = True
    Else
        Image1.Stretch = False
    End If
End Sub

Private Sub Form_Activate()

optSize_Click 0
lstBarCode_Click

End Sub

Private Sub optSize_Click(Index As Integer)

Call Text1_Change

End Sub

Private Sub Form_Load()
    Me.Width = 5298
    Me.Height = 6630
    P_CentraForm Me
End Sub

Private Sub Form_Resize()
    P_CentraObejct Frame1, Me
End Sub

Private Sub lstBarCode_Click()
    Text1.Text = lstBarCode.Text
End Sub

Private Sub lstBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub Text1_Change()
'Call DrawBarcode(Text1.Text, Picture1, "", "", False)
'
'MinWidth = 2 * Text1.Left + Text1.Width
'pw = 2 * Picture1.Left + Picture1.Width
'fw = MinWidth
'If pw > fw Then fw = pw
'Picture1.caption = Text1.Text
    
End Sub



Private Sub Text2_Change()

End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd, "###"
End Sub
