VERSION 5.00
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Begin VB.Form frmCadEti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Etiquetas"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtbxAltBar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   30
      Text            =   "0"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtbxTamBar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   28
      Text            =   "0"
      Top             =   4080
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2880
      ScaleHeight     =   1065
      ScaleWidth      =   2625
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
      Begin BARCODEXLib.BarcodeX brcdx 
         Height          =   615
         Left            =   240
         Top             =   120
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   1085
         _StockProps     =   13
         BackColor       =   -2147483639
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1234567890123"
         Title           =   ""
         Scaling         =   -1
         BarcodeType     =   8
      End
   End
   Begin VB.TextBox txtbxTamFonte 
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      Text            =   "10"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtbxFonte 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
   Begin VB.CommandButton cmndbtnPadrao 
      Caption         =   "Etiqueta Padrão"
      Height          =   735
      Left            =   120
      Picture         =   "frmCadEti.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6480
      Width           =   1815
   End
   Begin VB.PictureBox pcEti 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   2400
      ScaleHeight     =   48.948
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   67.998
      TabIndex        =   24
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtbxQtdLinha 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   21
      Text            =   "0"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtbxDistH 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "0"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtbxQtdCol 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Text            =   "0"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtbxMSup 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Text            =   "0"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtbxMEsq 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Text            =   "0"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtbxLargura 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Text            =   "0"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtbxDistV 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Text            =   "0"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtbxAltura 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "0"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtbxNome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fecha&r"
      Height          =   735
      Left            =   4800
      Picture         =   "frmCadEti.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Sal&var"
      Height          =   735
      Left            =   2880
      Picture         =   "frmCadEti.frx":08B4
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnVisualizar 
      Caption         =   "Visualizar"
      Height          =   255
      Left            =   5160
      TabIndex        =   26
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(em pixels)"
      Height          =   195
      Left            =   2520
      TabIndex        =   32
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Altura Barra"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   31
      Top             =   4440
      Width           =   1065
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tam. Barra"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   29
      Top             =   4080
      Width           =   1065
   End
   Begin VB.Label lblFonte 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fonte:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   2
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs...: Altura em milímetros"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   6120
      Width           =   1920
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Linhas:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   22
      Top             =   3720
      Width           =   1065
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Distância Hor:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   7
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Colunas:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   17
      Top             =   3360
      Width           =   1065
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Margem Sup:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   15
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Margem Esq:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   13
      Top             =   2640
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Largura:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   11
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Distância Vert:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   9
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Altura:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   5
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   0
      Top             =   480
      Width           =   1065
   End
End
Attribute VB_Name = "frmCadEti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnPadrao_Click()
    If Me.tag = "" Then
    
    Else
        If MsgBox("Deseja tornar esta etiqueta padrão?", vbYesNo + vbQuestion) = vbYes Then
            etiqueta.P_AlteraPadrao (Me.tag)
            
        End If
    End If
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.tag = "" Then
        If MsgBox("Confirma gravação da Etiqueta?", vbYesNo + vbQuestion) = vbYes Then
            Me.tag = etiqueta.F_GravaEti(txtbxNome.Text, txtbxAltura.Text, txtbxDistH.Text, _
            txtbxDistV.Text, txtbxLargura.Text, txtbxMEsq.Text, txtbxMSup.Text, _
            txtbxQtdCol.Text, txtbxQtdLinha.Text, txtbxFonte.Text, txtbxTamFonte.Text, _
            txtbxTamBar.Text, txtbxAltBar.Text)
            If Me.tag <> "" Then
                MsgBox "Etiqueta Adicionada com sucesso!", vbInformation
                Form_Activate
            End If
            
            
        End If
    Else
        If MsgBox("Confirma alteração da Etiqueta?", vbYesNo + vbQuestion) = vbYes Then
            
            If etiqueta.F_AlteraEti(Me.tag, txtbxNome.Text, txtbxAltura.Text, txtbxDistH.Text, _
            txtbxDistV.Text, txtbxLargura.Text, txtbxMEsq.Text, txtbxMSup.Text, _
            txtbxQtdCol.Text, txtbxQtdLinha.Text, _
            txtbxFonte.Text, txtbxTamFonte.Text, txtbxTamBar.Text, txtbxAltBar.Text) Then
                MsgBox "Etiqueta alterada com sucesso!", vbInformation
                
            End If
        End If
    End If
End Sub

Private Sub cmndbtnVisualizar_Click()
    ImprimeEti pcEti
End Sub

Sub ImprimeEti(obj As Object)
    sPos = txtbxMEsq.Text
    obj.Cls
    sRef = "00000000"
    obj.CurrentY = 0
    
   ' pcEti.Width = ScaleX(CSng(txtbxLargura.Text), vbTwips, vbMillimeters)
   ' pcEti.Height = CSng(txtbxAltura.Text)
    
    sNome = "Exemplo de Produto"
    
    obj.fontsize = txtbxTamFonte.Text
    obj.FontName = txtbxFonte.Text
    P_ImprQuebraTexto (sNome & " " & sRef), obj, Int(sPos), Int(sPos) + CSng(txtbxLargura.Text)
    'obj.CurrentY = obj.CurrentY + 3.4
    'P_ImprX obj, Trim(Mid(sNome, 26, 24)) & "-" & sRef, Int(sPos), False
    obj.Print
 '   MsgBox brcdx.Width
    brcdx.width = txtbxTamBar.Text
    brcdx.height = txtbxAltBar.Text
    'brcdx.Pictur
'    MsgBox brcdx.Width
    obj.PaintPicture brcdx.Picture, sPos, obj.CurrentY
    obj.CurrentY = obj.CurrentY + 0
                        obj.ScaleMode = Printer.ScaleMode
                        P_ImprX obj, Left("", 40), Int(sPos), False
    
    
    obj.DrawWidth = 1
    P_Line pcEti, 0, 0, CSng(txtbxLargura.Text), 0
    P_Line pcEti, 0, CSng(txtbxAltura.Text), CSng(txtbxLargura.Text), CSng(txtbxAltura.Text)
    P_Line pcEti, CSng(txtbxLargura.Text), 0, CSng(txtbxLargura.Text), CSng(txtbxAltura.Text)
    P_Line pcEti, 0, 0, 0, CSng(txtbxAltura.Text)
    obj.DrawWidth = 2
    P_Line pcEti, 0, CSng(txtbxAltura.Text) + 0.5, CSng(txtbxLargura.Text) + 0.5, CSng(txtbxAltura.Text) + 0.5
    P_Line pcEti, CSng(txtbxLargura.Text) + 0.5, 0.5, CSng(txtbxLargura.Text) + 0.5, CSng(txtbxAltura.Text) + 0.5
End Sub
Private Sub Form_Activate()
    If Me.tag <> "" Then
        carrega
        cmndbtnVisualizar_Click
    End If
    If Me.tag = "" Then
        cmndbtnPadrao.Visible = False
    Else
        
    End If
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_etiquetas where cod_eti='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        caption = !nome_eti
        txtbxAltura.Text = !altura_eti
        txtbxDistH.Text = !disth_eti
        txtbxDistV.Text = !distv_eti
        txtbxLargura.Text = !largura_eti
        txtbxMEsq.Text = !mesq_eti
        txtbxMSup.Text = !msup_eti
        txtbxNome.Text = !nome_eti
        txtbxQtdCol.Text = !qtdcol_eti
        txtbxQtdLinha.Text = !qtdlinha_eti
        txtbxFonte.Text = !fonte_eti
        txtbxTamFonte.Text = !tamfonte_eti
        txtbxAltBar.Text = !altbar_eti
        txtbxTamBar.Text = !tambar_eti
        ajustaTamEti
        .Close
    End With
End Sub
Private Sub txtbxCidade_GotFocus()
    
End Sub
Sub ajustaTamEti()
    'pcEti.ScaleMode = vbTwips
    pcEti.ScaleWidth = txtbxLargura.Text
    pcEti.ScaleHeight = txtbxAltura.Text
    
End Sub
Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub




Private Sub txtbxAltura_GotFocus()
    P_SelecTexto txtbxAltura
End Sub

Private Sub txtbxAltura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxAltura_LostFocus()
    P_ValorLostFocus txtbxAltura
End Sub

Private Sub txtbxDistH_GotFocus()
    P_SelecTexto txtbxDistH
End Sub

Private Sub txtbxDistH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDistH_LostFocus()
    P_ValorLostFocus txtbxDistH
End Sub

Private Sub txtbxDistV_GotFocus()
    P_SelecTexto txtbxDistV
End Sub

Private Sub txtbxDistV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxFonte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFonte
    End If
End Sub
Sub buscaFonte()
    Dim s As String
    On Error Resume Next
    mdiFrm.cmndlg.Flags = cdlCFPrinterFonts
    mdiFrm.cmndlg.ShowFont
    txtbxFonte.Text = mdiFrm.cmndlg.FontName
    txtbxTamFonte.Text = mdiFrm.cmndlg.fontsize
    
    
    
End Sub
Private Sub txtbxLargura_GotFocus()
    P_SelecTexto txtbxLargura
End Sub

Private Sub txtbxLargura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxLargura_LostFocus()
    P_ValorLostFocus txtbxLargura
End Sub

Private Sub txtbxMEsq_GotFocus()
    P_SelecTexto txtbxMEsq
End Sub

Private Sub txtbxMEsq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxMEsq_LostFocus()
    P_ValorLostFocus txtbxMEsq
End Sub

Private Sub txtbxMSup_GotFocus()
    P_SelecTexto txtbxMSup
End Sub

Private Sub txtbxMSup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxMSup_LostFocus()
    P_ValorLostFocus txtbxMSup
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub


Private Sub txtbxQtdCol_GotFocus()
    P_SelecTexto txtbxQtdCol
End Sub

Private Sub txtbxQtdCol_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxQtdCol_LostFocus()
    P_ValorLostFocus txtbxQtdCol
End Sub

Private Sub txtbxQtdLinha_GotFocus()
    P_SelecTexto txtbxQtdLinha
End Sub

Private Sub txtbxQtdLinha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxQtdLinha_LostFocus()
    P_ValorLostFocus txtbxQtdLinha
End Sub

