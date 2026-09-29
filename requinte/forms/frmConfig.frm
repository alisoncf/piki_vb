VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações da Aparência"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   4995
   Begin VB.Frame frControle 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   4935
      Begin VB.CommandButton cmndbtnConfirmar 
         BackColor       =   &H8000000C&
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   1200
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   24
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   375
         Left            =   3120
         TabIndex        =   23
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cores:"
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4695
         Begin VB.PictureBox pcCores 
            Height          =   255
            Index           =   0
            Left            =   1800
            ScaleHeight     =   195
            ScaleWidth      =   435
            TabIndex        =   14
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmndbtnCores 
            Caption         =   "Fundo DeskTop"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnCores 
            Caption         =   "Caixa de Texto"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnCores 
            Caption         =   "Fonte"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   1575
         End
         Begin VB.PictureBox pcCores 
            Height          =   255
            Index           =   1
            Left            =   1800
            ScaleHeight     =   195
            ScaleWidth      =   435
            TabIndex        =   10
            Top             =   600
            Width           =   495
         End
         Begin VB.PictureBox pcCores 
            Height          =   255
            Index           =   2
            Left            =   1800
            ScaleHeight     =   195
            ScaleWidth      =   435
            TabIndex        =   9
            Top             =   960
            Width           =   495
         End
         Begin VB.CommandButton cmndbtnPadrão 
            Caption         =   "Restaura Cores Padrão"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Width           =   4455
         End
         Begin VB.PictureBox pcCores 
            Height          =   255
            Index           =   3
            Left            =   4080
            ScaleHeight     =   195
            ScaleWidth      =   435
            TabIndex        =   7
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmndbtnCores 
            Caption         =   "Fonte Cx de Texto"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   6
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnCores 
            Caption         =   "Formulário"
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   5
            Top             =   600
            Width           =   1575
         End
         Begin VB.PictureBox pcCores 
            Height          =   255
            Index           =   4
            Left            =   4080
            ScaleHeight     =   195
            ScaleWidth      =   435
            TabIndex        =   4
            Top             =   600
            Width           =   495
         End
         Begin VB.PictureBox pcCores 
            Height          =   255
            Index           =   5
            Left            =   4080
            ScaleHeight     =   195
            ScaleWidth      =   435
            TabIndex        =   3
            Top             =   960
            Width           =   495
         End
         Begin VB.CommandButton cmndbtnCores 
            Caption         =   "Painéis"
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   2
            Top             =   960
            Width           =   1575
         End
      End
      Begin MSComDlg.CommonDialog CmnDlg 
         Left            =   1800
         Top             =   3720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Caption         =   "Figura de Fundo:"
         Height          =   2535
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   4695
         Begin VB.CommandButton cmndbtBuscaPic 
            Caption         =   "&Buscar Figura"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   1575
         End
         Begin VB.Frame Frame3 
            Caption         =   "Barra de Ferramentas:"
            Height          =   615
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   2055
            Begin VB.CheckBox chkVisib 
               BackColor       =   &H8000000C&
               Caption         =   "Visível"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Aparência:"
            Height          =   615
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Width           =   2055
            Begin VB.ComboBox cmbxAppearance 
               Height          =   315
               ItemData        =   "frmConfig.frx":08CA
               Left            =   120
               List            =   "frmConfig.frx":08D4
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.OptionButton optnNoPic 
            Caption         =   "Nenhuma"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optnPic 
            Caption         =   "Fundo com Figura"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   1575
         End
         Begin VB.Image imgFundo 
            Height          =   2055
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtBuscaPic_Click()
    cmndlg.filter = "Jpeg (*.jpg)|*.Jpg|BitMaps(*.bmp)|*.bmp"
    

    cmndlg.ShowOpen
    imgFundo.Picture = LoadPicture(cmndlg.FileName)
    imgFundo.tag = (cmndlg.FileName)
End Sub

Private Sub cmndbtnConfirmar_Click()
    Dim strSQL As String
    
    
    
    If pcCores(1).BackColor = pcCores(2).BackColor Then
        MsgBox "As cores da Fonte e da Caixa de Texto não podem ser iguais"
        Exit Sub
    End If
    If F_BuscaUmCampo("Cod_Usu", "tab_aparencia", " cod_usu='" & strCod_Usu & "'") <> "" Then
        alteraCor
    Else
        insereCor
    End If
    
    
End Sub
Sub insereCor()
    On Error GoTo trata_Erro
    strSQL = "insert into tab_aparencia(cod_usu, corfundo_par,corcaixa_par,"
    strSQL = strSQL & "Corfonte_par,corfontecaixa_par,corform_par,corframe_par,"
    strSQL = strSQL & "pathpic_par,toolvisible_par,appearance_par)Values("
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & pcCores(0).BackColor & "', "
    strSQL = strSQL & "'" & pcCores(1).BackColor & "', "
    strSQL = strSQL & "'" & pcCores(2).BackColor & "', "
    strSQL = strSQL & "'" & pcCores(3).BackColor & "', "
    strSQL = strSQL & "'" & pcCores(4).BackColor & "', "
    strSQL = strSQL & "'" & pcCores(5).BackColor & "', "
    If InStr(1, imgFundo.tag, "\") Then
        strSQL = strSQL & "'" & ContraBarraPipe(imgFundo.tag) & "', "
    Else
        strSQL = strSQL & "'', "
    End If
    If chkVisib.Value = 0 Then
        strSQL = strSQL & "'N' ,"
    Else
        strSQL = strSQL & "'S' ,"
    End If
    strSQL = strSQL & " '" & cmbxAppearance.ListIndex & "')"
    
    Conn.Execute strSQL
    If InStr(1, imgFundo.tag, "|") > 0 Then
        mdiFrm.Picture = LoadPicture(ContraBarraPipe(imgFundo.tag))
    Else
        mdiFrm.Picture = LoadPicture(imgFundo.tag)
    End If
    mdiFrm.BackColor = pcCores(0).BackColor
    'mdiFrm.Toolbar1.Visible = chkVisib.Value = 1
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub alteraCor()
    On Error GoTo trata_Erro
    strSQL = "update tab_aparencia set corfundo_par='" & pcCores(0).BackColor & "', "
    strSQL = strSQL & "corcaixa_par='" & pcCores(1).BackColor & "', "
    strSQL = strSQL & "corfonte_par='" & pcCores(2).BackColor & "', "
    strSQL = strSQL & "corfontecaixa_par='" & pcCores(3).BackColor & "', "
    strSQL = strSQL & "corform_par='" & pcCores(4).BackColor & "', "
    strSQL = strSQL & "corframe_par='" & pcCores(5).BackColor & "', "
    If InStr(1, imgFundo.tag, "\") Then
        strSQL = strSQL & "pathpic_par='" & ContraBarraPipe(imgFundo.tag) & "', "
    Else
        strSQL = strSQL & "pathpic_par='', "
    End If
    If chkVisib.Value = 0 Then
        strSQL = strSQL & "toolvisible_par='N' ,"
    Else
        strSQL = strSQL & "toolvisible_par='S' ,"
    End If
    strSQL = strSQL & " appearance_par = '" & cmbxAppearance.ListIndex & "'"
    strSQL = strSQL & " WHERE COD_USU='" & strCod_Usu & "' "
    Conn.Execute strSQL
    If InStr(1, imgFundo.tag, "|") > 0 Then
        mdiFrm.Picture = LoadPicture(ContraBarraPipe(imgFundo.tag))
    Else
        mdiFrm.Picture = LoadPicture(imgFundo.tag)
    End If
    mdiFrm.BackColor = pcCores(0).BackColor
    'mdiFrm.Toolbar1.Visible = chkVisib.Value = 1
Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub
    
Private Sub cmndbtnCores_Click(Index As Integer)
    Dim iCor As Long
    cmndlg.ShowColor
    iCor = cmndlg.color
    pcCores(Index).BackColor = iCor
End Sub

Private Sub cmndbtnPadrão_Click()
    pcCores(0).BackColor = vbApplicationWorkspace
    pcCores(1).BackColor = vbWindowBackground
    pcCores(2).BackColor = vbWindowText
    pcCores(3).BackColor = vbWindowText
    pcCores(4).BackColor = vbButtonFace
    pcCores(5).BackColor = vbButtonFace
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim iCount As Integer
    Dim rPesq As New ADODB.Recordset
    P_CentraForm Me
    On Error GoTo trata_Erro:
    'Me.Height = 5310
    'Me.Width = 4965
    strSQL = "select * from tab_aparencia"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            pcCores(0).BackColor = !corfundo_par
            pcCores(1).BackColor = !corcaixa_par
            pcCores(2).BackColor = !corfonte_par
            pcCores(3).BackColor = !corfontecaixa_par
            pcCores(4).BackColor = !corform_par
            If !pathpic_par <> "" Then
                imgFundo.tag = !pathpic_par
                
                imgFundo.Picture = LoadPicture(ContraBarraPipe(!pathpic_par))
                
                optnPic.Value = True
            End If
            If !toolvisible_par = "S" Then
                chkVisib.Value = 1
            Else
                chkVisib.Value = 0
            End If
            
            On Error Resume Next
            cmbxAppearance.ListIndex = !appearance_par
            
                
            
        End If
        .Close
    End With
    MudaAparencia Me
    Exit Sub
trata_Erro:
    
End Sub


Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub optnNoPic_Click()
    imgFundo.Visible = False
    cmndbtBuscaPic.Enabled = False
    imgFundo.tag = ""
End Sub

Private Sub optnPic_Click()
    imgFundo.Visible = True
    cmndbtBuscaPic.Enabled = True
End Sub
