VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmFoto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fotos de Produtos"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   6120
      Width           =   8895
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   7680
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "Salvar"
         Height          =   375
         Left            =   6480
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnGravaFoto 
         Caption         =   "&Apagar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frAss 
      Caption         =   "Imagem do Produto:"
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   5535
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8415
         Begin MSComDlg.CommonDialog cmndlg 
            Left            =   1440
            Top             =   2880
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.OLE OLE1 
            Appearance      =   0  'Flat
            Class           =   "PBrush"
            Height          =   5535
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   8415
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   5535
         LargeChange     =   500
         Left            =   8520
         SmallChange     =   100
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   500
         Left            =   120
         SmallChange     =   100
         TabIndex        =   1
         Top             =   5760
         Width           =   8415
      End
      Begin MSAdodcLib.Adodc adodcFirma 
         Height          =   450
         Left            =   5040
         Top             =   4440
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   794
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=firmas"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "firmas"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmFoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub grava()
     Dim strSQL As String
    Dim con As ADODB.Connection
    Set con = New ADODB.Connection
    Dim tabela As ADODB.Recordset
    Set tabela = New ADODB.Recordset
    'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
              App.Path & "\foto.MDB"

    con.Open "DSN=fotos_sty"
    tabela.Open "select * from tab_fotos", con, adOpenKeyset, adLockOptimistic
    tabela.AddNew
    'tabela!seq_fot = "00000001"
    'tabela!data_fir = Date
    'tabela!cod_usu = "000"
    tabela!cod_pro = "00000001"
    GravaBLOB cmndlg.FileName, tabela!foto_fot
    
    tabela.Update
    con.Close
    Unload Me
End Sub
Private Sub cmndbtnBuscar_Click()
    Dim sAux As String
    Dim strSQL As String
    Dim sChar As String
    Dim iCount As Long
    Dim sTrans As String
    Dim strTeste As Variant
    Dim fs
    Dim contPic As String
    Dim sPath As String
    cmndlg.Filter = "Pictures (*.bmp;*.jpg)|*.bmp;*.jpg"
    Set fs = CreateObject("Scripting.FileSystemObject")
    

    cmndlg.ShowOpen
    
    OLE1.CreateEmbed cmndlg.FileName
    
    
    If cmndlg.FileName = "" Then
        MsgBox "Não foi selecionado um arquivo!", vbExclamation
    Else
        If Dir(App.Path & "\pic", vbDirectory) = "" Then
            fs.createfolder App.Path & "\Pic"
        End If
        
        
        sPath = App.Path & "\" & "pic" & "\" & Me.Tag & "_" & Second(Now) & ".pki"
        fs.copyfile cmndlg.FileName, sPath, True
        
        cmndbtnSalvar.Tag = strSQL
        cmndbtnSalvar.Enabled = True
        
    End If

End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    grava
End Sub

Private Sub Command1_Click()
    R = TWAIN_SelectImageSource(Me.hwnd)
End Sub

Private Sub Command2_Click()
      R = TWAIN_AcquireToClipboard(Me.hwnd, t%)
    pcFirma.Picture = Clipboard.GetData(vbCFDIB)
    grava
End Sub

Private Sub Command3_Click()
    OLE1.CreateEmbed cmndlg.FileName
End Sub
Sub buscaFoto()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select "
End Sub
Private Sub HScroll1_Change()
    pcFirma.Left = 100 - HScroll1.Value
End Sub

Private Sub VScroll1_Change()
    pcFirma.Top = 100 - VScroll1.Value
End Sub
