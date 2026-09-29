VERSION 5.00
Begin VB.Form frmContratoUpload 
   Caption         =   "Contrato"
   ClientHeight    =   7785
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15165
   Icon            =   "frmContratoUpload.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   15165
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcDireita 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   6315
      Left            =   12270
      ScaleHeight     =   6285
      ScaleWidth      =   2865
      TabIndex        =   14
      Top             =   495
      Width           =   2895
      Begin VB.ListBox lst 
         Appearance      =   0  'Flat
         Height          =   3150
         Left            =   0
         TabIndex        =   15
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Documentos Salvos"
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
         Height          =   435
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   15105
      TabIndex        =   6
      Top             =   0
      Width           =   15165
      Begin VB.ComboBox cmbZoom 
         Height          =   315
         ItemData        =   "frmContratoUpload.frx":2E6C6
         Left            =   3135
         List            =   "frmContratoUpload.frx":2E6D3
         TabIndex        =   18
         Text            =   "75"
         Top             =   105
         Width           =   630
      End
      Begin VB.CommandButton cmdProx 
         Caption         =   ">"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdAnterior 
         Caption         =   "<"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox txtPagina 
         Height          =   285
         Left            =   375
         TabIndex        =   7
         Text            =   "1"
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Zoom (%):"
         Height          =   195
         Left            =   2295
         TabIndex        =   17
         Top             =   165
         Width           =   705
      End
      Begin VB.Label lblCaminho 
         AutoSize        =   -1  'True
         Caption         =   "Caminho"
         Height          =   195
         Left            =   3975
         TabIndex        =   12
         Top             =   165
         Width           =   615
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1800
         TabIndex        =   9
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "de"
         Height          =   195
         Left            =   1320
         TabIndex        =   8
         Top             =   120
         Width           =   180
      End
   End
   Begin VB.PictureBox pcBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   15135
      TabIndex        =   4
      Top             =   6810
      Width           =   15165
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   9840
         TabIndex        =   21
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Enabled         =   0   'False
         Height          =   735
         Left            =   5520
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCopia 
         Caption         =   "Salvar Cópia"
         Enabled         =   0   'False
         Height          =   735
         Left            =   3720
         TabIndex        =   19
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar"
         Height          =   735
         Left            =   1920
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddArquivo 
         Caption         =   "Selecionar PDF"
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcVisualiza 
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5595
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   480
      Width           =   9615
      Begin VB.HScrollBar hscrollPdf 
         Height          =   255
         LargeChange     =   500
         Left            =   0
         TabIndex        =   3
         Top             =   5400
         Width           =   9375
      End
      Begin VB.VScrollBar vscrollPDF 
         Height          =   6855
         LargeChange     =   500
         Left            =   9360
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picPDF 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5775
         Left            =   0
         ScaleHeight     =   5745
         ScaleWidth      =   5745
         TabIndex        =   1
         Top             =   0
         Width           =   5775
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnNOvo 
         Caption         =   "Adicionar novo PDF"
      End
      Begin VB.Menu mnAbrir 
         Caption         =   "Abrir documento"
      End
      Begin VB.Menu mnSave 
         Caption         =   "Salvar uma cópia"
      End
      Begin VB.Menu mnExcluir 
         Caption         =   "Excluir documento"
      End
   End
End
Attribute VB_Name = "frmContratoUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim maximizou As Boolean
Dim bActive As Boolean
Dim Arquivos_Gravados As New Collection
Dim zoom As Integer

Private Sub cmbZoom_Click()
    miniatura lblCaminho.caption
End Sub

Private Sub cmbZoom_LostFocus()
    cmbZoom.text = F_SoNumero(cmbZoom.text)
    If cmbZoom.text = "" Then
        cmbZoom.text = "50"
    End If
    If cmbZoom.text > "150" Then
        cmbZoom.text = "100"
    End If
    
End Sub

Private Sub cmdAddArquivo_Click()
    Dim caminho As String
    Dim thumb As String
    On Error GoTo trata_Erro:
    
    Dim s As String
    caminho = BuscaNomeArquivoExterno("Arquivos PDFs (*.PDF)|*.pdf")
    lblCaminho.caption = caminho
    lblCaminho.tag = caminho
    txtPagina.text = "1"
    On Error Resume Next
    lblPaginas.caption = ExtraiNumeroPaginasPDF(caminho)
    
    On Error GoTo trata_Erro
    
    cmdCopia.Enabled = False
    cmdExcluir.Enabled = False
    
    If Not miniatura(caminho) Then
        MsgBox "Não foi possível visualizar a miniatura. Instale o ghostscript para visualizar.", vbExclamation
    End If
    
    If caminho <> "" Then
        If MsgBox("Deseja gravar arquivo selecionado?", vbYesNo + vbQuestion) = vbYes Then
            cmdGravar_Click
        End If
    End If
    
trata_Erro:
    
End Sub
Function miniatura(caminho As String) As Boolean
    Dim thumb As String
    Dim Pic As IPictureDisp
    Dim IMG As New WIACtl.ImageFile
    On Error GoTo trata_Erro
    P_Status "gerando PDF"
    thumb = App.path & "\pic\pdf_" & F_SoNumero(Now) & "_" & txtPagina.text & ".tmp"
    
    If caminho <> "" Then
        ConvertPDFToThumbnail caminho, thumb, Int(txtPagina.text), Int(txtPagina.text), Int(cmbZoom.text)
        wait thumb, 5
        Set IMG = Nothing
        Sleep 1000
        If IsFileLocked(thumb) Then
            MsgBox ""
        Else
            IMG.LoadFile thumb
        End If
        Set picPDF = IMG.FileData.Picture
        P_Status
    End If
    miniatura = True
    P_Status
trata_Erro:
    P_Status
End Function
Function IsFileLocked(filePath As String) As Boolean
    On Error GoTo trata_Erro
    Open filePath For Input As #1
    
    Close #1
    Exit Function
trata_Erro:
    IsFileLocke = True
End Function
Private Function GetPdfInfo(ByVal pdfFilePath As String) As String
    Dim shell As Object
    Dim command As String
    Dim output As String
    Dim tempFilePath As String
    Dim fileNumber As Integer
    
    ' Cria instância do objeto WScript.Shell
    Set shell = CreateObject("WScript.Shell")

    ' Defina o caminho do arquivo PDF (garanta que não tenha aspas duplas adicionais)
    command = "gswin64c.exe -q -dNODISPLAY -dPDFINFO """ & pdfFilePath & """ -c quit"

    ' Defina o caminho do arquivo temporário
    tempFilePath = "C:\temp.txt"

    Set exec = shell.exec(command)

    
     Do While Not exec.StdOut.AtEndOfStream
        outputLine = exec.StdOut.ReadLine()
        output = output & outputLine & vbCrLf
    Loop
    
    ' Retorna a saída completa
    GetPdfInfo = output
        
        
        
    ' Aguarde para garantir que o arquivo seja criado
    DoEvents

    ' Lê a saída do arquivo temporário
    fileNumber = FreeFile
    Open tempFilePath For Input As #fileNumber
    output = Input$(LOF(fileNumber), #fileNumber)
    Close #fileNumber

    ' Retorna o conteúdo do output
    GetPdfInfo = output
End Function

Private Function getLastPage(pdfPath As String) As String
    Dim biopdfPath As String
    Dim ConteudoArquivo As String
    Dim oShell
    Set oShell = CreateObject("Shell.Application")

    'biopdfPath = "C:\Program Files\bioPDF\PDF Writer\gs\gswin64c.exe"  ' Altere para o caminho do BioPDF
    biopdfPath = "gswin64c.exe"  ' Altere para o caminho do BioPDF
    ' Comando para converter a primeira página do PDF em uma imagem PNG
    Dim command As String
    Dim numPages As Integer
    Dim infoPdf As String
    infoPdf = App.path & "\infopdf.tmp"
    command = biopdfPath & " -q -dNODISPLAY -dPDFINFO -c quit" & aspas & pdfPath & aspas
    Dim output As String
    
    output = shell(command, vbHide)
'    ConteudoArquivo = LeArquivoExterno(infoPdf)
    
    Set oExec = oShell.exec(command)
    
    ' Aguarda a execução do comando
    Do While oExec.status <> 0
        WScript.Sleep 100
    Loop
    
    ' Obtém a saída padrão do comando
    sOutput = oExec.StdOut.ReadAll
    
    ' Fecha o processo
    oExec.StdOut.Close
    
    ' Limpa o objeto Shell
    Set oExec = Nothing
    Set oShell = Nothing
    
    ' Exibe a saída no Visual Basic
    MsgBox sOutput



    Dim lastPage As Integer
    If InStr(output, "%%Pages:") > 0 Then
        lastPage = val(Mid(output, InStr(output, "%%Pages:") + 9))
    Else
        ' Lidar com o erro, se necessário
        lastPage = 9999 ' Valor padrão se não encontrar o número de páginas
    End If



    ' Executa o comando
    shell command, vbHide
End Function

Private Sub ConvertPDFToThumbnail(pdfPath As String, outputImage As String, _
    Optional fp As Integer = 1, Optional lp As Integer = 1, Optional zoom As Integer = 50)
    Dim biopdfPath As String
    On Error GoTo trata_Erro:
    'biopdfPath = "C:\Program Files\bioPDF\PDF Writer\gs\gswin64c.exe"  ' Altere para o caminho do BioPDF
    biopdfPath = "gswin64c.exe"  ' Altere para o caminho do BioPDF
    ' Comando para converter a primeira página do PDF em uma imagem PNG
    Dim command As String
    command = biopdfPath & " -sDEVICE=png16m -dFirstPage=" & fp & " -dLastPage=" & lp & " -dNOPAUSE -dBATCH -r" & zoom & " -sOutputFile=" & aspas & outputImage & aspas & " " & aspas & pdfPath & aspas
'    command = biopdfPath & " -sDEVICE=png16m -dFirstPage=" & fp & " -dNOPAUSE -dBATCH -r50 -sOutputFile=" & aspas & outputImage & aspas & " " & aspas & pdfPath & aspas

    ' Executa o comando
    shell command, vbHide
    Exit Sub
trata_Erro:
    MsgBox "Erro ao converter PDF... verifique se o gswin64c.exe está no path do windows" & Err.Description, vbExclamation
End Sub
Sub posicionaControles()
    On Error Resume Next
    pcVisualiza.width = pcDireita.Left - pcVisualiza.Left - 100
    pcVisualiza.height = pcBottom.top - pcVisualiza.top - 100
    
    
    vscrollPDF.height = pcVisualiza.height - 300
    posicionaADireita vscrollPDF, pcVisualiza, 50
    hscrollPdf.width = pcVisualiza.width - 280
    posicionaNoTopo hscrollPdf, pcVisualiza, 50
    posicionaADireita cmdSair, pcBottom
End Sub

Private Sub Form_WindowStateChange(ByVal NewState As Integer)
    ' NewState é um valor que indica o novo estado da janela:
    ' 0 = Normal, 1 = Minimizado, 2 = Maximizado

    Select Case NewState
        Case 1 ' Minimizado
            ' Código a ser executado quando o formulário é minimizado.
        Case 2 ' Maximizado
            ' Código a ser executado quando o formulário é maximizado.
        Case Else ' Normal ou outro estado
            ' Código a ser executado quando o formulário retorna ao estado normal.
    End Select
End Sub

Private Sub cmdAnterior_Click()
    If Int(txtPagina.text) > 1 Then
        txtPagina.text = Int(txtPagina.text) - 1
        miniatura lblCaminho.caption
    End If
End Sub

Private Sub cmdCopia_Click()
    Dim caminho As String
    On Error GoTo trata_Erro
    caminho = arquivo.BuscaNomeArquivoExternoSalvar("*.pdf")
    
    FileCopy lblCaminho.caption, caminho
    shell "explorer " & arquivo.extrai_pasta_arquivo(caminho)
    
trata_Erro:
End Sub

Private Sub cmdExcluir_Click()
    If strExcldoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir o documento selecionado?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    If Not F_senha Then
        Exit Sub
    End If
    Contrato.excluir_arquivo Me.tag, lst.text
    lst.RemoveItem lst.ListIndex
    Form_Activate
End Sub

Private Sub cmdGravar_Click()
    Dim size As Long
    If ja_gravou(lblCaminho.caption) Then
        MsgBox "Este arquivo já foi gravado!", vbExclamation
        Exit Sub
    End If
     
    
    If lblCaminho.tag <> "" Then
        size = GetFileSize(lblCaminho.tag)
        If size / 1024 > 5000 Then
            MsgBox "O arquivo é muito grande. Diminua o tamanho do PDF para no máximo 5Mb e tente novamente", vbExclamation
            Exit Sub
        End If
        GravaArquivo
    End If
End Sub
Function ja_gravou(caminho As String) As Boolean
    Dim i As Integer
    If Arquivos_Gravados.Count = 0 Then Exit Function
    For i = 1 To Arquivos_Gravados.Count
        If Arquivos_Gravados(i) = caminho Then
            ja_gravou = True
            Exit Function
        End If
    Next
End Function
Sub GravaArquivo()
    Dim pasta As String
    Dim arquivo As String
    Dim hash As String
    pasta = caminho_contratos & "\" & Me.tag
    arquivo = pasta & "\" & F_SoNumero(Now) & ".pdf"
    If Dir(pasta, vbDirectory) = "" Then
        MkDir pasta
    End If
    FileCopy lblCaminho.tag, arquivo
    Sleep 100
    hash = GetFileHash(arquivo)
    
    insere_documento_contrato Me.tag, arquivo, hash
    carrega_arquivos
    Arquivos_Gravados.Add arquivo
    Arquivos_Gravados.Add lblCaminho.caption
End Sub
Private Sub cmdProx_Click()
    On Error GoTo trata_Erro
    If Int(txtPagina.text) < Int(lblPaginas.caption) Then
        txtPagina.text = Int(txtPagina.text) + 1
        miniatura lblCaminho.caption
    End If
trata_Erro:
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    bActive = False
    carrega_arquivos
End Sub
Sub carrega_arquivos()
    Dim rPesq As ADODB.Recordset
    Set rPesq = Contrato.buscaDocs_Contrato(Me.tag)
    If rPesq.RecordCount > 0 Then
        lst.Clear
        
        Do While Not rPesq.EOF
            lst.AddItem extrai_nome_arquivo(rPesq!caminho)
            rPesq.MoveNext
        Loop
    End If
    
End Sub
Private Sub Form_Click()
    Form_Resize
End Sub

Private Sub Form_Load()
    resizeForm Me
    zoom = 50
    
    bActive = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If maximizou Then
        Form_Resize
    End If
End Sub

Private Sub Form_Resize()
    
    posicionaControles
    If Me.WindowState = vbMaximized Then
        maximizou = True
    Else
        maximizou = False
    End If
    
End Sub

Private Sub lblCaminho_Click()
    If lblCaminho.tag <> "" Then
        shell "explorer " & extrai_pasta_arquivo(lblCaminho.caption), vbNormalFocus
    End If
End Sub

Private Sub lst_Click()
    copia_arquivo_Temp
End Sub
Sub copia_arquivo_Temp()
    Dim tempFile As String
    Dim Source As String
    On Error GoTo trata_Erro:
    Source = caminho_contratos & "\" & Me.tag & "\" & lst.text
    'tempFile = Environ("TEMP") & "\" & lst.Text
    tempFile = App.path & "\" & lst.text
    
    FileCopy Source, tempFile
    lblCaminho.caption = tempFile
    Arquivos_Gravados.Add tempFile
    lblPaginas.caption = ExtraiNumeroPaginasPDF(tempFile)
    miniatura tempFile
    cmdCopia.Enabled = True
    cmdExcluir.Enabled = True
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub lst_DblClick()
    picPDF_Click
End Sub

Private Sub lst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnMenu
    End If
End Sub

Private Sub mnAbrir_Click()
    picPDF_Click
End Sub

Private Sub mnExcluir_Click()
    If cmdExcluir.Enabled = True Then
    
        cmdExcluir_Click
    End If
End Sub

Private Sub mnNovo_Click()
    cmdAddArquivo_Click
End Sub

Private Sub mnSave_Click()
    If cmdCopia.Enabled = True Then
    
        cmdCopia_Click
    End If
End Sub

Private Sub pcBottom_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If maximizou Then
        Form_Resize
    End If
End Sub

Private Sub pcDireita_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If maximizou Then
        Form_Resize
    End If
End Sub

Private Sub pcVisualiza_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If maximizou Then
        Form_Resize
    End If
End Sub

Private Sub picPDF_Click()
    If lblCaminho.caption <> "" And lblCaminho.caption <> "Caminho" Then
        shell "explorer " & lblCaminho.caption
    Else
        cmdAddArquivo_Click
    End If
    
End Sub



Private Sub scrollZoom_Scroll()

End Sub

Private Sub scrollZoom_Validate(Cancel As Boolean)

End Sub

Private Sub txtPagina_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        miniatura lblCaminho.caption
    End If
End Sub

Private Sub vscrollPDF_Change()
    picPDF.top = 0 - vscrollPDF.Value
    
End Sub
