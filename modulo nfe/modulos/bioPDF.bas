Attribute VB_Name = "bioPDF"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SETTINGS_PROGID = "biopdf.PDFSettings"
Const UTIL_PROGID = "biopdf.PDFUtil"

Private Function PrinterIndex(ByVal printerName As String) As Integer
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If LCase(Printers(i).deviceName) Like LCase(printerName) Then
            PrinterIndex = i
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function

Function ApontarPrinterToPDF(toPDF As String, Optional folder As String = "app.path\pdf", Optional abrePDF As Boolean = True) As String
    Dim prtidx As Integer
    Dim sPrinterName As String
    Dim settings As Object
    Dim util As Object
    If folder = "" Then
        folder = "app.path\pdf"
    End If
    If folder = "app.path\pdf" Then
        If Dir(App.Path & "\pdf", vbDirectory) = "" Then
            MkDir App.Path & "\pdf"
        
            
        End If
        toPDF = App.Path & "\pdf\" & toPDF
    Else
        If folder <> "" Then
            toPDF = folder & toPDF
        End If
    End If
    
    
    
    
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    
    Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.printerName = sPrinterName
    settings.setValue "Output", toPDF
    settings.setValue "ConfirmOverwrite", "no"
    settings.setValue "ShowSaveAS", "never"
    settings.setValue "ShowSettings", "never"
    If abrePDF = True Then
        settings.setValue "ShowPDF", "yes"
    Else
        settings.setValue "ShowPDF", "no"
    End If
    settings.setValue "RememberLastFileName", "no"
    settings.setValue "RememberLastFolderName", "no"
    settings.WriteSettings True
    
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then Err.Raise 1000, , _
        "No printer was found by the name of '" & sPrinterName & "'."
        
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
    ApontarPrinterToPDF = toPDF
End Function


Private Sub cmdPrint_Click()
    Dim prtidx As Integer
    Dim sPrinterName As String
    Dim settings As Object
    Dim util As Object
    
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    
    Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.printerName = sPrinterName
    settings.setValue "Output", "\myfile.pdf"
    settings.setValue "ConfirmOverwrite", "no"
    settings.setValue "ShowSaveAS", "never"
    settings.setValue "ShowSettings", "never"
    settings.setValue "ShowPDF", "yes"
    settings.setValue "RememberLastFileName", "no"
    settings.setValue "RememberLastFolderName", "no"
    settings.WriteSettings True
    
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then Err.Raise 1000, , _
        "No printer was found by the name of '" & sPrinterName & "'."
        
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
        
    Rem -- Print something
    'If optOrientation(0).Value Then
        Printer.Orientation = PrinterObjectConstants.vbPRORPortrait
    'Else
    '    Printer.Orientation = PrinterObjectConstants.vbPRORLandscape
    'End If
    
    Rem -- Set paper size
    Rem -- http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.powerpacks.printing.compatibility.vb6.printer.papersize.aspx
    Rem -- Note: Custom paper size is not supported by VB6 after Windows 98.
    Printer.PaperSize = vbPRPSA3
    
    Printer.fontsize = 50
    Printer.Print "Hello VB6..."
    Printer.fontsize = 20
    Printer.ForeColor = vbBlue
    Printer.Print "The time is " & Now
    Printer.EndDoc
    
    Rem -- Wait for runonce settings file to disappear
    Dim runonce As String
    runonce = settings.GetSettingsFilePath(True)
    While Dir(runonce, vbNormal) <> ""
        Sleep 100
    Wend
    
    MsgBox "myfile.pdf was saved on your desktop", vbInformation, "PDF Created"
End Sub

