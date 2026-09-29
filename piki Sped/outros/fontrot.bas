Attribute VB_Name = "fontrot"
Option Explicit

Public Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFacename As String * 33
End Type

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Sub CheckVals()
  'Command1.Enabled = ((val(txtDegree.Text) < 360) And val(txtSize.Text) > 7)
End Sub

Sub FontStuff(Tam As Integer, tamFonte As Integer, angDeg As Integer, obj As Object, Text As String, X As Single, _
    Y As Single, Optional NomeFonte As String = "Arial Black", Optional optFont As Integer = 0)
  
  'optfont = 0 normal, 1 negrito, 2 italico, 3 negrito e italico
  
  On Error GoTo GetOut
  Dim F As LOGFONT, hPrevFont As Long, hFont As Long, FontName As String
  Dim fontsize As Integer
  fontsize = val(tamFonte)

  F.lfEscapement = 10 * angDeg 'rotation angle, in tenths
  
  If optFont = 1 Then
    
  End If
  FontName = NomeFonte + Chr$(0) 'null terminated
  
  F.lfFacename = FontName
  F.lfHeight = (fontsize * -20) / Screen.TwipsPerPixelY
  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(obj.hdc, hFont)
  obj.CurrentX = X
  obj.CurrentY = Y
  obj.Print Text
  
'  Clean up, restore original font
  hFont = SelectObject(obj.hdc, hPrevFont)
  DeleteObject hFont
  
  Exit Sub
GetOut:
  Exit Sub

End Sub

