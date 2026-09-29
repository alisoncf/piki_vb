Attribute VB_Name = "Converte"

'........... Add these defines and declarations to your Global module ...........
' Image descriptor
Declare Function bmpinfo Lib "VIC32.DLL" (ByVal Fname As String, bdat As BITMAPINFOHEADER) As Long
Declare Function allocimage Lib "VIC32.DLL" (image As imgdes, ByVal wid As Long, ByVal leng As Long, ByVal BPPixel As Long) As Long
Declare Function loadbmp Lib "VIC32.DLL" (ByVal Fname As String, desimg As imgdes) As Long
Declare Sub freeimage Lib "VIC32.DLL" (image As imgdes)
Declare Function convert1bitto8bit Lib "VIC32.DLL" (srcimg As imgdes, desimg As imgdes) As Long
Declare Sub copyimgdes Lib "VIC32.DLL" (srcimg As imgdes, desimg As imgdes)
Declare Function savejpg Lib "VIC32.DLL" (ByVal Fname As String, srcimg As imgdes, ByVal quality As Long) As Long

Type imgdes
   ibuff As Long
   stx As Long
   sty As Long
   endx As Long
   endy As Long
   buffwidth As Long
   palette As Long
   colors As Long
   imgtype As Long
   bmh As Long
   hBitmap As Long
End Type

Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Sub ConvertePicture(PathPicBmp As String, pathPicJpg As String)
      Dim tmpimage As imgdes    ' Image descriptors
   Dim tmp2image As imgdes
   Dim rcode As Long
   Dim quality As Long
   Dim vbitcount As Long
   Dim bdat As BITMAPINFOHEADER ' Reserve space for BMP struct
   Dim bmp_fname As String
   Dim jpg_fname As String

   bmp_fname = PathPicBmp
   jpg_fname = pathPicJpg

   quality = 100
   ' Get info on the file we're to load
   rcode = bmpinfo(bmp_fname, bdat)
   If (rcode <> NO_ERROR) Then
      MsgBox "Cannot find file", 0, "Error encountered!"
      Exit Sub
   End If
    
   vbitcount = bdat.biBitCount
   If (vbitcount >= 16) Then  ' 16-, 24-, or 32-bit image is loaded into 24-bit buffer
      vbitcount = 24
   End If
   
   ' Allocate space for an image
   rcode = allocimage(tmpimage, bdat.biWidth, bdat.biHeight, vbitcount)
   If (rcode <> NO_ERROR) Then
     MsgBox "Not enough memory", 0, "Error encountered!"
     Exit Sub
   End If
   
   ' Load image
   rcode = loadbmp(bmp_fname, tmpimage)
   If (rcode <> NO_ERROR) Then
      freeimage tmpimage ' Free image on error
      MsgBox "Cannot load file", 0, "Error encountered!"
      Exit Sub
   End If

   If (vbitcount = 1) Then ' If we loaded a 1-bit image, convert to 8-bit grayscale
       ' because jpeg only supports 8-bit grayscale or 24-bit color images
     rcode = allocimage(tmp2image, bdat.biWidth, bdat.biHeight, 8)
     If (rcode = NO_ERROR) Then
         rcode = convert1bitto8bit(tmpimage, tmp2image)
         freeimage tmpimage  ' Replace 1-bit image with grayscale image
         copyimgdes tmp2image, tmpimage
     End If
   End If

   ' Save image
   rcode = savejpg(jpg_fname, tmpimage, quality)
   freeimage tmpimage
End Sub



