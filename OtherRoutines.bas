Attribute VB_Name = "OtherRoutines"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Type RGBTRIPLE
    red As Integer
    blue As Integer
    green As Integer
End Type
Public Function LongToRGB(ColorValue As Long) As RGBTRIPLE
    Dim rCol As Long, gCol As Long, bCol As Long
    Dim RGBT As RGBTRIPLE
    On Error GoTo ERRLONGTORGB
    RGBT.red = ColorValue And &H10000FF  'this uses binary comparason
    RGBT.green = (ColorValue And &H100FF00) / (2 ^ 8)
    RGBT.blue = (ColorValue And &H1FF0000) / (2 ^ 16)
    LongToRGB = RGBT
ERRLONGTORGB:
End Function



Public Function RGBDist(A As RGBTRIPLE, B As RGBTRIPLE, Technique As Integer) As Single
Dim temp As Single
If Technique = 0 Then
    temp = CLng(A.red - B.red) * CLng(A.red - B.red) + CLng(A.blue - B.blue) * CLng(A.blue - B.blue)
    temp = temp + CLng(A.green - B.green) * CLng(A.green - B.green)
    RGBDist = Sqr(temp)
Else
     RGBDist = A.red - B.red + A.green - B.green + A.blue - B.blue
End If
End Function
