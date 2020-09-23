VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Edge Trace"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   690
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   525
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   654
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   $"edge.frx":0000
      Height          =   735
      Left            =   6300
      TabIndex        =   9
      Top             =   7080
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use Threshold in GrayScale Edge Trace"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   7260
      Width           =   3195
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GrayScale Edge Trace (Use Threshold of 230)"
      Height          =   675
      Left            =   7560
      TabIndex        =   7
      Top             =   6300
      Width           =   2235
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   4020
      Top             =   7500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.bmp|*.bmp"
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2340
      TabIndex        =   5
      Text            =   "270"
      Top             =   7200
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restore Image"
      Height          =   675
      Left            =   2640
      TabIndex        =   4
      Top             =   6300
      Width           =   2535
   End
   Begin VB.PictureBox Image2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   0
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   653
      TabIndex        =   3
      Top             =   0
      Width           =   9795
   End
   Begin VB.PictureBox Image1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   0
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   653
      TabIndex        =   2
      Top             =   0
      Width           =   9795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "B&&W Edge Trace (Use Threshold of 270)"
      Height          =   675
      Left            =   5160
      TabIndex        =   1
      Top             =   6300
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Image"
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   6300
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Edge Threshold Value:"
      Height          =   255
      Left            =   660
      TabIndex        =   6
      Top             =   7260
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Text1.Text = 96
    Command2.Caption = "B&&W Edge Trace (Use Threshold of 96)"
    Command4.Caption = "GrayScale Edge Trace (Use Threshold of 96)"
Else
    Text1.Text = 270
    Command2.Caption = "B&&W Edge Trace (Use Threshold of 270)"
    Command4.Caption = "GrayScale Edge Trace (Use Threshold of 230)"
End If
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
Check2_Click
End Sub


Private Sub Command1_Click()
On Error GoTo COMMAND1EXIT
CMD.ShowOpen
If Trim$(CMD.FileName = "") Then Exit Sub
Image1.Picture = LoadPicture(CMD.FileName)
Image2.Picture = Image1.Picture
Image2.Top = Image1.Top
Image2.Left = Image1.Left
DoEvents
Command1.Top = Image1.Top + Image1.Height + 4
Command2.Top = Image1.Top + Image1.Height + 4
Command3.Top = Image1.Top + Image1.Height + 4
Command4.Top = Image1.Top + Image1.Height + 4
Label1.Top = Command1.Top + Command1.Height + 20
Text1.Top = Command1.Top + Command1.Height + 20
Check1.Top = Command1.Top + Command1.Height + 20
Check2.Top = Command1.Top + Command1.Height + 20
DoEvents
Form1.Height = (32 + Check2.Top + Check2.Height) * Screen.TwipsPerPixelY
W = (Image1.Left + Image1.Width)
If W < (Command4.Left + Command4.Width) Then
    W = (Command4.Left + Command4.Width + 8)
End If
Form1.Width = W * Screen.TwipsPerPixelX
COMMAND1EXIT:
End Sub


Private Sub Command2_Click()
Dim temp As Long
Dim IsEdge As Boolean
Dim Diff As Single
Dim Orig As RGBTRIPLE
Dim Other As RGBTRIPLE
For x% = 1 To Image1.Width - 1
    For y% = 1 To Image1.Height - 1
        temp = Image1.Point(x%, y%)
        If temp <> -1 Then
        Orig = LongToRGB(temp)
        IsEdge = False
        Diff = 0
        For xx% = x% - 1 To x% + 1
            For yy% = y% - 1 To y% + 1
            If (xx% <> x%) Or (yy% <> y%) Then
                temp = Image1.Point(xx%, yy%)
                If temp <> -1 Then
                Other = LongToRGB(temp)
                Diff = Diff + RGBDist(Other, Orig, Check2.Value)
                End If
            End If
            Next yy%
        Next xx%
        End If
If (Diff) > Abs(Val(Text1.Text)) Then
    Image2.PSet (x%, y%), 0
    IsEdge = True
Else
    Image2.PSet (x%, y%), RGB(255, 255, 255)
End If
    Next y%
    If x% Mod 10 = 0 Then
        Image2.Refresh
    End If
Next x%
End Sub


Private Sub Command3_Click()
Image2.Picture = Image1.Picture
End Sub


Private Sub Command4_Click()
Dim temp As Long
'Dim IsEdge As Boolean
Dim Diff As Single
Dim Orig As RGBTRIPLE
Dim Other As RGBTRIPLE
For x% = 1 To Image1.Width - 1
    For y% = 1 To Image1.Height - 1
        temp = Image1.Point(x%, y%)
        If temp <> -1 Then
        Orig = LongToRGB(temp)
        Diff = 0
        For xx% = x% - 1 To x% + 1
            For yy% = y% - 1 To y% + 1
            If (xx% <> x%) Or (yy% <> y%) Then
                temp = Image1.Point(xx%, yy%)
                If temp <> -1 Then
                Other = LongToRGB(temp)
                'Diff = Diff + ((Other.red - Orig.red))
                'Diff = Diff + ((Other.blue - Orig.blue))
                'Diff = Diff + ((Other.green - Orig.green))
                Diff = Diff + RGBDist(Other, Orig, Check2.Value)
                End If
            End If
            Next yy%
        Next xx%
        End If
If Check1.Value = 0 Then
    Diff = 255 - Diff
    If Diff < 0 Then Diff = 0
    If Diff > 255 Then Diff = 255
    Image2.PSet (x%, y%), RGB(Diff, Diff, Diff)
Else
    If Diff > Text1.Text Then
        Diff = 255 - (Diff - Text1.Text)
        If Diff < 0 Then Diff = 0
        If Diff > 255 Then Diff = 255
        Image2.PSet (x%, y%), RGB(Diff, Diff, Diff)
    Else
        Image2.PSet (x%, y%), RGB(255, 255, 255)
    End If
End If
'If IsEdge = False Then
'    Image2.PSet (x%, y%), RGB(255, 255, 255)
'End If
    Next y%
If x% Mod 10 = 0 Then
        Image2.Refresh
    End If
Next x%
End Sub


