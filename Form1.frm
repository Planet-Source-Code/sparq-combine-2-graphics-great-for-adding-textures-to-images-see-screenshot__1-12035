VERSION 5.00
Begin VB.Form y 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   870
      Left            =   3180
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   810
      ScaleWidth      =   2985
      TabIndex        =   3
      Top             =   60
      Width           =   3045
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   60
      Picture         =   "Form1.frx":194AA
      ScaleHeight     =   810
      ScaleWidth      =   2985
      TabIndex        =   0
      Top             =   60
      Width           =   3045
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3120
      Y1              =   1020
      Y2              =   1260
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   3120
      Y1              =   960
      Y2              =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   660
      TabIndex        =   2
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "y"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long

Dim r, g, b As Integer
Dim r2, g2, b2 As Integer
Dim percent As Integer
Dim Col As Long


Private Function rtnColorSplit(color1 As OLE_COLOR, color2 As OLE_COLOR) As OLE_COLOR
On Error Resume Next
r = color1 Mod 256
b = Int(color1 / 65536)
g = (color1 - (b * 65536) - r) / 256
r2 = color2 Mod 256
b2 = Int(color2 / 65536)
g2 = (color2 - (b2 * 65536) - r2) / 256

percent = 50 'Colors meet Halfway
r = r + ((r2 - r) * (percent / 100))
g = g + ((g2 - g) * (percent / 100))
b = b + ((b2 - b) * (percent / 100))
rtnColorSplit = RGB(r, g, b)
End Function


Private Sub Command1_Click()
  Dim X As Long
  Dim y As Long
    
    If Picture1.Width > Picture2.Width Then
        Picture3.Width = Picture1.Width
    Else
        Picture3.Width = Picture2.Width
    End If
    
    If Picture1.Height > Picture2.Height Then
        Picture3.Height = Picture1.Height
    Else
        Picture3.Height = Picture2.Height
    End If
    
    Picture1.Width = Picture3.Width
    Picture2.Width = Picture3.Width
    Picture1.Height = Picture3.Height
    Picture2.Height = Picture3.Height
    
    Picture3.Left = Picture1.Left + (Picture1.Width / 2) + 60
    Picture3.Visible = True
    For y = 0 To Picture3.ScaleWidth
        For X = 0 To Picture3.ScaleHeight
            SetPixelV Picture3.hdc, X, y, rtnColorSplit(GetPixel(Picture1.hdc, X, y), GetPixel(Picture2.hdc, X, y))
            DoEvents
        Next X
    Next y
    Command1.Caption = "true"
End Sub

Private Sub Form_Load()
    Picture1.ScaleMode = vbPixels
    Picture1.Picture = LoadPicture(App.Path & "\logo1.gif")
    Picture2.Picture = LoadPicture(App.Path & "\logo2.gif")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Label1 = "X:" & vbCrLf & "Y:"
    Label2 = X & vbCrLf & y
        Dim XX As Long, YY As Long, A As Long
        XX = X: YY = y
        'Set the picturebox' backcolor
End Sub
