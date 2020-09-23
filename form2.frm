VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Dim Down As Boolean

Private Sub Form_Load()
    ScaleMode = vbPixels
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Down = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Down Then
        SetPixelV Form1.hdc, X, Y, Not (GetPixel(Form1.hdc, X, Y))
        
        SetPixelV Form1.hdc, X - 10, Y - 10, Not (GetPixel(Form1.hdc, X - 10, Y - 10))
        SetPixelV Form1.hdc, X + 10, Y + 10, Not (GetPixel(Form1.hdc, X + 10, Y + 10))
        SetPixelV Form1.hdc, X - 10, Y + 10, Not (GetPixel(Form1.hdc, X - 10, Y + 10))
        SetPixelV Form1.hdc, X + 10, Y - 10, Not (GetPixel(Form1.hdc, X + 10, Y - 10))
        
        SetPixelV Form1.hdc, X - 10, Y, Not (GetPixel(Form1.hdc, X - 10, Y))
        SetPixelV Form1.hdc, X + 10, Y, Not (GetPixel(Form1.hdc, X + 10, Y))
        
        SetPixelV Form1.hdc, X, Y - 10, Not (GetPixel(Form1.hdc, X, Y - 10))
        SetPixelV Form1.hdc, X, Y + 10, Not (GetPixel(Form1.hdc, X, Y + 10))
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Down = False
End Sub
