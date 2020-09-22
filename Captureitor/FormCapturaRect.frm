VERSION 5.00
Begin VB.Form FormCapturaRect 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2265
   ClientLeft      =   3975
   ClientTop       =   5400
   ClientWidth     =   2685
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   DrawWidth       =   2
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   179
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FormCapturaRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XStart, YStart As Single
Dim XPrevious, YPrevious As Single
Private Sub Form_Activate()
Me.Left = -2
Me.Top = -2
Me.Width = Screen.Width + 2
Me.Height = Screen.Height + 2
Me.DrawStyle = 2
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   XStart = X: YStart = Y: XPrevious = XStart: YPrevious = YStart
   FormCapturaRect.AutoRedraw = False
Else
   FormCaptura.Show
   Unload Me
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
   FormCapturaRect.Line (XStart, YStart)-(XPrevious, YPrevious), , B
   FormCapturaRect.Refresh
   FormCapturaRect.Line (XStart, YStart)-(X, Y), , B
   XPrevious = X
   YPrevious = Y
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Dim X1 As Single, Y1 As Single
Dim CopyWidth As Single, CopyHeight As Single
Dim PictWidth As Single, PictHeight As Single

FormCapturaRect.Line (XStart, YStart)-(XPrevious, YPrevious), , B
FormCapturaRect.Refresh
If X > XStart Then X1 = XStart Else X1 = X
If Y > YStart Then Y1 = YStart Else Y1 = Y
CopyWidth = Abs(X - XStart)
CopyHeight = Abs(Y - YStart)

FormCaptura.Picture1 = CaptureWindow(FormCapturaRect.hwnd, X1, Y1, Abs(X - XStart), Abs(Y - YStart))
FormCaptura.Show
DoEvents

Unload Me
End If
End Sub


