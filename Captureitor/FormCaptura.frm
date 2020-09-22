VERSION 5.00
Begin VB.Form FormCaptura 
   Caption         =   " Captureitor"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCaptura.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdVentana 
      Caption         =   "Ventana Especifica"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdRectangulo 
      Caption         =   "Rectangulo"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdPantalla 
      Caption         =   "Pantalla Completa"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   6915
      TabIndex        =   1
      Top             =   720
      Width           =   6975
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6480
      Picture         =   "FormCaptura.frx":0ABA
      ScaleHeight     =   450
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "FormCaptura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Anterior As Long
Dim ElHandle As Long
Private Sub CmdPantalla_Click()
Me.Hide
DoEvents
Set Picture1 = CaptureScreen()
Me.Show
End Sub
Private Sub CmdRectangulo_Click()
Me.Hide
DoEvents
Set FormCapturaRect.Picture = CaptureScreen()
FormCapturaRect.Show
End Sub
Private Sub CmdVentana_Click()
picDrag.Visible = True
End Sub

Private Sub Form_Resize()
Picture1.Left = 150
Picture1.Top = 800
Picture1.Height = Me.Height - 1450
Picture1.Width = Me.Width - 400
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 99
Me.MouseIcon = picDrag.Picture
Me.Hide
End Sub
Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call GetWindowInformation(ElHandle)
If Anterior <> ElHandle Then Deshacer Anterior
Pintar_Borde ElHandle
Anterior = ElHandle
End Sub

Private Sub picDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call GetWindowInformation(ElHandle)
Deshacer ElHandle
DoEvents
Set Picture1 = CaptureVentana(ElHandle)
Me.Show
End Sub

