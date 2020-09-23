VERSION 5.00
Begin VB.Form frmSaver 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   3
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmSaver.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   211
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrBendTwister 
      Interval        =   1
      Left            =   90
      Top             =   90
   End
End
Attribute VB_Name = "frmSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BendTwisterSaver © May, 2004 - Ed Gabel (edgabel@comcast.net)
'Will run properly in Windows 9X, 2000, XP.  Written in VB6.
'**************************************************************************************
Option Explicit

Private Sub Form_Load()

    'in Windows 2000, non password-protected screen savers will start
    'minimized and the following line will fix that.
    WindowState = vbMaximized

    'find out if we are running under NT-type sytems (NT, Win2K, XP, etc.)
    Call GetVersion32

    'tell the system that this is a screen saver.  Ctrl-Alt-Del will be disabled
    'on Win9x systems.  NT handles password-protected screen savers at
    'the system level, so Ctrl-Alt-Del cannot be disabled.
    tempLong = SystemParametersInfo(SPI_SCREENSAVERRUNNING, 1&, 0&, 0&)
    
    Call LoadPrefs  'load configuration information
    Call AdjustScreenRes  'set for current screen resolution
    
End Sub

Private Sub Finish()
    If RunMode = rmScreenSaver Then ShowCursor True
    End
End Sub

Private Sub Form_Click()
    Call Finish
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call Finish
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Finish
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Finish
End Sub

Private Sub Form_Unload(Cancel As Integer)  'redisplay the cursor if we hid it in Sub Main
    Call Finish
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x1 As Single, y1 As Single)
Static Counts As Integer
    Counts = Counts + 1  'allow enough time for program to run
    If Counts > 5 Then If RunMode = rmScreenSaver Then Call Finish
End Sub

Private Sub ColorSet()  'set color and shade change

    r = r - 1: g = g - 1: b = b - 1
    If r < 40 Or g < 40 Or b < 40 Then  'minimize dark shading
        r = Rnd * 255
        g = Rnd * 255
        b = Rnd * 255
    End If
    
    ForeColor = RGB(r, g, b)
    
End Sub

'draws and rotates all of the shapes
Private Sub Redo(X As Single, Y As Single, TotalRedo As Integer, StartRedo As Integer, CurrentRedo As Integer)

Dim x2 As Single, y2 As Single
Dim i As Integer

For i = 0 To Shape
    If Odds(CurrentRedo) = 1 Then
        x2 = Cos((Angle + (360 / Shape) * i) / Complex) * (Size * 10) / ((CurrentRedo + 0.1) * 2) + X
        y2 = Sin((Angle + (360 / Shape) * i) / Complex) * (Size * 10) / ((CurrentRedo + 0.1) * 2) + Y
        If StartRedo - 1 < CurrentRedo Then Line (x2, y2)-(Cos((Angle + (360 / Shape) * (i - 1)) / Complex) * (Size * 10) / ((CurrentRedo + 1) * 2) + X, Sin((Angle + (360 / Shape) * (i - 1)) / Complex) * (Size * 10) / ((CurrentRedo + 1) * 2) + Y)
    Else
        x2 = Sin((Angle + (360 / Shape) * i) / Complex) * (Size * 10) / ((CurrentRedo + 1) * 2) + X
        y2 = Cos((Angle + (360 / Shape) * i) / Complex) * (Size * 10) / ((CurrentRedo + 1) * 2) + Y
        If StartRedo - 1 < CurrentRedo Then Line (x2, y2)-(Sin((Angle + (360 / Shape) * (i - 1)) / Complex) * (Size * 10) / ((CurrentRedo + 1) * 2) + X, Cos((Angle + (360 / Shape) * (i - 1)) / Complex) * (Size * 10) / ((CurrentRedo + 1) * 2) + Y)
    End If
    
    If CurrentRedo >= TotalRedo - 1 Then
    Else
        Redo x2, y2, TotalRedo, StartRedo, CurrentRedo + 1
        End If
Next i

End Sub

Private Sub tmrBendTwister_Timer()

Randomize Timer

If Shape = 5 Then
    Complex = 28.65  '1/2 a radian for the star shape
Else
    Complex = 57.3  '1 radian for all other shapes
End If

If SetParams = False Then
    
    Sleep 750  'pause before the next shape display
        
    Select Case Shaper  'set the shape to display
    Case 1
        Shape = Int(Rnd * 4) + 2  'random
    Case 2
        Shape = 2  'H
    Case 3
        Shape = 3  'triangle
    Case 4
        Shape = 4  'square
    Case 5
        Shape = 5  'star
    End Select
    
    SetParams = True
End If

If PreviewMode = True Then  'set for the preview mode
    Size = 6
    DrawWidth = 1
End If

Call ColorSet

Angle = Angle + 0.2  'adjust the shape bend and twist angle changes
Redo ScaleWidth / 2, ScaleHeight / 2, 3, 0, 0  'draw the shapes

recount:
CntrS = CntrS + 1  'increment the display speed delay counter
    If CntrS = 5000 * Speed Then CntrS = 0 'delay done, start again
    If CntrS > 0 Then GoTo recount  'continue the speed delay count

CntrCS = CntrCS + 1 'increment the clear screen counter
    If CntrCS > 25 * ClrScrn Then 'check counter value
        CntrCS = 0  'zero the counter
        Cls  'clear the screen
        SetParams = False  'set for another shape
    End If

End Sub
