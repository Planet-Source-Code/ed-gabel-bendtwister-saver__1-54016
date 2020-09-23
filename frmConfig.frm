VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   Caption         =   "BendTwister Saver Preferences"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   ControlBox      =   0   'False
   DrawWidth       =   2
   ForeColor       =   &H00000000&
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2535
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Graphic Shape Selection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3780
      TabIndex        =   14
      Top             =   1020
      Width           =   3500
      Begin VB.OptionButton opt_5 
         Caption         =   "Star"
         Height          =   195
         Left            =   2580
         TabIndex        =   19
         Top             =   390
         Width           =   600
      End
      Begin VB.OptionButton opt_4 
         Caption         =   "Square"
         Height          =   195
         Left            =   1480
         TabIndex        =   18
         Top             =   540
         Width           =   900
      End
      Begin VB.OptionButton opt_3 
         Caption         =   "Triangle"
         Height          =   195
         Left            =   1480
         TabIndex        =   17
         Top             =   270
         Width           =   900
      End
      Begin VB.OptionButton opt_2 
         Caption         =   "H"
         Height          =   195
         Left            =   330
         TabIndex        =   16
         Top             =   540
         Width           =   950
      End
      Begin VB.OptionButton opt_1 
         Caption         =   "Random"
         Height          =   195
         Left            =   330
         TabIndex        =   15
         Top             =   270
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Graphic Display Size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3780
      TabIndex        =   10
      Top             =   60
      Width           =   3500
      Begin MSComctlLib.Slider sldGrapSize 
         Height          =   285
         Left            =   150
         TabIndex        =   11
         Top             =   300
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   503
         _Version        =   393216
         Min             =   1
         Max             =   101
         SelStart        =   1
         TickFrequency   =   4
         Value           =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Large"
         Height          =   195
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Width           =   400
      End
      Begin VB.Label Label3 
         Caption         =   "Small"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   600
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clear Screen Delay"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   135
      TabIndex        =   6
      Top             =   1020
      Width           =   3500
      Begin MSComctlLib.Slider sldClrScrn 
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   503
         _Version        =   393216
         Min             =   1
         Max             =   41
         SelStart        =   1
         TickFrequency   =   2
         Value           =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Short"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Long"
         Height          =   195
         Left            =   3060
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Graphic Display Speed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   135
      TabIndex        =   2
      Top             =   60
      Width           =   3500
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Width           =   3220
         _ExtentX        =   5689
         _ExtentY        =   503
         _Version        =   393216
         Min             =   1
         Max             =   201
         SelStart        =   1
         TickFrequency   =   8
         Value           =   1
      End
      Begin VB.Label Label11 
         Caption         =   "Fast"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label12 
         Caption         =   "Slow"
         Height          =   195
         Left            =   3060
         TabIndex        =   3
         Top             =   630
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   350
      Left            =   4150
      TabIndex        =   1
      Top             =   2070
      Width           =   2805
   End
   Begin VB.CommandButton cmdSetPref 
      Caption         =   "Set Preferences"
      Height          =   350
      Left            =   480
      TabIndex        =   0
      Top             =   2070
      Width           =   2805
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Call LoadPrefs  'load current preferences from registry
    
    If Shaper = 1 Then opt_1.Value = True  'random shapes
    If Shaper = 2 Then opt_2.Value = True  'H shapes
    If Shaper = 3 Then opt_3.Value = True  'triangle shapes
    If Shaper = 4 Then opt_4.Value = True  'square shapes
    If Shaper = 5 Then opt_5.Value = True  'star shapes
        
    sldSpeed.Value = Speed  'restore the display speed value
    sldClrScrn.Value = ClrScrn 'restore the clear screen delay value
    sldGrapSize.Value = Size 'restore the size value

End Sub

Private Sub opt_1_Click()
    If opt_1.Value = True Then Shaper = 1  'random shapes
End Sub

Private Sub opt_2_Click()
    If opt_2.Value = True Then Shaper = 2  'H shapes
End Sub

Private Sub opt_3_Click()
    If opt_3.Value = True Then Shaper = 3  'triangle shapes
End Sub

Private Sub opt_4_Click()
    If opt_4.Value = True Then Shaper = 4  'square shapes
End Sub

Private Sub opt_5_Click()
    If opt_5.Value = True Then Shaper = 5  'star shapes
End Sub

Private Sub sldSpeed_Change()
    Speed = sldSpeed.Value  'change the display speed
End Sub

Private Sub sldClrScrn_Change()
    ClrScrn = sldClrScrn.Value  'change the clear screen delay
End Sub

Private Sub sldGrapSize_Change()
    Size = sldGrapSize.Value  'change the size value
End Sub

Private Sub cmdSetPref_Click()
    Call SavePrefs  'save preferences
    Unload Me  'unload this form
End Sub

Private Sub cmdExit_Click()
    Dim Msg, Style, Title, Response
    Msg = "Are you sure you want to exit without saving changes?"
    Style = vbYesNo + vbDefaultButton2   'define message box buttons
    Title = "Verify Exit"
    Response = MsgBox(Msg, Style, Title)
    If Response = vbNo Then GoTo Done
    Unload Me  'exit without changes
Done:
End Sub
