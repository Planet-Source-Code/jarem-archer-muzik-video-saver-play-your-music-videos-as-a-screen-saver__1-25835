VERSION 5.00
Begin VB.Form Menus 
   Caption         =   "Menus"
   ClientHeight    =   675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   675
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This form only is used for popup menus on borderless forms. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop up Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "&Previous"
      End
      Begin VB.Menu mnuRandom 
         Caption         =   "&Random"
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMute 
         Caption         =   "&Mute"
      End
      Begin VB.Menu b2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pa&use"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "&Start Over"
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSize 
         Caption         =   "Size"
         Begin VB.Menu mnuSizeFull 
            Caption         =   "Full Screen"
         End
         Begin VB.Menu mnuSizeOriginal 
            Caption         =   "Original Size"
         End
         Begin VB.Menu mnuSize2x 
            Caption         =   "Double Size"
         End
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "&Configure"
      End
      Begin VB.Menu b4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "&End"
      End
   End
End
Attribute VB_Name = "Menus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Hide
End Sub

Private Sub mnuEnd_Click()
    SaveSettings
    If PreviewMode Then Unload Main Else End
End Sub
Private Sub mnuConfig_Click()
    Unload Main
    Settings.Show
End Sub
Private Sub mnuNext_Click()
Main.PlayNext
End Sub
Private Sub mnuPrevious_Click()
Main.PlayPrev
End Sub
Private Sub mnuRandom_Click()
Main.PlayRnd
End Sub
Private Sub mnuMute_Click()
If mnuMute.Checked Then
    Main.Player.Mute = False
    mnuMute.Checked = False
    bMuteSound = False
Else
    Main.Player.Mute = True
    mnuMute.Checked = True
    bMuteSound = True
End If
End Sub
Private Sub mnuStart_Click()
    Main.Player.CurrentPosition = 0
End Sub
Private Sub mnuPause_Click()
If mnuPause.Caption = "Play" Then
    Main.Player.Play
    mnuPause.Caption = "Pause"
Else
    Main.Player.Pause
    mnuPause.Caption = "Play"
End If
End Sub
Private Sub mnuSizeOriginal_Click()
Main.Player.DisplaySize = mpDefaultSize
Menus.mnuSizeOriginal.Checked = True
Menus.mnuSize2x.Checked = False
Menus.mnuSizeFull.Checked = False
intVidSize = 0

End Sub
Private Sub mnuSize2x_Click()
Main.Player.DisplaySize = mpDoubleSize
Menus.mnuSizeOriginal.Checked = False
Menus.mnuSize2x.Checked = True
Menus.mnuSizeFull.Checked = False
intVidSize = 1
End Sub
Private Sub mnuSizeFull_Click()
Main.Player.DisplaySize = mpFullScreen
Menus.mnuSizeOriginal.Checked = False
Menus.mnuSize2x.Checked = False
Menus.mnuSizeFull.Checked = True
intVidSize = 2
End Sub
