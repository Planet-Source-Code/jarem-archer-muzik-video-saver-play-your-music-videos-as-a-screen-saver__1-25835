VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Main 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Main"
   ClientHeight    =   6870
   ClientLeft      =   420
   ClientTop       =   0
   ClientWidth     =   8445
   DrawMode        =   16  'Merge Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrDoEvents 
      Interval        =   1
      Left            =   5340
      Top             =   540
   End
   Begin VB.ListBox VidList 
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I'm sorry but there are no videos present in playlist. To add videos, right click and click ""Configure""."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   3000
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   6735
   End
   Begin MediaPlayerCtl.MediaPlayer Player 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   6060
      Width           =   1245
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   3
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   -1  'True
      SendMouseClickEvents=   -1  'True
      SendMouseMoveEvents=   -1  'True
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
      Select Case KeyAscii
    Case vbKeySpace
        If Player.PlayState = mpPlaying Then Player.Pause Else Player.Play
    Case vbKeyZ, 122
        PlayPrev
    Case vbKeyX, 120
        PlayNext
    Case vbKeyC, 99
        Player.CurrentPosition = 0
    Case vbKeyEscape
        End
    'Case vbkey
    '    End
End Select

End Sub

Private Sub Form_Load()
If App.PrevInstance Then End

'For Windows Screen Saver
   If Left(Command$, 2) = "/c" And Not PreviewMode Then 'Not previewmode so it wont unload by settings
     Settings.Show   ' display configuration form
    Unload Me
     Exit Sub
   ElseIf Left(Command$, 2) = "/p" Then
        End
   End If

Main.WindowState = vbMaximized


Player.Top = 0
Player.Left = 0



If Not PreviewMode Then LoadSettings

On Error Resume Next
If bDeskBack Then Me.Picture = LoadPicture(getstring(HKEY_USERS, ".default\control panel\desktop", "wallpaper"))


IntPlay

End Sub
Sub PlayRnd()
TryAgain:
    Randomize
    VidList.ListIndex = Int((VidList.ListCount - 1 + 1) * Rnd + 0)
If VidList.List(VidList.ListIndex) = strLastPlayed And strLastPlayed <> "" And VidList.ListCount > 1 Then GoTo TryAgain
    Player.FileName = VidList.List(VidList.ListIndex)
    strLastPlayed = Player.FileName
    Player.Play
End Sub
Sub PlayNext()
Skip:
If VidList.ListIndex = VidList.ListCount - 1 Then
    VidList.ListIndex = 0
Else
    VidList.ListIndex = VidList.ListIndex + 1
End If
If VidList.List(VidList.ListIndex) = strLastPlayed And strLastPlayed <> "" And VidList.ListCount > 1 Then GoTo Skip
Player.FileName = VidList.List(VidList.ListIndex)
Player.Play
strLastPlayed = VidList.List(VidList.ListIndex)
End Sub
Sub PlayPrev()
Skip:
If VidList.ListIndex = 0 Then
    VidList.ListIndex = VidList.ListCount - 1
Else
    VidList.ListIndex = VidList.ListIndex - 1
End If
If VidList.List(VidList.ListIndex) = strLastPlayed And strLastPlayed <> "" And VidList.ListCount > 1 Then GoTo Skip
Player.FileName = VidList.List(VidList.ListIndex)
Player.Play
strLastPlayed = VidList.List(VidList.ListIndex)
End Sub
Sub IntPlay()
If strPlaylist = "" Then GoTo None
Dim strFile As String, temp, Temp2, lngIndex As Integer

'Add all the files in the playlist
temp = Split(strPlaylist, "?")
For lngIndex = 0 To UBound(temp) - 1
    Temp2 = Split(temp(lngIndex), "*")
    VidList.AddItem Temp2(0)
Next lngIndex

'Set the Video Size
Select Case intVidSize
    Case 0
        Player.DisplaySize = mpDefaultSize
        Menus.mnuSizeOriginal.Checked = True
    Case 1
        Player.DisplaySize = mpDoubleSize
        Menus.mnuSize2x.Checked = True
    Case 2
        Player.DisplaySize = mpFullScreen
        Menus.mnuSizeFull.Checked = True
End Select

'Check for Mute
If bMuteSound Then
    Player.Mute = True
    Menus.mnuMute.Checked = True
Else
    Player.Mute = False
    Menus.mnuMute.Checked = False
End If



'Now play!
If bRandomize Then
    PlayRnd
Else
    PlayNext
End If


Exit Sub
None:
Player.Visible = False
Label1.Visible = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If bPopUpMenu And Button = vbRightButton Then Main.PopupMenu Menus.mnuPop

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If bPopUpMenu And Button = vbRightButton Then Main.PopupMenu Menus.mnuPop

End Sub

Private Sub Player_Click(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
  If bPopUpMenu And Button = vbRightButton Then Main.PopupMenu Menus.mnuPop

End Sub

Private Sub Player_DblClick(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
    If bEndOnClick Then
       End
    Else
        If Player.PlayState = mpPlaying Then
            Player.Pause
            Menus.mnuPause.Caption = "Play"
        Else
            Player.Play
            Menus.mnuPause.Caption = "Pause"
        End If
    End If
End Sub

Private Sub Player_KeyPress(CharacterCode As Integer)

On Error Resume Next
      Select Case CharacterCode
    Case vbKeySpace
        If Player.PlayState = mpPlaying Then Player.Pause Else Player.Play
    Case vbKeyZ, 122
        PlayPrev
    Case vbKeyX, 120
        PlayNext
    Case vbKeyC, 99
        Player.CurrentPosition = 0
    Case vbKeyEscape
        End
    'Case vbkey
    '    End
End Select
Main.SetFocus
End Sub

Private Sub Player_MouseMove(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
'must fix
'If bEndOnMove Then End
End Sub

Private Sub Player_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
If Player.PlayState = mpStopped Then
    If bRandomize Then
        PlayRnd
    ElseIf bvidloop Then
        Player.Play
    Else
        PlayNext
End If
End If
End Sub

Private Sub tmrDoEvents_Timer()
Player.Height = Main.Height
Player.Width = Main.Width

tmrDoEvents.Enabled = False
End Sub
