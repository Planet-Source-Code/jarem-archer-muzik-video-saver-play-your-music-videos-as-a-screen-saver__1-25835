VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Settings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Muzik Video Saver"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   675
      Left            =   0
      Picture         =   "settings.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   6315
      TabIndex        =   29
      Top             =   0
      Width           =   6375
   End
   Begin VB.Timer tmrPosition 
      Interval        =   100
      Left            =   5820
      Top             =   2460
   End
   Begin MSComctlLib.Slider Slider 
      Height          =   375
      Left            =   4800
      TabIndex        =   28
      Top             =   2460
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      _Version        =   393216
      TickStyle       =   3
   End
   Begin MSComDlg.CommonDialog cdbFile 
      Left            =   5760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Add Video"
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   0
      TabIndex        =   0
      Top             =   2775
      Width           =   6345
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control"
      Height          =   1230
      Left            =   15
      TabIndex        =   16
      Top             =   3660
      Width           =   2190
      Begin VB.CommandButton cmdEdit 
         Caption         =   "What Keys?"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   795
         Width           =   1920
      End
      Begin VB.CheckBox chkControlKeys 
         Caption         =   "Enable Control Keys"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1950
      End
      Begin VB.CheckBox chkPopUpMenu 
         Caption         =   "Enable Pop-Menu"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1920
      End
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Preview"
      Height          =   495
      Left            =   4980
      TabIndex        =   10
      Top             =   2940
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4980
      TabIndex        =   9
      Top             =   4380
      Width           =   1365
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   4980
      TabIndex        =   8
      Top             =   3780
      Width           =   1365
   End
   Begin VB.Frame Frame4 
      Caption         =   "Video Options"
      Height          =   2010
      Left            =   2235
      TabIndex        =   6
      Top             =   2880
      Width           =   2685
      Begin VB.CheckBox chkDeskBack 
         Caption         =   "Use Desktop Background"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkEndOnMove 
         Caption         =   "End when Mouse Moves"
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   1260
         Width           =   2100
      End
      Begin VB.OptionButton optEndOnClick 
         Caption         =   "End Saver"
         Height          =   225
         Left            =   1560
         TabIndex        =   21
         Top             =   1740
         Width           =   1065
      End
      Begin VB.OptionButton optPauseOnClick 
         Caption         =   "Pause Video"
         Height          =   225
         Left            =   90
         TabIndex        =   20
         Top             =   1740
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CheckBox chkMuteSound 
         Caption         =   "Mute Sound"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1365
      End
      Begin VB.ComboBox cmbVidSize 
         Height          =   315
         ItemData        =   "settings.frx":4EAB
         Left            =   600
         List            =   "settings.frx":4EB8
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1620
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   2580
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "When Double Clicked:"
         Height          =   210
         Left            =   105
         TabIndex        =   22
         Top             =   1500
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Size:"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Playlist"
      Height          =   750
      Left            =   15
      TabIndex        =   5
      Top             =   2880
      Width           =   2190
      Begin VB.CheckBox chkLoopVid 
         Caption         =   "Loop Same Video"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1725
      End
      Begin VB.CheckBox chkRandomize 
         Caption         =   "Randomize"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   285
      Left            =   2460
      TabIndex        =   4
      Top             =   2475
      Width           =   1230
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   285
      Left            =   1140
      TabIndex        =   3
      Top             =   2475
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   2475
      Width           =   975
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   1785
      Left            =   15
      TabIndex        =   1
      Top             =   660
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   3149
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   1412
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   0
      TabIndex        =   24
      Top             =   4920
      Width           =   6315
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click here to Vote!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2820
         MousePointer    =   10  'Up Arrow
         TabIndex        =   26
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "You Like?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Line Line2 
      X1              =   4980
      X2              =   6240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   2460
      Y2              =   2760
   End
   Begin MediaPlayerCtl.MediaPlayer Player 
      Height          =   2160
      Left            =   3720
      TabIndex        =   7
      Top             =   660
      Width           =   2685
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
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
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
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
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   -1  'True
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command7_Click()
EditKeys.Show
End Sub

Private Sub chkEndOnMove_Click()
If chkEndOnMove.Value = 1 Then
    Label2.Enabled = False
    optPauseOnClick.Enabled = False
    optEndOnClick.Enabled = False
    chkPopUpMenu.Enabled = False
Else
    Label2.Enabled = True
    chkPopUpMenu.Enabled = True
    optPauseOnClick.Enabled = True
    optEndOnClick.Enabled = True
End If
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
cdbFile.Filter = "Video Files (*.mpg, *.mpeg, *.avi, *.asf)|*.mpg;*.mpeg;*.avi;*.asf"
cdbFile.ShowOpen
If cdbFile.FileName = "" Then Exit Sub

Player.FileName = cdbFile.FileName

If Player.Duration = 0 Then 'if the duration is 0, then it cant play
    MsgBox "Error playing file. Cant Add.", vbCritical, "Error"
    Exit Sub
End If

'set the slider
Slider.Max = Player.Duration
Slider.Value = 0

Player.Play

lstFiles.ListItems.Add , , GetFileName(cdbFile.FileName)
lstFiles.ListItems.Item(lstFiles.ListItems.Count).Tag = cdbFile.FileName
lstFiles.ListItems.Item(lstFiles.ListItems.Count).SubItems(1) = Right(Duration(Player.Duration, 1), 5)


End Sub

Private Sub Command5_Click()


End Sub

Private Sub cmdCancel_Click()

End
End Sub

Private Sub cmdClear_Click()
lstFiles.ListItems.Clear
End Sub

Private Sub cmdEdit_Click()
MsgBox "Control Keys: " & vbCrLf & vbCrLf _
     & "  Spacebar     - Pause" & vbCrLf _
     & "  Z                  - Previous Video" & vbCrLf _
     & "  X                  - Next Video" & vbCrLf _
     & "  C                  - Start Video Over" & vbCrLf & vbCrLf _
     & "  ESC              - End Saver"
     
     
End Sub

Private Sub cmdOK_Click()
PreviewMode = False
    bRandomize = chkRandomize.Value
    bLoopVid = chkLoopVid.Value
    bPopUpMenu = chkPopUpMenu.Value
    bControlKeys = chkControlKeys.Value
    intVidSize = cmbVidSize.ListIndex
    bMuteSound = chkMuteSound.Value
    bDeskBack = chkDeskBack.Value
    bEndOnMove = chkEndOnMove.Value
    bPauseOnClick = optPauseOnClick.Value
    bEndOnClick = optEndOnClick.Value
    
    strPlaylist = ""
    Dim lngIndex
    For lngIndex = 1 To lstFiles.ListItems.Count
        strPlaylist = strPlaylist & lstFiles.ListItems(lngIndex).Tag & "*" & lstFiles.ListItems(lngIndex).SubItems(1) & "?"
    Next lngIndex
SaveSettings
End
End Sub

Private Sub cmdPrev_Click()
PreviewMode = True
    bRandomize = chkRandomize.Value
    bLoopVid = chkLoopVid.Value
    bPopUpMenu = chkPopUpMenu.Value
    bControlKeys = chkControlKeys.Value
    intVidSize = cmbVidSize.ListIndex
    bMuteSound = chkMuteSound.Value
    bDeskBack = chkDeskBack.Value
    bEndOnMove = chkEndOnMove.Value
    bPauseOnClick = optPauseOnClick.Value
    bEndOnClick = optEndOnClick.Value
    
    strPlaylist = ""
    Dim lngIndex
    For lngIndex = 1 To lstFiles.ListItems.Count
        strPlaylist = strPlaylist & lstFiles.ListItems(lngIndex).Tag & "*" & lstFiles.ListItems(lngIndex).SubItems(1) & "?"
    Next lngIndex

If Player.PlayState = mpPlaying Then Player.Stop
Main.Show
End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
lstFiles.ListItems.Remove lstFiles.SelectedItem.Index
End Sub

Private Sub Form_Load()
On Error Resume Next
LoadSettings
'Sets Controls to Settings
    chkRandomize.Value = bRandomize
    chkLoopVid.Value = bLoopVid
    chkPopUpMenu.Value = bPopUpMenu
    chkControlKeys.Value = bControlKeys
    cmbVidSize.ListIndex = intVidSize
    chkMuteSound.Value = bMuteSound
    chkDeskBack.Value = bDeskBack
    chkEndOnMove.Value = bEndOnMove
    optPauseOnClick.Value = bPauseOnClick
    optEndOnClick.Value = bEndOnClick

If Trim(strPlaylist) <> "" Then
    Dim lngIndex As Integer, temp, temp2
    temp = Split(strPlaylist, "?")
    For lngIndex = 0 To UBound(temp) - 1
      temp2 = Split(temp(lngIndex), "*")
      lstFiles.ListItems.Add , , GetFileName((temp2(0))) 'Short Filename
      lstFiles.ListItems.Item(lstFiles.ListItems.Count).Tag = (temp2(0)) 'Long Filename
      lstFiles.ListItems.Item(lstFiles.ListItems.Count).SubItems(1) = (temp2(1)) 'Duration
    Next lngIndex
End If

'Play the first file
If lstFiles.ListItems.Count = 0 Then Exit Sub
Player.FileName = lstFiles.ListItems.Item(1).Tag
DoEvents
Slider.Max = Player.Duration
Slider.Value = 0
Player.Play
End Sub


Private Sub Label4_Click()

OpenFile Me.hwnd, "https://www.planet-source-code.com/ads/authentication/login.asp?lngWId=1&blnOutsideOfVBSubWeb=False&txtReturnURL=http%3A%2F%2Fwww%2Epscode%2Ecom%2Fvb%2Fscripts%2Fvoting%2FVoteOnCodeRating%2Easp%3FlngWId%3D1%26optCodeRatingValue=5%26cmdRateIt=Rate%A0It%21%26txtCodeId=25835%26txtCodeName=AA%26intUserRatingTotal=5%26intNumOfUserRatings=1&txtCancelURL=%2Fvb%2Fdefault%2Easp%3FlngWId%3D1"
End Sub

Private Sub lstFiles_DblClick()
On Error Resume Next
Player.FileName = lstFiles.SelectedItem.Tag
Slider.Max = Player.Duration
Slider.Value = 0
Player.Play
End Sub

Private Sub lstFiles_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyAdd Then

  '  If lstFiles.SelectedItem.Index = 0 Or lstFiles.SelectedItem.Index = lstFiles.ListItems.Count Then Exit Sub
  '  Dim Temp(0, 1) 'the Backup
  '  Temp(0, 0) = lstFiles.ListItems(lstFiles.SelectedItem.Index + 1).Text
  '  Temp(0, 1) = lstFiles.ListItems(lstFiles.SelectedItem.Index + 1).SubItems(1)
    
    'make the one above, the one below
   ' lstFiles.ListItems(lstFiles.SelectedItem.Index + 1).Text = lstFiles.SelectedItem.Text
   ' lstFiles.ListItems(lstFiles.SelectedItem.Index + 1).SubItems(1) = lstFiles.SelectedItem.SubItems(1)
    
   ' 'Make the old the temp
   ' lstFiles.SelectedItem.Text = Temp(0, 0)
   ' lstFiles.SelectedItem.SubItems(1) = Temp(0, 1)
    
    
ElseIfKeyAscii = vbKeyPageDown
    If lstFiles.SelectedItem.Index = 0 Or lstFiles.SelectedItem.Index = lstFiles.ListItems.Count Then Exit Sub
    
End If

End Sub

Private Sub Slider_Scroll()
Player.CurrentPosition = Slider.Value
End Sub

Private Sub tmrPosition_Timer()
Slider.Value = Player.CurrentPosition
End Sub
