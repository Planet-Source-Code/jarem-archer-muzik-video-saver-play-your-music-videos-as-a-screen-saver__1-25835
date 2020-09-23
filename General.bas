Attribute VB_Name = "General"
Option Explicit
'Playlist Settings
Public bRandomize As Integer 'Randomize Playlist
Public bLoopVid As Integer 'Loop Same Video

'Control
Public bPopUpMenu As Integer 'Pop up Menu
Public bControlKeys As Integer 'Control Keys


'Video Settings
Public intVidSize As Integer 'Video Size
Public bMuteSound As Integer 'Mute Sound
Public bDeskBack As Integer 'Use Desktop Background
Public bEndOnMove As Integer 'End when mouse moves
Public bPauseOnClick As Integer 'Pause when clicked
Public bEndOnClick As Integer 'Pause when clicked



Public strPlaylist As String 'The Playlist

'Misc
Public strLastPlayed As String

Public PreviewMode As Boolean 'to preview in config (not saved)





Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&


Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1 ' Unicode nul terminated String
    Public Const REG_DWORD = 4 ' 32-bit number



Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE



Sub SaveSettings() 'Save Settings
Dim lngIndex As Integer
Dim fnum As Integer
On Error GoTo Nofile

fnum = FreeFile
Open App.Path & "\Muzik.dat" For Output As fnum
    
    Write #fnum, bRandomize
    Write #fnum, bLoopVid
    Write #fnum, bPopUpMenu
    Write #fnum, bControlKeys
    Write #fnum, intVidSize
    Write #fnum, bMuteSound
    Write #fnum, bDeskBack
    Write #fnum, bEndOnMove
    Write #fnum, bPauseOnClick
    Write #fnum, bEndOnClick
    Write #fnum, strPlaylist
    Write #fnum, strLastPlayed

Close fnum
Exit Sub
Nofile:
Close fnum
End Sub
Sub LoadSettings()
Dim txt As String
Dim fnum, lngIndex As Integer
fnum = FreeFile
On Error GoTo Nofile
Open App.Path & "\Muzik.dat" For Input As fnum
 lngIndex = 1
 
    Input #fnum, bRandomize
    Input #fnum, bLoopVid
    Input #fnum, bPopUpMenu
    Input #fnum, bControlKeys
    Input #fnum, intVidSize
    Input #fnum, bMuteSound
    Input #fnum, bDeskBack
    Input #fnum, bEndOnMove
    Input #fnum, bPauseOnClick
    Input #fnum, bEndOnClick
    Input #fnum, strPlaylist
    Input #fnum, strLastPlayed
    
Close fnum
Exit Sub
Nofile:
Close fnum
End Sub
Function GetFileName(strPath As String) As String
    If Trim(strPath) = "" Then Exit Function
    Dim temp
    temp = Split(strPath, "\")
    If temp(UBound(temp)) <> "" Then
        GetFileName = temp(UBound(temp))
    Else
        GetFileName = temp(UBound(temp)) - 1 'Path ends with a "\"
    End If
End Function
Public Sub Ontop(FormName As Form)
'Make a form always ontop of other windows
On Error GoTo error
Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
Public Sub NotOntop(FormName As Form)
'Make a form not always ontop of other windows
On Error GoTo error
Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
Public Function Duration(TotalSeconds As Long, UpFormat As Integer) As String
    
    ' Format = 0, 1, 2
    ' This determines the format of the time
    '     to be returned
    ' Type 0: 1d 4h 15m 47s
    ' Type 1: 1 day, 4:15:47
    ' Type 2: 1 day 4hrs 15mins 47secs
    ' Type else: Defaults to type 0
    
    Dim Seconds
    Dim Minutes
    Dim Hours
    Dim Days
    Dim DayString As String
    Dim HourString As String
    Dim MinuteString As String
    Dim SecondString As String
    
    Seconds = Int(TotalSeconds Mod 60)
    Minutes = Int(TotalSeconds \ 60 Mod 60)
    Hours = Int(TotalSeconds \ 3600 Mod 24)
    Days = Int(TotalSeconds \ 3600 \ 24)


    Select Case UpFormat
        Case 0
        DayString = "d "
        HourString = "h "
        MinuteString = "m "
        SecondString = "s"
        Case 1
        If Days = 1 Then DayString = " day, " _
    Else: DayString = " days, "
        HourString = ":"
        MinuteString = ":"
        SecondString = ""
        Case 2
        If Days = 1 Then DayString = " day " _
    Else: DayString = " days, "
        If Hours = 1 Then HourString = "hr " _
    Else: HourString = "hrs "
        If Minutes = 1 Then MinuteString = "min " _
    Else: MinuteString = "mins "
        If Seconds = 1 Then SecondString = "sec " _
    Else: SecondString = "secs"
        Case Else
        DayString = "d "
        HourString = "h "
        MinuteString = "m "
        SecondString = "s"
    End Select



Select Case Days
    Case 0
    Duration = Format(Hours, "0") & HourString & Format(Minutes, "00") & _
    MinuteString & Format(Seconds, "00") & SecondString
    Case Else
    Duration = Days & DayString & Format(Hours, "0") & HourString _
    & Format(Minutes, "00") & MinuteString & _
    Format(Seconds, "00") & SecondString
    End Select

End Function

Public Function OpenFile(hwnd As Long, ByVal file As String)
    
    Dim lRet As Long
    Const SW_SHOWNORMAL = 1
    
    lRet = ShellExecute(hwnd, vbNullString, file, vbNullString, App.Path, SW_SHOWNORMAL)
    'lRet now contains the hWnd of the docum
    '     ent
    'you just opened
End Function

Public Function getstring(Hkey As Long, strPath As String, strValue As String)
    'BY Kevin Mackey
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim r As Long
    Dim lValueType As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                getstring = Left$(strBuf, intZeroPos - 1)
            Else
                getstring = strBuf
            End If
        End If
    End If
End Function

