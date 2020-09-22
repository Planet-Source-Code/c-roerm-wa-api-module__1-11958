Attribute VB_Name = "Winamp"
'----------------------------------------------------------
'               [c-roerm] proudly presents....
'----------------------------------------------------------


'----------------------------------------------------------
'
'               THE DEFINITIVE WINAMP MODULE
'                           for
'                     Visual Basic 6.0
'
'----------------------------------------------------------


'----------------------------------------------------------
'                Made with lots of help from people at
'                      WWW.PLANET-SOURCE-CODE.COM
'                   the best VB-site on the net!
'----------------------------------------------------------

'----------------------------------------------------------
'Comments/feedback can be sent to c-roerm@online.no
'Also visit http://connect.to/moe (Norwegian)
'----------------------------------------------------------


Option Explicit
Option Base 1

' Find WinAmp Window.
' Finds the Window Handle for the window.
Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
               ByVal lpWindowName As Long) As Long

' PostMessage puts the message in a queue, and if the
' message could be put in the queue, then the function
' will return a non-zero value. If the queueing fails,
' zero will be returned. Then control will be returned to
' the calling code. The message itself will not be sent
' until Windows has gotten to the message in its queue.
' Meanwhile, the calling code may have proceeded to an
' unknown point. Useful for many WM_COMMAND messages.
Declare Function PostMessage Lib "user32" Alias _
"PostMessageA" (ByVal WndID As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

' SendMessage sends the message and will not return
' control to the calling code until the message has been
' sent and returned a value. Used for WM_USER messages.
Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal WndID As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

' SendMessage To Window (waits for reply)
' This is for sending a COPYDATASTRUCT structure.
Declare Function CopyDataSendMessage Lib "user32" Alias _
"SendMessageA" (ByVal WndID As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByRef lParam As COPYDATASTRUCT) As Long

' Find LPSTR address. Used when sending a
' COPYDATASTRUCT structure to find
' an address for the pointer to a string.
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
(ByVal lpString1 As String, ByVal lpString2 As String) _
As Long

'GetWindowText, used to e.g. find the current song title
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public hWndWinamp As Long
Public RetVal As Long

' Windows Const
Public Const WM_COMMAND = &H111
Public Const WM_COPYDATA = &H4A
Public Const WM_USER = &H400

' Winamp Constants
' You can add more constants and functions
' according to your needs.

'--------------------
'WM_COMMAND Messages:
'--------------------
Public Const waPlay                        As Long = 40045
Public Const waPause                       As Long = 40046
Public Const waStop                        As Long = 40047
Public Const waPreviousTrack               As Long = 40044
Public Const waNextTrack                   As Long = 40048
Public Const waFadeout_Stop                As Long = 40147
Public Const waStop_after_current_track    As Long = 40157
Public Const waFast_forward_5seconds       As Long = 40148
Public Const waFast_rewind_5seconds        As Long = 40144
Public Const waStart_of_playlist           As Long = 40154
Public Const waGoto_end_of_playlist        As Long = 40158
Public Const waOpen_file_dialog            As Long = 40029
Public Const waOpen_URL_dialog             As Long = 40155
Public Const waOpen_file_info_box          As Long = 40188
Public Const waTimedisplay2Elapsed         As Long = 40037
Public Const waTimedisplay2Remaining       As Long = 40038
Public Const waToggle_preferencesscreen    As Long = 40012
Public Const waOpenVisualizationOptions    As Long = 40190
Public Const waOpenPluginOptions           As Long = 40191
Public Const waExecuteCcurrentVisPlugIn    As Long = 40192
Public Const waToggleAboutBox              As Long = 40041
Public Const waToggle_Title_Autoscroll     As Long = 40189
Public Const waToggle_always_on_top        As Long = 40019
Public Const waToggle_Windowshade          As Long = 40064
Public Const waToggle_PL_Windowshade       As Long = 40266
Public Const waToggle_doublesize_mode      As Long = 40165
Public Const waToggle_EQ                   As Long = 40036
Public Const waToggle_playlist_editor      As Long = 40040
Public Const waToggle_main_window_visible  As Long = 40258
Public Const waToggle_minibrowser          As Long = 40298
Public Const waToggle_easymove             As Long = 40186
Public Const waRaise_volume_by_1_percent   As Long = 40058
Public Const waLower_volume_by_1_percent   As Long = 40059
Public Const waToggle_repeat               As Long = 40022
Public Const waToggle_shuffle              As Long = 40023
Public Const waOpen_jump_to_time_dialog    As Long = 40193
Public Const waOpen_jump_to_file_dialog    As Long = 40194
Public Const waOpen_skin_selector          As Long = 40219
Public Const waConfigure_cur_VIS_DLL       As Long = 40221
Public Const waReload_current_skin         As Long = 40291
Public Const waClose_Winamp                As Long = 40001

'--------------------
'WM_USER Messages
'--------------------
Public Const waClearPlaylist               As Long = 101
Public Const waGetStatus                   As Long = 104
Public Const waGetPosLen                   As Long = 105
Public Const waSetPos                      As Long = 106
Public Const waVolume                      As Long = 122
Public Const waAddFile                     As Long = 100

Public Const waOpenURL                     As Long = 241 ' Opens an new URL in the minibrowser. If the URL is NULL it will open the Minibrowser window
Public Const wabInetConnectionOpen         As Long = 242   'Returns 1 if the internet connecton is available for Winamp
Public Const waUpdateTitleInfo             As Long = 243 'Asks Winamp to update the information about the current title
Public Const waPLitem                      As Long = 245 'Sets the current playlist item
Public Const waGetCurURL                   As Long = 246 'Retrives the current Minibrowser URL into the buffer.
   '247 Flushes the playlist cache buffer
   '248 Blocks the Minibrowser from updates if value is set to 1
'   249 Works the same as 248 except that it will work even if 248 is set to 1
 '  250 Returns the status of the shuffle option (1 if set)
  ' 251 Returns the status of the repeat option (1 if set)
   '252 Sets the status of the suffle option (1 to turn it on)
   '253 Sets the status of the repeat option (1 to turn it on)




Public Const waGetVersion                  As Long = 0      'Retrieves the version of Winamp running. Version will be 0x20yx for 2.yx. This is a good way to determine if you did in fact find the right window, etc.
Public Const waStartPlayBack               As Long = 100    'Starts playback. A lot like hitting 'play' in Winamp, but not exactly the same
Public Const waSavePlayList                As Long = 120    'Writes out the current playlist to strPlayerDir\winamp.m3u, and returns the current position in the playlist.
Public Const waSelectPlSongNumber          As Long = 121    'Sets the playlist position to the position specified in tracks in 'data'.
Public Const waBalance                     As Long = 123    'Sets the panning to 'data', which can be between 0 (all left) and 255 (all right).
Public Const waGetPlNoOfSongs              As Long = 124    'Returns length of the current playlist, in tracks.
Public Const waGetPLCur                    As Long = 125    'Returns the position in the current playlist, in tracks (requires Winamp 2.05+).
Public Const waGetInfo                     As Long = 126    'Retrieves info about the current playing track. Returns samplerate (i.e. 44100) if 'data' is set to 0, bitrate if 'data' is set to 1, and number of channels if 'data' is set to 2. (requires Winamp 2.05+)
Public Const waGetSkinDir                  As Long = 260    'Retrieves the skin directory


' WinampPlayStatus Constants
Public Const waPlaying         As Long = 1
Public Const waPaused          As Long = 3
Public Const waStopped         As Long = 0
Public Const waPlayStatusError As Long = -1

' waclass is the class name for the Winamp window.
' Default is "Winamp v1.x", but can be modified with
' the /CLASS= switch. For example open Winamp with
' C:\path\to\winamp\winamp.exe /CLASS="myclassname"
' In this case, use "myclassname" for waclass when calling
' the functions.

Public Function GetPlayStatus(waClass As String) As Long
  ' Returns play status (playing, paused, stopped)
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'Can't find Winamp
      Exit Function
  End If
  
  GetPlayStatus = SendMessage(hWndWinamp, WM_USER, 0, _
                  waGetStatus)

End Function

Public Function TrackLength(waClass As String) As Long
  ' Returns length in seconds of playing track
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'Can't find Winamp
      Exit Function
  End If
      
  TrackLength = SendMessage(hWndWinamp, WM_USER, 1, _
                waGetPosLen)
  
End Function

Public Function GetPosition(waClass As String) As Long
  ' Gets position in milliseconds of playing track
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'Can't find Winamp
      Exit Function
  End If
      
  GetPosition = SendMessage(hWndWinamp, WM_USER, 0, _
             waGetPosLen)
  
End Function

Public Sub SetPosition(SetPos As Long, waClass As String)
  ' SetPos is desired position in milliseconds
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'Can't find Winamp
      Exit Sub
  End If
      
  RetVal = SendMessage(hWndWinamp, WM_USER, SetPos, _
            waSetPos)
End Sub

Public Sub ClearPlaylist(waClass As String)
  ' Clears playlist
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'Can't find Winamp
      Exit Sub
  End If
  
  RetVal = SendMessage(hWndWinamp, WM_USER, 0, waClearPlaylist)

End Sub

Public Sub CloseWinamp(waClass As String)
  ' Exit Winamp
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'Can't find Winamp
      Exit Sub
  End If
  
  RetVal = SendMessage(hWndWinamp, WM_COMMAND, waClose_Winamp, 0)
  
End Sub

Public Sub Volume(VolData As Long, waClass As String)
  ' Sets Volume (0-255)
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'can't find Winamp
      Exit Sub
  End If
  
  RetVal = SendMessage(hWndWinamp, WM_USER, VolData, waVolume)

End Sub

Public Sub AddFile(filename As String, waClass As String)
  ' Adds a file to the playlist
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'can't find Winamp
      Exit Sub
  End If
  
  Dim cds As COPYDATASTRUCT
  cds.dwData = waAddFile
  cds.lpData = lstrcpy(filename, filename)
  cds.cbData = Len(filename) + 1
  
  RetVal = CopyDataSendMessage(hWndWinamp, WM_COPYDATA, 0&, cds)

End Sub

' The following commands use by default PostMessage,
' but you can use SendMessage if you add a boolean
' parameter with the value False. They go via the
' private procedures PostIt or optionally SendIt.

Public Sub PlayTrack(waClass As String, _
                    Optional PostMess As Boolean = True)
  ' Plays a track
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'can't find Winamp
      Exit Sub
  End If
  
  If PostMess Then
    PostIt waPlay
  Else
    SendIt waPlay
  End If

End Sub



Public Sub PauseTrack(waClass As String, _
                    Optional PostMess As Boolean = True)
  ' Pauses track
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'can't find Winamp
      Exit Sub
  End If
  
  If PostMess Then
    PostIt waPause
  Else
    SendIt waPause
  End If

End Sub

Public Sub StopTrack(waClass As String, _
                    Optional PostMess As Boolean = True)
  ' Stops track
  hWndWinamp = FindWindow(waClass, 0)
  If hWndWinamp = 0 Then          'can't find Winamp
      Exit Sub
  End If
  
  If PostMess Then
    PostIt waStop
  Else
    SendIt waStop
  End If

End Sub

Public Function waGetCurrentSongTitle(waClass As String) As String

hWndWinamp = FindWindow(waClass, 0)
    If hWndWinamp = 0 Then
        Exit Function 'Can't find WA
    End If
    
    Dim llReturn As Long
    Dim lsTitle As String
    Dim lsBuffer As String
    lsTitle = ""
    lsBuffer = Space$(255)
    llReturn = GetWindowText(hWndWinamp, lsBuffer, 255)
    If llReturn Then
        lsTitle = Left$(lsBuffer, InStr(lsBuffer, Chr(0)) - 1)
    End If

    'This piece of code removes "Winamp" from the song title
    Dim i As Integer 'Counter
        For i = Len(lsTitle) To 1 Step -1
            If Mid(lsTitle, i, (Len("Winamp"))) = "Winamp" Then
                waGetCurrentSongTitle = Left(lsTitle, i - 3)
            End If
        Next i
    
End Function


Private Sub PostIt(Action As Long)
  ' Sends a message using PostMessage
  ' Checks to see if the message was successfully queued,
  ' otherwise it will be re-queued. There's a time-out
  ' so you won't get stuck forever.
  Dim n As Date
  
  n = Now
  Do
    RetVal = PostMessage(hWndWinamp, WM_COMMAND, Action, 0)
      Debug.Print "post", Now - n, 5 / 86400
    If RetVal = 0 Then
      If Now - n > 5 / 86400 Then
        MsgBox "Couldn't queue message for 5 seconds: " & Action
        Exit Do
      End If
    End If
  Loop Until RetVal <> 0
  
End Sub

Private Sub SendIt(Action As Long)
    'Sends a message using SendMessage
    RetVal = SendMessage(hWndWinamp, WM_COMMAND, Action, 0)
End Sub


