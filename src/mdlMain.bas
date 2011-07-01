Attribute VB_Name = "mdlMain"
Option Explicit

Const AppName = "SE"

'Win32 API's
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'BASS External Constants
Public BASSThread As Long
Public SEChannel As Long
Public SEext As String
Public streamURL As String
Public SEStreamName As String, tmpSEStreamName As String
Public TmpNameHold As String
Public TmpNameHold2 As String
Public WriteFile As clsFileIo
Public SEextn As String
Public SEchnfo As BASS_CHANNELINFO
Public FileIsOpen As Boolean, GotHeader As Boolean
Public DownloadStarted As Boolean, DoDownload As Boolean
Public DlOutput As String, SongNameUpdate As Boolean

Public Function StatusMsg(xMsg As String, Optional CRLF As Boolean = True)
    frmMain.txtStatus.text = frmMain.txtStatus.text & xMsg
    If CRLF = True Then
        frmMain.txtStatus.text = frmMain.txtStatus.text & vbCrLf & vbCrLf
    End If
    frmMain.txtStatus.SelStart = Len(frmMain.txtStatus)
End Function

Public Sub launchCC()
    If (MsgBox("This will launch the default browser to URL of Creative Commons Attribution License 2.5. Continue ?", vbOKCancel) = vbOK) Then
        Call ShellExecute(frmMain.hwnd, "Open", "http://creativecommons.org/licenses/by/2.5/", "", "", 99)
    End If
End Sub

Public Sub loadPlugins()
    Dim lngPlugin As Long, strPluginName As String
    
    'BASS AAC Plugin
    strPluginName = "bass_aac.dll"
    lngPlugin = BASS_PluginLoad(strPluginName, Len(strPluginName))
    If lngPlugin Then
        StatusMsg "BASS AAC core (" & strPluginName & ") loaded"
    End If

    'BASS WMA Plugin
    strPluginName = "basswma.dll"
    lngPlugin = BASS_PluginLoad(strPluginName, Len(strPluginName))
    If lngPlugin Then
        StatusMsg "BASS WMA core (" & strPluginName & ") loaded"
    End If
End Sub

Sub playStation(Miaw As Long)
        Call BASS_StreamFree(SEChannel) 'Close Old Stream
        SEChannel = BASS_StreamCreateURL(CStr(streamURL), 0, BASS_STREAM_STATUS, AddressOf SEDownloadProc, 0)
        If SEChannel = 0 Then
            StatusMsg "Can't play the stream"
        Else
            Do
                Dim progress As Long, len_ As Long
                len_ = BASS_StreamGetFilePosition(SEChannel, BASS_FILEPOS_END)
                If (len_ = -1) Then GoTo done ' something's gone wrong! (eg. BASS_Free called)
                progress = (BASS_StreamGetFilePosition(SEChannel, BASS_FILEPOS_DOWNLOAD) _
                    - BASS_StreamGetFilePosition(SEChannel, BASS_FILEPOS_CURRENT)) * 100 / len_ ' percentage of buffer filled
                If (progress > 75) Then Exit Do ' over 75% full, enough
                Call Sleep(50)
            Loop While 1

            Dim icyPTR As Long  ' a pointer where ICY info is stored
            Dim tmpICY As String

            'Get the broadcast name and bitrate
            icyPTR = BASS_ChannelGetTags(SEChannel, BASS_TAG_ICY)

            If (icyPTR) Then
                Do
                    tmpICY = VBStrFromAnsiPtr(icyPTR)
                    icyPTR = icyPTR + Len(tmpICY) + 1
                    SEStreamName = IIf(Mid(tmpICY, 1, 9) = "icy-name:", Mid(tmpICY, 10), SEStreamName)
                    'Call UpdateStation
                Loop While (tmpICY <> "")
            End If

            Call BASS_ChannelGetInfo(SEChannel, SEchnfo)
            Select Case Val(SEchnfo.ctype)
                Case 68352:
                    SEextn = "aac"
                Case 68352:
                    SEextn = "mp4"
                Case 66304:
                    SEextn = "wma"
                Case 66305:
                    SEextn = "wma"
                Case 65541:
                    SEextn = "mp3"
                Case 65540:
                    SEextn = "mp2"
                Case 65539:
                    SEextn = "mp1"
                Case 65538:
                    SEextn = "ogg"
                Case 262144:
                    SEextn = "wav"
                Case Else
                    SEextn = "mp3" 'assume its mp3; very rarely happens;
            End Select
            
            'Get the stream title and set sync for subsequent titles
            Call DoMeta(BASS_ChannelGetTags(SEChannel, BASS_TAG_META))
            
            'Sync the Tag
            Call BASS_ChannelSetSync(SEChannel, BASS_SYNC_META, 0, AddressOf MetaSync, 0)

            'Play
            Call BASS_ChannelPlay(SEChannel, BASSFALSE)
        End If
done:
    Call CloseHandle(BASSThread) 'Close the thread
    BASSThread = 0
End Sub

Sub MetaSync(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
    Call DoMeta(data)
End Sub

Sub DoMeta(ByVal meta As Long)
    Dim p As String, tmpMeta As String
    If meta = 0 Then Exit Sub
    tmpMeta = VBStrFromAnsiPtr(meta)
    Debug.Print tmpMeta
    If ((Mid(tmpMeta, 1, 13) = "StreamTitle='")) Then
        p = Mid(tmpMeta, 14)
        TmpNameHold = Mid(p, 1, InStr(p, ";") - 2)
        If TmpNameHold <> TmpNameHold2 Then
            TmpNameHold2 = TmpNameHold
            GotHeader = False
            DownloadStarted = False
        End If
        DlOutput = App.Path & "\" & RemoveSpecialChar(Mid(p, 1, InStr(p, ";") - 2)) & "." & SEextn
        Debug.Print DlOutput
        If DoDownload = True Then
            StatusMsg "saving as: " & RemoveSpecialChar(Mid(p, 1, InStr(p, ";") - 2)) & "." & SEextn
        Else
            StatusMsg "playing: " & RemoveSpecialChar(Mid(p, 1, InStr(p, ";") - 2))
        End If
    End If
End Sub

Sub UpdateStation()
    If SEStreamName <> tmpSEStreamName Then
        tmpSEStreamName = SEStreamName
        StatusMsg "now playing: " & SEStreamName
    End If
End Sub

Public Sub SEDownloadProc(ByVal buffer As Long, ByVal length As Long, ByVal user As Long)
    If (buffer And length = 0) Then
        'frmNetRadio.lblBPS.Caption = VBStrFromAnsiPtr(buffer) ' display connection status
        Exit Sub
    End If

    If (Not DoDownload) Then
        DownloadStarted = False
        Call WriteFile.CloseFile
        Exit Sub
    End If

    If (Trim(DlOutput) = "") Then Exit Sub

    If (Not DownloadStarted) Then
        DownloadStarted = True
        Call WriteFile.CloseFile
        If (WriteFile.OpenFile(DlOutput)) Then
            SongNameUpdate = False
        Else
            SongNameUpdate = True
            GotHeader = False
        End If
    End If

    If (Not SongNameUpdate) Then
        If (length) Then
            Call WriteFile.WriteBytes(buffer, length)
        Else
            Call WriteFile.CloseFile
            GotHeader = False
        End If
    Else
        DownloadStarted = False
        Call WriteFile.CloseFile
        GotHeader = False
    End If
End Sub

Public Function RemoveSpecialChar(strFileName As String)
    Dim i As Byte
    Dim SpecialChar As Boolean
    Dim SelChar As String, OutFileName As String

    For i = 1 To Len(strFileName)
        SelChar = Mid(strFileName, i, 1)
        SpecialChar = InStr(":/\?*|<>" & Chr$(34), SelChar) > 0

        If (Not SpecialChar) Then
            OutFileName = OutFileName & SelChar
            SpecialChar = False
        Else
            OutFileName = OutFileName
            SpecialChar = False
        End If
    Next i

    RemoveSpecialChar = OutFileName
End Function

Public Sub loadList()
Dim tmpCount, i As Integer
tmpCount = GetSetting(AppName, "playlist", "count")
If Val(Trim(tmpCount)) < 1 Then Exit Sub
For i = 0 To tmpCount - 1
    frmMain.lstStationName.AddItem GetSetting(AppName, "playlist", "stationname_" & i)
    frmMain.lstStations.AddItem GetSetting(AppName, "playlist", "stationurl_" & i)
Next
StatusMsg "playlist loaded successfully"
End Sub

Public Sub saveList()
Dim i As Integer
SaveSetting AppName, "playlist", "count", frmMain.lstStations.ListCount
For i = 0 To frmMain.lstStations.ListCount
    SaveSetting AppName, "playlist", "stationname_" & i, frmMain.lstStationName.List(i)
    SaveSetting AppName, "playlist", "stationurl_" & i, frmMain.lstStations.List(i)
Next i
StatusMsg "playlist saved successfully"
End Sub
