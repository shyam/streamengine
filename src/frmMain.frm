VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stream Engine"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top Playing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   8775
   End
   Begin VB.ListBox lstStationName 
      Height          =   1425
      ItemData        =   "frmMain.frx":6852
      Left            =   120
      List            =   "frmMain.frx":6854
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play Station"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox chkRecording 
      BackColor       =   &H00FFFFFF&
      Caption         =   "R&ecord Mode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Station"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Station"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox lstStations 
      Height          =   1425
      ItemData        =   "frmMain.frx":6856
      Left            =   2040
      List            =   "frmMain.frx":6858
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   3615
   End
   Begin VB.Frame fraStation 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Station Management && Play Controls "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8775
   End
   Begin VB.Frame fraDebug 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Debug Information "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   8775
      Begin VB.TextBox txtTmp 
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Label lblCC 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Licensed under Creative Commons Attribution 2.5. Some Rights Reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1800
      MouseIcon       =   "frmMain.frx":685A
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      Top             =   4820
      Width           =   6870
   End
   Begin VB.Image fraCC 
      Height          =   465
      Left            =   360
      MouseIcon       =   "frmMain.frx":D0AC
      MousePointer    =   1  'Arrow
      Picture         =   "frmMain.frx":138FE
      Top             =   4680
      Width           =   1320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkRecording_Click()
    If chkRecording.value = CheckBoxConstants.vbChecked Then
        DoDownload = True
    Else
        DoDownload = False
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim tmpStation As String, tmpName As String, ix As Integer
    tmpStation = InputBox("Enter Station URL", "Station URL", "http://128.177.3.80:4028")
    If Trim(tmpStation) <> "" Then
        If (LCase(Mid(tmpStation, 1, 7)) <> "http://") Then 'Check for Valid HTTP://
            MsgBox "URL should start with http://", vbInformation
            Exit Sub
        End If
        If lstStations.ListCount > 0 Then 'Check for Repetitions of URL's
            For ix = 0 To lstStations.ListCount - 1
                If lstStations.List(ix) = tmpStation Then
                    MsgBox "Station aleady in Queue", vbInformation
                    Exit Sub
                End If
            Next ix
        End If
    End If
    tmpStation = Trim(tmpStation)
    If tmpStation = "" Then Exit Sub 'Dont ask Station Name if user cancels URL
    tmpName = InputBox("Enter Station Name", "Station Name", "DI.FM Classics")
    If Trim(tmpName) = "" Then
        tmpName = "UnNamed Station"
    End If
    lstStationName.AddItem tmpName
    lstStations.AddItem tmpStation
    StatusMsg "... added station: " + tmpName
End Sub

Private Sub cmdPlay_Click()
    If lstStations.ListCount = 0 Then
        MsgBox "No Stations to Play", vbInformation
        Exit Sub
    End If
    If lstStations.SelCount <> 0 Then
        If (BASSThread) Then ' Already Connecting
            Call Beep
        Else
            ' Open URL in a New Thread (so that main thread is free)
            Dim threadid As Long
            streamURL = lstStations.List(lstStations.ListIndex)
            StatusMsg "playing station: " & lstStationName.List(lstStationName.ListIndex)
            BASSThread = CreateThread(ByVal 0&, 0, AddressOf playStation, lstStations.ListIndex, 0, threadid)     ' threadid param required on win9x
        End If
    Else
        MsgBox "Select a Station to Play ", vbInformation
    End If
End Sub

Private Sub cmdStop_Click()
    BASS_ChannelStop SEChannel 'stop playback
    Call CloseHandle(BASSThread) 'Close the thread
    BASSThread = 0
End Sub

Private Sub Form_Initialize()
Dim X As Long
    X = InitCommonControls
End Sub

Private Sub cmdRemove_Click()
    If lstStations.ListCount = 0 Then
        MsgBox "No Stations to Delete", vbInformation
        Exit Sub
    End If
    If lstStations.SelCount <> 0 Then
        lstStations.RemoveItem lstStations.ListIndex
        StatusMsg "... removed station: " + lstStationName.List(lstStationName.ListIndex)
        lstStationName.RemoveItem lstStationName.ListIndex
    Else
        MsgBox "Select an Station (Name/URL) to Delete ", vbInformation
    End If
End Sub

Private Sub Form_Load()
    Show
    DoEvents
    StatusMsg "Stream Engine " & App.Major & "." & App.Minor & "." & App.Revision & " by CS3"
    StatusMsg "--"
    If Not HiWord(BASS_GetVersion) = BASSVERSION Then
        MsgBox "BASS API version 2.3 not found", vbCritical, "Fatal Error"
        End
    Else
        StatusMsg "BASS Sound System version 2.3 loaded"
    End If
    If BASS_Init(-1, 44100, 0, frmMain.hwnd, 0) = 0 Then
        MsgBox "Cannot initialize Playback Device", vbCritical, "Fatal Error"
        End
    Else
        StatusMsg "using device: " & VBStrFromAnsiPtr(BASS_GetDeviceDescription(BASS_GetDevice())) & " @ 44100 Hz"
    End If
    Call BASS_SetConfig(BASS_CONFIG_NET_PREBUF, 0) 'Min Auto pre-buffering, so we can do it (and display it) instead
    StatusMsg "Setting BASS_CONFIG_NET_PREBUF as 0"
    Call loadPlugins
    Call loadList
    Set WriteFile = New clsFileIo 'For writing Streams to Files
    StatusMsg "--"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCC.ForeColor = vbBlack
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call saveList
    Call BASS_Free
    Call BASS_PluginFree(0) 'Free all loaded plugins
End Sub

Private Sub fraCC_Click()
    Call launchCC
End Sub

Private Sub fraCC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCC.ForeColor = vbBlack
End Sub

Private Sub lblCC_Click()
    Call launchCC
End Sub

Private Sub lblCC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCC.ForeColor = vbBlue
End Sub

Private Sub lstStationName_Click()
On Error Resume Next
    lstStations.Selected(lstStationName.ListIndex) = True
End Sub

Private Sub lstStations_Click()
On Error Resume Next
    lstStationName.Selected(lstStations.ListIndex) = True
End Sub

Private Sub txtStatus_GotFocus()
    txtTmp.SetFocus
End Sub

Private Sub txtStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCC.ForeColor = vbBlack
End Sub

