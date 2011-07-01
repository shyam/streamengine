Attribute VB_Name = "bass_aac"
Option Explicit

' Additional tags available from BASS_StreamGetTags
Global Const BASS_TAG_MP4 = 7       ' MP4/iTunes metadata

Global Const BASS_AAC_DOWNMATRIX = &H20000  ' downmatrix to stereo

' BASS_CHANNELINFO type
Global Const BASS_CTYPE_STREAM_AAC = &H10B00
Global Const BASS_CTYPE_STREAM_MP4 = &H10B01

Declare Function BASS_AAC_StreamCreateFile Lib "bass_aac.dll" (ByVal mem As Long, ByVal f As Any, ByVal offset As Long, ByVal length As Long, ByVal flags As Long) As Long
Declare Function BASS_AAC_StreamCreateURL Lib "bass_aac.dll" (ByVal url As String, ByVal offset As Long, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_AAC_StreamCreateFileUser Lib "bass_aac.dll" (ByVal buffered As Long, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_MP4_StreamCreateFile Lib "bass_aac.dll" (ByVal mem As Long, ByVal f As Any, ByVal offset As Long, ByVal length As Long, ByVal flags As Long) As Long
Declare Function BASS_MP4_StreamCreateFileUser Lib "bass_aac.dll" (ByVal buffered As Long, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long