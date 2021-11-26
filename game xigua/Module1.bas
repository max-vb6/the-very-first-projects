Attribute VB_Name = "Module1"
Public Ico As Integer
Public b As Boolean
Public c As Integer

Option Explicit

'  flag values for uFlags parameter
Private Const SND_SYNC = &H0              '  play synchronously (default)
Private Const SND_ASYNC = &H1             '  play asynchronously
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_MEMORY = &H4            '  lpszSoundName points to a memory file
Private Const SND_ALIAS = &H10000         '  name is a WIN.INI [sounds] entry
Private Const SND_FILENAME = &H20000      '  name is a file name
Private Const SND_RESOURCE = &H40004      '  name is a resource name or atom
Private Const SND_ALIAS_ID = &H110000     '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ALIAS_START = 0         '  must be > 4096 to keep strings in same section of resource file
Private Const SND_LOOP = &H8              '  loop the sound until next sndPlaySound
Private Const SND_NOSTOP = &H10           '  don't stop any currently playing sound
Private Const SND_VALID = &H1F            '  valid flags          / ;Internal /
Private Const SND_NOWAIT = &H2000         '  don't wait if the driver is busy
Private Const SND_VALIDFLAGS = &H17201F   '  Set of valid flag bits.  Anything outside this range will raise an error
Private Const SND_RESERVED = &HFF000000   '  In particular these flags are reserved
Private Const SND_TYPE_MASK = &H170007

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'  waveform audio error return values
Private Const WAVERR_BASE = 32
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)       '  unsupported wave format
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)    '  still something playing
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)      '  header not prepared
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)            '  device is synchronous
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)       '  last error in range

Private m_snd() As Byte

Public Function PlaySoundResource(ByVal SndID As Long) As Long
   Const Flags = SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
   m_snd = LoadResData(SndID, "CUSTOM")
   PlaySoundResource = PlaySoundData(m_snd(0), 0, Flags)
End Function

Public Sub XiGua()
If c >= 20 Then
Form1.Image1.Picture = Form2.Image1(0).Picture
End If
If c >= 30 Then
Form1.Image1.Picture = Form2.Image1(1).Picture
End If
If c >= 40 Then
Form1.Image1.Picture = Form2.Image1(2).Picture
End If
If c >= 50 Then
Form1.Image1.Picture = Form2.Image1(3).Picture
End If
If c >= 60 Then
Form1.Image1.Picture = Form2.Image1(4).Picture
End If
If c >= 70 Then
Form1.Image1.Picture = Form2.Image1(5).Picture
End If
If c >= 80 Then
Form1.Image1.Picture = Form2.Image1(6).Picture
End If
If c >= 100 Then
Form1.Image1.Picture = Form2.Image1(7).Picture
b = False
Form1.Label4.Caption = ("������" + vbCrLf + "��Ӯ��")
End If
End Sub

