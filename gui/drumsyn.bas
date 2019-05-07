Attribute VB_Name = "DRUMSYN1"
Global errorText, errorCount
Global DS_PATH 'current root directory
Global SaveWAV

'Synthesis
Global filename As String
Global filename_ds As String
Global filename_mute As String
Global AskFileName As String
Global playable As Integer
Global CurEnv                      'current envelope
Global envpts(7, 1, 32)            'envelope time/level
Global envbuf(1, 32)               'copy/paste buffer
Global envcol(7)                   'envelope colours
Global EnvRange                    'ms of envelope displayed

Dim envData(7, 4): Const MAX = 0: Const ENV = 1
Dim envmaxlength
Const PNT = 2: Const dENV = 3: Const NEXTT = 4
Dim timestretch

'DrumSynth DLL
Declare Function ds2wav Lib "ds2wav.dll" (ByVal dsfile As String, ByVal wavfile As String, ByVal hWnd As Long) As Long

'Screen colours
Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

'Help file
Global Const HELP_CONTEXT = &H1     ' Display topic identified by number in Data
Global Const HELP_QUIT = &H2        ' Terminate help
Global Const HELP_INDEX = &H3       ' Display index
'Global Const HELP_HELPONHELP = &H4  ' Display help on using help
'Global Const HELP_SETINDEX = &H5    ' Set an alternate Index for help file with more than one index
'Global Const HELP_KEY = &H101       ' Display topic for keyword in Data
'Global Const HELP_MULTIKEY = &H201  ' Lookup keyword in alternate table and display topic

Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
 
'Type MULTIKEYHELP
'    mkSize As Integer
'    mkKeylist As String * 1
'    szKeyphrase As String * 253
'End Type

'Ini File
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Global INI_FILENAME As String

'UINT GetTempFileName

    'LPCTSTR lpPathName, // address of directory name for temporary file
    'LPCTSTR lpPrefixString, // address of filename prefix
    'UINT uUnique,   // number used to create temporary filename
    'LPTSTR lpTempFileName   // address of buffer that receives the new filename
   ');

'Temp path
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

'Wave file
Type WAVEHEADER
    RiffHeader As String * 4
    RiffLength As Long
    RiffType As String * 4
    FormatHeader As String * 4
    FormatLength As Long
    FormatTag As Integer
    Channels As Integer 'Number of channels
    Fs As Long          'Sampling frequency
    Rate As Long        'Total bytes per second
    Bytes As Integer    'Total bytes per sample
    Bits As Integer     'Bits per sample per channel
    DataHeader As String * 4
    DataLength As Long
    Data As Integer
End Type

'Type Dat
'    Dat(1199) As Integer
'End Type

'Type Daf
'    Daf(1199) As Single
'End Type
          
'PlaySound
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Global Const SND_ASYNC = &H1     'play asynchronously
Global Const SND_NODEFAULT = &H2 'don't use default sound
'Global Const SND_MEMORY = &H4    'lpszSoundName points to a memory file
'Global Const SND_LOOP = &H8      'loop the sound until next sndPlaySound
'Global Const SND_NOSTOP = &H10   'don't stop any currently playing sound
'examples...
'rtn = PlaySound(ByVal 0&, ByVal 0&, SND_ASYNC) 'mute
'rtn = PlaySound(filenameString, ByVal 0&, SND_ASYNC Or SND_NODEFAULT)

'Pop-Up Menu
'Type Rect
'    Left As Integer
'    Top As Integer
'    Right As Integer
'    Bottom As Integer
'End Type
'Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu%, ByVal wFlags%, ByVal X%, ByVal Y%, ByVal r2%, ByVal hWd%, r As Rect) As Integer
'Declare Function GetMenu Lib "user32" (ByVal hWd%) As Integer
'Declare Function GetSubMenu Lib "user32" (ByVal hMenu%, ByVal nPos%) As Integer

Sub drawAllEnvs()

  'errText ("DrawAllEnvs()")
  mainform.envbox.AutoRedraw = True
  mainform.envbox.Cls
  For j = 1 To 7
    If envIsOn(j) Then Call DrawEnv(-j)
  Next j
  mainform.envbox.AutoRedraw = False
  mainform.envbox.ForeColor = &H0
  Call DrawEnv(CurEnv)

End Sub

Sub DrawEnv(ByVal e)  'draw an envelope
  
  On Error GoTo drawenverr

  cw = 0.005 * mainform.envbox.ScaleWidth
  tmp = 2 '1.5

  If e > 0 Then
    mainform.envbox.Cls
    circ = True
    mainform.envbox.ForeColor = &H0&
  Else
    e = -e
    circ = False
    mainform.envbox.ForeColor = envcol(e)
  End If
       
  tmpnext = 1
  If e = 7 Then tmpnext = 3 'don't show last filter point!
  ep = 0
  While envpts(e, 0, ep + tmpnext) >= 0
    mainform.envbox.Line (envpts(e, 0, ep), envpts(e, 1, ep))-(envpts(e, 0, ep + 1), envpts(e, 1, ep + 1))
    If circ Then
      mainform.envbox.Line (envpts(e, 0, ep) - cw, envpts(e, 1, ep) - tmp)-(envpts(e, 0, ep) + cw, envpts(e, 1, ep) + tmp), , B
    End If
    ep = ep + 1
  Wend
  If circ Then 'end point
    mainform.envbox.Line (envpts(e, 0, ep) - cw, envpts(e, 1, ep) - tmp)-(envpts(e, 0, ep) + cw, envpts(e, 1, ep) + 1.5), , B
  End If
  DoEvents 'to relieve dragging
  
Exit Sub

drawenverr:
  'Debug.Print drawenverr!
  Resume Next

End Sub

Function envIsOn(e)
  
  eok = False
  Select Case e
    Case 1: If mainform.chkOn(0) = 1 Then eok = True
    Case 2: If mainform.chkOn(1) = 1 Then eok = True
    Case 3, 4: If mainform.chkOn(2) = 1 Then eok = True
    Case 5: If mainform.chkOn(3) = 1 Then eok = True
    Case 6: If mainform.chkOn(4) = 1 Then eok = True
    Case 7: If (mainform.MasterFilter = 1) Or (mainform.chkOFilter = 1) Then eok = True
  End Select
  envIsOn = eok

End Function

Sub errText(t)
  errorText = t
  mainform.Caption = t
End Sub

Sub fillfilelist()

  'errText ("fillFileList()")
  mainform.fileList.Clear
  pp = DS_PATH & "\" & mainform.kitList.List(mainform.kitList.ListIndex)
  fl = Dir(pp & "\*.ds")
  While fl <> ""
    'fl = UCase(Left(fl, 1)) & LCase(Mid(fl, 2, Len(fl) - 4))
    fl = Left(fl, Len(fl) - 3)
    mainform.fileList.AddItem fl
    fl = Dir
  Wend

  'errText ("fillFileListWav()")
  fl = Dir(pp & "\*.wav")
  While fl <> ""
    WAVname = LCase(Left(fl, 1)) & LCase(Mid(fl, 2, Len(fl) - 5))
    found = False
    For srch = 0 To mainform.fileList.ListCount - 1
      If LCase(mainform.fileList.List(srch)) = WAVname Then found = True
    Next srch
    If found = False Then mainform.fileList.AddItem WAVname & " "
    fl = Dir
  Wend

End Sub

Sub getEnv(e, en)  'decodes envelope string to array

  p = 1
  ep = 0
  pmax = Len(en)

  s = ""
  While p <= pmax
    c = Mid(en, p, 1)
    If c = "," Then
      envpts(e, 0, ep) = Val(s)
      s = ""
      
    ElseIf c = " " Then
      envpts(e, 1, ep) = Val(s)
      s = ""
      ep = ep + 1
    Else
      s = s & c
    End If
    p = p + 1
  Wend
  envpts(e, 1, ep) = Val(s)
  envpts(e, 0, ep + 1) = -1

  'If envpts(e, 0, ep) > envmaxlength Then envmaxlength = envpts(e, 0, ep)

End Sub

Function getini(Section As String, Key As String, Default As String)
  
  Dim retVal As String, AppName As String, Worked As Integer
  
  retVal = String$(255, 0)
  Worked = GetPrivateProfileString(Section, Key, "", retVal, Len(retVal), INI_FILENAME)
  If Worked = 0 Then
    getini = Default
  Else
    getini = Left(retVal, Worked)
  End If

End Function

Sub Synth()
  
  On Error Resume Next
  
  Dim WH As WAVEHEADER
  
  WH.RiffHeader = "RIFF"
  WH.RiffLength = 38
  WH.RiffType = "WAVE"
  WH.FormatHeader = "fmt "
  WH.FormatLength = 16
  WH.FormatTag = 1 'PCM
  WH.Channels = 1
  WH.Fs = 44100
  WH.Rate = 88200
  WH.Bytes = 2
  WH.Bits = 16
  WH.DataHeader = "data"
  WH.DataLength = 2
  WH.Data = 0
  
  Open filename_mute For Binary As #7
    Put #7, , WH
  Close #7
  
End Sub

Sub GetWorkingPath()
  
  Dim rtn As String * 100
  
  filename = App.Path
  If Right(filename, 1) <> "\" Then filename = filename & "\"
  filename_mute = filename & "DrumSyn_.tmp"
  filename_ds = filename & "DrumSyn-.tmp"
  filename = filename & "DrumSyn~.tmp"
  
On Error GoTo FilePathErr
  
  Open filename For Output As #8

On Error GoTo FilePathErr2
  
  Close #8
  
FilePathEx:

Exit Sub

FilePathErr:
  'errText ("FilePathErr!")
  nul = GetTempFileName(0, "tmp", 1, rtn)
  filename_ds = Left(rtn, InStr(rtn, Chr(0)) - 13) & "DrumSyn-.tmp"
  filename = Left(rtn, InStr(rtn, Chr(0)) - 13) & "DrumSyn~.tmp"
  Resume FilePathEx

FilePathErr2:
  'errText ("FilePathErr2!")
  Resume Next

End Sub

Function Log10(v)
  Log10 = Log(v) / Log(10)
End Function

Function LongestEnv()

  L = 0
  For e = 1 To 6
    eon = e - 1: If eon > 2 Then eon = eon - 1
      p = 0
      While envpts(e, 0, p + 1) >= 0
        p = p + 1
      Wend
      envData(e, MAX) = Int(envpts(e, 0, p))
    If mainform.chkOn(eon) Then
      If envData(e, MAX) > L Then L = envData(e, MAX)
    End If
  Next e

  L = L * timestretch '***???
  LongestEnv = 2400 + (1200 * Int(L / 1200))

End Function

Function LoudestEnv()
  
  loudest = 0
  For i = 0 To 4
    If (mainform.chkOn(i)) Then
      If (mainform.sliLev(i)) > loudest Then
        loudest = (mainform.sliLev(i))
      End If
    End If
  Next i
  LoudestEnv = loudest ^ 2
                          
End Function

Function NeatName(f)  'gives home and file name

  pos = Len(f)
  While (pos > 0) And Mid(f, pos, 1) <> "\"
    pos = pos - 1
  Wend
  NameOnly = UCase(Mid(f, pos + 1, 1)) & LCase(Mid(f, pos + 2, Len(f) - pos - 4))

'  pos2 = pos - 1
'  While (pos2 > 0) And Mid(f, pos2, 1) <> "\"
'    pos2 = pos2 - 1
'  Wend
'
'  If pos2 > 1 Then
'    NeatName = Mid(f, pos2 + 1, pos - pos2 - 1) & " " & NameOnly
'  Else
    NeatName = NameOnly
'  End If

End Function

Sub openfile(filename)
  Dim sc As String
  Dim vvv As String 'without this, version gives type mismatch!
  sc = "General"
  INI_FILENAME = filename
  
  On Error Resume Next '///////////////
  DoEvents
  mainform.Caption = mainform.Caption & ".": DoEvents
  
  vvv = getini(sc, "Version", "")  '//////////////////

  'If vvv = "DrumSynth v1.0" Or vvv = "DrumSynth v2.0" Then
    mainform.Caption = "[" & vvv & "]"

    mainform.Caption = mainform.Caption & "1": DoEvents
    mainform.comment.Text = getini(sc, "Comment", "")

    mainform.Caption = mainform.Caption & "2": DoEvents
    mainform.MasterTune = getini(sc, "Tuning", "0.00")

    mainform.Caption = mainform.Caption & "3": DoEvents
    mainform.MasterLength = getini(sc, "Stretch", "100.0")
    
    mainform.Caption = mainform.Caption & "4": DoEvents
    mainform.sliMasterdB = getini(sc, "Level", "0")
    
    If mainform.MasterTune = 0 And mainform.MasterLength = 0 And mainform.sliMasterdB = 0 Then
      mainform.Caption = mainform.Caption & "5": DoEvents
      mainform.chkMasterOn = 0
    Else
      mainform.Caption = mainform.Caption & "6": DoEvents
      mainform.chkMasterOn = 1
    End If

mainform.Caption = mainform.Caption & ",": DoEvents '/////////////////////////
    mainform.MasterFilter = getini(sc, "Filter", "0")
    mainform.chkHPF = getini(sc, "HighPass", "0")
    mainform.sliFres = getini(sc, "Resonance", "0")
    en = getini(sc, "FilterEnv", "0,100 442000,100 443000,0"): Call getEnv(7, en)
    envmaxlength = 0 'find longest envelope except filter
mainform.Caption = mainform.Caption & "4": DoEvents
    sc = "Tone"
    mainform.chkOn(0) = Val(getini(sc, "On", "0"))
    mainform.sliLev(0) = Val(getini(sc, "Level", "128"))
    mainform.txtTF1 = getini(sc, "F1", "200")
    mainform.txtTF2 = getini(sc, "F2", "120")
    mainform.sliTDroop = Val(getini(sc, "Droop", "0"))
    mainform.txtTPhase = getini(sc, "Phase", "90")
    en = getini(sc, "Envelope", "0,100 100,30 200,0"): Call getEnv(1, en)
mainform.Caption = mainform.Caption & "5": DoEvents
    sc = "Noise"
    mainform.chkOn(1) = Val(getini(sc, "On", "0"))
    mainform.sliLev(1) = Val(getini(sc, "Level", "128"))
    mainform.sliNSlope = Val(getini(sc, "Slope", "0"))
    mainform.chkFixedRandom = Val(getini(sc, "FixedSeq", "0"))
    en = getini(sc, "Envelope", "0,100 100,30 200,0"): Call getEnv(2, en)
mainform.Caption = mainform.Caption & "6": DoEvents
    sc = "Overtones"
    mainform.chkOn(2) = Val(getini(sc, "On", "0"))
    mainform.sliLev(2) = Val(getini(sc, "Level", "128"))
    mainform.txtOF1 = getini(sc, "F1", "200")
    mainform.comOW1.ListIndex = Val(getini(sc, "Wave1", "0"))
    mainform.chkOTrack1 = getini(sc, "Track1", "0")
    mainform.txtOF2 = getini(sc, "F2", "120")
    mainform.comOW2.ListIndex = Val(getini(sc, "Wave2", "0"))
    mainform.chkOTrack2 = getini(sc, "Track2", "0")
    mainform.chkOFilter = getini(sc, "Filter", "0")
     If mainform.chkOFilter = 1 Then mainform.MasterFilter = 1
mainform.Caption = mainform.Caption & "7": DoEvents
    mainform.comOMethod.ListIndex = Val(getini(sc, "Method", "2"))
    mainform.sliOParam = getini(sc, "Param", "50")
    en = getini(sc, "Envelope1", "0,100 100,30 200,0"): Call getEnv(3, en)
    en = getini(sc, "Envelope2", "0,100 100,30 200,0"): Call getEnv(4, en)
mainform.Caption = mainform.Caption & "8": DoEvents
    sc = "NoiseBand"
    mainform.chkOn(3) = Val(getini(sc, "On", "0"))
    mainform.sliLev(3) = Val(getini(sc, "Level", "128"))
    mainform.txtNF = getini(sc, "F", "1000")
    mainform.sliNdF = Val(getini(sc, "dF", "50"))
    en = getini(sc, "Envelope", "0,100 100,30 200,0"): Call getEnv(5, en)
mainform.Caption = mainform.Caption & "9": DoEvents
    sc = "NoiseBand2"
    mainform.chkOn(4) = Val(getini(sc, "On", "0"))
    mainform.sliLev(4) = Val(getini(sc, "Level", "128"))
    mainform.txtNF2 = getini(sc, "F", "1000")
    mainform.sliNdF2 = Val(getini(sc, "dF", "50"))
    en = getini(sc, "Envelope", "0,100 100,30 200,0"): Call getEnv(6, en)
 mainform.Caption = mainform.Caption & "0": DoEvents
    sc = "Distortion"
    mainform.chkOn(5) = Val(getini(sc, "On", "0"))
    mainform.sliDist = getini(sc, "Clipping", "0")
    mainform.comDB.ListIndex = Val(getini(sc, "Bits", "0"))
    mainform.comDF.ListIndex = Val(getini(sc, "Rate", "0"))
    mainform.Caption = NeatName(filename) & " - " & App.Title
'  Else
'    mainform.warning.Visible = True
'    mainform.Caption = App.Title
'  End If

' If envmaxlength < 16640 Then
'   If envmaxlength < 4410 Then
'     mainform.txtEnvRange = "100 ms"
'     EnvRange = 4410
'   Else
'     mainform.txtEnvRange = "400 ms"
'     EnvRange = 17640
'   End If
' Else
'   If envmaxlength < 44100 Then
'     mainform.txtEnvRange = "1 sec"
'     EnvRange = 44100
'   Else
'     mainform.txtEnvRange = "3 sec"
'     EnvRange = 132300
'   End If
' End If
' mainform.envbox.ScaleWidth = EnvRange
End Sub

Function parsefreq(f)
  
  note = f
  freq = 1000
  
  Select Case Left(note, 1)
    Case "c", "C": midi = 60
    Case "d", "D": midi = 62
    Case "e", "E": midi = 64
    Case "f", "F": midi = 65
    Case "g", "G": midi = 67
    Case "a", "A": midi = 69
    Case "b", "B": midi = 71
  End Select

  If midi Then 'ie. starts with a note letter
    
    note = note & "  "
    If Mid(note, 2, 1) = "b" Then midi = midi - 1 'flat
    If Mid(note, 2, 1) = "#" Then midi = midi + 1 'sharp
    
    pos = 1
    Number = False
    Do
      pos = pos + 1
      ch = Asc(Mid(note, pos, 1))
      If ch > 47 And ch < 58 Then Number = True '0 to 9
      If ch = 43 Then Number = True '+
      If ch = 45 Then Number = True '-
    Loop While Number = False

    Number = True
    pos2 = pos
    Do
      pos2 = pos2 + 1
      ch = Asc(Mid(note, pos2, 1))
      If ch < 48 Then Number = False
      If ch > 57 Then Number = False
    Loop While Number = True

    octave = Val(Mid(note, pos, pos2 - pos))
    cents = Val(Right(note, Len(note) - pos2))
    midi = midi + 12 * (octave - 3) + (cents / 100)
    freq = 8.1757989 * (1.0594631 ^ midi)
    f = Format(freq, "0.00") 'put Hz value in text box!
    
  Else
    freq = Val(note)
  End If
    
  'If f > 22000 Then
  '  f = 22000
  'ElseIf f < 1 Then
  '  f = 1
  'End If
  
  finetune = 1.0594631 ^ (mainform.MasterTune)
  parsefreq = freq * finetune

End Function

Function putenv(e)  'encodes envelope points to string
  
  s = ""
  ep = 0
  While envpts(e, 0, ep) >= 0
    s = s & Format(envpts(e, 0, ep), "0") & "," & Format(envpts(e, 1, ep), "0") & " "
    ep = ep + 1
  Wend
  putenv = s

End Function

Sub savefile(filename)
  
  Dim sc As String
  'MsgBox "Saving " & filename

  INI_FILENAME = filename

  sc = "General"
  setini sc, "Version", "DrumSynth v2.0"
  setini sc, "Comment", (mainform.comment)
  
  If mainform.chkMasterOn = 1 Then
    setini sc, "Tuning", (mainform.MasterTune)
    setini sc, "Stretch", (mainform.MasterLength)
    setini sc, "Level", (mainform.sliMasterdB)
    Debug.Print (mainform.sliMasterdB)
  Else
    setini sc, "Tuning", 0
    setini sc, "Stretch", 100
    setini sc, "Level", 0
  End If
  
  If mainform.chkOFilter = 0 Then
    setini sc, "Filter", (mainform.MasterFilter)
  Else
    setini sc, "Filter", 0
  End If
  setini sc, "HighPass", (mainform.chkHPF)
  setini sc, "Resonance", (mainform.sliFres)
  setini sc, "FilterEnv", putenv(7)
    
  sc = "Tone"
  setini sc, "On", (mainform.chkOn(0))
  setini sc, "Level", (mainform.sliLev(0))
  setini sc, "F1", (mainform.txtTF1)
  setini sc, "F2", (mainform.txtTF2)
  setini sc, "Droop", (mainform.sliTDroop)
  setini sc, "Phase", (mainform.txtTPhase)
  setini sc, "Envelope", putenv(1)

  sc = "Noise"
  setini sc, "On", (mainform.chkOn(1))
  setini sc, "Level", (mainform.sliLev(1))
  setini sc, "Slope", (mainform.sliNSlope)
  setini sc, "Envelope", putenv(2)
  setini sc, "FixedSeq", (mainform.chkFixedRandom)
  
  sc = "Overtones"
  setini sc, "On", (mainform.chkOn(2))
  setini sc, "Level", (mainform.sliLev(2))
  setini sc, "F1", (mainform.txtOF1)
  setini sc, "Wave1", (mainform.comOW1.ListIndex)
  setini sc, "Track1", (mainform.chkOTrack1)
  setini sc, "F2", (mainform.txtOF2)
  setini sc, "Wave2", (mainform.comOW2.ListIndex)
  setini sc, "Track2", (mainform.chkOTrack2)
  setini sc, "Method", (mainform.comOMethod.ListIndex)
  setini sc, "Param", (mainform.sliOParam)
  setini sc, "Envelope1", putenv(3)
  setini sc, "Envelope2", putenv(4)
  If mainform.MasterFilter = 1 Then
    setini sc, "Filter", (mainform.chkOFilter)
  Else
    setini sc, "Filter", 0
  End If
  
  sc = "NoiseBand"
  setini sc, "On", (mainform.chkOn(3))
  setini sc, "Level", (mainform.sliLev(3))
  setini sc, "F", (mainform.txtNF)
  setini sc, "dF", (mainform.sliNdF)
  setini sc, "Envelope", putenv(5)
  
  sc = "NoiseBand2"
  setini sc, "On", (mainform.chkOn(4))
  setini sc, "Level", (mainform.sliLev(4))
  setini sc, "F", (mainform.txtNF2)
  setini sc, "dF", (mainform.sliNdF2)
  setini sc, "Envelope", putenv(6)
  
  sc = "Distortion"
  setini sc, "On", (mainform.chkOn(5))
  'If mainform.chkOn(5) = 1 Then
    setini sc, "Clipping", (mainform.sliDist)
    setini sc, "Bits", (mainform.comDB.ListIndex)
    setini sc, "Rate", (mainform.comDF.ListIndex)
  'Else
    
  'End If

  'reset wave playback (it thinks it knows this filename)
  rtn = PlaySound(ByVal 0&, ByVal 0&, SND_ASYNC Or SND_NODEFAULT)

  mainform.Caption = NeatName(filename) & " - " & App.Title
  'faliure: mainform.caption = App.Title

End Sub

Function ScreenColors() As Long
  Const PLANES = 14
  Const BITSPIXEL = 12
  ScreenColors = 2 ^ (GetDeviceCaps((hDC), PLANES) * GetDeviceCaps((hDC), BITSPIXEL)) 'mainform.hDC
End Function

Sub setini(Section As String, Key As String, Value As String)
  
  Dim retVal As String, AppName As String, Worked As Integer
  
  retVal = String$(255, 0)
  Worked = WritePrivateProfileString(Section, Key, Value, INI_FILENAME)

End Sub


'Sub Synth(filename)
'
'  playable = True
'
'End Sub

Sub UpdateEnv(e, t)
  
  envData(e, NEXTT) = envpts(e, 0, envData(e, PNT) + 1) * timestretch '***???
  If envData(e, NEXTT) < 0 Then envData(e, NEXTT) = 442000 * timestretch
  envData(e, ENV) = envpts(e, 1, envData(e, PNT)) / 100#
  endEnv = envpts(e, 1, envData(e, PNT) + 1) / 100#
  dT = envData(e, NEXTT) - t
  If dT < 1# Then dT = 1#
  envData(e, dENV) = (endEnv - envData(e, ENV)) / dT
  envData(e, PNT) = envData(e, PNT) + 1

End Sub

Function waveform(ph, Form)

  Select Case Form
    Case 0:
      waveform = Sin(ph)

    Case 1:
      waveform = Abs(2 * Sin(ph / 2)) - 1

    Case 2:
      phh = ph - twoPi * Int(ph / twoPi)
      phh = (0.6366197 * phh) - 1 '1/pi
      If phh > 1# Then
        waveform = 2# - phh
      Else
        waveform = phh
      End If
    
    Case 3:
      phh = ph - twoPi * Int(ph / twoPi)
      waveform = (0.3183098 * phh) - 1 '1/pi

    Case Else:
      waveform = Sgn(Sin(ph))


  End Select


End Function

