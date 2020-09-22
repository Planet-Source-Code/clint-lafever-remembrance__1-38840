Attribute VB_Name = "basMAIN"
Option Explicit
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0        ' Play synchronously (default).
Private Const SND_NODEFAULT = &H2    ' Do not use default sound.
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8         ' Loop the sound until next
Private Const SND_NOSTOP = &H10      ' Do not stop any currently
Private Const SND_ASYNC = &H1          '  play asynchronously
Private bytSound() As Byte ' Always store binary data in byte arrays!
Public Enum SoundFlags
    soundSYNC = SND_SYNC
    soundNO_DEFAULT = SND_NODEFAULT
    soundMEMORY = SND_MEMORY
    soundLOOP = SND_LOOP
    soundNO_STOP = SND_NOSTOP
    soundASYNC = SND_ASYNC
End Enum
Public Enum AppSounds
    sndSTAR_BANNER = 101
    sndMAX = 102
End Enum
Public Sub PlayWaveRes(vntResourceID As AppSounds, Optional vntFlags As SoundFlags = soundASYNC)
    bytSound = LoadResData(vntResourceID, "WAVE")
    If IsMissing(vntFlags) Then
        vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
    End If
    If (vntFlags And SND_MEMORY) = 0 Then
        vntFlags = vntFlags Or SND_MEMORY
    End If
    sndPlaySound bytSound(0), vntFlags
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever - [lafeverc@saic.com]
' Purpose:  Extracts a file from the custom resource file
'                to the local hard drive.
' Parameters:  resID=ID of resource  :  resSECTION=Section of custom resource ie. CUSTOM
'                     fEXT=Extension for new file  :  fPATH=Destination path, default is App.Path
'                     fNAME=Name for new file, default is TEMP
' Returns:  Full path and file name of file created
' Example:  retSTR=GenFileFromRes(101,"CUSTOM","JPG",,"IMAGE")
' Date: December,17 1999 @ 10:50:58
'------------------------------------------------------------
Public Function GenFileFromRes(resID As Long, resSECTION As String, fEXT As String, Optional fPath As String = "", Optional fNAME As String = "temp", Optional FullName As String = "") As String
    On Error GoTo ErrorGenFileFromRes
    Dim resBYTE() As Byte
    If fPath = "" Then fPath = App.Path
    If fNAME = "" Then fNAME = "temp"
    '------------------------------------------------------------
    ' Get the file out of the resource file
    '------------------------------------------------------------
    resBYTE = LoadResData(resID, resSECTION)
    '------------------------------------------------------------
    ' Open destination
    '------------------------------------------------------------
    If FullName = "" Then
        Open fPath & "\" & fNAME & "." & fEXT For Binary Access Write As #1
    Else
        Open FullName For Binary Access Write As #1
    End If
    '------------------------------------------------------------
    ' Write it out
    '------------------------------------------------------------
    Put #1, , resBYTE
    '------------------------------------------------------------
    ' Close it
    '------------------------------------------------------------
    Close #1
    If FullName = "" Then
        GenFileFromRes = fPath & "\" & fNAME & "." & fEXT
    Else
        GenFileFromRes = FullName
    End If
    Exit Function
ErrorGenFileFromRes:
    GenFileFromRes = ""
    MsgBox Err & ":Error in GenFileFromRes.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function
Public Sub Main()
    On Error Resume Next
    '------------------------------------------------------------
    ' Extract the .TXT file list of names.
    '------------------------------------------------------------
    GenFileFromRes 101, "TXT", "TXT", , , App.Path & "\091101.txt"
    GenFileFromRes 101, "JPG", "JPG", , , App.Path & "\flag.jpg"
    frmMAIN.Show
    frmMAIN.Refresh
    frmMAIN.BeginLoop
End Sub
