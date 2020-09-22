VERSION 5.00
Begin VB.Form frmMAIN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Remembrance"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Monotype Corsiva"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgFLAG 
      Height          =   495
      Left            =   1800
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'
' Press Escape to Quit out when you run it
'
'*********************************************
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const VK_ESCAPE = &H1B
Private QuitIt As Boolean
Private Const Ending = "All Gave Some...Some Gave All."
Private Sub Form_Load()
    On Error Resume Next
    Set Me.imgFLAG.Picture = LoadPicture(App.Path & "\flag.jpg")
    Me.WindowState = vbMaximized
End Sub
Public Sub BeginLoop()
    On Error Resume Next
    Dim SomeGaveAll, AllGaveSome
    Dim X As Long, Y As Long, tc As Long, Hero As String, z As Long, lORr As Long
    MakeFlag
    PlayWaveRes sndSTAR_BANNER, soundASYNC Or soundLOOP
    ShowCursor False
    Me.Refresh
    tc = GetTickCount
    While GetTickCount < tc + 1000 And Not QuitIt: DoEvents: Sleep 1: Wend
    '*********************************************
    ' Open our file
    '*********************************************
    Open App.Path & "\091101.txt" For Input As #1
    SomeGaveAll = ""
    lORr = 0
    '*********************************************
    ' Loop to end of file or until escape is pressed
    '*********************************************
    Do While Not EOF(1) And Not QuitIt
        '*********************************************
        ' Read a line
        '*********************************************
        Line Input #1, AllGaveSome
        '*********************************************
        ' Build a listing big enough to fill from top of form to bottom
        '*********************************************
        If Me.TextHeight(SomeGaveAll) + Me.TextHeight(AllGaveSome) < Me.ScaleHeight Then
            If SomeGaveAll = "" Then
                SomeGaveAll = AllGaveSome
            Else
                SomeGaveAll = SomeGaveAll & vbCrLf & AllGaveSome
            End If
        Else
            If lORr = 0 Then
                Me.Cls
                lORr = 1
            Else
                lORr = 0
            End If
            Y = Me.ScaleHeight / 2 - (Me.TextHeight(SomeGaveAll) / 2)
            Me.CurrentY = Y
            '*********************************************
            ' Prepare the string with a ~ as a delimiter
            '*********************************************
            SomeGaveAll = Replace(SomeGaveAll, vbCrLf, "~")
            '*********************************************
            ' For each Hero, loop and print out name centered
            '*********************************************
            For z = 1 To CountTokens(SomeGaveAll, "~")
                Hero = GetToken(SomeGaveAll, "~")
                If lORr = 1 Then
                    X = Me.ScaleWidth / 4 - (Me.TextWidth(Hero) / 2)
                Else
                    X = (Me.ScaleWidth / 2 + (Me.ScaleWidth / 4)) - (Me.TextWidth(Hero) / 2)
                End If
                Me.CurrentX = X + 70
                Y = Me.CurrentY
                Me.CurrentY = Y + 70
                Me.ForeColor = vbBlack
                Me.Print Hero
                Me.CurrentX = X
                Me.CurrentY = Y
                Me.ForeColor = &HC0C0C0
                Me.Print Hero
                tc = GetTickCount
                While GetTickCount < tc + 5 And Not QuitIt: Me.Refresh: DoEvents: Sleep 1: Wend
            Next z
            If lORr = 0 Then
                tc = GetTickCount
                While GetTickCount < tc + 1500 And Not QuitIt: DoEvents: Sleep 1: Wend
            End If
            '*********************************************
            ' Do we need to quit because user pressed escape?
            '*********************************************
            If GetAsyncKeyState(VK_ESCAPE) <> 0 Then QuitIt = True
            SomeGaveAll = ""
        End If
    Loop
    Close #1
    Me.Cls
    tc = GetTickCount
    While GetTickCount < tc + 2000 And Not QuitIt: DoEvents: Sleep 1: Wend
    Me.Font.Name = "Verdana"
    Me.FontBold = False
    Me.FontSize = 10
    Me.FontItalic = False
    Me.FontStrikethru = False
    Me.FontUnderline = False
    Y = Me.ScaleHeight / 2 - (Me.TextHeight(Ending) / 2)
    X = Me.ScaleWidth / 2 - (Me.TextWidth(Ending) / 2)
    For z = 0 To 255
        Me.ForeColor = RGB(z, z, z)
        Me.CurrentX = X
        Me.CurrentY = Y
        Me.Print Ending
        tc = GetTickCount
        While GetTickCount < tc + 5 And Not QuitIt: DoEvents: Sleep 1: Wend
    Next z
    '*********************************************
    'Wait for escape key to be pressed
    '*********************************************
    While GetAsyncKeyState(VK_ESCAPE) = 0 And Not QuitIt: DoEvents: Sleep 1: Wend
    ShowCursor True
    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    PlayWaveRes sndMAX
    ShowCursor True
    QuitIt = True
End Sub
'------------------------------------------------------------
' Author:  Clint M. LaFever - [lafeverc@saic.com]
' Date: September,10 2002 @ 14:05:47
'------------------------------------------------------------
Private Sub MakeFlag()
    On Error GoTo ErrorMakeFlag
    PaintPicture imgFLAG, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Set Me.Picture = Me.Image
    Exit Sub
ErrorMakeFlag:
    MsgBox Err & ":Error in call to MakeFlag()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
