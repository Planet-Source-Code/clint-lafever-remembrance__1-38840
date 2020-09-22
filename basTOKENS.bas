Attribute VB_Name = "basTOKENS"
Option Explicit
'------------------------------------------------------------
' Author:  Clint M. LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  Gets the first token off a delimited string.
'                 Returns the token and changes the passed string
'                with the first token removed.
' Parameters:  delimited string and delimiter
' Returns:  First Token off delimited string
' Date: April,30 1999 @ 11:14:01
'------------------------------------------------------------
Function GetToken(sSource, ByVal sDelim As String) As String
   Dim iDelimPos As Integer
    On Error GoTo ErrorGetToken
   '------------------------------------------------------------
   ' Find the first delimiter
   '------------------------------------------------------------
   iDelimPos = InStr(1, sSource, sDelim)
   '------------------------------------------------------------
   ' If no delimiter was found, return the existing
   ' string and set the source to an empty string.
   '------------------------------------------------------------
   If (iDelimPos = 0) Then
      GetToken = Trim$(sSource)
      sSource = ""
   '------------------------------------------------------------
   ' Otherwise, return everything to the left of the
   ' delimiter and return the source string with it
   ' removed.
   '------------------------------------------------------------
   Else
      GetToken = Trim$(Left$(sSource, iDelimPos - 1))
      sSource = Mid$(sSource, iDelimPos + 1)
   End If
   Exit Function
ErrorGetToken:
    MsgBox Err & ":Error in GetToken() Function.  Error Message:" & Error(Err), 48, "Warning"
    Exit Function
End Function
'------------------------------------------------------------
' Author:  Clint M. LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  Counts the number of tokens that are in a delimited
'                string.
' Parameters:  delimited string and delimiter
' Returns:  Number of tokens.
' Date: April,30 1999 @ 11:14:24
'------------------------------------------------------------
Function CountTokens(ByVal sSource, ByVal sDelim As String) As Integer
   Dim iDelimPos As Integer
   Dim iCount As Integer
    On Error GoTo ErrorCountTokens
    '------------------------------------------------------------
    ' Number of tokens = 0 if the source string is
    ' empty
    '------------------------------------------------------------
   If sSource = "" Then
      CountTokens = 0
   '------------------------------------------------------------
   ' Otherwise number of tokens = number of delimiters
   '  1
   '------------------------------------------------------------
   Else
      iDelimPos = InStr(1, sSource, sDelim)
      Do Until iDelimPos = 0
         iCount = iCount + 1
         iDelimPos = InStr(iDelimPos + 1, sSource, sDelim)
      Loop
      CountTokens = iCount + 1
   End If
   Exit Function
ErrorCountTokens:
    MsgBox Err & ":Error in CountTokens() Function.  Error Message:" & Error(Err), 48, "Warning"
    Exit Function
End Function
