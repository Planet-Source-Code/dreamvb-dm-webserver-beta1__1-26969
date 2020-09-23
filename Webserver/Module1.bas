Attribute VB_Name = "Module1"
Public FSO As FileSystemObject
Sub timeout(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub


Function addSlash(lzpath As String) As String
    If Right(lzpath, 1) <> "\" Then addSlash = lzpath & "\" Else addSlash = lzpath
    
End Function

Function ConvertWebSlash(lzUrl As String) As String
' not realy needed you chould just use the simple replace functions for this

Dim I As Integer
Dim StrL As String

    For I = 1 To Len(lzUrl)
        ch = Mid(lzUrl, I, 1)
        If ch = "/" Then
            StrL = StrL & "\"
        Else
            StrL = StrL & ch
        End If
    Next
    I = 0
    ConvertWebSlash = StrL
    lzUrl = ""
    
End Function

Function URL_Decode(TUrl As String) As String

Dim Xpos As Integer
Dim CGI_Str As String

    While (InStr(TUrl, "%") <> 0)
        Xpos = InStr(TUrl, "%")
         CGI_Str = Mid(TUrl, Xpos + 1, 2)
          TUrl = Replace(TUrl, "%" & CGI_Str, Chr("&H" & CGI_Str))
    Wend
        URL_Decode = Replace(TUrl, "+", " ")
        
End Function

