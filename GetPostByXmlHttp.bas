Attribute VB_Name = "GetPostByXmlHttp"
Option Explicit
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001

Function GetXmlHttp(ByVal GetUrl As String) As String
    Dim XmlHttp As Object
    Set XmlHttp = CreateObject("MSXML2.XMLHTTP")
    XmlHttp.Open "GET", GetUrl, True
    XmlHttp.send
    Do Until XmlHttp.readyState = 4
        DoEvents
    Loop
    GetXmlHttp = XmlHttp.ResponseText
    Set XmlHttp = Nothing
End Function

Function PostXmlHttp(ByVal PostUrl As String, ByVal PostData As String) As String
    Dim XmlHttp As Object
    Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
    With XmlHttp
       .Open "POST", PostUrl, True
       .SetRequestHeader "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-shockwave-flash, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, application/x-silverlight, */*"
       .SetRequestHeader "Referer", PostUrl
       .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
       .SetRequestHeader "Accept-Encoding", "gzip, deflate"
       .SetRequestHeader "Content-Length", Len(PostData)
       .SetRequestHeader "Connection", "Keep-Alive"
       .SetRequestHeader "Cache-Control", "no-cache"
       .send (PostData)
        Do Until .readyState = 4
           DoEvents
        Loop
        PostXmlHttp = .ResponseText
    End With
    Set XmlHttp = Nothing
End Function

Function UnicodeToUtf8(ByVal sData As String) As String
    Dim aRetn() As Byte, nSize As Long, ReturnStr As String, X As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim lResult As Long
    Dim abUTF8() As Byte
    lLength = Len(sData)
    If lLength = 0 Then Exit Function
    lBufferSize = lLength * 3 + 1
    ReDim aRetn(lBufferSize - 1)
    nSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sData), lLength, aRetn(0), lBufferSize, vbNullString, 0)
    If nSize = 0 Then Exit Function
    ReDim Preserve aRetn(0 To nSize - 1) As Byte
    For X = LBound(aRetn) To UBound(aRetn)
      ReturnStr = ReturnStr & "%" & String(2 - Len(Hex(aRetn(X))), "0") & Hex(aRetn(X))
    Next X
    Erase aRetn
    UnicodeToUtf8 = ReturnStr
End Function

Function RemoveHeadTail(ByVal Source As Variant, ByVal sStart As String, ByVal strEnd As String) As String
    On Error Resume Next
    Dim m As Long
    Dim n As Long
    RemoveHeadTail = ""
    m = InStr(1, Source, sStart)
    If m <> 0 Then
        n = InStr(m + Len(sStart) + 1, Source, strEnd)
        If n <> 0 Then
            RemoveHeadTail = Mid(Source, m + Len(sStart), n - m - Len(sStart))
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
End Function

Function RemoveChr(ByVal Source As String, ByVal UsrType As String) As String
    If UsrType = "GET" Then
        If InStr(Source, "]]]") <> 0 Then
            Source = RemoveHeadTail(Source, "]],[[", "]]]")
        Else
            If FormTranslate.Combo2.ListIndex = 1 Then
                Source = RemoveHeadTail(Source, "[[[""", """,") & "."
            Else
                Source = RemoveHeadTail(Source, "[[[""", """,")
            End If
        End If
        Source = Replace(Source, "]],", vbCrLf)
        Source = Replace(Source, ",[", ":")
        Source = Replace(Source, "[", "")
        Source = Replace(Source, """", "")
    Else
        Do Until InStr(Source, "&amp;quot;") = 0
            Source = Replace(Source, "&amp;quot;", Chr(34))
        Loop
        Source = Replace(Source, "&lt;", "")
        Source = Replace(Source, "br&gt;", "")
    End If
    RemoveChr = Source
End Function


