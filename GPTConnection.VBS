Attribute VB_Name = "Module1"
Option Explicit

Private Const OPENAI_API_KEY As String = "apikey"
Public Const LOG_FILE_PATH As String = "C:\OutlookGPT\log.txt"



' ----------------------
' Logowanie
' ----------------------
Public Sub WriteLog(msg As String)
    On Error Resume Next
    Dim fso As Object, ts As Object, logDir As String
    logDir = "C:\OutlookGPT"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(logDir) Then fso.CreateFolder (logDir)
    Set ts = fso.OpenTextFile(LOG_FILE_PATH, 8, True)
    ts.WriteLine Format(Now, "yyyy-mm-dd HH:nn:ss") & " - " & msg
    ts.Close
End Sub

' ----------------------
' POST JSON do OpenAI
' ----------------------
Private Function PostJson(url As String, json As String) As String
    On Error GoTo ErrHandler
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Authorization", "Bearer " & OPENAI_API_KEY
    http.Send json

    PostJson = http.responseText
    Exit Function
ErrHandler:
    WriteLog "Error in PostJson: " & Err.Description
    PostJson = ""
End Function

' ----------------------
' Parsowanie content z odpowiedzi API
' ----------------------
Public Function ExtractTextFromResponse(ByVal response As String) As String
    Dim startPos As Long, endPos As Long
    Dim contentStart As Long, contentEnd As Long
    Dim result As String

    ' Znajd  pozycj  "content"
    startPos = InStr(1, response, """content"":")
    If startPos = 0 Then
        ExtractTextFromResponse = ""
        Exit Function
    End If

    ' Znajd  pierwszy cudzys  w po "content":
    contentStart = InStr(startPos + 10, response, """") + 1
    If contentStart = 0 Then
        ExtractTextFromResponse = ""
        Exit Function
    End If

    ' Znajd  ko cowy cudzys  w (nie escapeowany)
    contentEnd = contentStart
    Do
        contentEnd = InStr(contentEnd + 1, response, """")
        If contentEnd = 0 Then
            ExtractTextFromResponse = ""
            Exit Function
        End If
        If Mid(response, contentEnd - 1, 1) <> "\" Then Exit Do
    Loop

    result = Mid(response, contentStart, contentEnd - contentStart)

    ' Odzyskaj znaki specjalne
    result = Replace(result, "\""", """")
    result = Replace(result, "\\", "\")
    result = Replace(result, "\n", vbCrLf)

    ExtractTextFromResponse = Trim(result)
End Function



' ----------------------
' Wykrywanie j zyka
' ----------------------
Public Function DetectLanguage(text As String) As String
    On Error GoTo ErrHandler

    Dim safeText As String, json As String, resp As String
    safeText = Replace(text, """", "\""")
    safeText = Replace(safeText, vbCrLf, "\n")
    safeText = Replace(safeText, vbCr, "\n")
    safeText = Replace(safeText, vbLf, "\n")

    json = "{""model"":""gpt-4o-mini"",""messages"":[{""role"":""user"",""content"":""Return ONLY the ISO language code (e.g. en, pl, fr, de) of this text:\n" & safeText & """}]}"

    resp = PostJson("https://api.openai.com/v1/chat/completions", json)
    WriteLog "DetectLanguage response: " & Left(resp, 200)
    WriteLog "RAW DetectLanguage response: " & resp
    If InStr(resp, """error""") > 0 Or Len(resp) = 0 Then
        WriteLog "DetectLanguage API error detected"
        DetectLanguage = "other"
        Exit Function
    End If

    Dim code As String
    code = ExtractTextFromResponse(resp)
    If Len(code) = 0 Then
    WriteLog "DetectLanguage: empty response from API"
    DetectLanguage = "other"
Else
    DetectLanguage = LCase(Trim(code))
End If
    Exit Function

ErrHandler:
    WriteLog "Error in DetectLanguage: " & Err.Description
    DetectLanguage = "other"
End Function

' ----------------------
' T umaczenie wiadomo ci
' ----------------------
Public Function TranslateText(txt As String) As String
    On Error GoTo ErrHandler

    Dim safeText As String, json As String, resp As String
    safeText = Replace(txt, """", "\""")
    safeText = Replace(safeText, vbCrLf, "\n")
    safeText = Replace(safeText, vbCr, "\n")
    safeText = Replace(safeText, vbLf, "\n")

    json = "{""model"":""gpt-4o-mini"",""messages"":[{""role"":""user"",""content"":""Translate this text to Polish (pl) or English (en), preserve formatting:\n" & safeText & """}]}"

    resp = PostJson("https://api.openai.com/v1/chat/completions", json)
    WriteLog "TranslateText response: " & Left(resp, 200)

    TranslateText = ExtractTextFromResponse(resp)
    Exit Function

ErrHandler:
    WriteLog "Error in TranslateText: " & Err.Description
    TranslateText = ""
End Function
Public Function StripHTML(ByVal html As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .Pattern = "<[^>]+>"   ' dopasowuje wszystko pomi dzy < i >
        StripHTML = .Replace(html, "")
    End With
    
    ' Zamie  typowe encje HTML na znaki
    StripHTML = Replace(StripHTML, "&nbsp;", " ")
    StripHTML = Replace(StripHTML, "&amp;", "&")
    StripHTML = Replace(StripHTML, "&lt;", "<")
    StripHTML = Replace(StripHTML, "&gt;", ">")
    StripHTML = Replace(StripHTML, "&quot;", """")
    StripHTML = Replace(StripHTML, "&#39;", "'")
End Function
Sub TestLanguageDetection()
    Dim sample As String
    sample = "Ceci est un test en fran ais."
    Debug.Print "Detected: " & DetectLanguage(sample)
End Sub
