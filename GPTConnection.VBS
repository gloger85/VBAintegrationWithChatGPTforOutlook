Option Explicit

Private Const OPENAI_API_KEY As String = "APIkey"
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
Sub OtworzZaznaczonyMailDoEdycji()

    
    Dim objMail As MailItem
    Dim objReply As MailItem

    ' Sprawdzenie, czy jest zaznaczony jakiś element
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Proszę zaznaczyć wiadomość e-mail.", vbExclamation
        Exit Sub
    End If

    ' Pobranie pierwszego zaznaczonego maila
    Set objMail = Application.ActiveExplorer.Selection(1)

    ' Sprawdzenie, czy zaznaczony element jest wiadomością e-mail
    If TypeName(objMail) <> "MailItem" Then
        MsgBox "Zaznaczony element nie jest wiadomością e-mail.", vbExclamation
        Exit Sub
    End If

    ' Utworzenie odpowiedzi
    Set objReply = objMail.Reply

    ' Wklejenie tekstu na początku wiadomości
    objReply.HTMLBody = "Hi,<br><br>" & _
"ich freue mich Ihnen mitteilen zu können, dass Ihr AllData-Konto freigeschaltet worden ist. Unten finden Sie Ihre Anmeldedaten:<br><br>" & _
"Benutzername:<br>" & _
"Password:                         changeme (bei erster Anmeldung wird aufgefordert, es auf ein eigenes Passwort zu ändern)<br><br>" & _
"Bitte prüfen und kurz bestätigen, ob alles zugreifbar ist. Danke!<br><br>" & _
"Am 1. Februar 2024 wurden die Alldata-Passwörter nach den neuen, globalen Passwortvorgaben geändert.<br><br>" & _
"Folgende Parameter ändern sich im Vergleich zur bisherigen Passwortvorgabe:<br><br>" & _
"- Änderung der Passwortlänge von 8 auf 12 Ziffern<br>" & _
"- Keine Verwendung der letzten 12 Passwörter <br><br>" & _
"Ablauf:<br>" & _
"- Am 31. Januar ab 20:00 Uhr wurden alle Alldata-Passwörter von der IT-Abteilung auf 'abgelaufen' gesetzt.<br>" & _
"- Wenn Alldata am 1. Februar startet, werden Sie aufgefordert, ein neues Passwort einzugeben.<br>" & _
"- Bitte beachten Sie bei der Vergabe, dass das neue Passwort 12-stellig und alphanumerisch sein muss<br>" & _
"- Gültige Zeichen für das Passwort sind: A-Z, Groß-/Kleinbuchstaben, 0-9 und die Sonderzeichen #, $, der Unterstrich (_) und §<br><br>" & _
"Viele Grüsse/Best regards/Pozdrawiam <br><br> " & objReply.HTMLBody

    ' Otwarcie okna odpowiedzi
    objReply.Display


End Sub
Sub TranslateEmail()

    Dim item As Outlook.MailItem
    Dim bodyToUse As String, lang As String, translated As String
    Dim userChoice As VbMsgBoxResult
    Dim filePath As String, fNum As Integer
    Dim sel As Selection
    Dim ns As Outlook.NameSpace
    Dim draftFolder As Outlook.Folder
    Dim i As Long

    On Error GoTo ErrHandler

    WriteLog "== Manual processing triggered =="

    Set sel = Application.ActiveExplorer.Selection
    Set ns = Application.GetNamespace("MAPI")
    Set draftFolder = ns.GetDefaultFolder(olFolderDrafts)

    ' --- Spróbuj pobrać zaznaczone maile ---
    If sel.Count > 0 Then
        For i = 1 To sel.Count
            If TypeName(sel.item(i)) = "MailItem" Then
                Set item = sel.item(i)
                Exit For
            End If
        Next i
    End If

    ' --- Jeśli brak zaznaczenia, pobierz pierwszy mail z Drafts ---
    If item Is Nothing Then
        If draftFolder.Items.Count = 0 Then
            MsgBox "Brak maili do przetworzenia (ani w zaznaczeniu, ani w Drafts).", vbExclamation
            Exit Sub
        End If

        ' Znajdź pierwszy element typu MailItem
        For i = 1 To draftFolder.Items.Count
            If TypeName(draftFolder.Items(i)) = "MailItem" Then
                Set item = draftFolder.Items(i)
                Exit For
            End If
        Next i

        If item Is Nothing Then
            MsgBox "Brak maili typu MailItem w folderze Drafts.", vbExclamation
            Exit Sub
        End If
    End If

    WriteLog "Mail selected: " & item.Subject & " From: " & item.SenderName

    ' --- Pobranie treści wiadomości ---
    If Len(item.HTMLBody) > 0 Then
        bodyToUse = item.HTMLBody
    ElseIf Len(item.Body) > 0 Then
        bodyToUse = item.Body
    ElseIf Len(item.RTFBody) > 0 Then
        bodyToUse = item.RTFBody
    Else
        bodyToUse = ""
    End If

    WriteLog "Body length: " & Len(bodyToUse)
    WriteLog "Body preview: " & Left(StripHTML(bodyToUse), 200)

    If Len(Trim(bodyToUse)) = 0 Then
        MsgBox "Brak treści wiadomości — przerwano.", vbExclamation
        WriteLog "Empty body — abort."
        Exit Sub
    End If

    ' --- Detekcja języka ---
    lang = DetectLanguage(Left(bodyToUse, 1500))
    WriteLog "Detected language: " & lang

    ' --- Tłumaczenie jeśli wymagane ---
    If lang <> "pl" And lang <> "en" And lang <> "" And lang <> "other" Then
        WriteLog "Starting translation..."
        translated = TranslateText(bodyToUse)

        If Len(translated) > 0 Then

            userChoice = MsgBox("Wybierz działanie dla tłumaczenia:" & vbCrLf & _
                                "Yes - Zapisz jako plik" & vbCrLf & _
                                "No - Dołącz pod oryginałem" & vbCrLf & _
                                "Cancel - Wyświetl w oknie", _
                                vbYesNoCancel + vbQuestion, "Opcje tłumaczenia")

            Select Case userChoice
                Case vbYes
                    filePath = Environ("USERPROFILE") & "\Desktop\" & _
                               "Tlumaczenie_" & Replace(item.Subject, " ", "_") & ".txt"
                    fNum = FreeFile
                    Open filePath For Output As #fNum
                    Print #fNum, StripHTML(translated)
                    Close #fNum
                    MsgBox "Tłumaczenie zapisano: " & filePath, vbInformation
                    WriteLog "Translation saved to file: " & filePath

                Case vbNo
                    item.HTMLBody = item.HTMLBody & _
                        "<hr><p><b>Tłumaczenie:</b></p>" & translated
                    item.Save
                    MsgBox "Tłumaczenie dołączone do wiadomości.", vbInformation
                    WriteLog "Translation appended."

                Case vbCancel
                    MsgBox StripHTML(translated), vbInformation, _
                           "Tłumaczenie: " & item.Subject
                    WriteLog "Translation displayed."
            End Select

        Else
            WriteLog "Translation empty."
        End If
    Else
        WriteLog "No translation required (pl/en/other)."
    End If

    WriteLog "== Manual processing finished =="
    Exit Sub

ErrHandler:
    MsgBox "Błąd: " & Err.Description, vbCritical
    WriteLog "Error: " & Err.Description
End Sub
















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

Sub test()
MsgBox "działa!"
End Sub


' ----------------------
' EscapeJSON - poprawna obsługa HTML i Plain Text
' ----------------------
Private Function EscapeJSON(ByVal str As String, Optional ByVal isHtml As Boolean = False) As String
    If Len(str) = 0 Then
        EscapeJSON = ""
        Exit Function
    End If

    str = Replace(str, "\", "\\")
    str = Replace(str, """", "\""")
    
    If Not isHtml Then
        ' Tylko dla zwykłego tekstu zamieniamy nowe linie na \n
        str = Replace(str, vbCrLf, "\n")
        str = Replace(str, vbCr, "\n")
        str = Replace(str, vbLf, "\n")
    End If
    
    ' usuń kontrolne znaki poniżej 32
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(str)
        ch = Mid$(str, i, 1)
        If Asc(ch) >= 32 Then out = out & ch
    Next i
    EscapeJSON = out
End Function

' ----------------------
' Clean HTML - usuwa śmieci z Worda
' ----------------------
Private Function CleanHtmlForTranslation(ByVal html As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True

    ' Usuń sekcje <style>...</style>
    regex.Pattern = "<style[\s\S]*?</style>"
    html = regex.Replace(html, "")

    ' Usuń <head>...</head>
    regex.Pattern = "<head[\s\S]*?</head>"
    html = regex.Replace(html, "")

    ' Usuń tagi xmlns="..."
    regex.Pattern = "xmlns(:\w+)?=""[^""]*"""
    html = regex.Replace(html, "")

    ' Usuń inline mso-style
    regex.Pattern = "mso-[^:]+:[^;""']+;?"
    html = regex.Replace(html, "")

    ' Usuń puste <o:p>
    html = Replace(html, "<o:p>&nbsp;</o:p>", "")
    html = Replace(html, "<o:p></o:p>", "")

    CleanHtmlForTranslation = Trim(html)
End Function

' ----------------------
' TranslateText - tłumaczenie z obsługą HTML i Plain Text
' ----------------------
Public Function TranslateText(ByVal sourceText As String) As String
    On Error GoTo ErrHandler
    If Len(Trim(sourceText)) = 0 Then
        TranslateText = ""
        Exit Function
    End If

    Dim sample As String
    sample = Left(sourceText, 4000) ' fragment do tłumaczenia

    ' --- wyczyść HTML Worda / MS Office ---
    sample = CleanHtmlForTranslation(sample)

    Dim safeText As String
    Dim wantHtml As Boolean
    wantHtml = (InStr(LCase(sample), "<p>") > 0 Or _
                InStr(LCase(sample), "<html") > 0 Or _
                InStr(LCase(sample), "<div") > 0)

    safeText = EscapeJSON(sample, wantHtml)

    Dim prompt As String
    If wantHtml Then
        prompt = "Translate the following HTML text into Polish. " & _
                 "Preserve the HTML structure but translate only visible text content. " & _
                 "Return only clean HTML, without Markdown formatting, triple backticks, or explanations." & vbCrLf & sample
    Else
        prompt = "Translate the following text into Polish. Preserve formatting and line breaks. " & _
                 "Do NOT wrap the result in Markdown or code blocks." & vbCrLf & sample
    End If

    Dim json As String
    json = "{""model"":""gpt-4o-mini"",""messages"":[{""role"":""user"",""content"":""" & EscapeJSON(prompt, False) & """}],""max_tokens"":3500}"

    Dim resp As String
    resp = PostJson("https://api.openai.com/v1/chat/completions", json)
    WriteLog "TranslateText response (head): " & Left(resp, 300)

    If Len(resp) = 0 Or InStr(resp, """error""") > 0 Then
        TranslateText = ""
        Exit Function
    End If

    Dim outText As String
    outText = ExtractTextFromResponse(resp)
    ' Usuń literalne \n w HTML (bo nie są potrzebne)
If wantHtml Then
    outText = Replace(outText, "\n", "")
Else
    ' dla zwykłego tekstu zamień \n na prawdziwe łamanie linii
    outText = Replace(outText, "\n", vbCrLf)
End If

    ' ?? usuń znaczniki Markdown typu ```html ... ```
    outText = Replace(outText, "```html", "")
    outText = Replace(outText, "```HTML", "")
    outText = Replace(outText, "```", "")
    outText = Trim(outText)

    ' Zamień \n na faktyczne nowe linie, tylko jeśli to nie HTML
    If Not wantHtml Then
        outText = Replace(outText, "\n", vbCrLf)
    End If

    TranslateText = outText
    Exit Function

ErrHandler:
    WriteLog "Error in TranslateText: " & Err.Description
    TranslateText = ""
End Function



' ----------------------
' POST JSON do OpenAI - z logowaniem JSON
' ----------------------
Private Function PostJson(url As String, json As String) As String
    On Error GoTo ErrHandler
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & OPENAI_API_KEY
    WriteLog "JSON sent: " & Left(json, 2000)
    http.Send json
    PostJson = http.responseText
    WriteLog "RAW response length: " & Len(PostJson)
    Exit Function
ErrHandler:
    WriteLog "Error in PostJson: " & Err.Description
    PostJson = ""
End Function

' ----------------------
' Bezpieczne wyciąganie content z odpowiedzi (robust)
' ----------------------
Public Function ExtractTextFromResponse(ByVal response As String) As String
    On Error GoTo ErrHandler
    Dim s As String
    s = response
    ' Znajdź "message": { ... "content": "...." } - najpierw spróbuj najczęstszego wzorca
    Dim pos As Long, posContent As Long, posStart As Long, posEnd As Long
    pos = InStr(1, s, """message"":")
    If pos = 0 Then pos = InStr(1, s, """content"":")
    If pos = 0 Then
        ExtractTextFromResponse = ""
        Exit Function
    End If
    posContent = InStr(pos, s, """content"":")
    If posContent = 0 Then
        ExtractTextFromResponse = ""
        Exit Function
    End If
    ' znajdź pierwszy nie-escaped quote po "content":
    posStart = InStr(posContent + 10, s, """")
    If posStart = 0 Then
        ExtractTextFromResponse = ""
        Exit Function
    End If
    posStart = posStart + 1
    posEnd = posStart
    Do
        posEnd = InStr(posEnd + 1, s, """")
        If posEnd = 0 Then
            ExtractTextFromResponse = ""
            Exit Function
        End If
        If Mid$(s, posEnd - 1, 1) <> "\" Then Exit Do
    Loop
    Dim result As String
    result = Mid$(s, posStart, posEnd - posStart)
    ' od-escape
    result = Replace(result, "\\n", vbCrLf)
    result = Replace(result, "\""", """")
    result = Replace(result, "\\", "\")
    result = Trim(result)
    ExtractTextFromResponse = result
    Exit Function
ErrHandler:
    WriteLog "Error in ExtractTextFromResponse: " & Err.Description
    ExtractTextFromResponse = ""
End Function

' ----------------------
' Backup: proste heurystyki językowe (jeśli model zawiedzie)
' ----------------------
Private Function HeuristicLanguageDetect(ByVal txt As String) As String
    Dim t As String
    t = LCase(Left(txt, 1000))
    If InStr(t, " le ") > 0 Or InStr(t, " la ") > 0 Or InStr(t, "bonjour") > 0 Or InStr(t, "merci") > 0 Then
        HeuristicLanguageDetect = "fr"
        Exit Function
    End If
    If InStr(t, " die ") > 0 Or InStr(t, " und ") > 0 Or InStr(t, "danke") > 0 Then
        HeuristicLanguageDetect = "de"
        Exit Function
    End If
    If InStr(t, " the ") > 0 Or InStr(t, " and ") > 0 Or InStr(t, "thank") > 0 Then
        HeuristicLanguageDetect = "en"
        Exit Function
    End If
    If InStr(t, " i ") > 0 Or InStr(t, " z ") > 0 Or InStr(t, " dzię") > 0 Or InStr(t, " proszę") > 0 Then
        HeuristicLanguageDetect = "pl"
        Exit Function
    End If
    HeuristicLanguageDetect = "other"
End Function

' ----------------------
' DetectLanguage (nowa, odporna wersja)
' ----------------------
Public Function DetectLanguage(ByVal text As String) As String
    On Error GoTo ErrHandler
    If Len(Trim(text)) = 0 Then
        WriteLog "DetectLanguage: empty input"
        DetectLanguage = "other"
        Exit Function
    End If

    Dim sample As String
    sample = Left(text, 1200) ' skróć próbkę do 1200 znaków
    ' przygotuj bezpieczny escaped tekst
    Dim safeText As String
    safeText = EscapeJSON(sample)

    ' Zbuduj JSON z rolą system i user, jasne instrukcje i brak pytań
    Dim json As String
    json = "{""model"":""gpt-4o-mini"",""messages"":["
    json = json & "{""role"":""system"",""content"":""You are a language detection assistant. Respond with ONLY a two-letter ISO 639-1 language code (like: en, pl, fr, de). DO NOT ask for clarification or produce any other text.""},"
    json = json & "{""role"":""user"",""content"":""Detect the language of the following text and respond ONLY with its two-letter ISO code. Text:\n" & safeText & """}"
    json = json & "],""max_tokens"":5}"

    Dim resp As String
    resp = PostJson("https://api.openai.com/v1/chat/completions", json)
    WriteLog "DetectLanguage response (head): " & Left(resp, 300)

    If Len(resp) = 0 Or InStr(resp, """error""") > 0 Then
        WriteLog "DetectLanguage API error or empty response"
        DetectLanguage = HeuristicLanguageDetect(text)
        Exit Function
    End If

    Dim code As String
    code = ExtractTextFromResponse(resp)
    If Len(code) = 0 Then
        WriteLog "DetectLanguage: empty content from model -> fallback heuristic"
        DetectLanguage = HeuristicLanguageDetect(text)
        Exit Function
    End If

    ' spróbuj normalizować i wyciągnąć sam kod (regexp-like)
    code = LCase(Trim(code))
    ' jeśli model odpisał całe zdanie, spróbuj znaleźć dwuliterowy token
    Dim candidate As String
    Dim i As Long
    For i = 1 To Len(code) - 1
        candidate = Mid$(code, i, 2)
        If candidate Like "[a-z][a-z]" Then
            DetectLanguage = candidate
            Exit Function
        End If
    Next i

    ' jeżeli nie znaleziono, sprawdź pełne słowa
    If InStr(code, "polish") > 0 Or InStr(code, "pl") = 1 Then
        DetectLanguage = "pl": Exit Function
    ElseIf InStr(code, "english") > 0 Or InStr(code, "en") = 1 Then
        DetectLanguage = "en": Exit Function
    ElseIf InStr(code, "french") > 0 Or InStr(code, "fr") = 1 Then
        DetectLanguage = "fr": Exit Function
    End If

    ' fallback heurystyka
    DetectLanguage = HeuristicLanguageDetect(text)
    Exit Function

ErrHandler:
    WriteLog "Error in DetectLanguage: " & Err.Description
    DetectLanguage = HeuristicLanguageDetect(text)
End Function




