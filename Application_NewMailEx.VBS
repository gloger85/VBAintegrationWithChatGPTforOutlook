Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim arr As Variant, i As Long
    Dim item As Object
    Dim bodyToUse As String, lang As String, translated As String
    Dim userChoice As String
    Dim filePath As String
    Dim fNum As Integer

    On Error GoTo ErrHandler

    WriteLog "== NewMailEx triggered =="
    arr = Split(EntryIDCollection, ",")

    For i = LBound(arr) To UBound(arr)
        Set item = Application.Session.GetItemFromID(arr(i))
        If Not item Is Nothing Then
            If TypeOf item Is Outlook.MailItem Then
                WriteLog "Mail received: " & item.Subject & " From: " & item.SenderName

                ' --- bezpieczne pobranie treści wiadomości ---
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

                ' --- jeśli nie ma treści, pomijamy ---
                If Len(Trim(bodyToUse)) = 0 Then
                    WriteLog "No message body found — skipping."
                    GoTo NextItem
                End If

                ' --- detekcja języka ---
                lang = DetectLanguage(Left(bodyToUse, 1500))
                WriteLog "Detected language: " & lang

                ' --- warunek tłumaczenia ---
                If lang <> "pl" And lang <> "en" And lang <> "" And lang <> "other" Then
                    WriteLog "Starting translation..."
                    translated = TranslateText(bodyToUse)

                    If Len(translated) > 0 Then
                        ' === wybór działania użytkownika ===
                        userChoice = InputBox( _
                            "Wybierz co zrobić z tłumaczeniem:" & vbCrLf & _
                            "1 - Zapisz jako plik" & vbCrLf & _
                            "2 - Dołącz pod oryginałem" & vbCrLf & _
                            "3 - Wyświetl w oknie", _
                            "Opcje tłumaczenia", "3")

                        Select Case userChoice
                            Case "1"
                                ' Zapisz tłumaczenie jako plik tekstowy na pulpicie
                                filePath = Environ("USERPROFILE") & "\Desktop\" & _
                                           "Tlumaczenie_" & Replace(item.Subject, " ", "_") & ".txt"
                                fNum = FreeFile
                                Open filePath For Output As #fNum
                                Print #fNum, StripHTML(translated)
                                Close #fNum
                                MsgBox "Tłumaczenie zapisano jako: " & filePath, vbInformation
                                WriteLog "Translation saved to file: " & filePath

                            Case "2"
                                ' Dołącz tłumaczenie pod oryginałem
                                item.HTMLBody = item.HTMLBody & _
                                    "<hr><p><b>Tłumaczenie:</b></p>" & translated
                                item.Save
                                MsgBox "Tłumaczenie zostało dołączone do wiadomości.", vbInformation
                                WriteLog "Translation appended to message."

                            Case Else
                                ' Domyślnie wyświetl w okienku
                                MsgBox StripHTML(translated), vbInformation, "Tłumaczenie: " & item.Subject
                                WriteLog "Translation displayed."
                        End Select
                    Else
                        WriteLog "Translation empty."
                    End If
                Else
                    WriteLog "No translation required (pl/en/other)."
                End If
            End If
        End If
NextItem:
    Next i

    WriteLog "== Processing finished =="
    Exit Sub

ErrHandler:
    WriteLog "Error in Application_NewMailEx: " & Err.Description
End Sub


