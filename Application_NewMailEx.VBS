VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim arr As Variant, i As Long
    Dim item As Object
    Dim bodyToUse As String, lang As String, translated As String

    On Error GoTo ErrHandler

    WriteLog "== NewMailEx triggered =="
    arr = Split(EntryIDCollection, ",")

    For i = LBound(arr) To UBound(arr)
        Set item = Application.Session.GetItemFromID(arr(i))
        If Not item Is Nothing Then
            If TypeOf item Is Outlook.MailItem Then
                WriteLog "Mail received: " & item.Subject & " From: " & item.SenderName

                If Len(item.HTMLBody) > 0 Then
                    bodyToUse = item.HTMLBody
                Else
                    bodyToUse = item.Body
                End If

                lang = DetectLanguage(Left(bodyToUse, 1500)) ' tylko fragment do detekcji
                WriteLog "Detected language: " & lang

                If lang <> "pl" And lang <> "en" And lang <> "" Then
                    WriteLog "Starting translation..."
                    translated = TranslateText(bodyToUse)
                    If Len(translated) > 0 Then
                        MsgBox StripHTML(translated), vbInformation, "T³umaczenie: " & item.Subject
                        WriteLog "Translation displayed."
                    Else
                        WriteLog "Translation empty."
                    End If
                Else
                    WriteLog "No translation required (pl/en)."
                End If
            End If
        End If
    Next i

    WriteLog "== Processing finished =="
    Exit Sub

ErrHandler:
    WriteLog "Error in Application_NewMailEx: " & Err.Description
End Sub

