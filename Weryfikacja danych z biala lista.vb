''
' Weryfikacja danych z biala lista v1
' (c) Szymon Bialka - s.bialka@outlook.com
'
' @authors:
' s.bialka@outlook.com
'
' @version 1.0.0
' @release date 2/24/2023
'
' Dokumentacja API rejestru WL:
' https://www.gov.pl/web/kas/api-wykazu-podatnikow-vat
' https://wl-test.mf.gov.pl/
' https://wl-api.mf.gov.pl/
''

Sub NipValidator()
    
    Sheets(1).Select
    Dim RequestDate As Variant
    RequestDate = Format(Range("C12").Value, "yyyy-mm-dd")

    Sheets(2).Select
    Range("B2:F10000").Clear
    
    Dim Index As Integer
    Index = 2
    
    Do
        Dim NIP As Variant
        NIP = Range("A" & Index).Value

        Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")
        Request.Open "GET", "https://wl-api.mf.gov.pl/api/search/nip/" & NIP & "?date=" & RequestDate
        Request.Send

        Dim Json As Object
        Set Json = JsonConverter.ParseJson(Request.ResponseText)
        
        If Json.Exists("code") Then
            Range("B" & Index).Value = "Napotkano blad - " & Json("code") & ": " & Json("message")
            GoTo ContinueDo
        End If
        
        If IsNull(Json("result")("subject")) Then
            Range("B" & Index).Value = "Nie figuruje w rejestrze VAT"
            Range("C" & Index).Value = "Brak"
            Range("D" & Index).Value = "00 00000 0000 0000 0000 0000 0000"
            Range("E" & Index).Value = RequestDate
            Range("F" & Index).Value = Json("result")("requestId")
            GoTo ContinueDo
        End If

        Range("B" & Index).Value = Json("result")("subject")("name")
        Range("C" & Index).Value = Json("result")("subject")("statusVat")
        
        Range("D" & Index).Value = ""
        
        Dim BankAccountsCount As Integer
        BankAccountsCount = Json("result")("subject")("accountNumbers").Count

        For BankAccountsCounter = 1 To BankAccountsCount
            Range("D" & Index).Value = Range("D" & Index).Value & Format(Json("result")("subject")("accountNumbers")(BankAccountsCounter), "## #### #### #### #### #### ####")

            If BankAccountsCounter < BankAccountsCount Then
                Range("D" & Index).Value = Range("D" & Index).Value & vbNewLine & vbNewLine
            End If
    
        Next BankAccountsCounter

        Range("E" & Index).Value = RequestDate
        Range("F" & Index).Value = Json("result")("requestId")

        Rows(Index & ":" & Index).EntireRow.AutoFit

ContinueDo:
        Index = Index + 1
    Loop While Range("A" & Index).Value <> ""
    
    MsgBox "Zakonczono sprawdzanie"
End Sub
