Attribute VB_Name = "Module1"
Public Const TEMPLATE As String = "Rek"
Public Const TOTALS As String = "Totaal"

' --- Copy template file and change sum formulas in 'Totaal' sheet
Sub copyTemplate()
Attribute copyTemplate.VB_ProcData.VB_Invoke_Func = " \n14"

Dim i, wantedSheets, totalSheets As Integer
Dim ws, wsNew As Worksheet

Dim line, firstLine, lastLine As Integer
Dim myFormula, myRange As String
Dim answer As Integer

' Totaal aantal bladen in het bestand
' Normaal gezien = 2, nl. 'Totaal' en TEMPLATE blad
totalSheets = countBills()

' Check bestaan van eventuele afrekeningen (reeds op start gedrukt)
' UPDATE: Start knop wordt op disabled gezet, dus overbodig...

If totalSheets > 0 Then
    answer = MsgBox("Wenst U met nieuwe gegevens aan de slag wenst te gaan?" & vbCrLf & "ALLES WORDT GEWIST!", vbQuestion + vbYesNo + vbDefaultButton2, "Opnieuw starten")

    If answer = vbNo Then
        Exit Sub
    Else
        ' Wis bestaande sheets and wis formules in Totaal berekeningen
        deleteBills
    End If
End If


' Empty template sheet
clearTemplate

Set wsTot = ThisWorkbook.Sheets(TOTALS)

' Verwachte klanten (wordt uit blad 'Totaal' gehaald
wantedSheets = wsTot.Range("B5")

If Not IsNumeric(wantedSheets) Then
    MsgBox "Het 'Gewenst aantal bladen' moet een getal zijn!", vbOKOnly, "Aantal bladen"
    wsTot.Range("B5").Value = 0
    wsTot.Range("B5:J5").Activate
    Exit Sub
End If

If wantedSheets = "" Or wantedSheets <= 0 Then
    MsgBox "Het 'Gewenst aantal bladen' moet groter zijn dan 0!", vbOKOnly, "Aantal bladen"
    wsTot.Range("B5:J5").Activate
    Exit Sub
End If


' Maak gewenst aantal kopijen
For i = 1 To wantedSheets

    totalSheets = ThisWorkbook.Sheets.count - 1
    
    Set ws = Sheets(TEMPLATE)
    ws.Copy After:=Sheets(totalSheets)
    
    Set wsNew = Sheets(Sheets(totalSheets).Index + 1)
    wsNew.Name = TEMPLATE & " " & i
    wsNew.Visible = xlSheetVisible
    wsNew.Activate
    
    If wsNew.Range("N4:P5").Locked = True Then
        wsNew.Unprotect
        wsNew.Range("N4:P5").Locked = False
        wsNew.Range("N4:P5").Value = "# " & i
        wsNew.Range("N4:P5").Locked = True
        wsNew.Protect
    End If
        
Next i

updateTotaal

' Reset 'Gewenst aantal bladen'
wsTot.Activate

' Reset betaald en geregistreerd veld
Range("V3") = 0
Range("V4") = 0

' Wis range met onbetaalde rekeningen
clearUnpaidRange

' Lock Naam, Jaar Activiteit en aantal bladen
wsTot.Activate

wsTot.Unprotect

If wsTot.Range("B3:J3").Locked = False Then
    wsTot.Range("B3:J3").Locked = True
End If
If wsTot.Range("B4:J4").Locked = False Then
    wsTot.Range("B4:J4").Locked = True
End If
If wsTot.Range("B5:J5").Locked = False Then
    wsTot.Range("B5:J5").Locked = True
End If

disableStartBtn

wsTot.Protect

Range("P3").Select
           
End Sub

' --- Remove all bills & clear template file + remove formulas in 'Totaal' sheet
Sub reset()

Dim totalSheets, clientSheetCount As Integer
Dim ws As Worksheet

Dim answer As Integer

answer = MsgBox("Wens je alle bladen te verwijderen?" & vbCrLf & "ALLES WORDT GEWIST!", vbQuestion + vbYesNo + vbDefaultButton2, "Bevestig verwijderen")

Set wsTot = Worksheets(TOTALS)
    
If answer = vbYes Then
    ' Unlock Naam en Jaar Activiteit
    wsTot.Activate
    
    ' Wis lijst onbetaalde rekeningen
    clearUnpaidRange
    ' Verwijdert alle bills
    deleteBills
    
    ' Unprotect Activiteit, Jaar en aantal bladen
    If wsTot.Range("B3:J3").Locked = True Then
        wsTot.Unprotect
        wsTot.Range("B3:J3").Locked = False
    End If
    If wsTot.Range("B4:J4").Locked = True Then
        wsTot.Unprotect
        wsTot.Range("B4:J4").Locked = False
    End If
    If wsTot.Range("B5:J5").Locked = True Then
        wsTot.Unprotect
        wsTot.Range("B5:J5").Locked = False
    End If
    
    ' Reset aantal bladen
    Range("B5").Value = 0
    
    ' Reset betaald en geregistreerd veld
    Range("V3").Value = 0
    Range("V4").Value = 0

    enableStartBtn
End If

wsTot.Protect
wsTot.Range("B3").Select

End Sub

' --- Add extra pages
Sub extra()

Dim extra As Variant
Dim i As Integer

' Empty template sheet
' clearTemplate

extra = InputBox("Hoeveel extra bladen wil je toevoegen?")

If extra > 0 And extra <> "" Then
    
    'Hoeveel bills zijn er?
    currentBills = countBills()
    
    
    For i = 1 To extra
    
        totalSheets = ThisWorkbook.Sheets.count - 1
            
        Set ws = Sheets(TEMPLATE)
        ws.Copy After:=Sheets(totalSheets)
    
        Set wsNew = Sheets(Sheets(totalSheets).Index + 1)
        myIndex = currentBills + i
        
        wsNew.Name = TEMPLATE & " " & myIndex
        wsNew.Visible = xlSheetVisible
        If wsNew.Range("N4:P5").Locked = True Then
            wsNew.Unprotect
            wsNew.Range("N4:P5").Locked = False
            wsNew.Range("N4:P5").Value = "# " & myIndex
            wsNew.Range("N4:P5").Locked = True
            wsNew.Protect
        End If
        
    Next i
    
    ' Set "Gewenst aantal bladen"
    Set wsTot = ThisWorkbook.Sheets(TOTALS)
    wsTot.Activate
    wsTot.Unprotect
    
    ' Unprotect Activiteit, Jaar en aantal bladen
    If wsTot.Range("B3:J3").Locked = False Then
        wsTot.Unprotect
        wsTot.Range("B3:J3").Locked = True
    End If
    If wsTot.Range("B4:J4").Locked = False Then
        wsTot.Unprotect
        wsTot.Range("B4:J4").Locked = True
    End If
    If wsTot.Range("B5:J5").Locked = True Then
        wsTot.Range("B5:J5").Locked = False
    End If
    wsTot.Range("B5").Value = countBills()
            
    wsTot.Range("B5:J5").Locked = True
    
    wsTot.Protect
    
    updateTotaal
    
End If

Worksheets(TOTALS).Activate
Range("B3").Select
    
End Sub

' --- Clear content of TEMPLATE sheet
Sub clearTemplate()

Worksheets(TEMPLATE).Range("k10:m36").ClearContents
Worksheets(TEMPLATE).Visible = xlSheetHidden

End Sub

' --- Change formulas in TOTALS sheet
Sub updateTotaal()

firstLine = 10
lastLine = 35

clientSheetCount = countBills()

For line = firstLine To lastLine
    
    myRange = "J" & line
    myFormula = "=SUM('" & TEMPLATE & " 1:" & TEMPLATE & " " & clientSheetCount
    myFormula = myFormula & "'!K" & line
    myFormula = myFormula & ")"
    
    Worksheets(TOTALS).Range(myRange).Formula = myFormula
    
Next line

End Sub

' --- Clear formulas in TOTALS sheet
Sub clearTotaal()

firstLine = 10
lastLine = 35

For line = firstLine To lastLine

    myRange = "J" & line
    Worksheets(TOTALS).Range(myRange).Value = ""
    
Next line

End Sub

' --- Find all copies ('bills') in a workbook based on TEMPLATE name and delete them
Sub deleteBills()
    Dim ws As Worksheet
    Dim counter, tNameLen As Integer
    
    Application.DisplayAlerts = False
    
    counter = 0
    
    For Each ws In ActiveWorkbook.Worksheets
        If isBill(ws) Then
            ws.Delete
            counter = counter + 1
        End If
    Next ws
    
    ' Wis formules in 'Totaal' gebaseed op de verwijderde bladen
    clearTotaal
End Sub

' Telt het aantal rekenbladen (beginnende met 'Afrekening XX')

Private Function countBills()

    cnt = 0
    
    For Each ws In ActiveWorkbook.Worksheets
        If isBill(ws) Then
            cnt = cnt + 1
        End If
    Next ws
    
    countBills = cnt
    
End Function


' --- Navigate to page
Sub NavToSheet()

sheetNumber = Range("P3").Value

' Check existence
If checkSheetName(sheetNumber) Then
    ' sheet with given number exists, clear lookup cell
    Range("P3").Value = ""
    Sheets(TEMPLATE & " " & sheetNumber).Activate
    Range("D4").Select
Else
    MsgBox ("Blad '" & TEMPLATE & " " & sheetNumber & "' bestaat niet!")
    Range("P3").Value = ""
End If

End Sub

' --- Navigate to TOTALS sheet
Sub NavToTotaal()
    ' Schrijf tijd weg
    Range("L47").Value = Now
    
    ' Set Tab color
    If Range("D4").Value <> "" Then
        If Range("G43").Value <> "" And IsNumeric(Range("G43").Value) Then
            ActiveSheet.Tab.Color = RGB(169, 208, 142) 'Green
        Else
            ActiveSheet.Tab.Color = RGB(237, 125, 49) 'Orange
        End If
    Else
        ActiveSheet.Tab.Color = xlNone
    End If
    
    Worksheets(TOTALS).Activate
    
    ' Check unpaid
    clearUnpaidRange
    ' Pas totaal registraties aan
    countRegistrations
    
    ' Pas totaal betaald aan
    countPayments
    
    Range("P3").Select
End Sub

Private Function checkSheetName(sheetNr)

checkSheetName = False

For Each sht In ThisWorkbook.Worksheets

    If sht.Name = TEMPLATE & " " & sheetNr Then
            checkSheetName = True
            Exit Function
    End If

Next sht

checkSheetName = False

End Function


' Controleert hoeveel bladen toegewezen zijn (veld 'Naam' - cel D4 wordt gecheckt)
Sub countRegistrations()
    
    Set wsTot = Sheets(TOTALS)
    
    checkCell = "D4"
    destCell = "V3"
    
    cnt = 0
    
    For Each ws In ActiveWorkbook.Worksheets
        If isBill(ws) Then
            If Len(ws.Range(checkCell).Value) > 0 Then
                cnt = cnt + 1
            End If
        End If
    Next ws
    
    ' Set total registration on TOTAAL
    wsTot.Range(destCell).Value = cnt
    
End Sub

' Controleert hoeveel bladen al afgerekend zijn (veld 'Ontvangen' - cel G43 wordt gecheckt)
Sub countPayments()

    Set wsTot = Sheets(TOTALS)
    
    checkCell = "G43"
    destCell = "V4"
    
    cnt = 0
    
    For Each ws In ActiveWorkbook.Worksheets
        If isBill(ws) Then
            If Len(ws.Range(checkCell).Value) > 0 Then
                cnt = cnt + 1
            End If
        End If
    Next ws
    
    ' Set total paid on TOTAAL
    wsTot.Range(destCell).Value = cnt
    
End Sub

' Check geregistreerd maar niet afgerekend
Sub showUnpaid()

    checkRegisteredCell = "D4"
    checkPaidCell = "G43"
    
    uPaidRow = 10
    uPaidCol = 17
    
    ' Clear previous stats
    clearUnpaidRange
    
    
    For Each ws In ActiveWorkbook.Worksheets
        
        If isBill(ws) Then
        
            If Len(ws.Range(checkRegisteredCell).Value) > 0 And Len(ws.Range(checkPaidCell).Value) = 0 Then
        
                clientName = ws.Range(checkRegisteredCell).Value
               
                ' activeer totaal en ga naar Q10
                Set wsTotal = Sheets(TOTALS)
                wsTotal.Activate
                wsTotal.Unprotect
                
                If Range(Cells(uPaidRow, uPaidCol), Cells(uPaidRow, uPaidCol + 3)).Locked = True Then
                
                    Range(Cells(uPaidRow, uPaidCol), Cells(uPaidRow, uPaidCol + 3)).Locked = False
                  
                    ' specify the range on the summary sheet to link to
                    Set targetRange = ws.Range("B1")
                
                    ' specify where on that sheet we'll create the hyperlink
                    Set linkRange = wsTotal.Cells(uPaidRow, uPaidCol)

                    wsTotal.Hyperlinks.Add Anchor:=linkRange, Address:="", SubAddress:= _
                        "'" & ws.Name & "'!" & targetRange.Address, _
                        TextToDisplay:=ws.Name
                    
                End If
                
                wsTotal.Activate
                If Range(Cells(uPaidRow, uPaidCol + 4), Cells(uPaidRow, uPaidCol + 9)).Locked = True Then
                     Range(Cells(uPaidRow, uPaidCol + 4), Cells(uPaidRow, uPaidCol + 9)).Locked = False
                     Cells(uPaidRow, uPaidCol + 4).Value = clientName
                End If
                
                formatUnpaidCells (uPaidRow)
                
                wsTotal.Activate
                If Range(Cells(uPaidRow, uPaidCol), Cells(uPaidRow, uPaidCol + 3)).Locked = False Then
                    Range(Cells(uPaidRow, uPaidCol), Cells(uPaidRow, uPaidCol + 3)).Locked = True
                End If
                
                If Range(Cells(uPaidRow, uPaidCol + 4), Cells(uPaidRow, uPaidCol + 9)).Locked = False Then
                    Range(Cells(uPaidRow, uPaidCol + 4), Cells(uPaidRow, uPaidCol + 9)).Locked = True
                End If
                
                uPaidRow = uPaidRow + 1
                
                wsTotal.Protect
                
            End If
            
        End If
    
    Next ws
    

End Sub

' Functie checkt of een gegeven werkblad start met de TEMPLATE naam gevolgd door XXX (een kopij van de tamplate)
Private Function isBill(ws)

    check = False
    tNameLen = Len(TEMPLATE)
    
    If Left(ws.Name, tNameLen) = TEMPLATE And Len(ws.Name) > tNameLen Then
        check = True
    End If
    
    isBill = check
End Function

Private Function formatUnpaidCells(row)
    
    uPaidCol1 = 17   ' Link column
    uPaidOffset1 = 3 ' Client Name column
    
    uPaidCol2 = 21
    uPaidOffset2 = 5
    
    Set ws = Sheets(TOTALS)
    
    ws.Range(Cells(row, uPaidCol1), Cells(row, uPaidCol1 + uPaidOffset1)).Select
      
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Strikethrough = False
        .Font.Superscript = False
        .Font.Subscript = False
        .Font.OutlineFont = False
        .Font.Shadow = False
        .Font.Underline = xlUnderlineStyleSingle
        .Font.ThemeColor = xlThemeColorHyperlink
        .Font.TintAndShade = 0
        .Font.ThemeFont = xlThemeFontMinor
    End With
    
    
    ws.Range(Cells(row, uPaidCol2), Cells(row, uPaidCol2 + uPaidOffset2)).Select

    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Strikethrough = False
        .Font.Superscript = False
        .Font.Subscript = False
        .Font.OutlineFont = False
        .Font.Shadow = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
        .Font.ThemeFont = xlThemeFontMinor
    End With
    
    Cells(row, uPaidCol2 - 1).Activate
    
End Function

' --- Wist lijst onbetaalde bladen op TOTALS

Sub clearUnpaidRange()
    Set ws = Sheets(TOTALS)
    ws.Activate
    ws.Unprotect
    
    uPaidRow = 10
    uPaidCol = 17
    
    row = uPaidRow
    
    Do Until IsEmpty(Cells(row, uPaidCol))
        Range(Cells(row, uPaidCol), Cells(row, uPaidCol + 8)).Select
        Selection.Delete Shift:=xlUp
        Cells(row, uPaidCol).Select
    Loop
    
    Range("P3").Select
    
    ws.Protect
    
End Sub

' --- Disable START button

Sub disableStartBtn()

Set ws = Sheets(TOTALS)

ws.Activate
ws.Unprotect

Set btn = ws.Buttons("btnStart")
btn.Font.ColorIndex = 15
btn.Enabled = False
Application.Cursor = xlDefault

ws.Protect

End Sub

' --- Enable START button

Sub enableStartBtn()

Set ws = Sheets(TOTALS)

ws.Activate
ws.Unprotect

Set btn = Sheets(TOTALS).Buttons("btnStart")
btn.Font.ColorIndex = 1
btn.Enabled = True
Application.Cursor = xlDefault

ws.Unprotect

End Sub
