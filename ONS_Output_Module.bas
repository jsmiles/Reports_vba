Attribute VB_Name = "ONS_Output"

'The publicly visible macro that calls the other three macros required

Public Sub ONS_Output()

    CreateResultsSheet
    NameResults
    CalculateResults

End Sub


'The macro required to create a new tab within the worksheet

Private Sub CreateResultsSheet()

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:= _
             ActiveWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "My_Results"

End Sub

'This macro outputs strings of text that describe the data output for ease of understanding

Private Sub NameResults()

    Sheets("My_Results").Range("C2") = "Number of people currently Self Employed"
    Sheets("My_Results").Range("C3") = "Count increase from last year"
    Sheets("My_Results").Range("C4") = "Percentage change since last year"
    Sheets("My_Results").Range("C5") = "Percentage of all those self of employed who are Women"

End Sub

'This macro calculates the various statistics and outputs them as appropriate

Private Sub CalculateResults()

    Dim value1 As Long
    Dim value2 As Long
    Dim value3 As Long

    value1 = ActiveWorkbook.Sheets("People (16+)").Range("D8").End(xlDown).Value
    ActiveWorkbook.Sheets("My_Results").Range("B2") = value1
    
    value1 = ActiveWorkbook.Sheets("People (16+)").Range("D8").End(xlDown).Value
    value2 = ActiveWorkbook.Sheets("People (16+)").Range("D8").End(xlDown).Offset(-12).Value
    ActiveWorkbook.Sheets("My_Results").Range("B3") = value1 - value2
    
    value3 = value1 - value2
    ActiveWorkbook.Sheets("My_Results").Range("B4") = value3 / value2
    
    value1 = ActiveWorkbook.Sheets("People (16+)").Range("D8").End(xlDown).Value
    value2 = ActiveWorkbook.Sheets("Women (16+)").Range("D8").End(xlDown).Value
    ActiveWorkbook.Sheets("My_Results").Range("B5") = value2 / value1
    
End Sub



