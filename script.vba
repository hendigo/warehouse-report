Private Sub NumberingCells()

Dim ws As Worksheet, rg As Range
Dim rowCount As Long, numberCount As Long

Set ws = ThisWorkbook.Worksheets("REPORT")

rowCount = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row - 10
numberCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 10

If numberCount > 0 Then

ws.Range("A11:A" & numberCount + 10).ClearContents
ws.Range("A11:A" & numberCount + 10).Borders.LineStyle = xlNone

End If

'Debug.Print numberCount

For i = 1 To rowCount
        Set rg = ws.Cells(10 + i, "A") ' Menetapkan rentang sel ke variabel rg
        rg.Value = i
        
        With rg.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
Next i

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Dim changeRange As Range
Set changeRange = Intersect(Target, Me.Range("C2:E2"))
    If Not changeRange Is Nothing Then
        Dim rgData As Range, rgCriteria As Range, rgDestination As Range, clrDestination As Range
        Dim rgColumn As Range

        Set clrDestination = ThisWorkbook.Worksheets("REPORT").Range("C11:R5000")

        clrDestination.Borders.LineStyle = xlNone
        clrDestination.Interior.ColorIndex = xlNone

        Set rgData = ThisWorkbook.Worksheets("DATA").Range("A5").CurrentRegion
        Set rgCriteria = ThisWorkbook.Worksheets("REPORT").Range("C1").CurrentRegion
        Set rgDestination = ThisWorkbook.Worksheets("REPORT").Range("C10").CurrentRegion.Resize(1)
    
        rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgDestination
        
        ThisWorkbook.Worksheets("REPORT").Range("L6").Formula = "=IF(ISBLANK($C$2),""Total Seluruh Pengeluaran"",""Total Pengeluaran Untuk : "" & $C$2)"
        ThisWorkbook.Worksheets("REPORT").Range("L7").Formula = "=IF(ISBLANK($C$2),""Total Seluruh Stok di Gudang"",""Total Stok : "" & $C$2)"
        ThisWorkbook.Worksheets("REPORT").Range("L8").Formula = "=IF(ISBLANK($C$2),""Jumlah Pengeluaran & Stok"",""Jumlah Pengeluaran & Stok "" & $C$2)"
        
    
        ThisWorkbook.Worksheets("REPORT").Range("M6").Formula = "=IF(ISBLANK(REPORT!$C$2),SUMIFS(DATA!$Q:$Q,DATA!$K:$K,"">=1""),SUMIFS(DATA!$Q:$Q,DATA!$K:$K,"">=1"",DATA!$F:$F,REPORT!$C$2))"
        ThisWorkbook.Worksheets("REPORT").Range("M7").Formula = "=IF(ISBLANK(REPORT!$C$2),SUMIFS(DATA!$Q:$Q,DATA!$K:$K,""=0""),SUMIFS(DATA!$Q:$Q,DATA!$K:$K,""=0"",DATA!$F:$F,REPORT!$C$2))"
        ThisWorkbook.Worksheets("REPORT").Range("M8").Formula = "=SUM($M$6:$M$7)"
        
        NumberingCells
        
    End If
End Sub
