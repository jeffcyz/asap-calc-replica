Option Explicit

Sub AddNegativeTotalToAggregateAudit()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim originalFormula As String
    Dim newFormula As String
    Dim outputPath As String
    Dim updatedCount As Long

    Dim oldCalculation As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean

    On Error GoTo ErrorHandler

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets("AggregateAudit")

    If Trim$(ws.Range("R1").Value) <> "SummaryTotal" Then
        MsgBox "R1 is not SummaryTotal. Please check.", vbExclamation
        Exit Sub
    End If

    If Trim$(ws.Range("S1").Value) = "Negative Total" Then
        MsgBox "Negative Total already exists.", vbExclamation
        Exit Sub
    End If

    oldCalculation = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    lastRow = ws.Cells(ws.Rows.Count, "R").End(xlUp).Row

    ' Insert column S after SummaryTotal
    ws.Columns("S").Insert Shift:=xlToRight

    ' Copy formatting
    ws.Range("R1:R" & lastRow).Copy
    ws.Range("S1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Columns("S").ColumnWidth = ws.Columns("R").ColumnWidth

    ' New header
    ws.Range("S1").Value = "Negative Total"

    ' Generate Negative Total formulas
    For i = 2 To lastRow

        If ws.Cells(i, "R").HasFormula Then

            originalFormula = ws.Cells(i, "R").Formula

            newFormula = AddNegativeSUMIFS(originalFormula)

            ws.Cells(i, "S").Formula = newFormula

            If newFormula <> originalFormula Then
                updatedCount = updatedCount + 1
            End If

        Else
            ws.Cells(i, "S").Value = ws.Cells(i, "R").Value
        End If

    Next i

    ' Save as a new XLSB file
    outputPath = Left$(wb.FullName, _
                 InStrRev(wb.FullName, ".") - 1) & _
                 "_NegativeTotal.xlsb"

    wb.SaveAs Filename:=outputPath, FileFormat:=xlExcel12

    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents

    MsgBox "Completed!" & vbCrLf & _
           "Updated formula rows: " & updatedCount & vbCrLf & _
           "Saved as: " & outputPath, vbInformation

    Exit Sub

ErrorHandler:

    On Error Resume Next
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents

    MsgBox "Error: " & Err.Description, vbCritical

End Sub


Function AddNegativeSUMIFS(ByVal formulaText As String) As String

    Dim p As Long
    Dim openPos As Long
    Dim closePos As Long
    Dim firstComma As Long
    Dim i As Long
    Dim depth As Long
    Dim inQuotes As Boolean
    Dim ch As String
    Dim sumRange As String

    ' Process SUMIFS from right to left
    p = InStrRev(UCase$(formulaText), "SUMIFS(")

    Do While p > 0

        openPos = p + 6
        closePos = 0
        firstComma = 0
        depth = 0
        inQuotes = False

        For i = openPos To Len(formulaText)

            ch = Mid$(formulaText, i, 1)

            If ch = """" Then
                inQuotes = Not inQuotes

            ElseIf Not inQuotes Then

                Select Case ch

                    Case "("
                        depth = depth + 1

                    Case ")"
                        depth = depth - 1

                        If depth = 0 Then
                            closePos = i
                            Exit For
                        End If

                    Case ","
                        If depth = 1 And firstComma = 0 Then
                            firstComma = i
                        End If

                End Select

            End If

        Next i

        If firstComma > 0 And closePos > 0 Then

            sumRange = Trim$(Mid$(formulaText, _
                       openPos + 1, _
                       firstComma - openPos - 1))

            formulaText = Left$(formulaText, closePos - 1) & _
                          "," & sumRange & ",""<""&0" & _
                          Mid$(formulaText, closePos)

        End If

        If p > 1 Then
            p = InStrRev(UCase$(formulaText), "SUMIFS(", p - 1)
        Else
            p = 0
        End If

    Loop

    AddNegativeSUMIFS = formulaText

End Function
