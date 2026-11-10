Option Explicit

Sub InsertTotalNegative()

    '========== PARAMETERS ==========
    Const SHEET_NAME As String = "AggregateAudit"
    Const TARGET_COL As String = "S"
    Const NEW_HEADER As String = "Total Negative"
    Const SOURCE_HEADER As String = "SummaryTotal"
    '================================

    Dim ws As Worksheet
    Dim targetColNum As Long
    Dim sourceColNum As Long
    Dim headerCell As Range
    Dim sourceHeaderCell As Range

    Dim headerRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim f As String
    Dim countUpdated As Long
    Dim isExisting As Boolean

    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean

    Set ws = ActiveWorkbook.Worksheets(SHEET_NAME)

    targetColNum = ws.Columns(TARGET_COL).Column
    sourceColNum = targetColNum - 1

    If sourceColNum < 1 Then Exit Sub

    ' Last row of preceding column
    lastRow = ws.Cells(ws.Rows.Count, _
                      sourceColNum).End(xlUp).Row

    ' Find Total Negative anywhere in target column
    Set headerCell = ws.Columns(targetColNum).Find( _
        What:=NEW_HEADER, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    isExisting = Not headerCell Is Nothing

    If isExisting Then

        ' Use actual row of existing header
        headerRow = headerCell.Row

    Else

        ' Find SummaryTotal header in preceding column
        Set sourceHeaderCell = ws.Columns(sourceColNum).Find( _
            What:=SOURCE_HEADER, _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)

        If sourceHeaderCell Is Nothing Then
            MsgBox "Source header not found."
            Exit Sub
        End If

        headerRow = sourceHeaderCell.Row

    End If

    If lastRow <= headerRow Then
        MsgBox "No source data below header."
        Exit Sub
    End If

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If Not isExisting Then

        ' Insert new column
        ws.Columns(targetColNum).Insert Shift:=xlToRight

        ' Copy formatting
        ws.Range(ws.Cells(headerRow, sourceColNum), _
                 ws.Cells(lastRow, sourceColNum)).Copy

        ws.Cells(headerRow, targetColNum).PasteSpecial _
            Paste:=xlPasteFormats

        Application.CutCopyMode = False

        ws.Columns(targetColNum).ColumnWidth = _
            ws.Columns(sourceColNum).ColumnWidth

        ' Write new header at source header row
        ws.Cells(headerRow, targetColNum).Value = NEW_HEADER

    End If

    ' Start immediately BELOW actual header row
    For r = headerRow + 1 To lastRow

        If ws.Cells(r, sourceColNum).HasFormula Then

            f = ws.Cells(r, sourceColNum).Formula

            f = AddNegativeCondition(f)

            ws.Cells(r, targetColNum).Formula = f

            countUpdated = countUpdated + 1

        Else

            ws.Cells(r, targetColNum).Value = _
                ws.Cells(r, sourceColNum).Value

        End If

    Next r

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    MsgBox "Completed!" & vbCrLf & _
           "Target column: " & TARGET_COL & vbCrLf & _
           "Existing column: " & isExisting & vbCrLf & _
           "Header row: " & headerRow & vbCrLf & _
           "Start row: " & headerRow + 1 & vbCrLf & _
           "Last row: " & lastRow & vbCrLf & _
           "Formula rows: " & countUpdated, vbInformation

    Exit Sub

ErrorHandler:

    Application.CutCopyMode = False
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    MsgBox "Error: " & Err.Description, vbExclamation

End Sub


Function AddNegativeCondition(ByVal f As String) As String

    Dim p As Long
    Dim openPos As Long
    Dim closePos As Long
    Dim commaPos As Long
    Dim i As Long
    Dim depth As Long
    Dim inQuotes As Boolean
    Dim ch As String
    Dim sumRange As String

    p = InStrRev(UCase$(f), "SUMIFS(")

    Do While p > 0

        openPos = p + Len("SUMIFS")
        closePos = 0
        commaPos = 0
        depth = 0
        inQuotes = False

        For i = openPos To Len(f)

            ch = Mid$(f, i, 1)

            If ch = """" Then

                If inQuotes And _
                   Mid$(f, i + 1, 1) = """" Then
                    i = i + 1
                Else
                    inQuotes = Not inQuotes
                End If

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
                        If depth = 1 And commaPos = 0 Then
                            commaPos = i
                        End If

                End Select

            End If

        Next i

        If commaPos > 0 And closePos > 0 Then

            sumRange = Trim$(Mid$(f, _
                openPos + 1, commaPos - openPos - 1))

            f = Left$(f, closePos - 1) & _
                "," & sumRange & ",""<""&0" & _
                Mid$(f, closePos)

        End If

        If p > 1 Then
            p = InStrRev(UCase$(f), "SUMIFS(", p - 1)
        Else
            p = 0
        End If

    Loop

    AddNegativeCondition = f

End Function
