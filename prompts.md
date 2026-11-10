
Option Explicit

Sub InsertNegativeTotal()

    '========== PARAMETERS ==========
    Const SHEET_NAME As String = "AggregateAudit"
    Const TARGET_COL As String = "S"
    Const NEW_HEADER As String = "Negative Total"
    '================================

    Dim ws As Worksheet
    Dim targetColNum As Long
    Dim sourceColNum As Long
    Dim lastRow As Long
    Dim r As Long
    Dim f As String
    Dim countUpdated As Long

    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean

    Set ws = ActiveWorkbook.Worksheets(SHEET_NAME)

    ' Determine target and preceding columns
    targetColNum = ws.Columns(TARGET_COL).Column
    sourceColNum = targetColNum - 1

    If sourceColNum < 1 Then Exit Sub

    ' Prevent duplicate insertion
    If ws.Cells(1, targetColNum).Value = NEW_HEADER Then
        MsgBox "Column already exists."
        Exit Sub
    End If

    ' Last used row of preceding column
    lastRow = ws.Cells(ws.Rows.Count, _
                      sourceColNum).End(xlUp).Row

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Insert target column
    ws.Columns(targetColNum).Insert Shift:=xlToRight

    ' Copy source formatting
    ws.Range(ws.Cells(1, sourceColNum), _
             ws.Cells(lastRow, sourceColNum)).Copy

    ws.Cells(1, targetColNum).PasteSpecial _
        Paste:=xlPasteFormats

    Application.CutCopyMode = False

    ws.Columns(targetColNum).ColumnWidth = _
        ws.Columns(sourceColNum).ColumnWidth

    ' Set header
    ws.Cells(1, targetColNum).Value = NEW_HEADER

    ' Generate formulas until last row of source column
    For r = 2 To lastRow

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

    MsgBox "Completed: " & TARGET_COL & vbCrLf & _
           "Last row: " & lastRow & vbCrLf & _
           "Formula rows processed: " & countUpdated

    Exit Sub

ErrorHandler:

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

                ' Handle escaped double quotes
                If inQuotes And Mid$(f, i + 1, 1) = """" Then
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

            ' First SUMIFS argument = sum_range
            sumRange = Trim$(Mid$(f, _
                openPos + 1, commaPos - openPos - 1))

            ' Append negative condition
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
