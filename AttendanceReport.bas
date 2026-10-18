Attribute VB_Name = "AttendanceReport"
Option Explicit

' ============================================================================
' AttendanceReport — BuildAttendanceReport macro
' Monthly attendance analytics for the "Attendance" sheet
' (15 employees x 30 days, codes P/A/L/H).
'
' What it does:
'   1. Color-codes the source grid: Absent (A) -> red, Late (L) -> amber.
'   2. Builds an "Attendance Report" sheet: per-employee % present, absents,
'      lates, half days, days tracked; flags chronic absentees (>=5) red and
'      frequent lates (>=3) amber.
'   3. Adds a SUMMARY dashboard: employees tracked, total absents, total
'      lates, and the employee with the most absences.
'
' How to use (Excel): open Attendance.xlsm, enable macros, press Alt+F8,
' select BuildAttendanceReport, click Run. Safe to re-run — the report sheet
' is deleted and rebuilt from the current data every time.
'
' 100% local. No network calls, no external files. by Muhammad Bilal.
' ============================================================================

Public Sub BuildAttendanceReport()
    On Error GoTo ErrHandler
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rep As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long, s As Long
    Dim emp As String, code As String
    Dim present As Long, absent As Long, late As Long, half As Long, tracked As Long
    Dim repRow As Long
    Dim maxAbsent As Long, maxAbsentEmp As String
    Dim totAbsent As Long, totLate As Long
    Dim headers(1 To 6) As String

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Attendance")

    lastRow = 1
    Do While Trim(CStr(ws.Cells(lastRow + 1, 1).Value)) <> ""
        lastRow = lastRow + 1
    Loop
    lastCol = 1
    Do While Trim(CStr(ws.Cells(1, lastCol + 1).Value)) <> ""
        lastCol = lastCol + 1
    Loop
    If lastRow < 2 Or lastCol < 2 Then
        MsgBox "No attendance data found on the 'Attendance' sheet.", vbExclamation, "Attendance Report"
        Exit Sub
    End If

    For r = 2 To lastRow
        For c = 2 To lastCol
            code = UCase(Trim(CStr(ws.Cells(r, c).Value)))
            If code = "A" Then
                ws.Cells(r, c).Interior.Color = RGB(255, 205, 205)
                ws.Cells(r, c).Font.Color = RGB(160, 20, 20)
                ws.Cells(r, c).Font.Bold = True
            ElseIf code = "L" Then
                ws.Cells(r, c).Interior.Color = RGB(255, 235, 180)
                ws.Cells(r, c).Font.Color = RGB(150, 90, 0)
                ws.Cells(r, c).Font.Bold = True
            End If
        Next c
    Next r

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("Attendance Report").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    Set rep = wb.Worksheets.Add(After:=ws)
    rep.Name = "Attendance Report"

    Call SetColWidth(rep, "A", 22)
    Call SetColWidth(rep, "B", 14)
    Call SetColWidth(rep, "C", 12)
    Call SetColWidth(rep, "D", 12)
    Call SetColWidth(rep, "E", 12)
    Call SetColWidth(rep, "F", 14)

    rep.Range("A1:F1").Merge
    rep.Range("A1").Value = "MONTHLY ATTENDANCE REPORT"
    rep.Range("A1").Font.Bold = True
    rep.Range("A1").Font.Size = 18
    rep.Range("A1").Font.Color = RGB(255, 255, 255)
    rep.Range("A1").Interior.Color = RGB(14, 76, 129)
    rep.Range("A2").Value = "Generated " & Format(Now, "yyyy-mm-dd hh:mm") & "   |   Codes: P = Present, A = Absent, L = Late, H = Half day"

    headers(1) = "Employee": headers(2) = "% Present": headers(3) = "Absents"
    headers(4) = "Late": headers(5) = "Half Days": headers(6) = "Days Tracked"
    For c = 1 To 6
        rep.Cells(4, c).Value = headers(c)
        rep.Cells(4, c).Font.Bold = True
        rep.Cells(4, c).Font.Color = RGB(255, 255, 255)
        rep.Cells(4, c).Interior.Color = RGB(60, 90, 130)
    Next c

    repRow = 5
    maxAbsent = -1
    maxAbsentEmp = ""
    totAbsent = 0
    totLate = 0
    For r = 2 To lastRow
        emp = CStr(ws.Cells(r, 1).Value)
        present = 0: absent = 0: late = 0: half = 0: tracked = 0
        For c = 2 To lastCol
            code = UCase(Trim(CStr(ws.Cells(r, c).Value)))
            If code <> "" Then
                tracked = tracked + 1
                If code = "P" Then present = present + 1
                If code = "A" Then absent = absent + 1
                If code = "L" Then late = late + 1
                If code = "H" Then half = half + 1
            End If
        Next c
        rep.Cells(repRow, 1).Value = emp
        If tracked > 0 Then
            rep.Cells(repRow, 2).Value = present / tracked
        Else
            rep.Cells(repRow, 2).Value = 0
        End If
        rep.Cells(repRow, 2).NumberFormat = "0.0%"
        rep.Cells(repRow, 3).Value = absent
        rep.Cells(repRow, 4).Value = late
        rep.Cells(repRow, 5).Value = half
        rep.Cells(repRow, 6).Value = tracked
        If absent >= 5 Then
            rep.Cells(repRow, 3).Interior.Color = RGB(255, 205, 205)
            rep.Cells(repRow, 3).Font.Bold = True
        End If
        If late >= 3 Then
            rep.Cells(repRow, 4).Interior.Color = RGB(255, 235, 180)
            rep.Cells(repRow, 4).Font.Bold = True
        End If
        totAbsent = totAbsent + absent
        totLate = totLate + late
        If absent > maxAbsent Then
            maxAbsent = absent
            maxAbsentEmp = emp
        End If
        repRow = repRow + 1
    Next r

    s = repRow + 2
    rep.Range("A" & s & ":F" & s).Merge
    rep.Range("A" & s).Value = "SUMMARY"
    rep.Range("A" & s).Font.Bold = True
    rep.Range("A" & s).Font.Size = 14
    rep.Range("A" & s).Font.Color = RGB(255, 255, 255)
    rep.Range("A" & s).Interior.Color = RGB(14, 76, 129)
    rep.Cells(s + 1, 1).Value = "Employees tracked"
    rep.Cells(s + 1, 2).Value = lastRow - 1
    rep.Cells(s + 2, 1).Value = "Total absents (all employees)"
    rep.Cells(s + 2, 2).Value = totAbsent
    rep.Cells(s + 3, 1).Value = "Total lates (all employees)"
    rep.Cells(s + 3, 2).Value = totLate
    rep.Cells(s + 4, 1).Value = "Most absences"
    rep.Cells(s + 4, 2).Value = maxAbsentEmp & " (" & maxAbsent & ")"
    rep.Range("A" & (s + 1) & ":A" & (s + 4)).Font.Bold = True
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Attendance Report"
End Sub

Private Sub SetColWidth(sh As Worksheet, colLetter As String, w As Double)
    On Error Resume Next
    sh.Columns(colLetter).ColumnWidth = w
End Sub
