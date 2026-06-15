Attribute VB_Name = "KoujiSchedule"
Option Explicit

' =====================================================================
' Generates a unit construction schedule table from a resident list
' (.xls/.xlsx).
'
' How to use:
'   1. Import this module into Excel
'      (VBE: Alt+F11 -> File -> Import File -> select this file)
'   2. Press Alt+F8, run "GenerateKoujiSchedule"
'   3. Select the resident list file
'   4. Specify where to save the schedule table
' =====================================================================

Private Function FwMonth() As String
    FwMonth = ChrW(&H6708)
End Function

Private Function FwDay() As String
    FwDay = ChrW(&H65E5)
End Function

Private Function FwVacant() As String
    FwVacant = ChrW(&H7A7A) & ChrW(&H5BA4)
End Function

Private Function FwOutOfPeriod() As String
    FwOutOfPeriod = ChrW(&H5DE5) & ChrW(&H671F) & ChrW(&H5916)
End Function

Private Function FwCamera() As String
    FwCamera = ChrW(&H30AB) & ChrW(&H30E1) & ChrW(&H30E9)
End Function

Private Function FwHandset() As String
    FwHandset = ChrW(&H53D7) & ChrW(&H8A71) & ChrW(&H5668)
End Function

Private Function Weekdays() As Variant
    Weekdays = Array(ChrW(&H6708), ChrW(&H706B), ChrW(&H6C34), ChrW(&H6728), _
                      ChrW(&H91D1), ChrW(&H571F), ChrW(&H65E5))
End Function

' Parses strings like "6/19(Fri)13:00" or "3-28(Sat)9:00" (with full-width
' parentheses/colon and optional kanji month/day suffix).
' Returns array(month, day, weekday, hour, minute) or Empty.
Private Function ParseScheduleText(ByVal txt As String) As Variant
    If Len(Trim(txt)) = 0 Then
        ParseScheduleText = Empty
        Exit Function
    End If

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = "([0-9]{1,2})[/" & FwMonth() & "]([0-9]{1,2})" & FwDay() & "?[" & _
                  ChrW(&HFF08) & "(](.)[" & ChrW(&HFF09) & ")]\s*([0-9]{1,2})[:" & _
                  ChrW(&HFF1A) & "]([0-9]{2})"

    If re.Test(txt) Then
        Dim m As Object
        Set m = re.Execute(txt)(0)
        Dim res(4) As Variant
        res(0) = CInt(m.SubMatches(0))
        res(1) = CInt(m.SubMatches(1))
        res(2) = m.SubMatches(2)
        res(3) = CInt(m.SubMatches(3))
        res(4) = CInt(m.SubMatches(4))
        ParseScheduleText = res
    Else
        ParseScheduleText = Empty
    End If
End Function

' Builds the A/B option suffix from the remark column (camera/handset).
Private Function GetOptionMarker(ByVal remark As String) As String
    Dim suffix As String
    suffix = ""
    If InStr(remark, FwCamera()) > 0 Then suffix = suffix & "A"
    If InStr(remark, FwHandset()) > 0 Then suffix = suffix & "B"
    GetOptionMarker = suffix
End Function

' Formats a room number, stripping a trailing ".0" if it was read as a number.
Private Function FormatRoom(ByVal v As Variant) As String
    If IsNumeric(v) And Not IsEmpty(v) And Trim(CStr(v)) <> "" Then
        If CDbl(v) = Int(CDbl(v)) Then
            FormatRoom = CStr(CLng(v))
            Exit Function
        End If
    End If
    FormatRoom = Trim(CStr(v))
End Function

' Builds the "common-area construction" cell text, optionally with a
' "(start~end)" time range when the period does not cover the full day.
Private Function CommonAreaLabel(ByVal startLabel As String, ByVal endLabel As String) As String
    Dim s As String
    s = ChrW(&H5171) & ChrW(&H3000) & ChrW(&H7528) & ChrW(&H3000) & ChrW(&H90E8) & _
        ChrW(&H3000) & ChrW(&H5DE5) & ChrW(&H3000) & ChrW(&H4E8B)
    If startLabel <> "" Then
        s = s & ChrW(&HFF08) & startLabel & ChrW(&H301C) & endLabel & ChrW(&HFF09)
    End If
    s = s & Chr(10) & _
        ChrW(&HFF08) & ChrW(&H304A) & ChrW(&H90E8) & ChrW(&H5C4B) & ChrW(&H306E) & _
        ChrW(&H5DE5) & ChrW(&H7A0B) & ChrW(&H306F) & ChrW(&H51FA) & ChrW(&H6765) & _
        ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&HFF09)
    CommonAreaLabel = s
End Function

' Parses "M/D H:MM" (or "M/D HH:MM") into array(month, day, hour, minute).
Private Function ParseDateTimeInput(ByVal s As String) As Variant
    Dim parts() As String
    parts = Split(Trim(s), " ")

    Dim dateParts() As String
    dateParts = Split(parts(0), "/")

    Dim timeParts() As String
    If UBound(parts) >= 1 And Trim(parts(1)) <> "" Then
        timeParts = Split(parts(1), ":")
    Else
        timeParts = Split("9:00", ":")
    End If

    Dim res(3) As Variant
    res(0) = CInt(dateParts(0))
    res(1) = CInt(dateParts(1))
    res(2) = CInt(timeParts(0))
    res(3) = CInt(timeParts(1))
    ParseDateTimeInput = res
End Function

' Sorts an array of (room, minute, option) entries ascending by minute.
Private Sub SortEntriesByMinute(ByRef entries() As Variant, ByVal n As Long)
    Dim i As Long, j As Long
    Dim tmp As Variant
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If entries(j)(1) < entries(i)(1) Then
                tmp = entries(i)
                entries(i) = entries(j)
                entries(j) = tmp
            End If
        Next j
    Next i
End Sub

Sub GenerateKoujiSchedule()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename( _
        "Excel Files (*.xls;*.xlsx),*.xls;*.xlsx", , "Select resident list file")
    If srcPath = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim srcWB As Workbook
    Set srcWB = Workbooks.Open(srcPath, ReadOnly:=True)

    Dim buildingName As String
    buildingName = CStr(srcWB.Sheets(1).Cells(1, 1).Value)

    ' --- Collect data ---
    Dim dateKeys As Object, slotData As Object
    Set dateKeys = CreateObject("Scripting.Dictionary")  ' key: month*100+day -> weekday
    Set slotData = CreateObject("Scripting.Dictionary")  ' key: (month*100+day)*100+hour -> Collection

    Dim vacantList As Collection, notSubmittedList As Collection, outOfPeriodList As Collection
    Set vacantList = New Collection
    Set notSubmittedList = New Collection
    Set outOfPeriodList = New Collection

    Dim ws As Worksheet
    For Each ws In srcWB.Worksheets
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 3 Then
            Dim r As Long
            For r = 3 To lastRow
                Dim roomVal As Variant
                roomVal = ws.Cells(r, 1).Value
                If Trim(CStr(roomVal)) <> "" Then
                    Dim room As String, nameVal As String, schedText As String, remarkVal As String
                    room = FormatRoom(roomVal)
                    nameVal = CStr(ws.Cells(r, 2).Value)
                    schedText = CStr(ws.Cells(r, 5).Value)
                    remarkVal = CStr(ws.Cells(r, 6).Value)

                    If nameVal = FwVacant() Then
                        vacantList.Add room
                    ElseIf InStr(remarkVal, FwOutOfPeriod()) > 0 Then
                        Dim oentry(1) As Variant
                        oentry(0) = room
                        oentry(1) = ParseScheduleText(schedText)
                        outOfPeriodList.Add oentry
                    Else
                        Dim parsed As Variant
                        parsed = ParseScheduleText(schedText)
                        If IsEmpty(parsed) Then
                            notSubmittedList.Add room
                        Else
                            Dim mon As Long, dayNum As Long, wd As String, hr As Long, mi As Long
                            mon = parsed(0): dayNum = parsed(1): wd = parsed(2)
                            hr = parsed(3): mi = parsed(4)

                            Dim dkey As Long
                            dkey = mon * 100 + dayNum
                            If Not dateKeys.Exists(dkey) Then dateKeys.Add dkey, wd

                            Dim skey As Long
                            skey = dkey * 100 + hr
                            If Not slotData.Exists(skey) Then slotData.Add skey, New Collection

                            Dim entry(2) As Variant
                            entry(0) = room
                            entry(1) = mi
                            entry(2) = GetOptionMarker(remarkVal)
                            slotData(skey).Add entry
                        End If
                    End If
                End If
            Next r
        End If
    Next ws

    srcWB.Close SaveChanges:=False

    ' --- Build the schedule sheet ---
    Dim outWB As Workbook
    Set outWB = Workbooks.Add
    Dim sht As Worksheet
    Set sht = outWB.Sheets(1)
    sht.Name = ChrW(&H5DE5) & ChrW(&H7A0B) & ChrW(&H8868)  ' "Schedule"

    Const TOTAL_COLS As Long = 9
    Dim timeSlots(7) As Long
    timeSlots(0) = 9: timeSlots(1) = 10: timeSlots(2) = 11
    timeSlots(3) = 13: timeSlots(4) = 14: timeSlots(5) = 15: timeSlots(6) = 16: timeSlots(7) = 17

    Dim YELLOW As Long, ORANGE As Long, GRAY As Long
    YELLOW = RGB(255, 255, 0)
    ORANGE = RGB(255, 192, 0)
    GRAY = RGB(217, 217, 217)

    ' Title
    With sht.Range(sht.Cells(1, 1), sht.Cells(1, TOTAL_COLS))
        .Merge
        .Value = buildingName & ChrW(&H3000) & ChrW(&H69D8) & ChrW(&H3000) & _
                 ChrW(&H4F4F) & ChrW(&H6238) & ChrW(&H5185) & ChrW(&H5DE5) & ChrW(&H7A0B) & _
                 ChrW(&H65E5) & ChrW(&H7A0B) & ChrW(&H8868)
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    sht.Rows(1).RowHeight = 30

    ' Note
    With sht.Range(sht.Cells(2, 1), sht.Cells(2, TOTAL_COLS))
        .Merge
        .Value = ChrW(&HFF0A) & ChrW(&H8868) & ChrW(&H4E2D) & ChrW(&H306E) & ChrW(&H90E8) & _
                 ChrW(&H5C4B) & ChrW(&H6570) & ChrW(&H306F) & ChrW(&H5DE5) & ChrW(&H7A0B) & _
                 ChrW(&H53EF) & ChrW(&H80FD) & ChrW(&H306A) & ChrW(&H90E8) & ChrW(&H5C4B) & _
                 ChrW(&H6570) & ChrW(&H3092) & ChrW(&H793A) & ChrW(&H3057) & ChrW(&H307E) & _
                 ChrW(&H3059) & ChrW(&H3002)
        .HorizontalAlignment = xlLeft
    End With

    Dim headerRow1 As Long, headerRow2 As Long
    headerRow1 = 3
    headerRow2 = 4

    ' "Construction date" header
    With sht.Range(sht.Cells(headerRow1, 1), sht.Cells(headerRow2, 1))
        .Merge
        .Value = ChrW(&H5DE5) & ChrW(&H4E8B) & ChrW(&H4E88) & ChrW(&H5B9A) & ChrW(&H65E5)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Morning header
    With sht.Range(sht.Cells(headerRow1, 2), sht.Cells(headerRow1, 4))
        .Merge
        .Value = ChrW(&H5348) & ChrW(&H524D) & ChrW(&HFF08) & "9" & ChrW(&H6642) & _
                 ChrW(&H301C) & "12" & ChrW(&H6642) & ChrW(&HFF09)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Afternoon header
    With sht.Range(sht.Cells(headerRow1, 5), sht.Cells(headerRow1, TOTAL_COLS))
        .Merge
        .Value = ChrW(&H5348) & ChrW(&H5F8C) & ChrW(&HFF08) & "13" & ChrW(&H6642) & _
                 ChrW(&H301C) & "18" & ChrW(&H6642) & ChrW(&HFF09)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Time slot labels
    Dim i As Long
    For i = 0 To 7
        With sht.Cells(headerRow2, 2 + i)
            .Value = timeSlots(i) & ChrW(&H6642) & ChrW(&H983C)  ' "n o'clock-ish"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i

    With sht.Range(sht.Cells(headerRow1, 1), sht.Cells(headerRow2, TOTAL_COLS))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' --- Data rows ---
    Dim row As Long
    row = headerRow2 + 1

    ' Sort date keys ascending
    Dim keysArr As Variant
    Dim nKeys As Long
    nKeys = dateKeys.Count
    Dim sortedKeys() As Long
    If nKeys > 0 Then
        keysArr = dateKeys.Keys
        ReDim sortedKeys(0 To nKeys - 1)
        Dim a As Long, b As Long, tmpL As Long
        For a = 0 To nKeys - 1
            sortedKeys(a) = keysArr(a)
        Next a
        For a = 0 To nKeys - 2
            For b = a + 1 To nKeys - 1
                If sortedKeys(b) < sortedKeys(a) Then
                    tmpL = sortedKeys(a)
                    sortedKeys(a) = sortedKeys(b)
                    sortedKeys(b) = tmpL
                End If
            Next b
        Next a
    End If

    ' Add the common-area construction day(s) before the first unit day.
    ' The period (start/end date+time) is entered by the user, so it can
    ' span multiple days. If the period ends partway through a day (e.g.
    ' 15:00) and that day also has unit work scheduled, the time slots
    ' from the end time onward are used for the unit work instead of
    ' being blocked out.
    Dim wds As Variant
    wds = Weekdays()

    Dim mergedDates As Object
    Set mergedDates = CreateObject("Scripting.Dictionary")  ' dkey -> True

    Dim wdIdxFirst As Long
    Dim defCaDate As Date
    Dim promptStart As String, promptEnd As String, titleCa As String
    Dim startInput As String, endInput As String
    Dim startParts As Variant, endParts As Variant
    Dim caStartDate As Date, caEndDate As Date
    Dim caStartHour As Long, caEndHour As Long
    Dim caStartLabel As String, caEndLabel As String
    Dim caDay As Date
    Dim effStartHour As Long, effEndHour As Long
    Dim effStartLabel As String, effEndLabel As String
    Dim dkeyForDay As Long
    Dim isFullDay As Boolean
    Dim diffDays As Long, wdIdxForDay As Long
    Dim firstBlockedCol As Long, lastBlockedCol As Long
    Dim mDkey As Long, mWd As String, mNRows As Long, mSkey As Long
    Dim col3 As Long, skey3 As Long
    Dim entries3() As Variant, nEntries3 As Long
    Dim eRoom3 As String, eMin3 As Long, eOpt3 As String
    Dim c3 As Range

    If nKeys > 0 Then
        Dim firstMon As Long, firstDay As Long, firstWd As String
        firstMon = sortedKeys(0) \ 100
        firstDay = sortedKeys(0) Mod 100
        firstWd = dateKeys(sortedKeys(0))

        wdIdxFirst = 0
        For a = 0 To 6
            If wds(a) = firstWd Then wdIdxFirst = a
        Next a

        defCaDate = DateSerial(2001, firstMon, firstDay) - 1

        titleCa = ChrW(&H5171) & ChrW(&H7528) & ChrW(&H90E8) & ChrW(&H5DE5) & ChrW(&H4E8B)
        promptStart = titleCa & ChrW(&H306E) & ChrW(&H958B) & ChrW(&H59CB) & ChrW(&H65E5) & _
                      ChrW(&H6642) & " (M/D H:MM)"
        promptEnd = titleCa & ChrW(&H306E) & ChrW(&H7D42) & ChrW(&H4E86) & ChrW(&H65E5) & _
                    ChrW(&H6642) & " (M/D H:MM)"

        startInput = InputBox(promptStart, titleCa, Month(defCaDate) & "/" & Day(defCaDate) & " 9:00")
        If Trim(startInput) = "" Then
            outWB.Close SaveChanges:=False
            Exit Sub
        End If

        endInput = InputBox(promptEnd, titleCa, Month(defCaDate) & "/" & Day(defCaDate) & " 18:00")
        If Trim(endInput) = "" Then
            outWB.Close SaveChanges:=False
            Exit Sub
        End If

        On Error GoTo CaInputError
        startParts = ParseDateTimeInput(startInput)
        endParts = ParseDateTimeInput(endInput)
        On Error GoTo 0

        caStartDate = DateSerial(2001, startParts(0), startParts(1))
        caEndDate = DateSerial(2001, endParts(0), endParts(1))
        caStartHour = startParts(2)
        caEndHour = endParts(2)
        caStartLabel = Format(startParts(2), "0") & ":" & Format(startParts(3), "00")
        caEndLabel = Format(endParts(2), "0") & ":" & Format(endParts(3), "00")

        For caDay = caStartDate To caEndDate Step 1
            If caDay = caStartDate Then
                effStartHour = caStartHour
                effStartLabel = caStartLabel
            Else
                effStartHour = 9
                effStartLabel = "9:00"
            End If
            If caDay = caEndDate Then
                effEndHour = caEndHour
                effEndLabel = caEndLabel
            Else
                effEndHour = 18
                effEndLabel = "18:00"
            End If

            dkeyForDay = Month(caDay) * 100 + Day(caDay)
            isFullDay = (effStartHour <= 9 And effEndHour >= 18)

            If isFullDay Then
                ' The whole day is common-area construction.
                diffDays = CLng(DateSerial(2001, firstMon, firstDay) - caDay)
                wdIdxForDay = ((wdIdxFirst - diffDays) Mod 7 + 7) Mod 7

                With sht.Range(sht.Cells(row, 1), sht.Cells(row, 1))
                    .Merge
                    .Value = Month(caDay) & FwMonth() & Day(caDay) & FwDay() & ChrW(&HFF08) & wds(wdIdxForDay) & ChrW(&HFF09)
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With

                With sht.Range(sht.Cells(row, 2), sht.Cells(row, TOTAL_COLS))
                    .Merge
                    .Value = CommonAreaLabel("", "")
                    .Font.Bold = True
                    .Font.Size = 12
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                    .Interior.Color = GRAY
                End With

                With sht.Range(sht.Cells(row, 1), sht.Cells(row, TOTAL_COLS))
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Interior.Color = GRAY
                End With
                sht.Rows(row).RowHeight = 40
                row = row + 1
            Else
                ' Partial day: some time slots are common-area, the rest
                ' (if any) are available for unit work.
                firstBlockedCol = -1
                lastBlockedCol = -1
                For i = 0 To 7
                    If timeSlots(i) >= effStartHour And timeSlots(i) < effEndHour Then
                        If firstBlockedCol = -1 Then firstBlockedCol = 2 + i
                        lastBlockedCol = 2 + i
                    End If
                Next i

                mDkey = dkeyForDay
                If dateKeys.Exists(mDkey) Then
                    mWd = dateKeys(mDkey)
                    mNRows = 1
                    For i = 0 To 7
                        mSkey = mDkey * 100 + timeSlots(i)
                        If slotData.Exists(mSkey) Then
                            If slotData(mSkey).Count > mNRows Then mNRows = slotData(mSkey).Count
                        End If
                    Next i
                    mergedDates.Add mDkey, True
                Else
                    diffDays = CLng(DateSerial(2001, firstMon, firstDay) - caDay)
                    wdIdxForDay = ((wdIdxFirst - diffDays) Mod 7 + 7) Mod 7
                    mWd = wds(wdIdxForDay)
                    mNRows = 1
                End If

                ' Date cell
                With sht.Range(sht.Cells(row, 1), sht.Cells(row + mNRows - 1, 1))
                    .Merge
                    .Value = Month(caDay) & FwMonth() & Day(caDay) & FwDay() & ChrW(&HFF08) & mWd & ChrW(&HFF09)
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With

                ' Common-area block (merged across the blocked columns/rows)
                If firstBlockedCol >= 2 Then
                    With sht.Range(sht.Cells(row, firstBlockedCol), sht.Cells(row + mNRows - 1, lastBlockedCol))
                        .Merge
                        .Value = CommonAreaLabel(effStartLabel, effEndLabel)
                        .Font.Bold = True
                        .Font.Size = 12
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Interior.Color = GRAY
                    End With
                End If

                ' Unit-work cells for the remaining (non-blocked) columns
                For i = 0 To 7
                    col3 = 2 + i
                    If col3 < firstBlockedCol Or col3 > lastBlockedCol Then
                        skey3 = mDkey * 100 + timeSlots(i)

                        nEntries3 = 0
                        If slotData.Exists(skey3) Then
                            nEntries3 = slotData(skey3).Count
                            ReDim entries3(0 To nEntries3 - 1)
                            For k = 1 To nEntries3
                                entries3(k - 1) = slotData(skey3)(k)
                            Next k
                            SortEntriesByMinute entries3, nEntries3
                        End If

                        For sub_ = 0 To mNRows - 1
                            Set c3 = sht.Cells(row + sub_, col3)
                            c3.HorizontalAlignment = xlCenter
                            c3.VerticalAlignment = xlCenter
                            c3.WrapText = True
                            If sub_ < nEntries3 Then
                                eRoom3 = entries3(sub_)(0)
                                eMin3 = entries3(sub_)(1)
                                eOpt3 = entries3(sub_)(2)
                                c3.Value = timeSlots(i) & ":" & Format(eMin3, "00") & Chr(10) & eRoom3 & eOpt3
                                c3.Font.Bold = True
                                If eOpt3 <> "" Then
                                    c3.Interior.Color = ORANGE
                                Else
                                    c3.Interior.Color = YELLOW
                                End If
                            End If
                        Next sub_
                    End If
                Next i

                With sht.Range(sht.Cells(row, 1), sht.Cells(row + mNRows - 1, TOTAL_COLS))
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With
                For r = row To row + mNRows - 1
                    sht.Rows(r).RowHeight = 40
                Next r
                row = row + mNRows
            End If
        Next caDay
    End If
    GoTo CaInputDone

CaInputError:
    MsgBox ChrW(&H5165) & ChrW(&H529B) & ChrW(&H5F62) & ChrW(&H5F0F) & ChrW(&H304C) & _
           ChrW(&H6B63) & ChrW(&H3057) & ChrW(&H304F) & ChrW(&H3042) & ChrW(&H308A) & _
           ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002) & " (M/D H:MM)" & vbCrLf & _
           ChrW(&H4F8B) & ChrW(&HFF1A) & "6/17 9:00"
    outWB.Close SaveChanges:=False
    Exit Sub
CaInputDone:

    ' One block of rows per date (dates already rendered together with the
    ' common-area construction row are skipped here).
    For a = 0 To nKeys - 1
        Dim dkey2 As Long, mon2 As Long, day2 As Long, wd2 As String
        dkey2 = sortedKeys(a)
        If mergedDates.Exists(dkey2) Then GoTo NextDate

        mon2 = dkey2 \ 100
        day2 = dkey2 Mod 100
        wd2 = dateKeys(dkey2)

        ' Number of rows needed = max number of bookings in any slot for this date
        Dim nRows As Long
        nRows = 1
        For i = 0 To 7
            Dim skey2 As Long
            skey2 = dkey2 * 100 + timeSlots(i)
            If slotData.Exists(skey2) Then
                If slotData(skey2).Count > nRows Then nRows = slotData(skey2).Count
            End If
        Next i

        ' Date cell
        With sht.Range(sht.Cells(row, 1), sht.Cells(row + nRows - 1, 1))
            .Merge
            .Value = mon2 & FwMonth() & day2 & FwDay() & ChrW(&HFF08) & wd2 & ChrW(&HFF09)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        For i = 0 To 7
            Dim col As Long
            col = 2 + i
            skey2 = dkey2 * 100 + timeSlots(i)

            Dim entries() As Variant
            Dim nEntries As Long
            nEntries = 0
            If slotData.Exists(skey2) Then
                nEntries = slotData(skey2).Count
                ReDim entries(0 To nEntries - 1)
                Dim k As Long
                For k = 1 To nEntries
                    entries(k - 1) = slotData(skey2)(k)
                Next k
                SortEntriesByMinute entries, nEntries
            End If

            Dim sub_ As Long
            For sub_ = 0 To nRows - 1
                Dim c As Range
                Set c = sht.Cells(row + sub_, col)
                c.HorizontalAlignment = xlCenter
                c.VerticalAlignment = xlCenter
                c.WrapText = True
                If sub_ < nEntries Then
                    Dim eRoom As String, eMin As Long, eOpt As String
                    eRoom = entries(sub_)(0)
                    eMin = entries(sub_)(1)
                    eOpt = entries(sub_)(2)
                    c.Value = timeSlots(i) & ":" & Format(eMin, "00") & Chr(10) & eRoom & eOpt
                    c.Font.Bold = True
                    If eOpt <> "" Then
                        c.Interior.Color = ORANGE
                    Else
                        c.Interior.Color = YELLOW
                    End If
                End If
            Next sub_
        Next i

        With sht.Range(sht.Cells(row, 1), sht.Cells(row + nRows - 1, TOTAL_COLS))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        row = row + nRows
NextDate:
    Next a

    ' Column widths / row heights
    sht.Columns(1).ColumnWidth = 16
    For i = 0 To 7
        sht.Columns(2 + i).ColumnWidth = 10
    Next i
    For r = headerRow2 + 1 To row - 1
        sht.Rows(r).RowHeight = 30
    Next r

    ' Option legend (camera/handset)
    Dim hasOption As Boolean
    hasOption = False
    Dim allKeys As Variant
    If slotData.Count > 0 Then
        allKeys = slotData.Keys
        For a = 0 To slotData.Count - 1
            Dim col2 As Collection
            Set col2 = slotData(allKeys(a))
            For k = 1 To col2.Count
                If col2(k)(2) <> "" Then hasOption = True
            Next k
        Next a
    End If
    If hasOption Then
        With sht.Cells(row, 1)
            .Value = "A" & ChrW(&HFF1A) & ChrW(&H30AB) & ChrW(&H30E1) & ChrW(&H30E9) & ChrW(&H4ED8) & _
                     ChrW(&H304D) & ChrW(&H3000) & "B" & ChrW(&HFF1A) & ChrW(&H53D7) & ChrW(&H8A71) & _
                     ChrW(&H5668) & ChrW(&H4ED8) & ChrW(&H304D) & ChrW(&H306E) & ChrW(&H304A) & _
                     ChrW(&H90E8) & ChrW(&H5C4B) & ChrW(&H3067) & ChrW(&H3059)
            .Interior.Color = ORANGE
            .Font.Bold = True
        End With
        row = row + 1
    End If

    ' Units that requested a schedule outside the construction period
    row = row + 1
    Dim oitem As Variant
    For Each oitem In outOfPeriodList
        Dim oroom As String
        oroom = oitem(0)
        Dim osched As Variant
        osched = oitem(1)
        Dim otxt As String
        otxt = oroom & ChrW(&H3000) & ChrW(&H5DE5) & ChrW(&H671F) & ChrW(&H5916) & ChrW(&H5E0C) & ChrW(&H671B)
        If Not IsEmpty(osched) Then
            otxt = otxt & ChrW(&HFF08) & osched(0) & "/" & osched(1) & osched(2) & " " & _
                   osched(3) & ":" & Format(osched(4), "00") & ChrW(&HFF09)
        End If
        With sht.Cells(row, 1)
            .Value = otxt
            .Font.Color = RGB(0, 0, 255)
            .Font.Bold = True
        End With
        row = row + 1
    Next oitem

    ' Units that have not yet submitted a schedule
    If notSubmittedList.Count > 0 Then
        Dim ns As String
        ns = ChrW(&H672A) & ChrW(&H63D0) & ChrW(&H51FA) & ChrW(&H3000)
        Dim item As Variant
        Dim first As Boolean
        first = True
        For Each item In notSubmittedList
            If Not first Then ns = ns & ", "
            ns = ns & item
            first = False
        Next item
        With sht.Cells(row, 1)
            .Value = ns
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
        End With
        row = row + 1
    End If

    ' Vacant units
    If vacantList.Count > 0 Then
        Dim vc As String
        vc = ChrW(&H7A7A) & ChrW(&H5BA4) & ChrW(&H3000)
        first = True
        For Each item In vacantList
            If Not first Then vc = vc & ", "
            vc = vc & item
            first = False
        Next item
        With sht.Cells(row, 1)
            .Value = vc
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
        End With
        row = row + 1
    End If

    ' --- Print setup: fit to one A4 landscape page ---
    With sht.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Zoom = False
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
    End With
    sht.PageSetup.PrintArea = sht.Range(sht.Cells(1, 1), sht.Cells(row - 1, TOTAL_COLS)).Address

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' --- Save the new workbook ---
    Dim outPath As Variant
    outPath = Application.GetSaveAsFilename( _
        InitialFileName:=ChrW(&H5DE5) & ChrW(&H7A0B) & ChrW(&H8868) & ".xlsx", _
        FileFilter:="Excel Workbook,*.xlsx")
    If outPath <> False Then
        outWB.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
        MsgBox ChrW(&H5DE5) & ChrW(&H7A0B) & ChrW(&H8868) & ChrW(&H3092) & ChrW(&H4F5C) & _
               ChrW(&H6210) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3057) & ChrW(&H305F) & _
               ChrW(&H3002) & vbCrLf & outPath
    End If
End Sub
