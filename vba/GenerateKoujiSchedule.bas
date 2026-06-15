Attribute VB_Name = "KoujiSchedule"
Option Explicit

' =====================================================================
' 入居者一覧(.xls/.xlsx)から住戸内工事日程表を生成するマクロ
'
' 使い方:
'   1. このモジュールをExcelにインポートする
'      (VBE: Alt+F11 -> ファイル -> ファイルのインポート -> このファイルを選択)
'   2. Alt+F8 で「GenerateKoujiSchedule」を実行する
'   3. 入居者一覧ファイルを選択する
'   4. 工程表の保存先を指定する
' =====================================================================

Private Function W月() As String
    W月 = ChrW(&H6708)
End Function

Private Function W日() As String
    W日 = ChrW(&H65E5)
End Function

Private Function W空室() As String
    W空室 = ChrW(&H7A7A) & ChrW(&H5BA4)
End Function

Private Function W工期外() As String
    W工期外 = ChrW(&H5DE5) & ChrW(&H671F) & ChrW(&H5916)
End Function

Private Function Wカメラ() As String
    Wカメラ = ChrW(&H30AB) & ChrW(&H30E1) & ChrW(&H30E9)
End Function

Private Function W受話器() As String
    W受話器 = ChrW(&H53D7) & ChrW(&H8A71) & ChrW(&H5668)
End Function

Private Function Weekdays() As Variant
    Weekdays = Array(ChrW(&H6708), ChrW(&H706B), ChrW(&H6C34), ChrW(&H6728), _
                      ChrW(&H91D1), ChrW(&H571F), ChrW(&H65E5))
End Function

' "6/19（金）13:00" や "3月28日(土)9：00" のような文字列を解析する
' 戻り値: 配列(月, 日, 曜日, 時, 分) または Empty
Private Function ParseScheduleText(ByVal txt As String) As Variant
    If Len(Trim(txt)) = 0 Then
        ParseScheduleText = Empty
        Exit Function
    End If

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = "([0-9]{1,2})[/" & W月() & "]([0-9]{1,2})" & W日() & "?[" & _
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

' 備考から「カメラ付」「受話器」のオプション表示(A/B)を作る
Private Function GetOptionMarker(ByVal remark As String) As String
    Dim suffix As String
    suffix = ""
    If InStr(remark, Wカメラ()) > 0 Then suffix = suffix & "A"
    If InStr(remark, W受話器()) > 0 Then suffix = suffix & "B"
    GetOptionMarker = suffix
End Function

' 部屋番号を整形する（数値で読み込まれた場合に小数点を除去）
Private Function FormatRoom(ByVal v As Variant) As String
    If IsNumeric(v) And Not IsEmpty(v) And Trim(CStr(v)) <> "" Then
        If CDbl(v) = Int(CDbl(v)) Then
            FormatRoom = CStr(CLng(v))
            Exit Function
        End If
    End If
    FormatRoom = Trim(CStr(v))
End Function

' 配列を分単位で昇順ソートする（要素は (room, minute, option) の配列）
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
        "Excelファイル (*.xls;*.xlsx),*.xls;*.xlsx", , "入居者一覧を選択してください")
    If srcPath = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim srcWB As Workbook
    Set srcWB = Workbooks.Open(srcPath, ReadOnly:=True)

    Dim buildingName As String
    buildingName = CStr(srcWB.Sheets(1).Cells(1, 1).Value)

    ' --- データ収集 ---
    Dim dateKeys As Object, slotData As Object
    Set dateKeys = CreateObject("Scripting.Dictionary")  ' key: 月*100+日 -> 曜日
    Set slotData = CreateObject("Scripting.Dictionary")  ' key: (月*100+日)*100+時 -> Collection

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

                    If nameVal = W空室() Then
                        vacantList.Add room
                    ElseIf InStr(remarkVal, W工期外()) > 0 Then
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

    ' --- 工程表シート作成 ---
    Dim outWB As Workbook
    Set outWB = Workbooks.Add
    Dim sht As Worksheet
    Set sht = outWB.Sheets(1)
    sht.Name = ChrW(&H5DE5) & ChrW(&H7A0B) & ChrW(&H8868)  ' 工程表

    Const TOTAL_COLS As Long = 9
    Dim timeSlots(7) As Long
    timeSlots(0) = 9: timeSlots(1) = 10: timeSlots(2) = 11
    timeSlots(3) = 13: timeSlots(4) = 14: timeSlots(5) = 15: timeSlots(6) = 16: timeSlots(7) = 17

    Dim YELLOW As Long, ORANGE As Long, GRAY As Long
    YELLOW = RGB(255, 255, 0)
    ORANGE = RGB(255, 192, 0)
    GRAY = RGB(217, 217, 217)

    ' タイトル
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

    ' 注記
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

    ' 工事予定日
    With sht.Range(sht.Cells(headerRow1, 1), sht.Cells(headerRow2, 1))
        .Merge
        .Value = ChrW(&H5DE5) & ChrW(&H4E8B) & ChrW(&H4E88) & ChrW(&H5B9A) & ChrW(&H65E5)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 午前
    With sht.Range(sht.Cells(headerRow1, 2), sht.Cells(headerRow1, 4))
        .Merge
        .Value = ChrW(&H5348) & ChrW(&H524D) & ChrW(&HFF08) & "9" & ChrW(&H6642) & _
                 ChrW(&H301C) & "12" & ChrW(&H6642) & ChrW(&HFF09)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 午後
    With sht.Range(sht.Cells(headerRow1, 5), sht.Cells(headerRow1, TOTAL_COLS))
        .Merge
        .Value = ChrW(&H5348) & ChrW(&H5F8C) & ChrW(&HFF08) & "13" & ChrW(&H6642) & _
                 ChrW(&H301C) & "18" & ChrW(&H6642) & ChrW(&HFF09)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 時間帯ラベル
    Dim i As Long
    For i = 0 To 7
        With sht.Cells(headerRow2, 2 + i)
            .Value = timeSlots(i) & ChrW(&H6642) & ChrW(&H983C)  ' n時頃
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i

    With sht.Range(sht.Cells(headerRow1, 1), sht.Cells(headerRow2, TOTAL_COLS))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' --- データ行 ---
    Dim row As Long
    row = headerRow2 + 1

    ' 日付キーを昇順ソート
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

    ' 初日の前日を共用部工事日として追加
    If nKeys > 0 Then
        Dim firstMon As Long, firstDay As Long, firstWd As String
        firstMon = sortedKeys(0) \ 100
        firstDay = sortedKeys(0) Mod 100
        firstWd = dateKeys(sortedKeys(0))

        Dim wds As Variant
        wds = Weekdays()
        Dim wdIdx As Long, prevWdIdx As Long
        wdIdx = 0
        For a = 0 To 6
            If wds(a) = firstWd Then wdIdx = a
        Next a
        prevWdIdx = (wdIdx - 1 + 7) Mod 7

        Dim caDate As Date
        caDate = DateSerial(2001, firstMon, firstDay) - 1

        With sht.Range(sht.Cells(row, 1), sht.Cells(row, 1))
            .Merge
            .Value = Month(caDate) & W月() & Day(caDate) & W日() & ChrW(&HFF08) & wds(prevWdIdx) & ChrW(&HFF09)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With sht.Range(sht.Cells(row, 2), sht.Cells(row, TOTAL_COLS))
            .Merge
            .Value = ChrW(&H5171) & ChrW(&H3000) & ChrW(&H7528) & ChrW(&H3000) & ChrW(&H90E8) & _
                     ChrW(&H3000) & ChrW(&H5DE5) & ChrW(&H3000) & ChrW(&H4E8B) & Chr(10) & _
                     ChrW(&HFF08) & ChrW(&H304A) & ChrW(&H90E8) & ChrW(&H5C4B) & ChrW(&H306E) & _
                     ChrW(&H5DE5) & ChrW(&H7A0B) & ChrW(&H306F) & ChrW(&H51FA) & ChrW(&H6765) & _
                     ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&HFF09)
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
    End If

    ' 各日付の行
    For a = 0 To nKeys - 1
        Dim dkey2 As Long, mon2 As Long, day2 As Long, wd2 As String
        dkey2 = sortedKeys(a)
        mon2 = dkey2 \ 100
        day2 = dkey2 Mod 100
        wd2 = dateKeys(dkey2)

        ' この日の最大同時刻件数を求める
        Dim nRows As Long
        nRows = 1
        For i = 0 To 7
            Dim skey2 As Long
            skey2 = dkey2 * 100 + timeSlots(i)
            If slotData.Exists(skey2) Then
                If slotData(skey2).Count > nRows Then nRows = slotData(skey2).Count
            End If
        Next i

        ' 日付セル
        With sht.Range(sht.Cells(row, 1), sht.Cells(row + nRows - 1, 1))
            .Merge
            .Value = mon2 & W月() & day2 & W日() & ChrW(&HFF08) & wd2 & ChrW(&HFF09)
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
    Next a

    ' 列幅・行高
    sht.Columns(1).ColumnWidth = 16
    For i = 0 To 7
        sht.Columns(2 + i).ColumnWidth = 10
    Next i
    For r = headerRow2 + 1 To row - 1
        sht.Rows(r).RowHeight = 30
    Next r

    ' オプション凡例
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

    ' 工期外希望
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

    ' 未提出
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

    ' 空室
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

    ' --- 印刷設定（A4横1ページ） ---
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

    ' --- 保存 ---
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
