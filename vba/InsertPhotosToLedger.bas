Attribute VB_Name = "InsertPhotosToLedger"
Option Explicit

' =====================================================================
' 写真台帳（インターホン設備改修工事）写真自動貼り付けマクロ
'
' 使い方:
'   1. 写真台帳 (.xlsm) を開く
'   2. Alt+F11 でVBエディタを開く
'   3. ファイル → ファイルのインポート → InsertPhotosToLedger.bas を選択
'   4. Alt+F8 で「InsertPhotosToLedger」を選んで「実行」
'   5. 写真フォルダを選択する
'
' 貼り付けルール:
'   写真をファイル名順に並べ、各シートの施工前→施工中→施工後の順に割り当てる
'   シート1の施工前, シート1の施工中, シート1の施工後,
'   シート2の施工前, ... という順で貼り付ける
'
' 写真スロット（各シート共通）:
'   施工前: セル A1 (結合 A1:A13)
'   施工中: セル A15 (結合 A15:A27)
'   施工後: セル A29 (結合 A29:A41)
' =====================================================================

' --- 設定 ---
Private Const PHOTO_ROW_BEFORE  As Long = 1    ' 施工前の写真行
Private Const PHOTO_ROW_DURING  As Long = 15   ' 施工中の写真行
Private Const PHOTO_ROW_AFTER   As Long = 29   ' 施工後の写真行
Private Const PHOTO_COL         As Long = 1    ' 写真を置く列 (A=1)

' 画像ファイルの拡張子（カンマ区切り）
Private Const IMAGE_EXTS As String = "jpg,jpeg,png,gif,bmp,tif,tiff"

' -----------------------------------------------------------------------
Public Sub InsertPhotosToLedger()

    ' 写真フォルダを選択
    Dim folderPath As String
    folderPath = SelectFolder("写真フォルダを選択してください")
    If folderPath = "" Then Exit Sub

    ' 画像ファイルを収集
    Dim files() As String
    Dim fileCount As Long
    fileCount = CollectImages(folderPath, files)

    If fileCount = 0 Then
        MsgBox "指定フォルダに画像ファイルが見つかりませんでした。" & vbCrLf & folderPath, vbExclamation, "写真台帳"
        Exit Sub
    End If

    ' 対象シート（非表示シートを除く）を収集
    Dim sheets() As Worksheet
    Dim sheetCount As Long
    sheetCount = CollectVisibleSheets(sheets)

    If sheetCount = 0 Then
        MsgBox "対象シートが見つかりません。", vbExclamation, "写真台帳"
        Exit Sub
    End If

    ' スロット数 = シート数 × 3
    Dim totalSlots As Long
    totalSlots = sheetCount * 3

    Dim msg As String
    msg = "写真: " & fileCount & " 枚" & vbCrLf
    msg = msg & "シート: " & sheetCount & " 枚" & vbCrLf
    msg = msg & "スロット: " & totalSlots & " 箇所（施工前・施工中・施工後 × " & sheetCount & " シート）" & vbCrLf & vbCrLf
    msg = msg & "ファイル名順に各スロットへ貼り付けます。" & vbCrLf & "よろしいですか？"

    If MsgBox(msg, vbOKCancel Or vbQuestion, "写真台帳 - 確認") <> vbOK Then Exit Sub

    ' 貼り付け実行
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim photoIdx As Long
    photoIdx = 0

    Dim phaseRows(2) As Long
    phaseRows(0) = PHOTO_ROW_BEFORE
    phaseRows(1) = PHOTO_ROW_DURING
    phaseRows(2) = PHOTO_ROW_AFTER

    Dim si As Long, pi As Long
    Dim pasteCount As Long
    pasteCount = 0

    For si = 0 To sheetCount - 1
        For pi = 0 To 2
            If photoIdx >= fileCount Then GoTo Done
            Call PastePhoto(sheets(si), files(photoIdx), PHOTO_COL, phaseRows(pi))
            photoIdx = photoIdx + 1
            pasteCount = pasteCount + 1
        Next pi
    Next si

Done:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "完了しました。" & pasteCount & " 枚の写真を貼り付けました。", vbInformation, "写真台帳"

End Sub

' -----------------------------------------------------------------------
' 写真1枚をセル (row, col) に貼り付ける（既存写真は削除）
Private Sub PastePhoto(ws As Worksheet, filePath As String, col As Long, row As Long)
    Dim topCell As Range
    Set topCell = ws.Cells(row, col)

    ' 同位置の既存の写真を削除
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then
            If Not Intersect(shp.TopLeftCell, topCell) Is Nothing Then
                shp.Delete
            End If
        End If
    Next shp

    ' セルの範囲（結合セル全体）を取得して写真を収める
    Dim targetArea As Range
    Set targetArea = topCell.MergeArea

    Dim cellLeft   As Double: cellLeft   = targetArea.Left   + 3
    Dim cellTop    As Double: cellTop    = targetArea.Top    + 3
    Dim cellWidth  As Double: cellWidth  = targetArea.Width  - 6
    Dim cellHeight As Double: cellHeight = targetArea.Height - 6

    ' 写真を挿入
    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture( _
        Filename:=filePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=cellLeft, _
        Top:=cellTop, _
        Width:=cellWidth, _
        Height:=cellHeight)

    ' アスペクト比を保ちつつセル内に収める
    pic.LockAspectRatio = msoTrue
    pic.Width = cellWidth
    If pic.Height > cellHeight Then
        pic.Height = cellHeight
        pic.LockAspectRatio = msoFalse
        pic.Width = cellWidth
        pic.Height = cellHeight
    End If

    ' 中央に配置
    pic.Left = cellLeft + (cellWidth - pic.Width) / 2
    pic.Top  = cellTop  + (cellHeight - pic.Height) / 2

    pic.Name = "Photo_" & ws.Name & "_r" & row
End Sub

' -----------------------------------------------------------------------
' フォルダ選択ダイアログ
Private Function SelectFolder(title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = title
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        SelectFolder = fd.SelectedItems(1)
    Else
        SelectFolder = ""
    End If
End Function

' -----------------------------------------------------------------------
' 指定フォルダ内の画像ファイルをファイル名順で収集して返す
Private Function CollectImages(folderPath As String, ByRef files() As String) As Long
    Dim exts() As String
    exts = Split(IMAGE_EXTS, ",")

    ' まず件数をカウント
    Dim count As Long
    count = 0
    Dim ext As Variant
    Dim f As String
    For Each ext In exts
        f = Dir(folderPath & "\" & "*." & Trim(CStr(ext)))
        Do While f <> ""
            count = count + 1
            f = Dir()
        Loop
    Next ext

    If count = 0 Then
        CollectImages = 0
        Exit Function
    End If

    ReDim files(count - 1)
    Dim idx As Long
    idx = 0
    For Each ext In exts
        f = Dir(folderPath & "\" & "*." & Trim(CStr(ext)))
        Do While f <> ""
            files(idx) = folderPath & "\" & f
            idx = idx + 1
            f = Dir()
        Loop
    Next ext

    ' バブルソート（ファイル名昇順）
    Dim i As Long, j As Long, tmp As String
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If LCase(files(i)) > LCase(files(j)) Then
                tmp = files(i): files(i) = files(j): files(j) = tmp
            End If
        Next j
    Next i

    CollectImages = count
End Function

' -----------------------------------------------------------------------
' 非表示シートを除く全シートを収集
Private Function CollectVisibleSheets(ByRef sheets() As Worksheet) As Long
    Dim count As Long
    count = 0
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        If ws.Visible = xlSheetVisible Then
            count = count + 1
        End If
    Next ws

    If count = 0 Then
        CollectVisibleSheets = 0
        Exit Function
    End If

    ReDim sheets(count - 1)
    Dim idx As Long
    idx = 0
    For Each ws In ActiveWorkbook.Sheets
        If ws.Visible = xlSheetVisible Then
            sheets(idx) = ws
            idx = idx + 1
        End If
    Next ws

    CollectVisibleSheets = count
End Function
