Attribute VB_Name = "InsertPhotosToLedger"
Option Explicit

' =====================================================================
' Photo Ledger - Auto Insert Photos
'
' Usage:
'   1. Open the photo ledger (.xlsm)
'   2. Alt+F11 -> File -> Import File -> select InsertPhotosToLedger.bas
'   3. Alt+F8 -> run "InsertPhotosToLedger"
'   4. Select the photo folder
'
' Photos are inserted in filename order into each sheet's slots:
'   Before(A1) -> During(A15) -> After(A29), then next sheet, etc.
' =====================================================================

Private Const PHOTO_ROW_BEFORE  As Long = 1
Private Const PHOTO_ROW_DURING  As Long = 15
Private Const PHOTO_ROW_AFTER   As Long = 29
Private Const PHOTO_COL         As Long = 1

Private Const IMAGE_EXTS As String = "jpg,jpeg,png,gif,bmp,tif,tiff"

' -----------------------------------------------------------------------
Public Sub InsertPhotosToLedger()

    Dim folderPath As String
    folderPath = SelectFolder()
    If folderPath = "" Then Exit Sub

    Dim files() As String
    Dim fileCount As Long
    fileCount = CollectImages(folderPath, files)

    If fileCount = 0 Then
        MsgBox "No image files found in:" & vbCrLf & folderPath, vbExclamation, "Photo Ledger"
        Exit Sub
    End If

    Dim sheets() As Worksheet
    Dim sheetCount As Long
    sheetCount = CollectVisibleSheets(sheets)

    If sheetCount = 0 Then
        MsgBox "No visible sheets found.", vbExclamation, "Photo Ledger"
        Exit Sub
    End If

    Dim msg As String
    msg = "Photos : " & fileCount & vbCrLf
    msg = msg & "Sheets : " & sheetCount & vbCrLf
    msg = msg & "Slots  : " & sheetCount * 3 & " (Before / During / After x " & sheetCount & " sheets)" & vbCrLf & vbCrLf
    msg = msg & "Insert photos in filename order?"

    If MsgBox(msg, vbOKCancel Or vbQuestion, "Photo Ledger") <> vbOK Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim phaseRows(2) As Long
    phaseRows(0) = PHOTO_ROW_BEFORE
    phaseRows(1) = PHOTO_ROW_DURING
    phaseRows(2) = PHOTO_ROW_AFTER

    Dim photoIdx As Long
    Dim pasteCount As Long
    Dim si As Long, pi As Long
    photoIdx = 0
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

    MsgBox "Done. " & pasteCount & " photo(s) inserted.", vbInformation, "Photo Ledger"

End Sub

' -----------------------------------------------------------------------
Private Sub PastePhoto(ws As Worksheet, filePath As String, col As Long, row As Long)
    Dim topCell As Range
    Set topCell = ws.Cells(row, col)

    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then
            If Not Intersect(shp.TopLeftCell, topCell) Is Nothing Then
                shp.Delete
            End If
        End If
    Next shp

    Dim targetArea As Range
    Set targetArea = topCell.MergeArea

    Dim cellLeft   As Double: cellLeft   = targetArea.Left   + 3
    Dim cellTop    As Double: cellTop    = targetArea.Top    + 3
    Dim cellWidth  As Double: cellWidth  = targetArea.Width  - 6
    Dim cellHeight As Double: cellHeight = targetArea.Height - 6

    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture( _
        Filename:=filePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=cellLeft, _
        Top:=cellTop, _
        Width:=cellWidth, _
        Height:=cellHeight)

    pic.LockAspectRatio = msoTrue
    pic.Width = cellWidth
    If pic.Height > cellHeight Then
        pic.Height = cellHeight
        pic.LockAspectRatio = msoFalse
        pic.Width = cellWidth
        pic.Height = cellHeight
    End If

    pic.Left = cellLeft + (cellWidth - pic.Width) / 2
    pic.Top  = cellTop  + (cellHeight - pic.Height) / 2
    pic.Name = "Photo_" & ws.Name & "_r" & row
End Sub

' -----------------------------------------------------------------------
Private Function SelectFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select Photo Folder"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        SelectFolder = fd.SelectedItems(1)
    Else
        SelectFolder = ""
    End If
End Function

' -----------------------------------------------------------------------
Private Function CollectImages(folderPath As String, ByRef files() As String) As Long
    Dim exts() As String
    exts = Split(IMAGE_EXTS, ",")

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
Private Function CollectVisibleSheets(ByRef sheets() As Worksheet) As Long
    Dim count As Long
    count = 0
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        If ws.Visible = xlSheetVisible Then count = count + 1
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
