Attribute VB_Name = "Module1"
'===========================================
' 写真貼付・部屋番号管理マクロ
'===========================================

Public currentPage As Long

Sub 写真貼付()
    Dim filePath As String
    Dim pic As Shape
    Dim targetCell As Range
    Dim cellLeft As Double, cellTop As Double
    Dim cellWidth As Double, cellHeight As Double
    Dim newWidth As Double, newHeight As Double
    Dim picRatio As Double
    Dim margin As Double
    Dim availW As Double, availH As Double

    On Error Resume Next
    Set targetCell = Selection.MergeArea
    If targetCell Is Nothing Then
        Set targetCell = Selection
    End If
    On Error GoTo 0

    filePath = Application.GetOpenFilename( _
        FileFilter:="画像ファイル (*.jpg;*.jpeg;*.png;*.bmp;*.gif),*.jpg;*.jpeg;*.png;*.bmp;*.gif", _
        Title:="貼り付ける写真を選択してください")

    If filePath = "False" Or filePath = "" Then Exit Sub

    cellLeft = targetCell.Left
    cellTop = targetCell.Top
    cellWidth = targetCell.Width
    cellHeight = targetCell.Height

    Set pic = ActiveSheet.Shapes.AddPicture( _
        Filename:=filePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=0, _
        Top:=0, _
        Width:=-1, _
        Height:=-1)

    picRatio = pic.Width / pic.Height
    margin = 3
    availW = cellWidth - margin * 2
    availH = cellHeight - margin * 2

    If picRatio > (availW / availH) Then
        newWidth = availW
        newHeight = newWidth / picRatio
    Else
        newHeight = availH
        newWidth = newHeight * picRatio
    End If

    pic.LockAspectRatio = msoTrue
    pic.Width = newWidth
    pic.Height = newHeight
    pic.Left = cellLeft + (cellWidth - pic.Width) / 2
    pic.Top = cellTop + (cellHeight - pic.Height) / 2
    pic.Placement = xlMoveAndSize

    On Error Resume Next
    targetCell.Value = ""
    On Error GoTo 0
End Sub

Sub 部屋番号セット()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim roomList() As String
    Dim roomCount As Long
    Dim i As Long
    Dim r As Long, c As Long
    Dim cellVal As String
    
    roomCount = 0
    
    ' E-H列の4行目以降から部屋番号を収集
    For c = 5 To 8  ' E=5, F=6, G=7, H=8
        For r = 4 To 100
            cellVal = Trim(CStr(ws.Cells(r, c).Value))
            If cellVal = "" Or cellVal = "0" Then
                If c < 8 Then Exit For
                If c = 8 Then Exit For
            End If
            roomCount = roomCount + 1
            ReDim Preserve roomList(1 To roomCount)
            roomList(roomCount) = cellVal
        Next r
    Next c
    
    If roomCount = 0 Then
        MsgBox "部屋番号リスト（E-H列）が空です。", vbExclamation
        Exit Sub
    End If
    
    ' 号室セルの位置（C2, C12, C22 = 3住戸分）
    Dim unitRows(1 To 3) As Long
    unitRows(1) = 2
    unitRows(2) = 12
    unitRows(3) = 22
    
    If currentPage = 0 Then currentPage = 1
    
    Dim startIdx As Long
    startIdx = (currentPage - 1) * 3 + 1
    
    Dim totalPages As Long
    totalPages = Application.WorksheetFunction.RoundUp(roomCount / 3, 0)
    
    If startIdx > roomCount Then
        MsgBox "全ての部屋番号をセット済みです。（" & totalPages & "ページ中 最終）", vbInformation
        Exit Sub
    End If
    
    For i = 1 To 3
        Dim idx As Long
        idx = startIdx + i - 1
        If idx <= roomCount Then
            ws.Cells(unitRows(i), 3).Value = roomList(idx)
        Else
            ws.Cells(unitRows(i), 3).Value = ""
        End If
    Next i
    
    MsgBox "ページ " & currentPage & " / " & totalPages & vbCrLf & _
           "（部屋 " & startIdx & "～" & Application.WorksheetFunction.Min(startIdx + 2, roomCount) & " / " & roomCount & "）", vbInformation
    
    currentPage = currentPage + 1
End Sub

Sub 部屋番号リセット()
    currentPage = 1
    MsgBox "ページを1に戻しました。", vbInformation
End Sub

Sub 全写真削除()
    Dim shp As Shape
    Dim count As Long
    count = 0
    For Each shp In ActiveSheet.Shapes
        If shp.Type = msoPicture Then
            shp.Delete
            count = count + 1
        End If
    Next shp
    MsgBox count & " 枚の写真を削除しました。", vbInformation
End Sub
