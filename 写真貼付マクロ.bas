Attribute VB_Name = "Module1"
'===========================================
' 写真貼付マクロ
' 選択したセル（結合セル）に写真を貼り付け、
' セルサイズに合わせて自動リサイズします
'===========================================

Sub 写真貼付()
    Dim filePath As String
    Dim pic As Shape
    Dim targetCell As Range
    Dim cellLeft As Double, cellTop As Double
    Dim cellWidth As Double, cellHeight As Double
    Dim picRatio As Double, cellRatio As Double
    Dim newWidth As Double, newHeight As Double
    Dim offsetX As Double, offsetY As Double

    Set targetCell = Selection.MergeArea

    filePath = Application.GetOpenFilename( _
        FileFilter:="画像ファイル (*.jpg;*.jpeg;*.png;*.bmp;*.gif),*.jpg;*.jpeg;*.png;*.bmp;*.gif", _
        Title:="貼り付ける写真を選択してください")

    If filePath = "False" Then Exit Sub

    cellLeft = targetCell.Left
    cellTop = targetCell.Top
    cellWidth = targetCell.Width
    cellHeight = targetCell.Height

    Set pic = ActiveSheet.Shapes.AddPicture( _
        Filename:=filePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=cellLeft, _
        Top:=cellTop, _
        Width:=-1, _
        Height:=-1)

    picRatio = pic.Width / pic.Height
    cellRatio = cellWidth / cellHeight

    Dim margin As Double
    margin = 3

    Dim availW As Double, availH As Double
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

    offsetX = (cellWidth - newWidth) / 2
    offsetY = (cellHeight - newHeight) / 2
    pic.Left = cellLeft + offsetX
    pic.Top = cellTop + offsetY

    pic.Placement = xlMoveAndSize

    targetCell.Value = ""
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
