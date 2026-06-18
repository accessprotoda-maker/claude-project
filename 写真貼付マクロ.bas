Attribute VB_Name = "Module1"
'===========================================
' 写真貼付マクロ
'===========================================

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
