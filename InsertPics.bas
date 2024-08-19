Attribute VB_Name = "Module5"
Public colSize As Double, rowSize As Double, picTotal%, partTotal%, _
        massData As Range, nameData As Range, rowTotal As Integer, colTotal
'Total page height value 750
'Total page width value 120


Sub InsertImage()
    
    Dim myDialog As FileDialog, myFolder As String, _
        myFile As String, myImg As String
    Set myDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    rowTotal = Cells(14, 7).Value
    colTotal = 3
    
    If myDialog.Show = -1 Then
        Application.ScreenUpdating = False
        Application.Sheets("Masses").Activate
        Set massData = getActiveRange("C")
        Set nameData = getActiveRange("B")
        picTotal = 0
        partTotal = 0
        
        myFolder = myDialog.SelectedItems(1) & Application.PathSeparator
        colSize = 120 / colTotal
        rowSize = 750 / rowTotal
        myFile = Dir(myFolder)
        
        Dim currentRow%, currentCol%, columnArr(2) As String, currentImg As Shape
        columnArr(0) = "B"
        columnArr(1) = "H"
        columnArr(2) = "N"
        currentRow = 1
        currentCol = 0
        
        
        Application.Sheets("Images").Activate
        Call Reset
        Call formatPage(currentRow)
        currentRow = currentRow + 1
        
        Do While myFile <> ""
            picTotal = picTotal + 1
            If currentCol > 2 Then
                currentRow = currentRow + rowTotal + 1
                currentCol = 0
                partTotal = partTotal + 3
                Call formatPage(currentRow)
                currentRow = currentRow + 1
            End If
            
            Range(columnArr(currentCol) & CStr(currentRow)).Select
            
            With Selection
                x = .Left
                y = .Top
                w = .width
                h = .height
            End With
            
            myImg = myFolder & myFile
            Application.Sheets("Images").Shapes.AddPicture (myImg), _
                    msoFalse, msoTrue, x, y, -1, -1
            Application.Sheets("Images").Shapes.SelectAll
            Set currentImg = Selection.ShapeRange(picTotal)
            Call scalingImg(currentImg, w, h)
            Call centerImg(currentImg, x, y, w, h)
            myFile = Dir
            currentRow = currentRow + 1
            
            
            If currentRow Mod (rowTotal + 2) = 0 Then
                currentRow = currentRow - rowTotal
                currentCol = currentCol + 1
            End If
            
        Loop
        
        
    
    End If
    
    Application.ScreenUpdating = False
    Cells(1, 1).Select
    
End Sub

Sub Reset()

    ' Reset all formatting of the sheet back to standard
    Cells.Select
    With Selection
    
        .ColumnWidth = 8.11
        .RowHeight = 14.4
        .Value = ""
        .UnMerge
        
        .Borders.ThemeColor = 1
        .Borders.TintAndShade = -0.149998474074526
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    
    
        With .Interior
        
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
            .ThemeColor = xlThemeColorDark1
            
        End With
    End With
    
    ActiveSheet.Pictures.Delete

End Sub

Sub formatPage(ByVal start As Integer)

' Formats one row of cells for image entry, making the header, and 3 larger cells for entry
'Need to pass the ROW in which the labels start

' Starting with looping through each column of image entries
    For imgcol = 1 To 3
    
        'Creating a header row
        'Shading cells blue
        For i = 1 To 6
            
            Application.Sheets("Images").Activate
            Cells(start, (imgcol - 1) * 6 + i).Select
            If i <> 1 Then
                With Selection.Interior
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
                End With
            End If
            
            Select Case i
            Case 1
                Selection.ColumnWidth = colSize * 0.01
            Case 2
                Selection.ColumnWidth = colSize * 0.08
            Case 3
                Selection.ColumnWidth = colSize * 0.14
            Case 4
                Selection.ColumnWidth = colSize * 0.28
            Case 5
                Selection.ColumnWidth = colSize * 0.21
            Case 6
                Selection.ColumnWidth = colSize * 0.28
            End Select
            
        Next i
        
        Cells(start, (imgcol - 1) * 6 + 2).Value = "#" & CStr(partTotal + imgcol)
        Cells(start, (imgcol - 1) * 6 + 3).Value = "Name:"
        Cells(start, (imgcol - 1) * 6 + 4).Value = nameData(partTotal + imgcol)
        Cells(start, (imgcol - 1) * 6 + 5).Value = "Mass (kg):"
        Cells(start, (imgcol - 1) * 6 + 6).Value = massData(partTotal + imgcol)
        
        'Now formatting the cells that the images should be going into
        For imgRow = 1 To rowTotal
            Range(Cells(start + imgRow, (imgcol - 1) * 6 + 2), _
                    Cells(start + imgRow, imgcol * 6)).Select
            Selection.Merge
            With Selection
                .RowHeight = rowSize
            End With
            
        Next imgRow
        
    Next imgcol


End Sub

Sub scalingImg(ByRef img As Shape, ByVal width As Variant, ByVal height As Variant)
    
    img.LockAspectRatio = msoTrue
    img.Rotation = 0
    If ((img.height - height) / height) > ((img.width - width) / width) Then
        img.height = height
    Else
        img.width = width
    End If

End Sub

Sub centerImg(ByRef img As Shape, ByVal x As Double, _
                ByVal y As Double, ByVal w As Double, ByVal h As Double)
    
    Dim xPos As Double, yPos As Double
    
    xPos = x + ((w - img.width) / 2)
    yPos = y + ((h - img.height) / 2)
    
    img.Top = yPos
    img.Left = xPos
    
End Sub

Function getActiveRange(ByVal column As String) As Range
    Dim row As Integer, cellName As String, counter As Integer
    row = rowTotal
    
    If Range(column & CStr(row)).Value <> "" Then
    
        Do While Range(column & CStr(row)).Value <> ""
            row = row + 1
        Loop
        
    End If
    
    Set getActiveRange = Range(column & "3:" & column & CStr(row - 1))
    
End Function


