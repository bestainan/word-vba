
Sub addTable(position As String, rowNum As Integer, colNum As Integer, tableContent As String)
    
    Dim lib As New jsonlib
    MsgBox tableContent
    Set myRange = ActiveDocument.Content

    myRange.Find.Execute FindText:=position
    If myRange.Find.Found = True Then myRange.text = ""
    
    myRange.Collapse Direction:=wdCollapseEnd
    Set TableNew = ActiveDocument.Tables.Add(Range:=myRange, NumRows:=rowNum, NumColumns:=colNum)
    With TableNew
        For intX = 1 To rowNum
        For intY = 1 To colNum
        .Cell(intX, intY).Range.InsertAfter "Cell: R" & intX & ", C" & intY
        Next intY
        Next intX
        .Columns.AutoFit
    End With
End Sub