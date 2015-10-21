Sub insertImg(position As String, img_path As String)
    
    Set myRange = ActiveDocument.Content
    myRange.Find.Execute FindText:=position
    If myRange.Find.Found = True Then myRange.text = ""
    myRange.Collapse Direction:=wdCollapseEnd

    Set MyPic = Selection.InlineShapes.AddPicture(Range:=myRange, FileName:=img_path, SaveWithDocument:=True)
    
End Sub

Sub test()
    insertImg "{{12img}}", "C:\Users\Robot\Desktop\1.png"
End Sub
