Attribute VB_Name = "ToggleImagesVisibility"
Sub ToggleImagesVisibility()
    Dim i As Integer
    
    ' Toggle visibility for inline images
    For i = ActiveDocument.InlineShapes.Count To 1 Step -1
        If ActiveDocument.InlineShapes(i).Type = wdInlineShapePicture Then
            If ActiveDocument.InlineShapes(i).Range.Font.Hidden = False Then
                ActiveDocument.InlineShapes(i).Range.Font.Hidden = True ' Hide the image
            Else
                ActiveDocument.InlineShapes(i).Range.Font.Hidden = False ' Unhide the image
            End If
        End If
    Next i
    
    ' Toggle visibility for floating images
    For i = ActiveDocument.Shapes.Count To 1 Step -1
        If ActiveDocument.Shapes(i).Type = msoPicture Then
            If ActiveDocument.Shapes(i).Visible = True Then
                ActiveDocument.Shapes(i).Visible = False ' Hide the image
            Else
                ActiveDocument.Shapes(i).Visible = True ' Unhide the image
            End If
        End If
    Next i
    
    MsgBox "All images have been toggled!"
End Sub
