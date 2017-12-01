Private Sub CommandButton1_Click()

Call AutomaticStyleSeparatorOptions
    
    Unload StyleSeparatorForm1
    
    If AutoStyleSeparator.InsertionCount > 1 Then
        MsgBox AutoStyleSeparator.InsertionCount & " Style Separators Inserted"
    ElseIf AutoStyleSeparator.InsertionCount = 1 Then
        MsgBox AutoStyleSeparator.InsertionCount & " Style Separator Inserted"
    ElseIf AutoStyleSeparator.InsertionCount = 0 Then
        MsgBox "No Style Separators Inserted"
    End If
    
AutoStyleSeparator.InsertionCount = 0
    
End Sub

Private Sub closeButton_Click()
Unload StyleSeparatorForm1
End Sub


Sub sbPositionForm()
' position form in middle of Oulook window
Dim lngLeft As LongPtr, lngTop As LongPtr
Dim lngWidth As LongPtr, lngHeight As LongPtr
Dim lngFrmWidth As LongPtr, lngFrmHeight As LongPtr

    ' grab Outlook main window stuff
    With Outlook.Application.ActiveExplorer
        lngLeft = .Left
        lngTop = .Top
        lngWidth = .Width
        lngHeight = .Height
    End With

    ' grab form stuff
    lngFrmWidth = Me.Width
    lngFrmHeight = Me.Height

    ' set values of left and top
    lngLeft = (lngWidth / 2) - (lngFrmWidth / 2)
    lngTop = (lngHeight / 2) - (lngFrmHeight / 2)

    ' position the from
    Me.Move lngLeft, lngTop

End Sub
