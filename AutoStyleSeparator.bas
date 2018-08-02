Attribute VB_Name = "AutoStyleSeparator"
Option Explicit

Public InsertionCount As Long

Public strApplyStyle As String

Sub AutomaticStyleSepForm()
'This Sub allows the Form to be called from the Macros Menu.
    StyleSeparatorForm1.Show
    
End Sub

Sub AutomaticStyleSeparatorOptions()
Dim intI As Integer
Dim arValues() As String

    arValues = fnShowSelections
    
    For intI = LBound(arValues) To UBound(arValues)
        RunHeadingAutoStyleSep arValues(intI)
    Next intI

 

    
End Sub


Sub RunHeadingAutoStyleSep(strFind As String)
'inserts the style separators to the specific headings based on the string passed.

Dim rng As Range
Dim rng1 As Range
Dim rngAll As Range
Dim rngParagraph As Range

'    Set rng = ActiveDocument.Range
    Set rngAll = ActiveDocument.Content
    Set rng = rngAll.Duplicate
    With rng.Find
        .ClearFormatting
        .Forward = True
        .Style = strFind
        Do While .Execute()
            rng.Select
           If Not rng.Paragraphs(1).IsStyleSeparator Then
                rng.MoveStartUntil Cset:="."
                rng.Collapse Direction:=wdCollapseStart
                rng.Select
                insertStyleSep
'                Application.Run MacroName:="LWmacros.basHNum.InsertStyleSeparator"
                InsertionCount = InsertionCount + 1
                Selection.Move Unit:=wdCharacter, count:=1
            Else
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub

Sub insertStyleSep()
Dim rgSel As Range
Dim rgNew As Range

    Set rgSel = Selection.Range
    
    rgSel.InsertParagraph
    Set rgNew = rgSel.Duplicate
    
    rgNew.Move wdParagraph, 1
    rgNew.Select
    
    If StyleExists = True Then
        rgNew.Style = ActiveDocument.Styles("BodyText 1")
    Else
        rgNew.Style = ActiveDocument.Styles(wdStyleBodyText)
    End If
    
    
    rgSel.Select
    Selection.InsertStyleSeparator
    
    rgSel.Collapse wdCollapseEnd
    rgSel.Delete
            
End Sub


Function StyleExists() As Boolean
    Dim t
    On Error Resume Next
    StyleExists = True
    Set t = ActiveDocument.Styles("BodyText 1")
    If Err.Number <> 0 Then StyleExists = False
    Err.Clear
End Function



Function fnGetStyles() As String()
Dim para As Paragraph, strStyle As String
Dim arStyles() As String

    With ActiveDocument
        For Each para In .Paragraphs
            If (VBA.InStr(strStyle, para.Style) = 0) Then
                strStyle = strStyle & para.Style & ","
            End If
        Next para
    End With
    
    strStyle = VBA.Left(strStyle, VBA.Len(strStyle) - 1)
    
    arStyles = VBA.Split(strStyle, ",")
        
    fnGetStyles = arStyles
    
End Function

'Function fnGetApplyStyles() As String()
'
'    For i = 0 To objStylesBox.ListCount - 1
'          If objStylesBox.Selected(i) Then
'              strApplyStyle = objStylesBox.List(i)
''          End If
''      Next i
'
'
'End Function
'

Function fnShowSelections() As String()
Dim arItems() As String
Dim intI As Integer, intNext As Integer

    ReDim arItems(0)
    
    With StyleSeparatorForm1
        With .objStylesBox
            For intI = 0 To .ListCount - 1
                If (.Selected(intI)) Then
                    intNext = fnNextElement(arItems)
                    arItems(intNext) = .List(intI)
                End If
            Next intI
        End With
    End With
    
    fnShowSelections = arItems
    
End Function

Function fnNextElement(arX) As Integer

    If (VBA.IsArray(arX)) Then
        If (arX(LBound(arX)) = "") Then
            fnNextElement = 0
        Else
            ReDim Preserve arX(UBound(arX) + 1)
            fnNextElement = UBound(arX)
        End If
    End If
    
End Function

Sub runNormal()

    InsertionCount = 0
        
    AutomaticStyleSeparatorOptions
    
    Unload StyleSeparatorForm1
    
    'Custom messagebox text per amount of style separators inserted
    If AutoStyleSeparator.InsertionCount > 1 Then
        MsgBox AutoStyleSeparator.InsertionCount & " Style Separators Inserted"
    ElseIf AutoStyleSeparator.InsertionCount = 1 Then
        MsgBox AutoStyleSeparator.InsertionCount & " Style Separator Inserted"
    ElseIf AutoStyleSeparator.InsertionCount = 0 Then
        MsgBox "No Style Separators Inserted"
    End If
    
    End
    
End Sub

'Sub runSaveSelected()
'
'Dim userResponse As Boolean
'Dim strDocName As String
'
'''put current file name into strDocName variable?
'
'On Error Resume Next
'userResponse = Application.Dialogs(wdDialogFileSaveAs).Show("strDocName")
'On Error GoTo 0
'
'    ''Stop running Macro if Dialog box is cancelled:
'    If userResponse = False Then
'        Unload StyleSeparatorForm1
'
'    ''Runs normally if file is saved:
'    Else
'        runNormal
'    End If
'
'End Sub

Function SaveCopyAs() As Boolean
    Const lCancelled_c As Long = 0
    Dim sSaveAsPath As String
    sSaveAsPath = GetSaveAsPath
    If (VBA.Len(sSaveAsPath) = 0) Then
        SaveCopyAs = False
        Exit Function
    Else
        SaveCopyAs = True
    End If
    
    If VBA.LenB(sSaveAsPath) = lCancelled_c Then Exit Function
     'Save changes to original document
    ActiveDocument.Save
     'the next line copies the active document
    Application.Documents.Add ActiveDocument.FullName
     'the next line saves the copy to your location and name
    ActiveDocument.SaveAs sSaveAsPath
     'next line closes the copy leaving you with the original document
    ActiveDocument.Close

End Function

Public Function GetSaveAsPath() As String
    Dim fd As Office.FileDialog
    Set fd = Word.Application.FileDialog(msoFileDialogSaveAs)
    fd.InitialFileName = ActiveDocument.FullName
    If fd.Show Then GetSaveAsPath = fd.SelectedItems(1)
End Function
