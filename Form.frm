VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StyleSeparatorForm1 
   Caption         =   "Select Heading Levels"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2985
   OleObjectBlob   =   "StyleSeparatorForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StyleSeparatorForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit


Private Sub btnApply_Click()
     
    Me.Hide
    
    If Me.chkSaveAs Then
        'Nested If below unloads if "Save As" dialog box is cancelled
        If SaveCopyAs Then
            runNormal
        
        Else
            Unload Me
            End
        
        End If
    
    Else
        runNormal
    End If
    
   
End Sub

Private Sub closeButton_Click()
    Unload StyleSeparatorForm1
End Sub

Sub sbPositionForm()
' position form in middle of window
Dim lngLeft As LongPtr, lngTop As LongPtr
Dim lngWidth As LongPtr, lngHeight As LongPtr
Dim lngFrmWidth As LongPtr, lngFrmHeight As LongPtr

    ' grab main window stuff
    With ActiveDocument.Windows(1)
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

Private Sub objStylesBox_Change()

    Me.btnApply.Enabled = fnEnableDisableApply
    
End Sub

Function fnEnableDisableApply() As Boolean
Dim intI As Integer

    fnEnableDisableApply = False
    
    With Me.objStylesBox
        For intI = 0 To .ListCount - 1
            If (.Selected(intI)) Then
                fnEnableDisableApply = True
                Exit Function
            End If
        Next intI
    End With

End Function

Private Sub UserForm_Initialize()
Dim intI As Integer, arStyles() As String

    arStyles = fnGetStyles
    
    For intI = LBound(arStyles) To UBound(arStyles)
        objStylesBox.AddItem arStyles(intI)
    Next intI

End Sub


