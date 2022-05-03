Attribute VB_Name = "ShapesLoop"
Option Explicit
' ----------------------------------------------------------------
' Purpose: Adds buttons, to all the procedures in all the modules of the project
' ----------------------------------------------------------------
Sub RunCreateButton()
    Dim topPosition As Long
    Dim arry() As String
    Dim k As Long
    On Error Resume Next
arry() = ProceduresListPrint()
    For k = LBound(arry) To UBound(arry)
       Call CreateButton(10, topPosition, 120, 30, arry(k), arry(k), arry(k))
        topPosition = topPosition + 40
    Next
End Sub

' ----------------------------------------------------------------
' Purpose: called by RunCreateButton
' ----------------------------------------------------------------
Sub CreateButton(left As Long, top As Long, width As Long, height As Long, BtnName As String, caption As String, MacroName As String)
    On Error Resume Next
        ActiveSheet.Buttons(BtnName).name = BtnName
'Assuret that the name for the button is unique.
        If Err.Number = 0 Then
            MsgBox "Button names have to be unique"
            End
        End If
    On Error GoTo 0

    With ActiveSheet.Buttons.Add(left, top, width, height)
        .name = BtnName
        .caption = caption
        .OnAction = MacroName
    End With
End Sub

'' ----------------------------------------------------------------
'' Purpose: Loop through all shapes in a active sheet and delete them
'' ----------------------------------------------------------------
Sub DeleteShapes()
'ActiveSheet.Buttons.Delete
    Dim shp As Shape
    For Each shp In ThisWorkbook.ActiveSheet.Shapes
    If shp.name <> "CreateButtons" And shp.name <> "DeleteButtons" And shp.name <> "TextBox 8" Then
        shp.Delete
    End If
    Next shp
End Sub

''https://docs.microsoft.com/en-us/office/vba/api/Office.MsoShapeType

