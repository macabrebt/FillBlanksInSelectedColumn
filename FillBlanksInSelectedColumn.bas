Attribute VB_Name = "FillBlanks"
Sub FillBlanksInSelectedColumn()
    ' This macro fills blank cells in the currently selected column
    ' with the value from the non-blank cell directly above it.

    Dim rSelectedColumn As Range
    Dim rCell As Range
    Dim vPreviousValue As Variant

    ' Check if a column (or part of a column) is selected.
    ' If not, inform the user and exit.
    If Selection.Columns.Count > 1 Then
        MsgBox "Please select only one column to fill blanks.", vbExclamation, "Multiple Columns Selected"
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range within a column.", vbExclamation, "No Range Selected"
        Exit Sub
    End If

    Set rSelectedColumn = Selection ' Get the currently selected range

    Application.ScreenUpdating = False ' Turn off screen updating for faster execution

    ' Loop through each cell in the selected column, starting from the second row
    ' of the selection, as the first cell will be the 'previous value' source.
    For Each rCell In rSelectedColumn.Cells
        If rCell.Row = rSelectedColumn.Row Then
            ' This is the first cell in the selection.
            ' Set the initial previous value.
            vPreviousValue = rCell.Value
        Else
            ' For subsequent cells:
            If IsEmpty(rCell.Value) Then
                ' If the cell is blank, fill it with the previous non-blank value.
                rCell.Value = vPreviousValue
            Else
                ' If the cell is not blank, update the previous value.
                vPreviousValue = rCell.Value
            End If
        End If
    Next rCell

    Application.ScreenUpdating = True ' Turn screen updating back on

    MsgBox "Blank cells filled successfully in the selected column!", vbInformation, "Macro Complete"

End Sub
