Attribute VB_Name = "Module1"
Sub SmartFillDown()
    Dim rng As Range, n As Long
    Set rng = ActiveCell.Offset(0, -1).CurrentRegion
    If rng.Cells.Count > 1 Then
        n = rng.Cells(1).Row + rng.Rows.Count - ActiveCell.Row
        ActiveCell.AutoFill Destination:=ActiveCell.Resize(n, 1), Type:=xlFillValues
    End If
End Sub

Sub SmartFillRight()
    Dim rng As Range, n As Long
    Set rng = ActiveCell.Offset(-1, 0).CurrentRegion
    If rng.Cells.Count > 1 Then
        n = rng.Cells(1).Column + rng.Columns.Count - ActiveCell.Column
        ActiveCell.AutoFill Destination:=ActiveCell.Resize(1, n), Type:=xlFillValues
    End If
End Sub
