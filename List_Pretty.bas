Attribute VB_Name = "List_Pretty"


Sub 上と同じセルの上罫線無し灰文字化()
    Dim Row As Long
    Dim Clm As Long
    
    For Row = Selection(1).Row To Selection(Selection.Count).Row
        For Clm = Selection(1).Column To Selection(Selection.Count).Column
            If Cells(Row, Clm).Value <> Cells(Row, Clm).Offset(-1, 0).Value Then
                Exit For
            Else
                Cells(Row, Clm).Font.ColorIndex = 15
                Cells(Row, Clm).Borders(xlEdgeTop).LineStyle = xlNone
            End If    
        Next Clm
    Next Row
End Sub
