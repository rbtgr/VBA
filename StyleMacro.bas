Attribute VB_Name = "StyleMacro"
Sub 使っていないスタイルを削除()
    Dim WB As Workbook
    Dim tmpWB As Workbook
    Dim StyleDic As Object
    Set StyleDic = CreateObject("scripting.dictionary")
    
    Set WB = ActiveWorkbook
    Set tmpWB = Workbooks.Add
    WB.Activate
    WB.Sheets.Copy Before:=tmpWB.Sheets(1)
    tmpWB.Activate
    
    For Each STL In tmpWB.Styles
       StyleDic.Add STL.Name, STL.Name
    Next STL
    
    WB.Activate
    
    For Each STL In WB.Styles
        If Not StyleDic.Exists(STL.Name) Then
            Debug.Print STL.Name
            STL.Delete
        End If
    Next STL

    tmpWB.Close SaveChanges:=False
    Set tmpWB = Nothing
    
End Sub

