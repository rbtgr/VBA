Attribute VB_Name = "commentMacro"
Sub �R�����g�̈ʒu�C���}�N��()
Attribute �R�����g�̈ʒu�C���}�N��.VB_ProcData.VB_Invoke_Func = " \n14"
    
    For Each CL In Selection
    
        If Not CL.Comment Is Nothing Then
            CL.Comment.Shape.Top = CL.Top + 10
            CL.Comment.Shape.Height = 50
            'CL.Comment.Shape.TextFrame.AutoSize = True
            
            CL.Comment.Shape.Placement = xlMove
            CL.Comment.Shape.Placement = xlMoveAndSize
            CL.Comment.Shape.Placement = xlFreeFloating
         
        End If
    Next CL
    
    
End Sub
