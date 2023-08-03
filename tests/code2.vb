Option Explicit
'---------------------------------------------------------------------------------------
' Function with multiline statements to test VBAToCodePlan.combine_split_lines()
'
' JDL 8/1/23   Modified: 8/1/23 JDL
'
Function ExampleProcedure(cls, ByVal arg1, Optional arg2) As Boolean

    SetErrorHandle ExampleProcedure: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer, j As Integer
    
    With cls
        If Not _
            .Method1(cls, arg1) _
            Then GoTo ErrorExit

    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, _
            "ExampleProcedure", ExampleProcedure
End Function