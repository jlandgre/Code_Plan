Option Explicit
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' A procedure
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' A docstring for a procedure
'
' JDL 12/13/21   Modified: 8/1/23 JDL
'
Function ExampleProcedure(cls, ByVal arg1, Optional arg2) As Boolean

    SetErrorHandle ExampleProcedure: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer, j As Integer
    
    With cls
        If Not .Method1(cls, arg1) Then GoTo ErrorExit
        If Not .Method2(cls, i, j, arg2) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "ExampleProcedure", ExampleProcedure
End Function
'---------------------------------------------------------------------------------------
' Method1 docstring is
' multiline
'
' JDL 8/1/23
'
Function Method1(cls, arg1) As Boolean

    SetErrorHandle Method1: If errs.IsHandle Then On Error GoTo ErrorExit
    
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "Method1", Method1
End Function
'---------------------------------------------------------------------------------------
' Method2 docstring
'
' JDL 8/1/23
'
Function Method2(cls, i As Integer, j As Integer, arg2)

    SetErrorHandle Method2: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim R As Variant, R_MI As Object
    
    i = 2 + 3
    j = 6
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "Method2", Method2
End Function
