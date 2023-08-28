'Version 6/30/22 - refactor mdlRow; add ListValidation option
'Version 9/21/22 - change to .Columns.Ungroup to avoid clearing row outline
'version 10/26/22 - make SetRngFormulaRows robust to Excel errors in Text-formatted cells
'Version 3/8/23 - TransferToMdl validated
'Version 3/15/23 - SwapModels validated
'Version 6/8/23 - Add wksht attribute
'Version 6/27/23 - Switch to use of .sModelName within SwapModels
'Version 7/19/23 - Init function called by Provision and callable separately
'Version 8/28/23 - Refactoring and validation of SwapModels Procedures

Option Explicit
Public cellHome As Range
Public rngRows As Range
Public rngHeader As Range
Public wkbk As Workbook
Public sht As String
Public wksht As Worksheet
Public IsCalc As Boolean
Public IsSuppHeader As Boolean
Public IsLiteModel As Boolean
Public IsMdlNmPrefix As Boolean
Public IsRngNames As Boolean
Public sModelName As String
Public nrows As Integer
Public sNamePrefix As String
Public rngPopRows As Range
Public rngFormulaRows As Range
Public rngPopCols As Range
Public rngStepsVars As Range
Public icolCalc As Integer
Public rngMdl As Range
Public rowCur As Range
Public colCur As Range

'Column Ranges
Public colrngHeader As Range 'Multi-column
Public colrngHeaderFmt As Range 'Multi-column
Public colrngModel As Range 'Multi-column if not .IsCalc
Public colrngGrp As Range
Public colrngSubgrp As Range
Public colrngDesc As Range
Public colrngVarNames As Range
Public colrngUnits As Range
Public colrngNumFmt As Range
Public colrngFormulas As Range
'---------------------------------------------------------------------------------------
' Initialize Scenario Model location within workbook and set params from arguments
'
' JDL 7/19/23
'
Public Function Init(ByRef mdl As mdlScenario, wkbk, Optional sht, Optional IsCalc, _
        Optional IsSuppHeader, Optional IsRngNames, Optional cellHome, _
        Optional sModelName, Optional sDefn, Optional nrows, Optional IsMdlNmPrefix, _
        Optional IsLiteModel) As Boolean
            
    SetErrorHandle Init: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim tblS As New tblRowsCols
    
    With mdl
        Set .wkbk = wkbk
        If Not IsMissing(sModelName) Then .sModelName = sModelName
        
        'Init model using sDefn or use sModelName to look up definition from setting
        If (Not IsMissing(sModelName)) Or (Not IsMissing(sDefn)) Then
            If Not .ParseMdlScenDefn(mdl, sDefn) Then GoTo ErrorExit
            
        'Init model from args (covers case of default model specified only by sht)
        Else
            If Not .SetAttsFromArgs(mdl, sht, IsLiteModel, IsSuppHeader, IsRngNames, _
                    IsCalc, IsMdlNmPrefix, nrows, sModelName) Then GoTo ErrorExit
            If Not .SetCellHome(mdl, cellHome) Then GoTo ErrorExit
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "Init", Init
End Function
'---------------------------------------------------------------------------------------
' Set CellHome range
'
' JDL 7/17/23
'
Public Function SetCellHome(mdl, cellHome) As Boolean
    SetErrorHandle SetCellHome: If errs.IsHandle Then On Error GoTo ErrorExit
    With mdl
    
        'CellHome specified from Init argument
        If Not IsMissing(cellHome) Then
            Set .cellHome = cellHome
            Exit Function
        End If
        
        'Default CellHome
        Set .cellHome = .wksht.Cells(2, 1)
        If .IsSuppHeader Then Set .cellHome = .wksht.Cells(1, 1)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "SetCellHome", SetCellHome
End Function
'---------------------------------------------------------------------------------------
' Populate Scenario Model properties and ranges
' Inputs: mdl [mdlScenario Class instance] Provision returns this populated
'         sht [String] sheet name with model (either sht or sModelName required)
'         wkbk [Workbook] workbook object containing the model
'         IsCalc [Boolean] True if single-column model (cells named instead of rows)
'         IsSuppHeader [Boolean] True to suppress writing a header row
'         IsRngNames [Boolean] True to create row and column range names
'         cellHome [Range] Home cell range just below header row (top left corner)
'         sModelName [String] Model name - for reading config from Settings
'         nRows [Integer] Restrict model to fixed number of rows
'         IsMdlNmPrefix [Boolean] Add model name (sheet name default) to range names
'         IsLiteModel [Boolean] True if compact header columns + ExcelSteps
'                     for variable metadata
'
' Created:   1/11/21 JDL Modified: 2/28/23 Boolean function; add error handling
'                         3/3/23 add sDefn argument modify ParseScenModelDefn
'                         6/16/23 Set .nrows (Bug fix; was =0 previously
'                         7/17/23 Refactor to use Init
'
Public Function Provision(ByRef mdl As mdlScenario, wkbk, Optional sht, _
            Optional IsCalc, Optional IsSuppHeader, _
            Optional IsRngNames, Optional cellHome, Optional sModelName, _
            Optional sDefn, Optional nrows, Optional IsMdlNmPrefix, _
            Optional IsLiteModel) As Boolean
            
    SetErrorHandle Provision: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rng As Range, tblS As New tblRowsCols
    
    With mdl
    
        If Not Init(mdl, wkbk, sht, IsCalc, IsSuppHeader, IsRngNames, cellHome, _
            sModelName, sDefn, nrows, IsMdlNmPrefix, IsLiteModel) Then GoTo ErrorExit
                
        'Set header column ranges - Full and "Lite" versions
        If Not SetColRanges(mdl) Then GoTo ErrorExit
        
        'Set model starting column
        Set .colrngModel = .colrngUnits.Offset(0, 2)
        If Not .IsLiteModel Then Set .colrngModel = .colrngFormulas.Offset(0, 2)
        
        'Set range for rows containing variables
        Set rng = Intersect(.cellHome.EntireRow, .colrngVarNames)
        If .nrows = 0 Then
            Set .rngRows = rngToExtent(rng, IsRows:=True)
        Else
            Set .rngRows = Range(.cellHome, .cellHome.Offset(.nrows - 1, 0)).EntireRow
        End If
        
        'Set .nrows as gross number of rows (not just populated) 6/16/23
        If Not .rngRows Is Nothing Then .nrows = .rngRows.Rows.Count
        
        'Set range for model's columns
        If Not .IsCalc Then Set .colrngModel = _
            rngToExtent(Intersect(.cellHome.EntireRow, .colrngModel), IsRows:=False)
        If .IsCalc Then .icolCalc = .colrngModel.Column

        'Label the Scenario Name row for multi-column models
        If Not .IsCalc Then
            Intersect(.cellHome.EntireRow, .colrngDesc) = "Scenario Names"
            Intersect(.cellHome.EntireRow, .colrngVarNames) = "Scenario"
        ElseIf Not .IsSuppHeader Then
            Set rng = Intersect(.cellHome.EntireRow, .colrngModel)
            If Len(rng.Value) = 0 Then rng.Value = "Calculator"
        End If
        
        'Set multicell ranges for model rows (variables) and columns (scenarios)
        Set .rngPopRows = BuildMultiCellRange(.rngRows, .colrngVarNames)
        Set .rngPopCols = .colrngModel
        If Not .IsCalc Then Set .rngPopCols = BuildMultiCellRange(.colrngModel, _
            .cellHome.EntireRow)
        
        'xxx move PrepSteps into SetRngFormulaRows or split it into separate IsLite version
        'Set multicell range of rows whose variables are calculated by formula
        If .IsLiteModel Then If Not .PrepStepsForMdl(wkbk, tblS) Then GoTo ErrorExit
        If Not .rngPopRows Is Nothing Then .SetRngFormulaRows mdl, tblS

        'Set the header range
        If Not .IsSuppHeader And .cellHome.Row > 1 Then _
            Set .rngHeader = Intersect(.cellHome.Offset(-1, 0).EntireRow, .colrngHeader)
            
        'Set range for entire model (not incl. header row)
        Set .rngMdl = .colrngModel.Columns(.colrngModel.Columns.Count)
        Set .rngMdl = Intersect(Range(.cellHome.EntireColumn, .rngMdl), .rngRows)
            
        'Create prefix for variable names
        .sNamePrefix = ""
        If .IsMdlNmPrefix Then .sNamePrefix = xlName(.sModelName) & "_"
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "Provision", Provision
End Function
'---------------------------------------------------------------------------------------
' Initialize Steps table for use in Lite mdl Refresh
'
' Modified JDL 3/8/23 Set rowCur
'
Public Function PrepStepsForMdl(wkbk, ByRef tblS) As Boolean

    SetErrorHandle PrepStepsForMdl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim refr As New Refresh, rng As Range
    
    With refr
        If Not .Init(refr, wkbkR:=wkbk) Then GoTo ErrorExit
        If Not .PrepExcelStepsSht(refr, tblS) Then GoTo ErrorExit
    End With
    
    'Set rowCur attribute to first unused row
    With tblS.wkbk.Sheets(tblS.sht)
        Set tblS.rowCur = tblS.colrngCol.Cells(.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "PrepStepsForMdl", PrepStepsForMdl
End Function
'---------------------------------------------------------------------------------------
' Format the Scenario Model
' Modified JDL 3/15/23 fix problem with ColWidthAutoFit
'
Public Sub FormatScenModelClass(mdl)
    Dim rng As Range, str As String
    With mdl
    
        'Set or refresh header strings and header column formats
        str = ScenHeader
        If .IsLiteModel Then str = ScenHeaderLite
        If Not .IsSuppHeader Then .rngHeader = Split(str, ",")
        
        'Set text format, column widths and column outline for specific columns
        .colrngModel.Columns(1).Offset(0, -1).ColumnWidth = 4
        If Not .IsLiteModel Then
            Intersect(.rngRows, .colrngNumFmt).NumberFormat = "@"
            Intersect(.rngRows, .colrngFormulas).NumberFormat = "@"
            .colrngGrp.ColumnWidth = 4
            .colrngSubgrp.ColumnWidth = 7
            ColWidthAutofit .colrngDesc
            
            With Range(.colrngNumFmt, .colrngFormulas)
                If HasColOutlining(mdl.wkbk, mdl.sht) Then
                    If mdl.colrngNumFmt.OutlineLevel > 1 Or _
                        mdl.colrngFormulas.OutlineLevel > 1 Then
                        .Columns.Ungroup
                    End If
                End If
                .Columns.Group
            End With
        End If

        'Format the rows' header columns
        If Not .rngPopRows Is Nothing Then
            For Each rng In .rngPopRows.EntireRow
                SetHeaderStyle Intersect(rng, .colrngHeaderFmt)
                SetBorders Intersect(rng, .colrngHeaderFmt), xlContinuous, True
            Next rng
        End If
               
        'Format header row and scenario name row
        If Not .IsSuppHeader Then .rngHeader.Style = "Accent1"
        If Not .rngPopCols Is Nothing Then
            For Each rng In .rngPopCols.EntireColumn
                If Not .IsSuppHeader Then _
                    Intersect(.rngHeader.EntireRow, rng).Style = "Accent1"
                If Not .IsCalc Then SetHeaderStyle Intersect(.cellHome.EntireRow, rng)
            Next rng
        End If
        
        ColWidthAutofit .colrngDesc, iMaxWidth:=40
        .colrngDesc.WrapText = True
        
        For Each rng In .colrngHeaderFmt
            ColWidthAutofit rng, iMaxWidth:=15, iMinWidth:=10
        Next rng
        If Not .rngPopCols Is Nothing Then
            For Each rng In .rngPopCols
                ColWidthAutofit rng, iMinWidth:=10
            Next rng
        End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Format column header cells and Scenario Name cells
'
Public Sub SetHeaderStyle(rng)
    rng.Font.Size = 9
    rng.Style = "Note"
End Sub
'---------------------------------------------------------------------------------------
' Set column ranges
'
Function SetColRanges(ByRef mdl) As Boolean

    SetErrorHandle SetColRanges: If errs.IsHandle Then On Error GoTo ErrorExit
    With mdl
        If Not .IsLiteModel Then
            Set .colrngGrp = .cellHome.Offset(0, 0).EntireColumn
            Set .colrngSubgrp = .cellHome.Offset(0, 1).EntireColumn
            Set .colrngDesc = .cellHome.Offset(0, 2).EntireColumn
            Set .colrngVarNames = .cellHome.Offset(0, 3).EntireColumn
            Set .colrngUnits = .cellHome.Offset(0, 4).EntireColumn
            Set .colrngNumFmt = .cellHome.Offset(0, 5).EntireColumn
            Set .colrngFormulas = .cellHome.Offset(0, 6).EntireColumn
            Set .colrngHeader = Range(.colrngGrp, .colrngFormulas)
            Set .colrngHeaderFmt = Range(.colrngVarNames, .colrngFormulas)
        Else
            Set .colrngGrp = .cellHome.Offset(0, 0).EntireColumn
            Set .colrngDesc = .cellHome.Offset(0, 1).EntireColumn
            Set .colrngVarNames = .cellHome.Offset(0, 2).EntireColumn
            Set .colrngUnits = .cellHome.Offset(0, 3).EntireColumn
            Set .colrngHeader = Range(.colrngDesc, .colrngUnits)
            Set .colrngHeaderFmt = Range(.colrngVarNames, .colrngUnits)
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "SetColRanges", SetColRanges
End Function
'---------------------------------------------------------------------------------------
'Parse metadata for a Scenario Model in the workbook
'
'Inputs: mdl [mdlScenario Object]
'      wkbk [Workbook object] workbook containing the Scenario Model
'      Model [String] user-assigned name of the Scenario Model
'
' This returns array of Scenario Model metadata. The function parses a string
' Such definitions can be manually created when the model is created or they can be
' written [refreshed] to Settings sheet and read from there
'
' Scenario Model setting format:
' Setting name: mdl_ModelName (Model function argument)
' Setting value: sht:r,c:IsCalc:IsSuppHeader:IsMdlNmPrefix:IsLiteModel
'              sht - worksheet name where model resides
'              r,c,nrows - row,col home cell on sht; specified nrows (0 to ignore)
'              Booleans - represent as either "T" or "F" in setting
'
'              sModelName is needed for setting range name prefixes and for naming the model's overall range.
'              It must be specified directly as Init arg (.sModelName set in Init) or, if not used here
'              to read Setting with sDefn string, it must be included in sDefn as optional 9th element
'              (parsed and set here by ParseMdlScenDefn)
'
' Example Definition w/o non-default sName: Process:8,31:0:T:T:T:T:T
'                        non-default sName: Process:8,31:0:T:T:T:T:T:mdlProcess
'
' 4/1/21 JDL    Modified 1/6/22 Add IsRngNames T/F param
'                      3/5/23 Add sDefn argument and code; 7/17/23 cleanup
'                      7/17/23 Refactor and convert to Boolean function; add sName
'                      7/27/23 Change criteria for setting sModelName from parsed string
'
Function ParseMdlScenDefn(mdl, sDefn) As Boolean
    
    SetErrorHandle ParseMdlScenDefn: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer, aryRaw As Variant, aryParams As Variant, aryCellHome As Variant
    With mdl
        
        'Read Setting if model definition not specified as Provision arg
        If IsMissing(sDefn) Then sDefn = ReadSetting(.wkbk, .sModelName)
        
        'Parse sDefn and set locations
        aryRaw = Split(sDefn, ":")
        .sht = aryRaw(0)
        
        'Allow for initializing mdl before adding its sheet
        If SheetExists(.wkbk, .sht) Then
            Set .wksht = .wkbk.Sheets(.sht)
            aryCellHome = Split(aryRaw(1), ",")
            'xxx Error if ubound > 1
            Set .cellHome = .wksht.Cells(CInt(aryCellHome(0)), CInt(aryCellHome(1)))
        End If
    
        'nrows = 0 --> not specified: use extent of var names to set nrows and rngRows)
        .nrows = CInt(aryRaw(2))

        'Booleans: IsCalc, IsSuppHeader, IsRngNames, IsMdlNmPrefix, IsLiteModel
        aryParams = Array(False, False, False, False, False)
        For i = 3 To 7
            If aryRaw(i) = "T" Then aryParams(i - 3) = True
        Next i
    
        'Translate aryParams to Class attributes
        .IsCalc = aryParams(0)
        .IsSuppHeader = aryParams(1)
        .IsRngNames = aryParams(2)
        .IsMdlNmPrefix = aryParams(3)
        .IsLiteModel = aryParams(4)
        
        'Set model name; don't override sModelName set from arg
        If UBound(aryRaw) = 8 And Len(.sModelName) = 0 Then .sModelName = aryRaw(8)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "ParseMdlScenDefn", ParseMdlScenDefn
End Function
'---------------------------------------------------------------------------------------
' Set Class Attributes from specified, optional arguments
'
' JDL 1/6/22; Modified 7/17/23 convert to Boolean function
'
Public Function SetAttsFromArgs(ByRef mdl, sht, IsLiteModel, IsSuppHeader, IsRngNames, IsCalc, _
        IsMdlNmPrefix, nrows, sModelName) As Boolean
    
    SetErrorHandle SetAttsFromArgs: If errs.IsHandle Then On Error GoTo ErrorExit
    
    'Uninitialized Params default to False
    With mdl
        
        'Default to name cell (IsCalc) or row
        .IsRngNames = True
        
        If Not IsMissing(sht) Then .sht = sht
        If Not IsMissing(IsLiteModel) Then .IsLiteModel = IsLiteModel
        If Not IsMissing(IsSuppHeader) Then .IsSuppHeader = IsSuppHeader
        If Not IsMissing(IsRngNames) Then .IsRngNames = IsRngNames
        If Not IsMissing(IsCalc) Then .IsCalc = IsCalc
        If Not IsMissing(IsMdlNmPrefix) Then .IsMdlNmPrefix = IsMdlNmPrefix
        .nrows = 0
        If Not IsMissing(nrows) Then .nrows = nrows
        
        'If/Endif to allow for possibility of adding sheet post-init
        If SheetExists(.wkbk, .sht) Then
            Set .wksht = .wkbk.Sheets(.sht)
            If Not .SetCellHome(mdl, cellHome) Then GoTo ErrorExit
        End If
    
        'Set model name
        .sModelName = .sht
        If Not IsMissing(sModelName) Then .sModelName = sModelName

    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "SetAttsFromArgs", SetAttsFromArgs
End Function
'---------------------------------------------------------------------------------------
' Set Class Attributes from specified, optional arguments
'
' JDL 1/6/22; Modified 3/8/23 comments and docstring
'
Sub SetRngFormulaRows(ByRef mdl, tblS)
    Dim w As Variant, rowsSteps As Range, IsFormula As Boolean, R As Range
    With mdl
    
        'Set a row range (multirange of entire rows) for mdl's rows in ExcelSteps
        If .IsLiteModel Then
            Set .rngStepsVars = KeyColRng(tblS, Array(tblS.colrngSht), Array(.sModelName))
            If .rngStepsVars Is Nothing Then Exit Sub
            Set .rngStepsVars = Intersect(.rngStepsVars, tblS.colrngCol)
        End If

        'Iterate over rows that contain variables
        For Each w In .rngPopRows
            IsFormula = False
            
            'If not Lite, can just look at Formulas column in mdl
            If Not .IsLiteModel Then
                Set R = Intersect(w.EntireRow, .colrngFormulas)
                
                'Check for a formula in Formula/Row Type column (mod 10/26/22)
                If Not IsEmpty(R) Then
                    R = R.Formula 'Reset in case Text=-formatted cell evals to
                                  '#Name, #N/A or #Spill error
                    If Left(R, 1) = "=" Then IsFormula = True
                End If
            
            'If Lite, need to search mdl's rows in ExcelSteps  (mod 10/26/22)
            Else
                Set R = FindInRange(.rngStepsVars, w.Value)
                If Not R Is Nothing Then _
                    IsFormula = (Left(TableLoc(R, tblS.colrngStrInput), 1) = "=")
            End If
            
            'Add variable to multirange forcalculated variables
            If IsFormula Then
                If .rngFormulaRows Is Nothing Then Set .rngFormulaRows = w
                Set .rngFormulaRows = Union(.rngFormulaRows, w)
            End If
        Next w
    End With
End Sub
'---------------------------------------------------------------------------------------
' Apply border around model
'
' JDL mod 5/8/23
Sub ApplyBorderAroundModel(mdl, Optional IsBufferRow = False, Optional IsBufferCol = False)
    Dim xlEdge As Variant, rng As Range
    
    Set rng = mdl.rngMdl
    If IsBufferRow Then Set rng = Union(rng, mdl.rngMdl.Offset(1, 0))
    If IsBufferCol Then Set rng = Union(rng, rng.Offset(0, 1))
    
    For Each xlEdge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With rng.Borders(xlEdge)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    Next xlEdge
End Sub
'---------------------------------------------------------------------------------------
' Clear model cell values and outline
'
' Modified: 3/13/23 JDL fix bug with IsBufferCol
'
Function ClearModel(mdl, Optional IsBufferRow = False, Optional IsBufferCol = False) As Boolean

    SetErrorHandle ClearModel: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rng As Range
    With mdl
    
        'Clear model cells
        Set rng = .rngMdl
        If IsBufferRow Then Set rng = Union(rng, .rngMdl.Offset(1, 0))
        If IsBufferCol Then Set rng = Union(rng, rng.Offset(0, 1))
        rng.Clear
        
        'Clear header cells
        If Not .IsSuppHeader Then
            .rngHeader.Clear
            Intersect(.rngHeader.EntireRow, .colrngModel).Clear
        End If
        
        'Clear column outline
        If Not .IsLiteModel Then
            If HasColOutlining(.wkbk, .sht) Then _
                Range(.colrngNumFmt, .colrngFormulas).Columns.Ungroup
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "ClearModel", ClearModel
End Function
'-----------------------------------------------------------------------------------------------
' Delete mdl Range names
' The .Visible property is used to skip hidden _xlfn.SINGLE name created by dynamic array
' glitch from circa 2020 Excel update. See: https://stackoverflow.com/questions/59121799
'
' JDL 6/20/23
'
Function DeleteMdlRangeNames(mdl, Optional sPrefix As String) As Boolean
    
    SetErrorHandle DeleteMdlRangeNames: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim w As Variant, nchars As Integer
    
    'If user doesn't override prefix, use ModelName
    If IsMissing(sPrefix) Then sPrefix = mdl.sModelName
    
    nchars = Len(sPrefix)
    For Each w In mdl.wkbk.Names
        If (Left(w.Name, nchars) = sPrefix) And w.Visible Then w.Delete
    Next w
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "DeleteMdlRangeNames", DeleteMdlRangeNames
End Function
'------------------------------------------------------------------------------------------------
' Add a dropdown to a Scenario Model variable value(s)
'
' Created:   1/3/22 JDL      Modified: 2/1/22 FindInRange
'
Sub AddDropdownToVariable(mdl, sVar, sDropdownFormula)
    Dim c As Range
    Set c = FindInRange(mdl.colrngVarNames, sVar)
    If c Is Nothing Then Exit Sub
    
    Set c = Intersect(c.EntireRow, mdl.colrngModel)
    
    'Skip if error such as named range not existing
    On Error GoTo 0
    AddValidationList c, sDropdownFormula
End Sub
'---------------------------------------------------------------------------------------
'Lookup and return Scenario Model value
'Inputs:    sVar [String] Scenario Model variable name for lookup
'         rngCol [Range] table column range
'
' Created: 2/4/21 JDL    Modified: 1/4/22 Class Method; 1/17/22 switch to FindInRange
'
Function ScenModelLoc(mdl, sVar, Optional rngCol) As Range
    Dim rngRow As Range
    Set rngRow = FindInRange(mdl.colrngVarNames, sVar)
    If rngRow Is Nothing Then Exit Function
    If IsMissing(rngCol) Then Set rngCol = mdl.colrngModel
    Set ScenModelLoc = Intersect(rngRow.EntireRow, rngCol)
End Function
'---------------------------------------------------------------------------------------
' Set value for specified Scenario Model variable
' Inputs: sVar [String] Scenario Model variable name for lookup
'         rngCol [Range] table column range
'         val [variant] value to set at intersection of rngCell.row and rngCol
'
' Created: 2/4/21 JDL    Modified: 7/7/21 to make rngcol optional; 1/17/22 FindInRange
'
Sub SetScenModelLoc(mdl, sVar, val, Optional rngCol)
    Dim rngRow As Range
    Set rngRow = FindInRange(mdl.colrngVarNames, sVar)
    If rngRow Is Nothing Then Exit Sub
    If IsMissing(rngCol) Then Set rngCol = mdl.colrngModel
    Intersect(rngRow.EntireRow, rngCol) = val
End Sub
'---------------------------------------------------------------------------------------
' Refresh a Scenario model
'
' Created: JDL      Modified: 1/7/22 - Refactor to use mdlRow Class
'                          6/30/22 - refactor mdlRow Class
'                          3/8/23 cleanup
Function Refresh(mdl) As Boolean

    SetErrorHandle Refresh: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim tblS As New tblRowsCols, R As Object, rngRow As Variant
            
    With mdl
        If .IsLiteModel Then If Not PrepStepsForMdl(wkbk, tblS) Then GoTo ErrorExit
        If Not CkVarAndScenNames(mdl) Then GoTo ErrorExit
        If .IsRngNames Then NameMdlColumns mdl
    End With
    
    'Loop through rows and apply the model - exit if no model rows
    If Not mdl.rngPopRows Is Nothing Then
        For Each rngRow In mdl.rngPopRows
        
            'Instance/initialize mdlRow Class to hold row attributes
            Set R = New mdlRow
            R.Init R, mdl, rngRow, tblS
            
            'Name cell/row
            If mdl.IsRngNames Then R.NameRow R, mdl
        
            'Set Number Format and formula
            If Not R.rngMdlCells Is Nothing Then
                R.FormatRow R
                If R.HasLstValidation Then R.AddListValidation R
                If Not R.WriteRowFormula(R, mdl) Then GoTo ErrorExit
            End If
        Next rngRow
    End If
    
    mdl.FormatScenModelClass mdl
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "Refresh", Refresh
End Function
'---------------------------------------------------------------------------------------
' Check row and column name strings are Excel compatible and non-redundant
'
' Created: 1/7/22 JDL
'
Function CkVarAndScenNames(mdl) As Boolean
    CkVarAndScenNames = True
    With mdl
        If CheckNames(.rngPopRows) > 0 Then CkVarAndScenNames = False
        If Not .IsCalc Then
            If CheckNames(.rngPopCols) > 0 Then CkVarAndScenNames = False
        End If
    End With
End Function
'---------------------------------------------------------------------------------------
' Name multi-column model column ranges
'
' Created: 1/6/22 JDL   Modified 3/8/23 Cleanup
'
Sub NameMdlColumns(mdl)
    Dim c As Range, rngRow As Range
    With mdl
        Set rngRow = .cellHome.EntireRow
        If Not .IsCalc And Not .rngPopCols Is Nothing Then
            For Each c In .rngPopCols
                MakeXLName .wkbk, Intersect(rngRow, c), _
                    MakeRefNameString(.sht, icol1:=c.Column)
            Next c
        ElseIf .IsCalc Then
            MakeXLName .wkbk, .sModelName, _
                MakeRefNameString(.sht, 0, 0, .icolCalc, .icolCalc)
        End If
    End With
End Sub
'---------------------------------------------------------------------------------------
'Methods related to SwapModel capability to replace a Scenario Model with a second one
'stored on tblImport Sheet
'---------------------------------------------------------------------------------------
' SwapModels master procedure for transferring a Scenario Model to a rows/cols "input
' deck" on shtTblImport and transferring a rows/cols version to Scenario Model as
' a replacement
'
' Inputs: ModelNew [String] Name of model to swap to mdlDest from tblImport sheet table
'         ModelDest [String] dest mdl name (for defn lookup if not specified)
'         ModelDefnDest [Optional String] destination model Defn (if pre-set)
'
'JDL 3/14/23 JDL    Modified: 7/14/23 Major refactoring
'
Function SwapModels(wkbk As Workbook, Optional ByVal ModelNew As String, _
        Optional ByVal ModelDest As String, Optional ByVal ModelDefnDest As String) As Boolean

    SetErrorHandle SwapModels: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim mdlDest As New mdlScenario, tblImp As New tblRowsCols, tblS As New tblRowsCols
    
    'Initialize and Provision Classes for the swap
    If Len(ModelDefnDest) < 1 Then ModelDefnDest = ReadSetting(wkbk, ModelDest)
    If Not InitSwapModels(mdlDest, tblImp, tblS, wkbk, ModelDefnDest) Then GoTo ErrorExit
    
    'If there is a previous model, transfer it to tblImport sheet
    If mdlDest.nrows > 1 Then
        If Not TransferToTblImport(mdlDest, tblImp, tblS) Then GoTo ErrorExit
    End If
    
    'If requested, transfer new mdl from tblImport sheet table
    If Len(ModelNew) > 0 Then
        If Not TransferToMdlDest(mdlDest, tblImp, tblS, ModelNew, ModelDefnDest) Then GoTo ErrorExit
        ApplyBorderAroundModel mdlDest, IsBufferRow:=True, IsBufferCol:=True
    End If
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "SwapModels", SwapModels
End Function
'---------------------------------------------------------------------------------------
' Initialize a swap or other move between Scenario Model and tblImport
'
'JDL 3/14/23 JDL    Modified: 7/14/23 Refactoring
'
Function InitSwapModels(mdlDest As mdlScenario, tblImp As tblRowsCols, _
        tblS As tblRowsCols, ByVal wkbk As Workbook, ModelDefnDest As String) As Boolean

    SetErrorHandle InitSwapModels: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim refr As New Refresh
    
    'Provision/initialize destination model
    With mdlDest
        If Not .Provision(mdlDest, wkbk, sDefn:=ModelDefnDest) Then GoTo ErrorExit
        Set .rowCur = .cellHome.EntireRow
    End With
        
    'Provision/initialize tblImport sheet table (xxx mod to use .End(xlUp) for .rowCur???
    With tblImp
        If Not .Provision(tblImp, wkbk, False, shtTblImp, ncols:=10) Then GoTo ErrorExit
        
        'Initialize rowCur to first blank row; initialize rngRows anyway if no data
        If .rngRows Is Nothing Then
            Set .rowCur = .cellHome.EntireRow
            Set .rngRows = .rowCur
        Else
            Set .rowCur = .rngRows.Rows(.rngRows.Count).Offset(1, 0)
        End If
    End With
     
    'Provision/Initialize ExcelSteps sheet (w/ Prep because could be missing/blank)
    If Not refr.Init(refr, wkbkR:=wkbk) Then GoTo ErrorExit
    If Not refr.PrepExcelStepsSht(refr, tblS) Then GoTo ErrorExit
    With tblS
        Set .rowCur = .colrngCol.Cells(.wksht.Rows.Count, 1).End(xlUp).Offset(2, 0).EntireRow
    End With
    Exit Function

ErrorExit:
    errs.RecordErr errs, "InitSwapModels", InitSwapModels
End Function
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' TransferToTblImport Procedure - transfer model from mdlDest region to tblImp rows/cols
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' Procedure - Transfer model from mdlDest Scenario Model region to tblImport sheet rows/cols
'
' JDL 8/22/23
'
Function TransferToTblImport(ByVal mdlDest As mdlScenario, ByRef tblImp As tblRowsCols, _
        ByVal tblS As tblRowsCols) As Boolean

    SetErrorHandle TransferToTblImport: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim ModelPrev As String
    With mdlDest
        If Not .ReadModelName(mdlDest, tblS, ModelPrev) Then GoTo ErrorExit
        If Not .TblImportDeleteModel(tblImp, ModelPrev) Then GoTo ErrorExit
        If Not .TransferMdlDestRows(mdlDest, tblImp, tblS, ModelPrev) Then GoTo ErrorExit
        If Not .DeleteTblImpTrailingBlankRows(tblImp) Then GoTo ErrorExit
        If Not .ClearModel(mdlDest, IsBufferRow:=True, IsBufferCol:=True) Then GoTo ErrorExit
        If Not .StepsDeleteMdl(mdlDest, tblS) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "TransferToTblImport", TransferToTblImport
End Function
'-----------------------------------------------------------------------------------------------
' Set ModelPrev by reading mdl_name variable value from mdlDest; set mdlDest.rngStepsVars
'
' JDL 7/27/23   Modified 8/22/23 Add set mdlDest.sModelName and .rngStepsVars
'
Function ReadModelName(ByVal mdlDest As mdlScenario, ByVal tblS As tblRowsCols, _
        ModelPrev As String) As Boolean

    SetErrorHandle ReadModelName: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rng As Range
    
    With mdlDest
        Set rng = FindInRange(.colrngVarNames, "mdl_name")
        If errs.IsFail(rng Is Nothing, 1) Then GoTo ErrorExit
        
        ModelPrev = Intersect(rng.EntireRow, .colrngModel)
        
        'Set ExcelSteps range for ModelPrev
        .sModelName = ModelPrev
        Set .rngStepsVars = KeyColRng(tblS, Array(tblS.colrngSht), Array(.sModelName))
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "ReadModelName", ReadModelName
End Function
'---------------------------------------------------------------------------------------
' Delete a model from tblImport sheet
'
' Created: 3/6/23 JDL Modified argument and tblImp name 8/17/23
'
Function TblImportDeleteModel(tblImp, Model) As Boolean

    SetErrorHandle TblImportDeleteModel: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rngModel As Range, wkbk As Workbook

    'Exit if tblImport sheet is empty
    If tblImp.rngRows Is Nothing Then Exit Function
    
    'Set range for model on tblImport sheet and delete its rows if any
    Set rngModel = rngKeycolRows(tblImp, tblImp.colrngMdlName, Model)
    If rngModel Is Nothing Then Exit Function
    rngModel.Delete
    
    'Re-initialize tblImp
    Set wkbk = tblImp.wkbk
    Set tblImp = New tblRowsCols
    If Not tblImp.Provision(tblImp, wkbk, False, sht:=shtTblImp, ncols:=10) _
            Then GoTo ErrorExit
    Exit Function
ErrorExit:
    errs.RecordErr errs, "tblImportDeleteModel", TblImportDeleteModel
End Function
'-----------------------------------------------------------------------------------------------
' Transfer mdlDest Scenario Model rows to tblImport rows/columns table
'
' JDL 7/27/23
'
Function TransferMdlDestRows(mdlDest, tblImp As tblRowsCols, ByVal tblS As tblRowsCols, _
        ByVal ModelPrev As String) As Boolean

    SetErrorHandle TransferMdlDestRows: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim R_MI As New mdlImportRow, R As Range
   
    With mdlDest
        Set .rowCur = .cellHome.EntireRow
        R_MI.Model = ModelPrev
        
        'Iterate over mdlDest rows; Use mdlImportRow Class to transfer to tblImport sheet
        For Each R In .rngRows
            With R_MI
                If Not .Init(R_MI) Then GoTo ErrorExit
                If Not .ReadMdlDestRow(R_MI, mdlDest) Then GoTo ErrorExit
                If Not .ReadStepsRow(R_MI, tblS, mdlDest.rngStepsVars) Then GoTo ErrorExit
                If Not .SetBooleanFlags(R_MI) Then GoTo ErrorExit
                If Not .ToTblWriteRow(R_MI, mdlDest, tblImp) Then GoTo ErrorExit
            End With
        Next R
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "TransferMdlDestRows", TransferMdlDestRows
End Function
'---------------------------------------------------------------------------------------
' Delete trailing blank rows, if any, from tblImp
'
' JDL 3/14/23 JDL    Modified 8/22/23
'
Function DeleteTblImpTrailingBlankRows(ByRef tblImp As tblRowsCols) As Boolean

    SetErrorHandle DeleteTblImpTrailingBlankRows: If errs.IsHandle Then On Error GoTo ErrorExit

    With tblImp
        Set .rowCur = .rngRows.Rows(.rngRows.Rows.Count)
        Do While Intersect(.rowCur, .colrngVarName) = "<blank>"
            .rowCur.Clear
            Set .rowCur = .rowCur.Offset(-1, 0)
            Set .rngRows = Range(.rngRows.Rows(1), .rowCur)
        Loop
    End With
    Exit Function

ErrorExit:
    errs.RecordErr errs, "DeleteTblImpTrailingBlankRows", DeleteTblImpTrailingBlankRows
End Function
'---------------------------------------------------------------------------------------
' Delete Steps rows for a model
'
' Created: 3/14/23 JDL
'
Function StepsDeleteMdl(mdl, tblS) As Boolean

    SetErrorHandle StepsDeleteMdl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer
    With mdl
    
        'Set multirange for model's rows on ExcelSteps sheet
        If Not .SetStepsRowRange(mdl, tblS) Then GoTo ErrorExit
        If .rngStepsVars Is Nothing Then Exit Function
        
        'Loop in reverse order and delete
        For i = .rngStepsVars.Rows.Count To 1 Step -1
            .rngStepsVars.Rows(i).Delete
        Next i
    End With
Exit Function
    
ErrorExit:
    errs.RecordErr errs, "StepsDeleteMdl", StepsDeleteMdl
End Function
'---------------------------------------------------------------------------------------
' Set a row range (multirange of entire rows) for mdl's rows in ExcelSteps
'
' Modified: 6/20/23 for "ProcessParams2" (sMdlProcess2) in ExcelSteps Sheet column
'
Function SetStepsRowRange(mdl, tblS) As Boolean

    SetErrorHandle SetStepsRowRange: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With mdl
        Set .rngStepsVars = KeyColRng(tblS, Array(tblS.colrngSht), Array(.sModelName))
    End With
    Exit Function
ErrorExit:
    errs.RecordErr errs, "SetStepsRowRange", SetStepsRowRange
End Function
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' TransferToMdlDest Procedure - transfer model from tblImport sheet to mdlDest region
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Transfer a model from tblImport sheet rows/cols "input deck" to mdlDest Scenario Model
'
' JDL 12/13/21   Modified: 8/22/23 JDL
'
Function TransferToMdlDest(mdlDest, tblImp, tblS, ModelNew, ModelDefnDest) As Boolean

    SetErrorHandle TransferToMdlDest: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim R_MI As New mdlImportRow
    
    With mdlDest
        
        'Clear previous mdlDest and set rng and rowCur in tblImport table
        If Not .InitTransferToMdl(mdlDest, tblImp, ModelNew) Then GoTo ErrorExit
        
        'Transfer ModelNew rows from tblImport table to mdlDest
        If Not .TransferTblImportRows(R_MI, mdlDest, tblImp, tblS) Then GoTo ErrorExit
        
        'Post-Transfer, delete ModelNew rows from tblImport table; Refresh mdlDest
        If Not .ResetPostTransfer(mdlDest, tblImp.rngRowsPopulated, ModelNew, _
                ModelDefnDest) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "TransferToMdlDest", TransferToMdlDest
End Function
'---------------------------------------------------------------------------------------
' Init transferring a model from tblImport sheet to Scenario Model
'
' JDL 7/25/23
'
Function InitTransferToMdl(ByVal mdlDest As mdlScenario, ByRef tblImp As tblRowsCols, _
        ByVal ModelNew As String)

    SetErrorHandle InitTransferToMdl: If errs.IsHandle Then On Error GoTo ErrorExit
    
    'Clear mdlDest region and reset .rowCur
    With mdlDest
        .ClearModel mdlDest
        Set .rowCur = .cellHome.EntireRow
    End With
        
    'Set range for model in tblImport (.rngRowsPopulated attribute) - Err if not found
    With tblImp
        Set .rngRowsPopulated = rngKeycolRows(tblImp, .colrngMdlName, ModelNew)
        If errs.IsFail(.rngRowsPopulated Is Nothing, 1) Then GoTo ErrorExit
        Set .rowCur = .rngRowsPopulated.Rows(1)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "InitTransferToMdl", InitTransferToMdl
End Function
'---------------------------------------------------------------------------------------
' Transfer ModelNew rows from tblImport table to mdlDest
'
' JDL 7/25/23   Modified 8/24/23
'
Function TransferTblImportRows(ByRef R_MI As mdlImportRow, ByRef mdlDest As mdlScenario, _
        ByVal tblImp As tblRowsCols, ByRef tblS As tblRowsCols)

    SetErrorHandle TransferTblImportRows: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim R As Variant
    
    'Iterate over tblImp rows for model; write table vals to mdlDest and tblS
    For Each R In tblImp.rngRowsPopulated.Rows
        With R_MI
            If Not .Init(R_MI) Then GoTo ErrorExit
            If Not .ReadRow(R_MI, tblImp) Then GoTo ErrorExit
            If Not .SetBooleanFlags(R_MI) Then GoTo ErrorExit
            If Not .SetStepType(R_MI) Then GoTo ErrorExit
            If Not .WriteRowToMdl(R_MI, mdlDest) Then GoTo ErrorExit
            If Not .WriteRowToSteps(R_MI, tblS) Then GoTo ErrorExit
        End With
        Set mdlDest.rowCur = mdlDest.rowCur.Offset(1, 0)
        Set tblImp.rowCur = tblImp.rowCur.Offset(1, 0)
    Next R
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "TransferTblImportRows", TransferTblImportRows
End Function
'---------------------------------------------------------------------------------------
' Post-Transfer, delete ModelNew rows from tblImport table; Refresh mdlDest
'
' JDL 7/25/23   Modified 8/22/23
'
Function ResetPostTransfer(ByRef mdlDest As mdlScenario, ByVal rngModel As Range, _
        ByVal ModelNew As String, ByVal ModelDefnDest As String) As Boolean

    SetErrorHandle ResetPostTransfer: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkbk As Workbook
    
    'Delete the model from tblImport sheet
    rngModel.Delete
    
    'Re-initialize mdlDest to clear sModelName and other params from pre-transfer
    Set wkbk = mdlDest.wkbk
    Set mdlDest = New mdlScenario
    
    'Re-provision/refresh dest model with customized model name
    With mdlDest
        If Not .Provision(mdlDest, wkbk, sModelName:=ModelNew, _
                sDefn:=ModelDefnDest) Then GoTo ErrorExit
        If Not .Refresh(mdlDest) Then GoTo ErrorExit
        .ApplyBorderAroundModel mdlDest, IsBufferRow:=True, IsBufferCol:=True
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr errs, "ResetPostTransfer", ResetPostTransfer
End Function
