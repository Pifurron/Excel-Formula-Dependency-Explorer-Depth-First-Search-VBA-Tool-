Attribute VB_Name = "Module1"
Function ExtractNumberFromString(str As String) As String
    Dim i As Integer
    Dim extractedNumber As String
    extractedNumber = ""

    For i = 1 To Len(str)
        If IsNumeric(Mid(str, i, 1)) Then
            extractedNumber = extractedNumber & Mid(str, i, 1)
        End If
    Next i

    ExtractNumberFromString = extractedNumber
End Function


Function ExtractNonNumericFromString(str As String) As String
    Dim i As Integer
    Dim nonNumeric As String
    nonNumeric = ""

    For i = 1 To Len(str)
        If Not IsNumeric(Mid(str, i, 1)) Then
            nonNumeric = nonNumeric & Mid(str, i, 1)
        End If
    Next i

    ExtractNonNumericFromString = nonNumeric
End Function



Sub WriteLastCellRefToColumnAV()
    Dim LastRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Feuil4") ' Change to your sheet name
    
    LastRow = ws.Cells(ws.Rows.Count, "BA").End(xlUp).Row

    ' Check if there is any cell reference in the Collection
    If cellRefs.Count > 0 Then
        ws.Cells(LastRow + 1, 2).Value = cellRefs(cellRefs.Count) ' Write the last cell reference
        cellRefs.Remove cellRefs.Count
    End If
End Sub


Function ExtractCellRefsWithCollection(Rg As Range) As Collection
    Dim xRegEx As Object
    Dim cellRefs As Collection
    Set cellRefs = New Collection
    Dim xMatches As Object
    Dim i As Long

    Application.Volatile
    Set xRegEx = CreateObject("VBSCRIPT.REGEXP")
    With xRegEx
        .Pattern = "(\[+[\w\s\.יטאחשפ0-9]{1,99}\]+)?('?[a-zA-Z0-9\s\._\-]{1,99})?'?!?\$?[A-Z]{1,3}\$?[0-9]{1,7}(:\$?[A-Z]{1,3}\$?[0-9]{1,7})?"
        '.Pattern = "(?:'([^']+)'\!|\b)(\$?[A-Z]{1,3}\$?[0-9]{1,7})(?::(\$?[A-Z]{1,3}\$?[0-9]{1,7}))?"
                    

        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With

    Set xMatches = xRegEx.Execute(Rg.Value)

    If xMatches.Count > 0 Then
        For i = 0 To xMatches.Count - 1
            cellRefs.Add xMatches.Item(i).Value
        Next i
        Set ExtractCellRefsWithCollection = cellRefs
    Else
        Set ExtractCellRefsWithCollection = Nothing
    End If
End Function

Sub AdderFifoMioV2(TargetCell As Range, Pyramidcol As Integer, Depth As Long, ws As Worksheet, FormulaInput As Integer, FormulaInputNum As Integer)
    Dim cellRefCollection As Collection
    Dim newRefs As Collection
    Dim i As Long
    Dim cell As Range
    Dim cell2 As Range
    Dim ref As Variant
    Dim splitString() As String
    Dim leftPart As String
    Dim rightPart As String
    Dim leftNumber As Integer
    Dim rightNumber As Integer
    Dim result As Integer
    Dim minimum As Integer
    Dim leftNonNumeric As String
    Dim f As Integer
    Dim L As Long
    Dim U As Long
    Dim N As Long
    Dim M As Long
    Dim CurrentDepth As Long
    Dim Depthofsit As Long
    Dim oldcellref As Collection
    Dim rowData As Variant
    Dim Row As Variant
    Dim cellsaved As Variant
    Dim CellSaved2 As Variant
    Dim Depthactivated As Boolean
    Dim Celldepth As Integer
    Dim Lastvalue As Boolean
    
    CurrentDepth = 0
    Lastvalue = False
    Depthofsit = 0
    L = 0
    Depthactivated = False
    
    Application.ScreenUpdating = False

    Set cellRefCollection = New Collection
    Set oldcellref = New Collection
        
    While CurrentDepth <= Depth
        For Each cell In TargetCell
            'ws.Cells(cell.Row, 55).Value = CurrentDepth
            If CurrentDepth = Depth And Celldepth = Depth Then
                
                t = 0
                While t <= Val(cellRefCollection.Count - oldcellref.Count)
                    M = 1
                    i = Pyramidcol   ' Celda a cambiar piramide
                    While M <= cellRefCollection.Count
                        Row = cellRefCollection(M)
                        cellsaved = Row(1)
                        ws.Cells(cell.Row + t, i).Value = cellsaved
                        i = i + 1
                        M = M + 1
                    Wend
                    On Error GoTo Finalcellwhenfive
                    Row = cellRefCollection(cellRefCollection.Count)
                    On Error GoTo 0
                    cellsaved = Row(1)
                    CellSaved2 = Row(0)
                    ws.Cells(ws.Cells(ws.Rows.Count, FormulaInput).End(xlUp).Row + 1, FormulaInput).Value = cellsaved  ' celda a cambiar input formulas
                    ws.Cells(ws.Cells(ws.Rows.Count, FormulaInputNum).End(xlUp).Row + 1, FormulaInputNum).Value = CellSaved2
                    
                    Row = cellRefCollection(cellRefCollection.Count) 'error
                    Depthofsit = Row(0)
                    CurrentDepth = Depthofsit
                    Depthofsit = 0
                    
                    cellRefCollection.Remove cellRefCollection.Count
                    t = t + 1
                    'ws.Cells(cell.Row, 55).Value = CurrentDepth
                Wend
                If cellRefCollection.Count = 0 Then
                    Lastvalue = True
                    
                Else
                    Row = cellRefCollection(cellRefCollection.Count) 'error
                    Depthofsit = Row(0)
                    CurrentDepth = Depthofsit
                End If

            Else
                Set newRefs = New Collection
                Set newRefs = ExtractCellRefsWithCollection(cell)
                If newRefs Is Nothing And Depthactivated = False Then
                Else
                    Depthactivated = False
                    If Depthofsit <> 0 Then
                        
                    Else
                        CurrentDepth = CurrentDepth + 1
                    End If
                End If
                Set oldcellref = New Collection
                For Each ref In cellRefCollection
                    oldcellref.Add ref
                Next ref
                If Depthofsit <> 0 Then
                    Depthofsit = 0
                    Depthactivated = True
                Else
                    If Not newRefs Is Nothing Then
                        For Each ref In newRefs
                            rowData = Array(CurrentDepth, ref)
                            cellRefCollection.Add rowData
                        Next ref
                    End If
                    Depthofsit = 0
                End If
                If cellRefCollection.Count > 0 Or Lastvalue = True Then
                    Lastvalue = False
                    i = Pyramidcol
                    U = 1
        
                    While U <= cellRefCollection.Count
                        Row = cellRefCollection(U)
                        cellsaved = Row(1)
                            'If InStr(1, cellsaved, ":") > 0 Then
                             '   splitString = Split(cellsaved, ":")
                            '
                            '    leftPart = splitString(0)
                            '    rightPart = splitString(1)
                           '
                            '    leftNumber = Val(ExtractNumberFromString(leftPart))
                             '   rightNumber = Val(ExtractNumberFromString(rightPart))
                            '
                            '    result = Abs(leftNumber - rightNumber)
                            '    minimum = Application.WorksheetFunction.Min(leftNumber, rightNumber)
                            '
                            '    leftNonNumeric = ExtractNonNumericFromString(leftPart)
                            '
                            '    For f = 1 To result + 1
                            '        rowData = Array(CurrentDepth, leftNonNumeric & minimum + f - 1)
                             '       cellRefCollection.Add rowData
                             '       'ws.Cells(cell.Row, i).Value = leftNonNumeric & minimum + f - 1
                             '   Next f
                              '  'cellRefCollection.Remove U
                              '  L = L + 1
                            'End If
                        'ws.Cells(cell.Row, i).Value = cellRefCollection(U)
                        'i = i + 1
                        U = U + 1
                    Wend
                    
                    N = 1
                    
                    If L > 0 Then
                        While N <= cellRefCollection.Count
                            Row = cellRefCollection(N)
                            cellsaved = Row(1)
                            If InStr(1, cellsaved, ":") > 0 Then
                                cellRefCollection.Remove N
                                N = N - 1
                            End If
                            N = N + 1
                            
                        Wend
                        L = 0
                    End If
        
                    M = 1
        
                    While M <= cellRefCollection.Count
                        Row = cellRefCollection(M)
                        cellsaved = Row(1)
                        ws.Cells(cell.Row, i).Value = cellsaved
                        i = i + 1
                        M = M + 1
                    Wend
                    Row = cellRefCollection(cellRefCollection.Count)
                    cellsaved = Row(1)
                    CellSaved2 = Row(0)
                    ws.Cells(ws.Cells(ws.Rows.Count, FormulaInput).End(xlUp).Row + 1, FormulaInput).Value = cellsaved
                    ws.Cells(ws.Cells(ws.Rows.Count, FormulaInputNum).End(xlUp).Row + 1, FormulaInputNum).Value = CellSaved2
                    Celldepth = CellSaved2
                    CurrentDepth = CellSaved2
                    cellRefCollection.Remove cellRefCollection.Count
                Else
                    'MsgBox "No Matches"
                    Exit Sub
                End If
            End If
        Next cell
    Wend
    Application.ScreenUpdating = True
Finalcellwhenfive:
    MsgBox "Finalized"
    Exit Sub

End Sub


Sub Button1_Click()
    Dim ws As Worksheet
    Dim TargetCell As Range
    Dim Depth As Long
    Dim Pyramidcol As Integer
    Dim FormulaInput As Integer
    Dim FormulaInputNum As Integer
    Dim Tableleftcol As Integer
    Dim Tablerightcol As Integer
    
    
    Set ws = ThisWorkbook.Worksheets("Feuil6")
    Set TargetCell = ws.Range("P8:P300")
    Depth = ws.Range("M4").Value
    Pyramidcol = ws.Range("M4").Column + 5
    FormulaInput = ws.Range("M4").Column - 2
    FormulaInputNum = ws.Range("M4").Column - 3
    

    'AdderFifoMioV2 targetcell, Pyramidcol, Depth, ws, FormulaInput, FormulaInputNum
    
    
    Tableleftcol = ws.Range("M4").Column - 6
    Tablerightcol = ws.Range("M4").Column + 3
    
    Formatcells ws, TargetCell, Tableleftcol, Tablerightcol, FormulaInputNum
    
    
End Sub




Sub LoopPrecedents()
    Dim refs As Collection
    Dim newRefs As Collection
    Dim cell As Range
    Dim Selection As Range
    Dim ws As Worksheet
    Dim sheet As String
    Dim startingformulacell As Range
    Dim secondformulacell As Range
    Dim Depthcell As Range
    Dim Depth As Long
    Dim rowData As Variant
    Dim cellsaved As Variant
    Dim Pyramidcol As Integer
    Dim FormulaInput As Integer
    Dim FormulaInputNum As Integer
    Dim TargetCell As Range
    Dim Tableleftcol As Integer
    Dim Tablerightcol As Integer
    Dim LastRow As Long
    Dim thex As Range
    Dim i As Integer
    
    'sheet = Application.InputBox("What is the name of the sheet you want the calculations to be done", Type:=2)
    Set ws = ThisWorkbook.Worksheets("32 - trim sans JV BE")
    
    Depth = 4
    
    Set refs = New Collection
    
    Set Selection = Application.InputBox("Select the cells you want analyzed", Type:=8)
    Set startingformulacell = Application.InputBox("Where do you want the starting formula cell", Type:=8)
    

    For Each cell In Selection
        If Not cell Is Nothing Then
            Set newRefs = ExtractCellRefsWithCollection(cell)
            For Each ref In newRefs
                rowData = ref
                refs.Add rowData
            Next ref
            
            'First month process
            cellsaved = refs(1)
            
            startingformulacell = cellsaved
            Set thex = ws.Cells(startingformulacell.Row, startingformulacell.Column - 1)
            thex = "x"
            
            Set Depthcell = ws.Cells(startingformulacell.Row - 3, startingformulacell.Column - 7)
            
            Depthcell = Depth
            
            Set TargetCell = ws.Range(ws.Cells(startingformulacell.Row, startingformulacell.Column + 5), ws.Cells(startingformulacell.Row + 150, startingformulacell.Column + 5))
            
            Pyramidcol = startingformulacell.Column + 7
            FormulaInputNum = startingformulacell.Column - 1
            FormulaInput = startingformulacell.Column
            
            AdderFifoMioV2 TargetCell, Pyramidcol, Depth, ws, FormulaInput, FormulaInputNum
            
            Set thex = Nothing
            Set Depthcell = Nothing
            
            Tableleftcol = startingformulacell.Column - 4
            Tablerightcol = startingformulacell.Column + 5
            
            Formatcells ws, TargetCell, Tableleftcol, Tablerightcol, FormulaInputNum
            
            Set TargetCell = Nothing

            
            'Second month process (exactly the same, its better to cut it and substitute with a loop)
            cellsaved = refs(2)
            
            Set secondformulacell = ws.Cells(startingformulacell.Row, startingformulacell.Column + 30)
            
            
            secondformulacell = cellsaved
            
            Set thex = ws.Cells(secondformulacell.Row, secondformulacell.Column - 1)
            thex = "x"
            
            Set Depthcell = ws.Cells(secondformulacell.Row - 3, secondformulacell.Column - 7)
            
            Depthcell = Depth
            
            Set TargetCell = ws.Range(ws.Cells(secondformulacell.Row, secondformulacell.Column + 5), ws.Cells(secondformulacell.Row + 150, secondformulacell.Column + 5))
            
            Pyramidcol = secondformulacell.Column + 7
            FormulaInputNum = secondformulacell.Column - 1
            FormulaInput = secondformulacell.Column
            
            AdderFifoMioV2 TargetCell, Pyramidcol, Depth, ws, FormulaInput, FormulaInputNum
            
            
            Set thex = Nothing
            Set Depthcell = Nothing
            
            Tableleftcol = secondformulacell.Column - 4
            Tablerightcol = secondformulacell.Column + 5
            
            Formatcells ws, TargetCell, Tableleftcol, Tablerightcol, FormulaInputNum
            
            Set TargetCell = Nothing
            
            
            'Reset starting cell to repeat code
            
            LastRow = ws.Cells(ws.Rows.Count, FormulaInput).End(xlUp).Row + 10
            Set startingformulacell = ws.Cells(LastRow, startingformulacell.Column)
            
            refs.Remove 2
            refs.Remove 1
        End If
        
        
    Next cell

End Sub


Sub FullCodeRunner()
    Dim ws As Worksheet
    Dim TargetCell As Range
    Dim Depth As Long
    Dim Pyramidcol As Integer
    Dim FormulaInput As Integer
    Dim FormulaInputNum As Integer
    Dim Tableleftcol As Integer
    Dim Tablerightcol As Integer
    Dim cell As Range
    Dim SelectedCells As Range
    Dim startingformulacell As Range
    Dim StartingformulaCol As Integer
    Dim LastRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Feuil6")
    Depth = 3
    
    Set SelectedCells = Application.InputBox("Select the cells you want analyzed", Type:=8)
    Set startingformulacell = Application.InputBox("Where do you want the starting formula cell", Type:=8)
    StartingformulaCol = startingformulacell.Column

    Pyramidcol = startingformulacell.Column + 2
    FormulaInput = startingformulacell.Column - 5
    FormulaInputNum = startingformulacell.Column - 6
    Tableleftcol = ws.Range(StartingformulaCol).Column - 9
    Tablerightcol = StartingformulaCol

    For Each cell In SelectedCells
        Set TargetCell = ws.Range(ws.Cells(startingformulacell.Row, StartingformulaCol), ws.Cells(startingformulacell.Row + 300, StartingformulaCol))

        AdderFifoMioV2 TargetCell, Pyramidcol, Depth, ws, FormulaInput, FormulaInputNum
        Formatcells ws, TargetCell, Tableleftcol, Tablerightcol, FormulaInputNum
        
        LastRow = ws.Cells(ws.Rows.Count, FormulaInput).End(xlUp).Row + 20
        Set startingformulacell = ws.Cells(LastRow, StartingformulaCol)
    Next cell
End Sub




Sub Formatcells(ws As Worksheet, TargetCell As Range, Tableleftcol As Integer, Tablerightcol As Integer, FormulaInputNum As Integer)
    Dim rng As Range
    Dim cell As Range
    
    For Each cell In TargetCell
        Set rng = ws.Range(ws.Cells(cell.Row, Tableleftcol), ws.Cells(cell.Row, Tablerightcol))
        
        ' Common formatting for all cells
        With rng.Borders
            .LineStyle = xlContinuous
            .Color = RGB(255, 255, 255) ' Light Grey color
            .Weight = xlThin
        End With

        Select Case ws.Cells(cell.Row, FormulaInputNum).Value
            Case 1 ' Title
                rng.Font.Bold = True
                rng.Font.Size = 16
                rng.Font.Color = RGB(248, 248, 250) ' Black
                rng.Interior.Color = RGB(83, 99, 131) ' Light Blue
                
            Case 2 ' Subtitle
                rng.Font.Bold = False
                rng.Font.Size = 12
                rng.Font.Color = RGB(250, 250, 252) ' Dark Blue
                rng.Interior.Color = RGB(122, 138, 170) ' Moderate Blue

            Case 3 ' Subtitle 2
                rng.Font.Bold = False
                rng.Font.Size = 11
                rng.Font.Color = RGB(255, 255, 255) ' Black
                rng.Interior.Color = RGB(171, 181, 201) ' Soft Green

            Case 4 ' Section
                rng.Font.Bold = False
                rng.Font.Size = 11
                rng.Font.Color = RGB(31, 37, 49) ' Dark Grey
                rng.Interior.Color = RGB(214, 217, 229) ' Light Orange

            Case 5 ' Section 2
                rng.Font.Bold = False
                rng.Font.Size = 10
                rng.Font.Color = RGB(31, 37, 49) ' Black
                rng.Interior.Color = RGB(158, 179, 194) ' Light Red
        End Select
    Next cell
End Sub


