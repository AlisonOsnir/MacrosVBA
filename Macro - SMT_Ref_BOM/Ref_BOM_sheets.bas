Attribute VB_Name = "Ref_BOM_sheets"
Sub createSheet()
Attribute createSheet.VB_Description = "Macro developed 15/11/202 by Alison Osnir"
Attribute createSheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
' createSheet Macro
' Macro developed 15/11/202 by Alison Osnir
'
'
    Application.ScreenUpdating = False
    
    Dim FileName As String
    Dim ExcelBook1 As Workbook
    Dim ExcelSheet1 As Worksheet
    Dim ExcelSheet2 As Worksheet
    Dim ExcelSheet3 As Worksheet

    'OPEN RAW EXPORTED CRYSTAL WORKBOOK
    xlsChoosed = Application.Dialogs(xlDialogOpen).Show

    If xlsChoosed = True Then
        Set ExcelBook1 = Application.ActiveWorkbook
        Set ExcelSheet1 = ExcelBook1.ActiveSheet
        ExcelSheet1.Name = "CRYSTAL"
        
        Set ExcelSheet2 = Sheets.Add
        ExcelSheet2.Name = "MAIN"
        ExcelSheet2.Tab.ColorIndex = 23
        
        Set ExcelSheet3 = Sheets.Add(After:=Sheets(Sheets.Count))
        ExcelSheet3.Name = "CSV"
    
    'IMPORT CSV OR TXT
        
        FileName = Application.GetOpenFilename("CSV or Text Files (*.csv;*.txt),*.csv;*.txt", , "Provide Text or CSV File:")
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & FileName, _
                            Destination:=ActiveSheet.Range("A1"))       ' change to suit
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .TextFileTabDelimiter = True
            .TextFileSpaceDelimiter = True
            .Refresh
        End With
        
    'FORMAT CRYSTAL SHEET
        
        ExcelSheet1.Activate
            
        Rows("1:2").Delete Shift:=xlUp
        Columns("C:D").Delete Shift:=xlToLeft
        Columns("E:E").Delete Shift:=xlToLeft
        
        Range("C2:C5000").Cut Destination:=Range("C1:C4999")
        Range("D2:D5000").Cut Destination:=Range("D1:D4999")
        
        Range("A1:A5000").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        For i = lastRow To 1 Step -1
            If (Cells(i, "A").Value) = "Fab/Forn:" Then
                Cells(i, "A").EntireRow.Delete
            End If
        Next i
        
        Columns("A:A").Insert Shift:=xlToRight
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "1"
        lastRow = Cells(Rows.Count, "B").End(xlUp).Row
        Selection.AutoFill Destination:=Range("A1:A" & lastRow), Type:=xlFillSeries
        
        Dim rLastRow As Range
        Set rLastRow = Cells(Rows.Count, "B").End(xlUp)
        rLastRow.Offset(-1).Resize(2).EntireRow.Delete
        
        Cells.EntireColumn.AutoFit
        Columns("A:A").ColumnWidth = 6
        Columns("B:B").ColumnWidth = 16
        
        Range("A1").Select
        
    'FORMAT CSV SHEET
    
        ExcelSheet3.Activate
        
        'Rows("1:12").Delete Shift:=xlUp
        'Columns("C").Delete Shift:=xlToLeft
        
        Rows("1:1").Resize(5).Insert Shift:=xlDown
        Range("A1:E2").Interior.ColorIndex = 44
        Range("A1:E2").Font.ColorIndex = 2
        
        Range("A1").FormulaR1C1 = "AS COLUNAS DEVEM SEGUIR A SEGUINTE ORDEM:"
        Range("A2").FormulaR1C1 = "Designator"
        Range("B2").FormulaR1C1 = "Center-X"
        Range("C2").FormulaR1C1 = "Center-Y"
        Range("D2").FormulaR1C1 = "Rotation"
        Range("E2").FormulaR1C1 = "Layer"
        
        Cells.ColumnWidth = 12
        
        Cells.Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        Range("A1").Select
    
    'CREATE MAIN SHEET SEPARATING POSITIONS BASED ON CRYSTAL
    
        ExcelSheet2.Activate
        
        HeaderRow = 0
        ColLevel = 1
        ColPN = 2
        ColQty = 4
        ColRefDes = 5
        
        LineCounter2 = 1
        For LineCounter = (HeaderRow + 1) To 65536
            If ExcelSheet1.Cells(LineCounter, CInt(ColLevel)) <> "" Then
                RefDes = ""
                If Len(ExcelSheet1.Cells(LineCounter, CInt(ColRefDes))) <> 0 Then
                    For CharCounter = 1 To Len(ExcelSheet1.Cells(LineCounter, CInt(ColRefDes)))
                        If Mid(ExcelSheet1.Cells(LineCounter, CInt(ColRefDes)), CharCounter, 1) <> "," Then
                            RefDes = RefDes & Mid(ExcelSheet1.Cells(LineCounter, CInt(ColRefDes)), CharCounter, 1)
                        End If
                        If Mid(ExcelSheet1.Cells(LineCounter, CInt(ColRefDes)), CharCounter, 1) = "," Or CharCounter = Len(ExcelSheet1.Cells(LineCounter, CInt(ColRefDes))) Then
                            ExcelSheet2.Cells(LineCounter2, 1) = RefDes
                            ExcelSheet2.Cells(LineCounter2, 2) = ExcelSheet1.Cells(LineCounter, CInt(ColPN))
                            RefDes = ""
                            LineCounter2 = LineCounter2 + 1
                        End If
                    Next
                Else
                    ExcelSheet2.Cells(LineCounter2, 1) = RefDes
                    ExcelSheet2.Cells(LineCounter2, 2) = ExcelSheet1.Cells(LineCounter, CInt(ColPN))
                    RefDes = ""
                    LineCounter2 = LineCounter2 + 1
                End If
            Else
                Exit For
            End If
        Next
        
        ExcelSheet2.Range("A1:A5000").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        
    'ADD VLOOKUP TO MAIN SHEET

        ExcelSheet2.Cells(1, 3).Formula = "=VLookup($B1, CRYSTAL!$B:$C, 2, False)"
        ExcelSheet2.Cells(1, 4).Formula = "=VLookup($A1, CSV!$A:$E, 2, False)"
        ExcelSheet2.Cells(1, 5).Formula = "=VLookup($A1, CSV!$A:$E, 3, False)"
        ExcelSheet2.Cells(1, 6).Formula = "=VLookup($A1, CSV!$A:$E, 4, False)"
        ExcelSheet2.Cells(1, 7).Formula = "=VLookup($A1, CSV!$A:$E, 5, False)"
    
        Range("C1:G1").Select
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        Selection.AutoFill Destination:=Range("C1:G" & lastRow), Type:=xlFillSeries
        
    'FORMAT MAIN SHEET
    
        Cells.ColumnWidth = 15
        Columns("B:B").ColumnWidth = 18
        Columns("C:C").EntireColumn.AutoFit
        
        Cells.Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        Columns("D:E").NumberFormat = "#,##0"
        
        'ADD HEADER TO MAIN SHEET
        
        Rows("1:1").Insert Shift:=xlDown
        Range("A1").FormulaR1C1 = "Designator"
        Range("B1").FormulaR1C1 = "P/N"
        Range("C1").FormulaR1C1 = "Description"
        Range("D1").FormulaR1C1 = "Center-X(mm)"
        Range("E1").FormulaR1C1 = "Center-Y(mm)"
        Range("F1").FormulaR1C1 = "Rotation"
        Range("G1").FormulaR1C1 = "Layer"
        
        Range("A1:G1").Interior.ColorIndex = 23
        Range("A1:G1").Font.ColorIndex = 2
        Range("A1").Select
        
        'ADD FILTER AND SORT BY LAYER
        
        Range("G1").AutoFilter
        ActiveWorkbook.Worksheets("MAIN").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("MAIN").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
            "G1:G373"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("MAIN").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        'MANUAL ADJUSTMENT OF THE COLUMNS
        ExcelSheet3.Activate
        Range("A1").Select
        
    End If
End Sub
