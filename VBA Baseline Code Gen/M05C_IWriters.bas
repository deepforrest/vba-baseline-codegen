Option Explicit

' READY FOR TESTING
Sub ClearWbTable()

    SortTable (iCbgIiWbMainNameColumn)
    InitializeRowScannerAndTable (sCbgIiWbRowScanner)

    Do Until oCbgIiWbMainNameCell = sBlankCell

        ResetTableCells (sCbgIiWbRowScanner)

    Loop

End Sub

' READY FOR TESTING
Sub FillWsTable()

    SortTable (iCbgIiWsMainNameColumn)
    InitializeAllTableScanners

    Do Until oCbgIiWsMainNameCell = sBlankCell

        oCbgIiWsWbCell.Value = oCbgIiWbMainNameCell
        oCbgIiWsInitCell = GetCodenameInit(oCbgIiWsMainNameCell)
        If IsEmpty(oCbgIiWsTypeCell) Then oCbgIiWsTypeCell.Value = sConstant

        SetNextRowTableScanner (sCbgIiWsRowScanner)

    Loop

End Sub

' READY FOR TESTING
Sub ClearWsTable()

    SortTable (iCbgIiWsMainNameColumn)
    InitializeRowScannerAndTable (sCbgIiWsRowScanner)

    Do Until oCbgIiWsMainNameCell = sBlankCell

        ResetTableCells (sCbgIiWsRowScanner)

    Loop

End Sub

' READY FOR TESTING
Sub FillTblTable()

    SortTable (iCbgIiTblMainNameColumn)
    InitializeAllTableScanners

    Do Until oCbgIiTblMainNameCell = sBlankCell

        ' Find Associated Worksheet:
        MatchPairsUpTo (sTable)

        ' Fill In The Blanks:
        oCbgIiTblWsCell.Value = oCbgIiWsMainNameCell

        If IsEmpty(oCbgIiTblInitCell) Then oCbgIiTblInitCell.Value = GetCodenameInit(oCbgIiTblMainNameCell)
        If IsEmpty(oCbgIiTblHeaderRowCell) Then oCbgIiTblHeaderRowCell.Value = iHeaderRowDefaultValue
        If IsEmpty(oCbgIiTblTypeCell) Then oCbgIiTblTypeCell.Value = sConstant

        ' Reset And Keep Moving:
        Call InitializeRowScannerAndTable(sCbgIiWsRowScanner)
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

End Sub

' READY FOR TESTING
Sub ClearTblTable()

    InitializeRowScannerAndTable (sCbgIiTblRowScanner)
    SortTable (iCbgIiTblMainNameColumn)

    Do Until oCbgIiTblMainNameCell = sBlankCell

        ResetTableCells (sCbgIiTblRowScanner)

    Loop

End Sub

' READY FOR TESTING
Sub FillClmnTable()

    ' Preliminary Options:
    SortTable (iCbgIiClmnMainNameColumn)
    InitializeAllTableScanners

    Do Until oCbgIiClmnMainNameCell = sBlankCell

        ' Find Associated Table
        MatchPairsUpTo (sColumn)

        ' Fill Worksheets and Other Optional Items:  (Use PopulateCell(Arg1, Arg2))
        oCbgIiClmnWsCell.Value = oCbgIiWsMainNameCell
        If IsEmpty(oCbgIiClmnTypeCell) Then oCbgIiClmnTypeCell.Value = sConstant

        'Reset For Next Round
        Call InitializeRowScannerAndTable(sCbgIiWsRowScanner)
        Call InitializeRowScannerAndTable(sCbgIiTblRowScanner)
        SetNextRowTableScanner (sCbgIiClmnRowScanner)

    Loop

End Sub

' READY FOR TESTING
Sub ClearClmnTable()

    SortTable (iCbgIiClmnMainNameColumn)
    InitializeRowScannerAndTable (sCbgIiClmnRowScanner)

    Do Until oCbgIiClmnMainNameCell = sBlankCell

        ResetTableCells (sCbgIiClmnRowScanner)

    Loop

End Sub

' DO NOT TEST
Sub CreateDefaultConsts(bCreate As Boolean)

    If bCreate = True Then

        SortTable (iCbgIiConstNameColumn)
        InitializeRowScannerAndTable (sCbgIiConstRowScanner)

        Do Until oCbgIiConstNameCell = sBlankCell

            SetNextRowTableScanner (sCbgIiConstRowScanner)

        Loop

        ' PopulateConstantTableRow("Name", Type, "Value")
'        Call PopulateConstantTableRow("TableSkipFactor", sTypeByte, "2")
'        Call PopulateConstantTableRow("WorkbookNotFound", sTypeString, sWorkbookNotFound)
'        Call PopulateConstantTableRow("WorksheetNotFound", sTypeString, sWorksheetNotFound)
'        Call PopulateConstantTableRow("TableNotFound", sTypeString, sTableNotFound)
'        Call PopulateConstantTableRow("ColumnNotFound", sTypeString, sColumnNotFound)
'        Call PopulateConstantTableRow("RowNotFound", sTypeString, sRowNotFound)
'        Call PopulateConstantTableRow("VariableNotFound", sTypeString, sVariableNotFound)
'        Call PopulateConstantTableRow("ConstantNotFound", sTypeString, sConstantNotFound)


    End If

End Sub

' READY FOR TESTING
Sub ClearConstTable()

    SortTable (iCbgIiConstNameColumn)
    InitializeRowScannerAndTable (sCbgIiConstRowScanner)

    Do Until oCbgIiConstNameCell = sBlankCell

        ResetTableCells (sCbgIiConstRowScanner)
 
    Loop

End Sub

' DO NOT TEST
Sub CreateDefaultVars(bCreate As Boolean)

    If bCreate = True Then

        InitializeRowScannerAndTable (sCbgIiVarRowScanner)
        SortTable (iCbgIiVarNameColumn)

        Do Until iCbgIiVarNameCell = sBlankCell

            SetNextRowTableScanner (sCbgIiVarRowScanner)

        Loop

        ' PopulateVariableTableRow("Name", "Type")

    End If

End Sub

' READY FOR TESTING
Sub ClearVarTable()

    ' Sort the Table So That All Non-Blanks Are At The Top
    SortTable (iCbgIiVarNameColumn)
    InitializeRowScannerAndTable (sCbgIiVarRowScanner)

    Do Until oCbgIiVarNameCell = sBlankCell

        ResetTableCells (sCbgIiVarRowScanner)

    Loop

End Sub
