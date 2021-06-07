Option Explicit

' READY FOR TESTING
Sub SetWorkbookAndWorksheets()

    ' Workbook Setters
    Set oCodeBitsGen = ThisWorkbook

    ' Worksheet Setters
    Set oCbgInputsInterface = oCodeBitsGen.Worksheets(sCbgInputsInterface)
    Set oCbgDeclarationsOutput = oCodeBitsGen.Worksheets(sCbgDeclarationsOutput)
    Set oCbgSettersOutput = oCodeBitsGen.Worksheets(sCbgSettersOutput)

End Sub

' READY FOR TESTING
Sub SetTablesAndHeaders(sDummyTable As String)

    Select Case sDummyTable
    
        'Case sWbWsTABLE

            'Table
            ' Set oWbWsTABLE = oWbWORKSHEET.Cells(sWbWsTABLE)

            'Header Setters
            ' Set oWbWsREPLACEHeader = oWbInputsInterface.Cells(oWbWsTblHeaderRow, oWbWsCOLUMNColumn):       oCbgIiTblCOLUMNHeader.Value = sCbgIiCOLUMNHeader
    
        Case sCbgIiWorkbooks
        
            ' Table
            Set oCbgIiWorkbooks = oCbgInputsInterface.ListObjects(sCbgIiWorkbooks)

            ' Header Setters
            Set oCbgIiWbMainNameHeader = oCbgInputsInterface.Cells(iCbgIiWbHeaderRow, iCbgIiWbMainNameColumn):       oCbgIiWbMainNameHeader.Value = sCbgIiWbMainNameHeader
            Set oCbgIiWbRankHeader = oCbgInputsInterface.Cells(iCbgIiWbHeaderRow, iCbgIiWbRankColumn):               oCbgIiWbRankHeader.Value = sCbgIiWbRankHeader
            Set oCbgIiWbCodeNameHeader = oCbgInputsInterface.Cells(iCbgIiWbHeaderRow, iCbgIiWbCodeNameColumn):       oCbgIiWbCodeNameHeader.Value = sCbgIiWbCodeNameHeader
            Set oCbgIiWbInitHeader = oCbgInputsInterface.Cells(iCbgIiWbHeaderRow, iCbgIiWbInitColumn):               oCbgIiWbInitHeader.Value = sCbgIiWbInitHeader
            Set oCbgIiWbFileNameHeader = oCbgInputsInterface.Cells(iCbgIiWbHeaderRow, iCbgIiWbFileNameColumn):       oCbgIiWbFileNameHeader.Value = sCbgIiWbFileNameHeader
        
        Case sCbgIiWorksheets
        
            ' Table
            Set oCbgIiWorksheets = oCbgInputsInterface.ListObjects(sCbgIiWorksheets)
            
            ' Header Setters
            Set oCbgIiWsWbHeader = oCbgInputsInterface.Cells(iCbgIiWsHeaderRow, iCbgIiWsWbColumn):                   oCbgIiWsWbHeader.Value = sCbgIiWsWbHeader
            Set oCbgIiWsRankHeader = oCbgInputsInterface.Cells(iCbgIiWsHeaderRow, iCbgIiWsRankColumn):               oCbgIiWsRankHeader.Value = sCbgIiWsRankHeader
            Set oCbgIiWsMainNameHeader = oCbgInputsInterface.Cells(iCbgIiWsHeaderRow, iCbgIiWsMainNameColumn):       oCbgIiWsMainNameHeader.Value = sCbgIiWsMainNameHeader
            Set oCbgIiWsCodeNameHeader = oCbgInputsInterface.Cells(iCbgIiWsHeaderRow, iCbgIiWsCodeNameColumn):       oCbgIiWsCodeNameHeader.Value = sCbgIiWsCodeNameHeader
            Set oCbgIiWsInitHeader = oCbgInputsInterface.Cells(iCbgIiWsHeaderRow, iCbgIiWsInitColumn):               oCbgIiWsInitHeader.Value = sCbgIiWsInitHeader
            Set oCbgIiWsTypeHeader = oCbgInputsInterface.Cells(iCbgIiWsHeaderRow, iCbgIiWsTypeColumn):               oCbgIiWsTypeHeader.Value = sCbgIiWsTypeHeader


        Case sCbgIiTables
        
            ' Table
            Set oCbgIiTables = oCbgInputsInterface.ListObjects(sCbgIiTables)
            
            ' Header Setters
            Set oCbgIiTblWsHeader = oCbgInputsInterface.Cells(iCbgIiTblHeaderRow, iCbgIiTblWsColumn):                   oCbgIiTblWsHeader.Value = sCbgIiTblWsHeader
            Set oCbgIiTblMainNameHeader = oCbgInputsInterface.Cells(iCbgIiTblHeaderRow, iCbgIiTblMainNameColumn):       oCbgIiTblMainNameHeader.Value = sCbgIiTblMainNameHeader
            Set oCbgIiTblCodeNameHeader = oCbgInputsInterface.Cells(iCbgIiTblHeaderRow, iCbgIiTblCodeNameColumn):       oCbgIiTblCodeNameHeader.Value = sCbgIiTblCodeNameHeader
            Set oCbgIiTblRankHeader = oCbgInputsInterface.Cells(iCbgIiTblHeaderRow, iCbgIiTblRankColumn):               oCbgIiTblRankHeader.Value = sCbgIiTblRankHeader
            Set oCbgIiTblInitHeader = oCbgInputsInterface.Cells(iCbgIiTblHeaderRow, iCbgIiTblInitColumn):               oCbgIiTblInitHeader.Value = sCbgIiTblInitHeader
            Set oCbgIiTblHeaderRowHeader = oCbgInputsInterface.Cells(iCbgIiTblHeaderRow, iCbgIiTblHeaderRowColumn):     oCbgIiTblHeaderRowHeader.Value = sCbgIiTblHeaderRowHeader
            Set oCbgIiTblTypeHeader = oCbgInputsInterface.Cells(iCbgIiTblHeaderRow, iCbgIiTblTypeColumn):               oCbgIiTblTypeHeader.Value = sCbgIiTblTypeHeader
                    
        Case sCbgIiColumns
        
            ' Table
             Set oCbgIiColumns = oCbgInputsInterface.ListObjects(sCbgIiColumns)
            
            ' Header Setters
            Set oCbgIiClmnRankHeader = oCbgInputsInterface.Cells(iCbgIiClmnHeaderRow, iCbgIiClmnRankColumn):            oCbgIiClmnRankHeader.Value = sCbgIiClmnRankHeader
            Set oCbgIiClmnWsHeader = oCbgInputsInterface.Cells(iCbgIiClmnHeaderRow, iCbgIiClmnWsColumn):                oCbgIiClmnWsHeader.Value = sCbgIiClmnWsHeader
            Set oCbgIiClmnTblHeader = oCbgInputsInterface.Cells(iCbgIiClmnHeaderRow, iCbgIiClmnTblColumn):              oCbgIiClmnTblHeader.Value = sCbgIiClmnTblHeader
            Set oCbgIiClmnMainNameHeader = oCbgInputsInterface.Cells(iCbgIiClmnHeaderRow, iCbgIiClmnMainNameColumn):    oCbgIiClmnMainNameHeader.Value = sCbgIiClmnMainNameHeader
            Set oCbgIiClmnCodeNameHeader = oCbgInputsInterface.Cells(iCbgIiClmnHeaderRow, iCbgIiClmnCodeNameColumn):    oCbgIiClmnCodeNameHeader.Value = sCbgIiClmnCodeNameHeader
            Set oCbgIiClmnTypeHeader = oCbgInputsInterface.Cells(iCbgIiClmnHeaderRow, iCbgIiClmnTypeColumn):            oCbgIiClmnTypeHeader.Value = sCbgIiClmnTypeHeader
                  
        
        Case sCbgIiConstants

            'Table
            Set oCbgIiConstants = oCbgInputsInterface.ListObjects(sCbgIiConstants)

            ' Header Setters
            Set oCbgIiConstNameHeader = oCbgInputsInterface.Cells(iCbgIiConstHeaderRow, iCbgIiConstNameColumn):       oCbgIiConstNameHeader.Value = sCbgIiConstNameHeader
            Set oCbgIiConstTypeHeader = oCbgInputsInterface.Cells(iCbgIiConstHeaderRow, iCbgIiConstTypeColumn):       oCbgIiConstTypeHeader.Value = sCbgIiConstTypeHeader
            Set oCbgIiConstValueHeader = oCbgInputsInterface.Cells(iCbgIiConstHeaderRow, iCbgIiConstValueColumn):     oCbgIiConstValueHeader.Value = sCbgIiConstValueHeader

        Case sCbgIiVariables

            ' Table
            Set oCbgIiVariables = oCbgInputsInterface.Cells(sCbgIiVariables)

            ' Header Setters
            Set oCbgIiVarNameHeader = oCbgInputsInterface.Cells(iCbgIiVarHeaderRow, iCbgIiVarNameColumn):       oCbgIiVarNameHeader.Value = sCbgIiVarNameHeader
            Set oCbgIiVarTypeHeader = oCbgInputsInterface.Cells(iCbgIiVarHeaderRow, iCbgIiVarTypeColumn):       oCbgIiVarTypeHeader.Value = sCbgIiVarTypeHeader
        
        Case Else: MsgBox (NotSupported(sTable)): End

    End Select

End Sub

' READY FOR TESTING
Sub SetTableScanner(sDummyRow As String)

    Select Case sDummyRow
    
    ' Set oWbWsTblCOLUMNCell = oWbWORKSHEET.Cells(iWbWsTblRowScanner, iWbWsCOLUMNColumn)
    
        Case sCbgIiWbRowScanner
        
            Set oCbgIiWbMainNameCell = oCbgInputsInterface.Cells(iCbgIiWbRowScanner, iCbgIiWbMainNameColumn)
            Set oCbgIiWbRankCell = oCbgInputsInterface.Cells(iCbgIiWbRowScanner, iCbgIiWbRankColumn)
            Set oCbgIiWbCodeNameCell = oCbgInputsInterface.Cells(iCbgIiWbRowScanner, iCbgIiWbCodeNameColumn)
            Set oCbgIiWbInitCell = oCbgInputsInterface.Cells(iCbgIiWbRowScanner, iCbgIiWbInitColumn)
            Set oCbgIiWbFileNameCell = oCbgInputsInterface.Cells(iCbgIiWbRowScanner, iCbgIiWbFileNameColumn)
        
        Case sCbgIiWsRowScanner
        
            Set oCbgIiWsWbCell = oCbgInputsInterface.Cells(iCbgIiWsRowScanner, iCbgIiWsWbColumn)
            Set oCbgIiWsRankCell = oCbgInputsInterface.Cells(iCbgIiWsRowScanner, iCbgIiWsRankColumn)
            Set oCbgIiWsMainNameCell = oCbgInputsInterface.Cells(iCbgIiWsRowScanner, iCbgIiWsMainNameColumn)
            Set oCbgIiWsCodeNameCell = oCbgInputsInterface.Cells(iCbgIiWsRowScanner, iCbgIiWsCodeNameColumn)
            Set oCbgIiWsInitCell = oCbgInputsInterface.Cells(iCbgIiWsRowScanner, iCbgIiWsInitColumn)
            Set oCbgIiWsTypeCell = oCbgInputsInterface.Cells(iCbgIiWsRowScanner, iCbgIiWsTypeColumn)
        
        Case sCbgIiTblRowScanner
        
            Set oCbgIiTblWsCell = oCbgInputsInterface.Cells(iCbgIiTblRowScanner, iCbgIiTblWsColumn)
            Set oCbgIiTblMainNameCell = oCbgInputsInterface.Cells(iCbgIiTblRowScanner, iCbgIiTblMainNameColumn)
            Set oCbgIiTblCodeNameCell = oCbgInputsInterface.Cells(iCbgIiTblRowScanner, iCbgIiTblCodeNameColumn)
            Set oCbgIiTblRankCell = oCbgInputsInterface.Cells(iCbgIiTblRowScanner, iCbgIiTblRankColumn)
            Set oCbgIiTblInitCell = oCbgInputsInterface.Cells(iCbgIiTblRowScanner, iCbgIiTblInitColumn)
            Set oCbgIiTblHeaderRowCell = oCbgInputsInterface.Cells(iCbgIiTblRowScanner, iCbgIiTblHeaderRowColumn)
            Set oCbgIiTblTypeCell = oCbgInputsInterface.Cells(iCbgIiTblRowScanner, iCbgIiTblTypeColumn)

        Case sCbgIiClmnRowScanner
        
            Set oCbgIiClmnRankCell = oCbgInputsInterface.Cells(iCbgIiClmnRowScanner, iCbgIiClmnRankColumn)
            Set oCbgIiClmnWsCell = oCbgInputsInterface.Cells(iCbgIiClmnRowScanner, iCbgIiClmnWsColumn)
            Set oCbgIiClmnTblCell = oCbgInputsInterface.Cells(iCbgIiClmnRowScanner, iCbgIiClmnTblColumn)
            Set oCbgIiClmnMainNameCell = oCbgInputsInterface.Cells(iCbgIiClmnRowScanner, iCbgIiClmnMainNameColumn)
            Set oCbgIiClmnCodeNameCell = oCbgInputsInterface.Cells(iCbgIiClmnRowScanner, iCbgIiClmnCodeNameColumn)
            Set oCbgIiClmnTypeCell = oCbgInputsInterface.Cells(iCbgIiClmnRowScanner, iCbgIiClmnTypeColumn)

        Case sCbgIiConstRowScanner

            Set oCbgIiConstNameCell = oCbgInputsInterface.Cells(iCbgIiConstRowScanner, iCbgIiConstNameColumn)
            Set oCbgIiConstTypeCell = oCbgInputsInterface.Cells(iCbgIiConstRowScanner, iCbgIiConstTypeColumn)
            Set oCbgIiConstValueCell = oCbgInputsInterface.Cells(iCbgIiConstRowScanner, iCbgIiConstValueColumn)

        Case sCbgIiVarRowScanner

            Set oCbgIiVarNameCell = oCbgInputsInterface.Cells(iCbgIiVarRowScanner, iCbgIiVarNameColumn)
            Set oCbgIiVarTypeCell = oCbgInputsInterface.Cells(iCbgIiVarRowScanner, iCbgIiVarTypeColumn)
        
        Case sCbgDoRowScanner
        
            Set oCbgDoCodeCell = oCbgDeclarationsOutput.Cells(iCbgDoRowScanner, iCbgDoCodeColumn)
        
        Case sCbgSoRowScanner
        
            Set oCbgSoCodeCell = oCbgSettersOutput.Cells(iCbgSoRowScanner, iCbgSoCodeColumn)
        
        Case Else: MsgBox (NotSupported(sTable)): End
    
    End Select

End Sub

' READY FOR TESTING
Sub InitializeRowScanners(sDummyWorksheet As String)

    Select Case sDummyWorksheet

        Case sCbgInputsInterface

            iCbgIiWbRowScanner = iCbgIiWbInitialRow
            iCbgIiWsRowScanner = iCbgIiWsInitialRow
            iCbgIiTblRowScanner = iCbgIiTblInitialRow
            iCbgIiClmnRowScanner = iCbgIiClmnInitialRow
            iCbgIiConstRowScanner = iCbgIiConstInitialRow
            iCbgIiVarRowScanner = iCbgIiVarInitialRow

        Case sCbgDeclarationsOutput
    
            iCbgDoRowScanner = iCbgDoInitialRow

        Case sCbgSettersOutput
        
            iCbgSoRowScanner = iCbgSoInitialRow

        Case Else: MsgBox (NotSupported(sWorksheet)): End

    End Select

End Sub

' READY FOR TESTING
Sub InitializeRowScannerAndTable(sDummyRow As String)

    Select Case sDummyRow

        ' --------STRING--------    ---------------NUMERIC---------------
        Case sCbgIiWbRowScanner:    iCbgIiWbRowScanner = iCbgIiWbInitialRow
        Case sCbgIiWsRowScanner:    iCbgIiWsRowScanner = iCbgIiWsInitialRow
        Case sCbgIiTblRowScanner:   iCbgIiTblRowScanner = iCbgIiTblInitialRow
        Case sCbgIiClmnRowScanner:  iCbgIiClmnRowScanner = iCbgIiClmnInitialRow
        Case sCbgIiConstRowScanner: iCbgIiConstRowScanner = iCbgIiConstInitialRow
        Case sCbgIiVarRowScanner:   iCbgIiVarRowScanner = iCbgIiVarInitialRow
        
        Case Else: NotSupported (sStrRowScanner): End

    End Select

    SetTableScanner (sDummyRow)

End Sub

' READY FOR TESTING
Sub InitializeAllTableScanners()

    ' Could Be Done With An Array
    InitializeRowScanners (sCbgInputsInterface)

    ' Do Until iArrIndex = UBound(sArrRowScanner)  (May require an Array Length Function)

        ' SetTableScanner(sArrRowScanner(iArrIndex))
        ' iArrIndex = iArrIndex + 1

    ' Loop

    SetTableScanner (sCbgIiWbRowScanner)
    SetTableScanner (sCbgIiWsRowScanner)
    SetTableScanner (sCbgIiTblRowScanner)
    SetTableScanner (sCbgIiClmnRowScanner)
    SetTableScanner (sCbgIiConstRowScanner)
    SetTableScanner (sCbgIiVarRowScanner)

End Sub

' READY FOR TESTING
Sub SetNextRowTableScanner(sDummyRow As String, Optional iStepValue As Integer)

    If IsMissing(iStepValue) Or iStepValue < 1 Then iStepValue = 1  'Automatically goes to next row

    Select Case sDummyRow

        'Case sWbWsTblRowScanner:   iWbWsTblRowScanner = iWbWsTblRowScanner + iStepValue
        
        Case sCbgIiWbRowScanner:    iCbgIiWbRowScanner = iCbgIiWbRowScanner + iStepValue
        Case sCbgIiWsRowScanner:    iCbgIiWsRowScanner = iCbgIiWsRowScanner + iStepValue
        Case sCbgIiTblRowScanner:   iCbgIiTblRowScanner = iCbgIiTblRowScanner + iStepValue
        Case sCbgIiClmnRowScanner:  iCbgIiClmnRowScanner = iCbgIiClmnRowScanner + iStepValue
        Case sCbgIiConstRowScanner: iCbgIiConstRowScanner = iCbgIiConstRowScanner + iStepValue
        Case sCbgIiVarRowScanner:   iCbgIiVarRowScanner = iCbgIiVarRowScanner + iStepValue

        Case sCbgDoRowScanner:      iCbgDoRowScanner = iCbgDoRowScanner + iStepValue
        Case sCbgSoRowScanner:      iCbgSoRowScanner = iCbgSoRowScanner + iStepValue
        
        Case Else:                  MsgBox (NotSupported(sStrRowScanner)): End
            
    End Select

    SetTableScanner (sDummyRow)

End Sub

' READY FOR TESTING
Sub ResetTableCells(sDummyRowScanner As String)

    Select Case sDummyRowScanner

        ' Would be better to generate code for CodeName Getters as a function of Main Name
        Case sCbgIiWbRowScanner
        
            oCbgIiWbMainNameCell.Value = sBlankCell
            oCbgIiWbRankCell.Value = sBlankCell
            oCbgIiWbCodeNameCell.Formula = "=SUBSTITUTE([@Workbook], "" "", """") "
            oCbgIiWbInitCell.Value = sBlankCell
            oCbgIiWbFileNameCell.Value = sBlankCell
        
        Case sCbgIiWsRowScanner
        
            oCbgIiWsWbCell.Value = sBlankCell
            oCbgIiWsRankCell.Value = sBlankCell
            oCbgIiWsMainNameCell.Value = sBlankCell
            oCbgIiWsCodeNameCell.Formula = "=SUBSTITUTE([@Worksheet], "" "", """") "
            oCbgIiWsInitCell.Value = sBlankCell
            oCbgIiWsTypeCell.Value = sBlankCell
        
        Case sCbgIiTblRowScanner
        
            oCbgIiTblWsCell.Value = sBlankCell
            oCbgIiTblRankCell.Value = sBlankCell
            oCbgIiTblMainNameCell.Value = sBlankCell
            oCbgIiTblCodeNameCell.Formula = "=SUBSTITUTE([@Table], "" "", """")"
            oCbgIiTblHeaderRowCell.Value = sBlankCell
            oCbgIiTblInitCell.Value = sBlankCell
            oCbgIiTblTypeCell.Value = sBlankCell

        Case sCbgIiClmnRowScanner
        
            oCbgIiClmnRankCell.Value = sBlankCell
            oCbgIiClmnWsCell.Value = sBlankCell
            oCbgIiClmnTblCell.Value = sBlankCell
            oCbgIiClmnMainNameCell.Value = sBlankCell
            oCbgIiClmnCodeNameCell.Formula = "=SUBSTITUTE([@Name], "" "", """") "
            oCbgIiClmnRankCell.Value = sBlankCell
            oCbgIiClmnTypeCell.Value = sBlankCell

        Case sCbgIiConstRowScanner

            oCbgIiConstNameCell.Value = sBlankCell
            oCbgIiConstTypeCell.Value = sBlankCell
            oCbgIiConstValueCell.Value = sBlankCell

        Case sCbgIiVarRowScanner

            oCbgIiVarNameCell.Value = sBlankCell
            oCbgIiVarTypeCell.Value = sBlankCell

        Case Else: MsgBox (NotSupported(sTable)): End

    End Select

    SetNextRowTableScanner (sDummyRowScanner)

End Sub

'*******************ARRAYS*******************'

' READY FOR TESTING
Sub SetColumnArrays(sTypeObject As String)

    Select Case sTypeObject

        Case sColumn
            ' Workbooks
                iArrWorkbookColumns(iCbgIiWbMainNameIndex) = iCbgIiWbMainNameColumn
                iArrWorkbookColumns(iCbgIiWbRankIndex) = iCbgIiWbRankColumn
                iArrWorkbookColumns(iCbgIiWbCodenameIndex) = iCbgIiWbCodeNameColumn
                iArrWorkbookColumns(iCbgIiWbInitIndex) = iCbgIiWbInitColumn
                iArrWorkbookColumns(iCbgIiWbFileNameIndex) = iCbgIiWbFileNameColumn

            ' Worksheets
                iArrWorksheetColumns(iCbgIiWsWbIndex) = iCbgIiWsWbColumn
                iArrWorksheetColumns(iCbgIiWsRankIndex) = iCbgIiWsRankColumn
                iArrWorksheetColumns(iCbgIiWsMainNameIndex) = iCbgIiWsMainNameColumn
                iArrWorksheetColumns(iCbgIiWsCodeNameIndex) = iCbgIiWsCodeNameColumn
                iArrWorksheetColumns(iCbgIiWsInitIndex) = iCbgIiWsInitColumn
                iArrWorksheetColumns(iCbgIiWsTypeIndex) = iCbgIiWsTypeColumn

            ' Tables
                iArrTableColumns(iCbgIiTblWsIndex) = iCbgIiTblWsColumn
                iArrTableColumns(iCbgIiTblMainNameIndex) = iCbgIiTblMainNameColumn
                iArrTableColumns(iCbgIiTblCodeNameIndex) = iCbgIiTblCodeNameColumn
                iArrTableColumns(iCbgIiTblRankIndex) = iCbgIiTblRankColumn
                iArrTableColumns(iCbgIiTblInitIndex) = iCbgIiTblInitColumn
                iArrTableColumns(iCbgIiTblHeaderRowIndex) = iCbgIiTblHeaderRowColumn
                iArrTableColumns(iCbgIiTblTypeIndex) = iCbgIiTblTypeColumn

            ' Columns
                iArrColumnColumns(iCbgIiClmnRankIndex) = iCbgIiClmnRankColumn
                iArrColumnColumns(iCbgIiClmnWsIndex) = iCbgIiClmnWsColumn
                iArrColumnColumns(iCbgIiClmnTblIndex) = iCbgIiClmnTblColumn
                iArrColumnColumns(iCbgIiClmnMainNameIndex) = iCbgIiClmnMainNameColumn
                iArrColumnColumns(iCbgIiClmnCodeNameIndex) = iCbgIiClmnCodeNameColumn
                iArrColumnColumns(iCbgIiClmnTypeIndex) = iCbgIiClmnTypeColumn

            ' Constants
                iArrConstantsColumns(iCbgIiConstNameIndex) = iCbgIiConstNameColumn
                iArrConstantsColumns(iCbgIiConstTypeIndex) = iCbgIiConstTypeColumn
                iArrConstantsColumns(iCbgIiConstValueIndex) = iCbgIiConstValueColumn

            ' Variables
                iArrVariablesColumns(iCbgIiVarNameIndex) = iCbgIiVarNameColumn
                iArrVariablesColumns(iCbgIiVarTypeIndex) = iCbgIiVarTypeColumn

            ' Super Array
                iArrAllTableColumns(iArrWorkbookIndex) = iArrWorkbookColumns
                iArrAllTableColumns(iArrWorksheetIndex) = iArrWorksheetColumns
                iArrAllTableColumns(iArrTableIndex) = iArrTableColumns
                iArrAllTableColumns(iArrColumnIndex) = iArrColumnColumns
                iArrAllTableColumns(iArrConstantsIndex) = iArrConstantsColumns
                iArrAllTableColumns(iArrVariableIndex) = iArrVariablesColumns

    
        Case sCell

            ' Workbooks
                oArrWorkbookCells(iCbgIiWbMainNameIndex) = oCbgIiWbMainNameCell
                oArrWorkbookCells(iCbgIiWbRankIndex) = oCbgIiWbRankCell
                oArrWorkbookCells(iCbgIiWbCodenameIndex) = oCbgIiWbCodeNameCell
                oArrWorkbookCells(iCbgIiWbInitIndex) = oCbgIiWbInitCell
                oArrWorkbookCells(iCbgIiWbFileNameIndex) = oCbgIiWbFileNameCell

            ' Worksheets
                oArrWorksheetCells(iCbgIiWsWbIndex) = oCbgIiWsWbCell
                oArrWorksheetCells(iCbgIiWsRankIndex) = oCbgIiWsRankCell
                oArrWorksheetCells(iCbgIiWsMainNameIndex) = oCbgIiWsMainNameCell
                oArrWorksheetCells(iCbgIiWsCodeNameIndex) = oCbgIiWsCodeNameCell
                oArrWorksheetCells(iCbgIiWsInitIndex) = oCbgIiWsInitCell
                oArrWorksheetCells(iCbgIiWsTypeIndex) = oCbgIiWsTypeCell

            ' Tables
                oArrTableCells(iCbgIiTblWsIndex) = oCbgIiTblWsCell
                oArrTableCells(iCbgIiTblMainNameIndex) = oCbgIiTblMainNameCell
                oArrTableCells(iCbgIiTblCodeNameIndex) = oCbgIiTblCodeNameCell
                oArrTableCells(iCbgIiTblRankIndex) = oCbgIiTblRankCell
                oArrTableCells(iCbgIiTblInitIndex) = oCbgIiTblInitCell
                oArrTableCells(iCbgIiTblHeaderRowIndex) = oCbgIiTblHeaderRowCell
                oArrTableCells(iCbgIiTblTypeIndex) = oCbgIiTblTypeCell

            ' Columns
                oArrColumnCells(iCbgIiClmnRankIndex) = oCbgIiClmnRankCell
                oArrColumnCells(iCbgIiClmnWsIndex) = oCbgIiClmnWsCell
                oArrColumnCells(iCbgIiClmnTblIndex) = oCbgIiClmnTblCell
                oArrColumnCells(iCbgIiClmnMainNameIndex) = oCbgIiClmnMainNameCell
                oArrColumnCells(iCbgIiClmnCodeNameIndex) = oCbgIiClmnCodeNameCell
                oArrColumnCells(iCbgIiClmnTypeIndex) = oCbgIiClmnTypeCell

            ' Constants
                oArrConstantsCells(iCbgIiConstNameIndex) = oCbgIiConstNameCell
                oArrConstantsCells(iCbgIiConstTypeIndex) = oCbgIiConstTypeCell
                oArrConstantsCells(iCbgIiConstValueIndex) = oCbgIiConstValueCell

            ' Variables
                oArrVariablesCells(iCbgIiVarNameIndex) = oCbgIiVarNameCell
                oArrVariablesCells(iCbgIiVarTypeIndex) = oCbgIiVarTypeCell

            ' Super Array
                oArrAllTableCells(iArrWorkbookIndex) = oArrWorkbookCells
                oArrAllTableCells(iArrWorksheetIndex) = oArrWorksheetCells
                oArrAllTableCells(iArrTableIndex) = oArrTableCells
                oArrAllTableCells(iArrColumnIndex) = oArrColumnCells
                oArrAllTableCells(iArrConstantsIndex) = oArrConstantsCells
                oArrAllTableCells(iArrVariableIndex) = oArrVariablesCells

        Case sHeader
            ' I - OBJECTS
                ' Workbooks
                    oArrWorkbookHeaders(iCbgIiWbMainNameIndex) = oCbgIiWbMainNameHeader
                    oArrWorkbookHeaders(iCbgIiWbRankIndex) = oCbgIiWbRankHeader
                    oArrWorkbookHeaders(iCbgIiWbCodenameIndex) = oCbgIiWbCodeNameHeader
                    oArrWorkbookHeaders(iCbgIiWbInitIndex) = oCbgIiWbInitHeader
                    oArrWorkbookHeaders(iCbgIiWbFileNameIndex) = oCbgIiWbFileNameHeader

                ' Worksheets
                    oArrWorksheetHeaders(iCbgIiWsWbIndex) = oCbgIiWsWbHeader
                    oArrWorksheetHeaders(iCbgIiWsRankIndex) = oCbgIiWsRankHeader
                    oArrWorksheetHeaders(iCbgIiWsMainNameIndex) = oCbgIiWsMainNameHeader
                    oArrWorksheetHeaders(iCbgIiWsCodeNameIndex) = oCbgIiWsCodeNameHeader
                    oArrWorksheetHeaders(iCbgIiWsInitIndex) = oCbgIiWsInitHeader
                    oArrWorksheetHeaders(iCbgIiWsTypeIndex) = oCbgIiWsTypeHeader

                ' Tables
                    oArrTableHeaders(iCbgIiTblWsIndex) = oCbgIiTblWsHeader
                    oArrTableHeaders(iCbgIiTblMainNameIndex) = oCbgIiTblMainNameHeader
                    oArrTableHeaders(iCbgIiTblCodeNameIndex) = oCbgIiTblCodeNameHeader
                    oArrTableHeaders(iCbgIiTblRankIndex) = oCbgIiTblRankHeader
                    oArrTableHeaders(iCbgIiTblInitIndex) = oCbgIiTblInitHeader
                    oArrTableHeaders(iCbgIiTblHeaderRowIndex) = oCbgIiTblHeaderRowHeader
                    oArrTableHeaders(iCbgIiTblTypeIndex) = oCbgIiTblTypeHeader

                ' Columns
                    oArrColumnHeaders(iCbgIiClmnRankIndex) = oCbgIiClmnRankHeader
                    oArrColumnHeaders(iCbgIiClmnWsIndex) = oCbgIiClmnWsHeader
                    oArrColumnHeaders(iCbgIiClmnTblIndex) = oCbgIiClmnTblHeader
                    oArrColumnHeaders(iCbgIiClmnMainNameIndex) = oCbgIiClmnMainNameHeader
                    oArrColumnHeaders(iCbgIiClmnCodeNameIndex) = oCbgIiClmnCodeNameHeader
                    oArrColumnHeaders(iCbgIiClmnTypeIndex) = oCbgIiClmnTypeHeader

                ' Constants
                    oArrConstantsHeaders(iCbgIiConstNameIndex) = oCbgIiConstNameHeader
                    oArrConstantsHeaders(iCbgIiConstTypeIndex) = oCbgIiConstTypeHeader
                    oArrConstantsHeaders(iCbgIiConstValueIndex) = oCbgIiConstValueHeader

                ' Variables
                    oArrVariablesHeaders(iCbgIiVarNameIndex) = oCbgIiVarNameHeader
                    oArrVariablesHeaders(iCbgIiVarTypeIndex) = oCbgIiVarTypeHeader

                ' Super Array
                    oArrAllTableHeaders(iArrWorkbookIndex) = oArrWorkbookHeaders
                    oArrAllTableHeaders(iArrWorksheetIndex) = oArrWorksheetHeaders
                    oArrAllTableHeaders(iArrTableIndex) = oArrTableHeaders
                    oArrAllTableHeaders(iArrColumnIndex) = oArrColumnHeaders
                    oArrAllTableHeaders(iArrConstantsIndex) = oArrConstantsHeaders
                    oArrAllTableHeaders(iArrVariableIndex) = oArrVariablesHeaders
    
            ' II - Strings
                ' Workbooks
                    sArrWorkbookHeaders(iCbgIiWbMainNameIndex) = sCbgIiWbMainNameHeader
                    sArrWorkbookHeaders(iCbgIiWbRankIndex) = sCbgIiWbRankHeader
                    sArrWorkbookHeaders(iCbgIiWbCodenameIndex) = sCbgIiWbCodeNameHeader
                    sArrWorkbookHeaders(iCbgIiWbInitIndex) = sCbgIiWbInitHeader
                    sArrWorkbookHeaders(iCbgIiWbFileNameIndex) = sCbgIiWbFileNameHeader

                ' Worksheets
                    sArrWorksheetHeaders(iCbgIiWsWbIndex) = sCbgIiWsWbHeader
                    sArrWorksheetHeaders(iCbgIiWsRankIndex) = sCbgIiWsRankHeader
                    sArrWorksheetHeaders(iCbgIiWsMainNameIndex) = sCbgIiWsMainNameHeader
                    sArrWorksheetHeaders(iCbgIiWsCodeNameIndex) = sCbgIiWsCodeNameHeader
                    sArrWorksheetHeaders(iCbgIiWsInitIndex) = sCbgIiWsInitHeader
                    sArrWorksheetHeaders(iCbgIiWsTypeIndex) = sCbgIiWsTypeHeader

                ' Tables
                    sArrTableHeaders(iCbgIiTblWsIndex) = sCbgIiTblWsHeader
                    sArrTableHeaders(iCbgIiTblMainNameIndex) = sCbgIiTblMainNameHeader
                    sArrTableHeaders(iCbgIiTblCodeNameIndex) = sCbgIiTblCodeNameHeader
                    sArrTableHeaders(iCbgIiTblRankIndex) = sCbgIiTblRankHeader
                    sArrTableHeaders(iCbgIiTblInitIndex) = sCbgIiTblInitHeader
                    sArrTableHeaders(iCbgIiTblHeaderRowIndex) = sCbgIiTblHeaderRowHeader
                    sArrTableHeaders(iCbgIiTblTypeIndex) = sCbgIiTblTypeHeader

                ' Columns
                    sArrColumnHeaders(iCbgIiClmnRankIndex) = sCbgIiClmnRankHeader
                    sArrColumnHeaders(iCbgIiClmnWsIndex) = sCbgIiClmnWsHeader
                    sArrColumnHeaders(iCbgIiClmnTblIndex) = sCbgIiClmnTblHeader
                    sArrColumnHeaders(iCbgIiClmnMainNameIndex) = sCbgIiClmnMainNameHeader
                    sArrColumnHeaders(iCbgIiClmnCodeNameIndex) = sCbgIiClmnCodeNameHeader
                    sArrColumnHeaders(iCbgIiClmnTypeIndex) = sCbgIiClmnTypeHeader

                ' Constants
                    sArrConstantsHeaders(iCbgIiConstNameIndex) = sCbgIiConstNameHeader
                    sArrConstantsHeaders(iCbgIiConstTypeIndex) = sCbgIiConstTypeHeader
                    sArrConstantsHeaders(iCbgIiConstValueIndex) = sCbgIiConstValueHeader

                ' Variables
                    sArrVariablesHeaders(iCbgIiVarNameIndex) = sCbgIiVarNameHeader
                    sArrVariablesHeaders(iCbgIiVarTypeIndex) = sCbgIiVarTypeHeader

                ' Super Array
                    sArrAllTableHeaders(iArrWorkbookIndex) = sArrWorkbookHeaders
                    sArrAllTableHeaders(iArrWorksheetIndex) = sArrWorksheetHeaders
                    sArrAllTableHeaders(iArrTableIndex) = sArrTableHeaders
                    sArrAllTableHeaders(iArrColumnIndex) = sArrColumnHeaders
                    sArrAllTableHeaders(iArrConstantsIndex) = sArrConstantsHeaders
                    sArrAllTableHeaders(iArrVariableIndex) = sArrVariablesHeaders
        
        Case sRowScanner

            ' I - NUMERICAL
                iArrRowScanners(iArrWorkbookIndex) = iCbgIiWbRowScanner
                iArrRowScanners(iArrWorksheetIndex) = iCbgIiWsRowScanner
                iArrRowScanners(iArrTableIndex) = iCbgIiTblRowScanner
                iArrRowScanners(iArrColumnIndex) = iCbgIiClmnRowScanner
                iArrRowScanners(iArrConstantsIndex) = iCbgIiConstRowScanner
                iArrRowScanners(iArrVariableIndex) = iCbgIiVarRowScanner
                iArrRowScanners(iArrDeclarationsIndex) = iCbgSoRowScanner
                iArrRowScanners(iArrSettersIndex) = iCbgDoRowScanner

            ' II - STRING
                sArrRowScanners(iArrWorkbookIndex) = sCbgIiWbRowScanner
                sArrRowScanners(iArrWorksheetIndex) = sCbgIiWsRowScanner
                sArrRowScanners(iArrTableIndex) = sCbgIiTblRowScanner
                sArrRowScanners(iArrColumnIndex) = sCbgIiClmnRowScanner
                sArrRowScanners(iArrConstantsIndex) = sCbgIiConstRowScanner
                sArrRowScanners(iArrVariableIndex) = sCbgIiVarRowScanner
                sArrRowScanners(iArrDeclarationsIndex) = sCbgSoRowScanner
                sArrRowScanners(iArrSettersIndex) = sCbgDoRowScanner

        Case Else: NotSupported (sTypeObject): End
    
    End Select

End Sub


' DEFER TESTING
Sub SetObjectTypeArray()

    Dim sArrObjectTypes(6) As String

    sArrObjectTypes(0) = sWorkbook
    sArrObjectTypes(1) = sWorksheet
    sArrObjectTypes(2) = sTable
    sArrObjectTypes(3) = sNumRowScanner
    sArrObjectTypes(4) = sStrRowScanner
    sArrObjectTypes(5) = sHeaderRow
    sArrObjectTypes(6) = sInitialRow

End Sub
