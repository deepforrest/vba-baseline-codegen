Option Explicit

' READY TO TEST
Sub EraseCurrentDeclarations()

    iRowCounter = iCounterInitialized
    
    Do Until iRowCounter = iRowStop
    
        ' Optimizes The Counter To Always Restart So It Only Finds A "RowStop" Number of Blanks Before Breaking:
        If oCbgDoCodeCell.Value <> sBlankCell Then iRowCounter = iCounterInitialized

        Call WriteDeclarationOutput(iNoIndent, sBlankCell, 1)
        iRowCounter = iRowCounter + 1
    
    Loop

    ' Resets Table
    iCbgDoRowScanner = iCbgDoInitialRow
    SetTableScanner (sCbgDoRowScanner)

End Sub

' READY TO TEST
Sub WriteWorkbookDeclarations()

    Call WriteDeclarationOutput(iNoIndent, CommentOut(iNoIndent, sHeaderWorkbooks), 1)

    iArrWbIndexCounter = iArrLengthInitializer

    InitializeAllTableScanners

    Do Until oCbgIiWbCodeNameCell = sBlankCell
    
        Call WriteDeclarationOutput(iSingleIndent, DeclareWorkbook(oCbgIiWbCodeNameCell), 1)
        SetNextRowTableScanner (sCbgIiWbRowScanner)
        iArrWbIndexCounter = iArrWbIndexCounter + 1

    Loop

    ' Criteria to Determine whether an Array Index is necessary:
    If iArrWbIndexCounter > iArrLengthInitializer Then

        sLeftHandSide = sDeclareConst + sNumPre + sArr + sWbInit + sFinalIndex + sAs + sTypeByte
        sRightHandSide = CStr(iArrWbIndexCounter)

        Call WriteDeclarationOutput(iSingleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)

    End If

    SetNextRowTableScanner (sCbgDoRowScanner)

End Sub

' READY TO TEST
Sub WriteWorksheetDeclarations()

    InitializeAllTableScanners
    iArrWsIndexCounter = iArrLengthInitializer

    Call WriteDeclarationOutput(iNoIndent, CommentOut(iNoIndent, sHeaderWorksheets), 1)

    Do Until oCbgIiWsCodeNameCell = sBlankCell
        
        MatchPairsUpTo (sWorksheet)
        Call WriteDeclarationOutput(iSingleIndent, DeclareWorksheet(oCbgIiWsCodeNameCell), 1)
        
        SetNextRowTableScanner (sCbgIiWsRowScanner)
        InitializeRowScannerAndTable (sCbgIiWbRowScanner)
        iArrWsIndexCounter = iArrWsIndexCounter + 1

    Loop

    ' Criteria to Determine whether an Array Index is necessary:
    If iArrWsIndexCounter > iCounterInitialized Then

        sLeftHandSide = sDeclareConst + sNumPre + sArr + sWsInit + sFinalIndex + sAs + sTypeByte
        sRightHandSide = CStr(iArrWsIndexCounter)

        Call WriteDeclarationOutput(iSingleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)

    End If

    SetNextRowTableScanner (sCbgDoRowScanner)

End Sub

' READY TO TEST
Sub WriteTableDeclarations()

    Call WriteDeclarationOutput(iNoIndent, CommentOut(iNoIndent, sHeaderTables), 1)

    InitializeAllTableScanners
    iArrTblIndexCounter = iArrLengthInitializer

    Do Until oCbgIiTblCodeNameCell = sBlankCell

        MatchPairsUpTo (sTable)
        Call WriteDeclarationOutput(iSingleIndent, DeclareTable(oCbgIiTblCodeNameCell), 1)
 
        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        SetNextRowTableScanner (sCbgIiTblRowScanner)
        iArrTblIndexCounter = iArrTblIndexCounter + 1

    Loop

    ' Criteria to Determine whether an Array Index is necessary:
    If iArrTblIndexCounter > iCounterInitialized Then

        sLeftHandSide = sDeclareConst + sNumPre + sArr + sTblInit + sFinalIndex + sAs + sTypeByte
        sRightHandSide = CStr(iArrTblIndexCounter)

        Call WriteDeclarationOutput(iSingleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)

    End If
    
    SetNextRowTableScanner (sCbgDoRowScanner)

End Sub

' READY TO TEST
Sub WriteTableElementDeclarations(sDummySuffix As String, sDataType As String)

    ' Getters
    Call WriteDeclarationOutput(iNoIndent, CommentOut(iNoIndent, GetTitle(sDummySuffix)), 1)

    InitializeAllTableScanners

    Do Until oCbgIiTblMainNameCell = sBlankCell

        ' Finds The Worksheet and Workbook For Initials
        MatchPairsUpTo (sTable)

        ' Creates A Header For The Table Selected
        sStatementToWrite = CStr(oCbgIiTblMainNameCell) + sComma + sTable + sBlankSpace + CStr(oCbgIiTblRankCell) + " from " + CStr(oCbgIiWsMainNameCell) + sBlankSpace + sWorksheet
        Call WriteDeclarationOutput(iNoIndent, CommentOut(iNoIndent, sStatementToWrite), 1)
        iColumnCounter = iCounterInitialized

        ' Finds Columns That Belong To Associated Table And Creates Their Assignments As Appropriate
        Do Until oCbgIiClmnMainNameCell = sBlankCell

            ' Checks To Make Sure Column Belongs To Selected Table
            If oCbgIiClmnTblCell = oCbgIiTblMainNameCell Then
            
                Call WriteDeclarationOutput(iSingleIndent, WriteTableElement(sDummySuffix, sDataType), 1)

            End If

            SetNextRowTableScanner (sCbgIiClmnRowScanner)

        Loop
        
        If sDummySuffix = sColumn Then
        
            sDataInit = GetDataLetter(sDataType)
            sPrefixInit = GetPrefixInitials(sDummySuffix)

            sFullPre = sDataInit + sPrefixInit

            SetNextRowTableScanner (sCbgDoRowScanner)

            sLeftHandSide = sDeclareConst + sFullPre + "TableLength" + sAs + sTypeByte
            sRightHandSide = sFinalColumn + sMinus + sStartingColumn + sPlusOne

            Call WriteDeclarationOutput(iSingleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 2)
        
        Else
        
            Call SetNextRowTableScanner(sCbgDoRowScanner, 2)
        
        End If
        
        InitializeRowScannerAndTable (sCbgIiClmnRowScanner)
        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

    SetNextRowTableScanner (sCbgDoRowScanner)

End Sub

' READY TO TEST
Sub WriteRowScannerDeclarations()

    Call WriteDeclarationOutput(iNoIndent, CommentOut(iNoIndent, sHeaderRowScanners), 1)

    InitializeAllTableScanners
    SetTableScanner (sCbgIiTblRowScanner)

    Do Until oCbgIiTblCodeNameCell = sBlankCell

        MatchPairsUpTo (sTable)

        ' Creates the Header Row:
        sRightHandSide = CStr("1")
        sStatementToWrite = DeclareConstant(oCbgIiTblCodeNameCell, sTypeByte, sRightHandSide, sHeaderRow)
        Call WriteDeclarationOutput(iSingleIndent, sStatementToWrite, 1)

        ' Creates the Initial Row:
        sRightHandSide = sNumPre + GetPrefixInitials(sColumn) + sHeaderRow + sPlusOne
        sStatementToWrite = DeclareConstant(oCbgIiTblCodeNameCell, sTypeByte, sRightHandSide, sInitialRow)
        Call WriteDeclarationOutput(iSingleIndent, sStatementToWrite, 1)

        ' Creates the RowScanner:
        oCbgDoCodeCell.Value = DeclareVariable(oCbgIiTblCodeNameCell, sTypeInteger, sRowScanner)

        If oCbgIiTblTypeCell = sConstant Then
            
            sRightHandSide = sStrPre + GetPrefixInitials(sColumn) + sRowScanner
            sExtendedStatement = DeclareConstant(oCbgIiTblCodeNameCell, sTypeString, sRightHandSide, sRowScanner)
            oCbgDoCodeCell.Value = ExtendStatement(CStr(oCbgDoCodeCell), sExtendedStatement)

        End If

        Call WriteDeclarationOutput(iSingleIndent, CStr(oCbgDoCodeCell), 2)

        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

    Call SetNextRowTableScanner(sCbgDoRowScanner, 2)

End Sub

' READY TO TEST
Sub WriteConstantDeclarations()

    Call WriteDeclarationOutput(iNoIndent, CommentOut(iNoIndent, sHeaderConstants), 1)

    InitializeAllTableScanners

    Do Until oCbgIiConstNameCell = sBlankCell

        sStatementToWrite = DeclareConstant(oCbgIiConstNameCell, oCbgIiConstTypeCell, oCbgIiConstValueCell)
        
        Call WriteDeclarationOutput(iSingleIndent, sStatementToWrite, 1)
        
        SetNextRowTableScanner (sCbgIiConstRowScanner)

    Loop

    SetNextRowTableScanner (sCbgDoRowScanner)

End Sub

' READY TO TEST
Sub WriteVariableDeclarations()

    Call WriteDeclarationOutput(iNoIndent, CommentOut(iNoIndent, sHeaderVariables), 1)

    InitializeAllTableScanners

    Do Until oCbgIiVarNameCell = sBlankCell

        sStatementToWrite = DeclareVariable(oCbgIiVarNameCell, oCbgIiVarTypeCell)
        
        Call WriteDeclarationOutput(iSingleIndent, sStatementToWrite, 1)
        
        SetNextRowTableScanner (sCbgIiVarRowScanner)

    Loop

    SetNextRowTableScanner (sCbgDoRowScanner)

End Sub

' CAN BE TESTED, BUT NOT OPTIMIZED
Sub WriteArrayDeclarations()

    Call WriteDeclarationOutput(iNoIndent, CommentOut(iNoIndent, "ARRAYS"), 1)

    If iArrWbIndexCounter > iCounterInitialized Then
    
        sStatementToWrite = "Public sArrWorkbooks(iArrWbFinalIndex) As String"
        Call WriteDeclarationOutput(iSingleIndent, sStatementToWrite, 1)
    
    End If

    If iArrWsIndexCounter > iCounterInitialized Then

        sStatementToWrite = "Public sArrWorksheets(iArrWsFinalIndex) As String"
        Call WriteDeclarationOutput(iSingleIndent, sStatementToWrite, 1)

    End If


    If iArrTblIndexCounter > iCounterInitialized Then
    
        Call WriteDeclarationOutput(iSingleIndent, "Public SArrTables(iArrTblFinalIndex) As String", 2)

        Call WriteDeclarationOutput(iSingleIndent, "Public iArrHeaderRows(iArrTblFinalIndex) As Byte", 1)
        Call WriteDeclarationOutput(iSingleIndent, "Public sArrHeaders(iArrTblFinalIndex) As String", 1)
        Call WriteDeclarationOutput(iSingleIndent, "Public iArrInitialRows(iArrTblFinalIndex) As Byte", 2)

        Call WriteDeclarationOutput(iSingleIndent, "Public iArrNumRowScanners(iArrTblFinalIndex) As Byte", 1)
        Call WriteDeclarationOutput(iSingleIndent, "Public iArrStrRowScanners(iArrTblFinalIndex) As Byte", 1)


    End If


    ' What this writer should produce:

        ' iArrTblNameClmns As Integer
        ' iArrTblNameHeaders As Object
        ' iArrTblNameCells As Object

End Sub
