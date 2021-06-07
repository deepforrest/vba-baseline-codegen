Option Explicit

' READY TO TEST
Sub EraseCurrentSetters()

    iRowCounter = iCounterInitialized
    
    Do Until iRowCounter = iRowStop ' = 4
    
        ' Optimizes The Counter To Always Restart So It Only Finds A "RowStop" Number of Blanks Before Breaking:
        If oCbgSoCodeCell.Value <> sBlankCell Then iRowCounter = iCounterInitialized

        Call WriteSetterOutput(iNoIndent, sBlankCell, 1)
        iRowCounter = iRowCounter + 1
        
    Loop

    'oCbgSoCodeCell.Select
    
    ' Resets Table
    iCbgSoRowScanner = iCbgSoInitialRow
    SetTableScanner (sCbgSoRowScanner)

End Sub

' READY TO TEST
Sub WriteWbWsSetters()

    sSubTitle = "SetWorkbooksAndWorksheets"

    ' Writes Header
    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle), 2)

    InitializeAllTableScanners

    Call WriteSetterOutput(iSingleIndent, CommentOut(iSingleIndent, sWorkbookSetters), 1)

    ' Workbook Setters
    Do Until oCbgIiWbMainNameCell = sBlankCell

        ' ******Down The Road, Make Sure The File Name String Is Set Up******
        sLeftHandSide = sSet + sObjPre + CStr(oCbgIiWbCodeNameCell)
        sRightHandSide = sThisWorkbook

        Call WriteSetterOutput(iDoubleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)
        SetNextRowTableScanner (sCbgIiWbRowScanner)

    Loop

    SetNextRowTableScanner (sCbgSoRowScanner)
    Call InitializeRowScannerAndTable(sCbgIiWbRowScanner)

    Call WriteSetterOutput(iSingleIndent, CommentOut(iSingleIndent, sWorksheetSetters), 1)

    Do Until oCbgIiWsMainNameCell = sBlankCell

        ' ******When The Time Comes, Create Do Loop For Multiple Workbooks. Not Necessary At This Time.******

        If oCbgIiWsTypeCell = sConstant Then
        
            sWsPrefixInit = GetPrefixInitials(sWorksheet)

            sLeftHandSide = sSet + sObjPre + sWsPrefixInit + CStr(oCbgIiWsCodeNameCell)
            sRightHandSide = sObjPre + CStr(oCbgIiWbCodeNameCell) + sDotWorksheets + sLeftPar + sStrPre + sWsPrefixInit + CStr(oCbgIiWsCodeNameCell) + sRightPar
        
            Call WriteSetterOutput(iDoubleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)
            
        End If

        SetNextRowTableScanner (sCbgIiWsRowScanner)

    Loop

    WriteEndSub

End Sub

' READY TO TEST
Sub WriteSetTablesAndHeaders()

    sSubTitle = "SetTableAndHeaders"
    sArgName_1 = "sDummyTable"
    
    ' Preliminary Set Up
    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle, sArgName_1, sTypeString), 2)
    Call WriteSetterOutput(iSingleIndent, sSelectOpener + sArgName_1, 2)

    ' Manual Replacement Code
    Call WriteSetterOutput(iDoubleIndent, CommentOut(iDoubleIndent, sCase + "sWbWsTABLE"), 2)
    Call WriteSetterOutput(iTripleIndent, CommentOut(iTripleIndent, sTableSetters), 1)

    Call WriteSetterOutput(iQuadIndent, CommentOut(iQuadIndent, sSourceCodeTable), 2)
    Call WriteSetterOutput(iTripleIndent, CommentOut(iTripleIndent, sHeaderSetters), 1)

    Call WriteSetterOutput(iQuadIndent, CommentOut(iQuadIndent, sSourceCodeCell), 2)

    InitializeAllTableScanners

    Do Until oCbgIiTblMainNameCell = sBlankCell

        ' Verify Worksheet Association for Initials
        MatchPairsUpTo (sTable)

        sPrefixInit = sStrPre + GetPrefixInitials(sTable)
        sStatementToWrite = sCase + sPrefixInit + CStr(oCbgIiTblCodeNameCell)

        Call WriteSetterOutput(iDoubleIndent, sStatementToWrite, 2)

        Call WriteSetterOutput(iTripleIndent, CommentOut(iTripleIndent, sTableSetters), 1)

        sStatementToWrite = WriteSetTable(oCbgIiTblCodeNameCell)
        Call WriteSetterOutput(iQuadIndent, sStatementToWrite, 2)

        Call WriteSetterOutput(iTripleIndent, CommentOut(iTripleIndent, sHeaderSetters), 1)

        ' Write A Setter For The Headers if the Table Matches
        Do Until oCbgIiClmnMainNameCell = sBlankCell

            If oCbgIiClmnTblCell = oCbgIiTblMainNameCell Then
            
                sStatementToWrite = WriteSetHeader(oCbgIiClmnCodeNameCell, oCbgIiClmnTypeCell)
                Call WriteSetterOutput(iQuadIndent, sStatementToWrite, 1)

            End If

            SetNextRowTableScanner (sCbgIiClmnRowScanner)

        Loop

        SetNextRowTableScanner (sCbgSoRowScanner)
        
        ' Takes Worksheets and Columns To The Beginning For the Next Go
        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        InitializeRowScannerAndTable (sCbgIiClmnRowScanner)

        ' Moves To Next Table
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

    Call WriteSetterOutput(iDoubleIndent, "NotSupported" + sLeftPar + sArgName_1 + sRightPar, 2)

    Call WriteSetterOutput(iSingleIndent, sSelectCloser, 1)

    WriteEndSub

End Sub

' READY TO TEST
Sub WriteSetTableScanner()

    sSubTitle = "SetTableScanner"
    sArgName_1 = "sDummyRow"

    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle, sArgName_1, sTypeString), 2)
    Call WriteSetterOutput(iSingleIndent, sSelectOpener + sArgName_1, 2)

    Call WriteSetterOutput(iDoubleIndent, CommentOut(iDoubleIndent, sSourceCodeCell), 2)

    InitializeAllTableScanners

    Do Until oCbgIiTblMainNameCell = sBlankCell

        ' Find The Relevant Worksheet First:
        MatchPairsUpTo (sColumn)

        sStatementToWrite = sCase + sStrPre + GetPrefixInitials(sColumn) + sRowScanner
        Call WriteSetterOutput(iDoubleIndent, sStatementToWrite, 2)

        Do Until oCbgIiClmnMainNameCell = sBlankCell

            sPrefixInit = GetPrefixInitials(sColumn)
            
            If oCbgIiClmnTblCell = oCbgIiTblMainNameCell Then
            
                sStatementToWrite = WriteSetCell(oCbgIiClmnCodeNameCell, sCell)
                Call WriteSetterOutput(iTripleIndent, sStatementToWrite, 1)

            End If

            SetNextRowTableScanner (sCbgIiClmnRowScanner)

        Loop

        Call InitializeRowScannerAndTable(sCbgIiWsRowScanner)
        Call InitializeRowScannerAndTable(sCbgIiClmnRowScanner)

        SetNextRowTableScanner (sCbgSoRowScanner)
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

    SetNextRowTableScanner (sCbgSoRowScanner)
    Call WriteSetterOutput(iDoubleIndent, WriteCaseElse(sSubTitle), 2)
    Call WriteSetterOutput(iSingleIndent, sSelectCloser, 1)

    WriteEndSub

End Sub

' READY TO TEST
Sub WriteIndividualRSInit()

    sSubTitle = "InitializeRowScanner"
    sArgName_1 = "sDummyRow"

    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle, sArgName_1, sTypeString), 2)

    InitializeAllTableScanners

    Call WriteSetterOutput(iSingleIndent, sSelectOpener + sArgName_1, 2)

    Do Until oCbgIiTblMainNameCell = sBlankCell

        MatchPairsUpTo (sTable)

        sStatementToWrite = sCase + sStrPre + GetPrefixInitials(sRowScanner) + sRowScanner
        Call WriteSetterOutput(iDoubleIndent, ExtendStatement(sStatementToWrite, WriteInitializeRowScanner), 1)
        
        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

    Call WriteSetterOutput(iDoubleIndent, WriteCaseElse(sSubTitle), 2)

    Call WriteSetterOutput(iSingleIndent, sSelectCloser, 1)

    WriteEndSub

End Sub

' READY TO TEST
Sub WriteInitAllRS()

    sSubTitle = "InitializeAllRowScanners"

    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle), 2)

    InitializeAllTableScanners

    Do Until oCbgIiTblMainNameCell = sBlankCell

        MatchPairsUpTo (sTable)

        Call WriteSetterOutput(iSingleIndent, WriteInitializeRowScanner, 1)
        
        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

    WriteEndSub

End Sub

' READY TO TEST
Sub WriteInitializeAllTables()

    sSubTitle = "InitializeAllTableScanners"

    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle), 2)

    InitializeAllTableScanners
    
    sCallSub = "InitializeAllRowScanners"

    Call WriteSetterOutput(iSingleIndent, sCallSub, 2)

    Do Until oCbgIiTblMainNameCell = sBlankCell

        MatchPairsUpTo (sTable)

        sPrefixInit = GetPrefixInitials(sRowScanner)

        sStatementToWrite = "SetTableScanner" + sLeftPar + sStrPre + sPrefixInit + sRowScanner + sRightPar
        Call WriteSetterOutput(iSingleIndent, sStatementToWrite, 1)

        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

    WriteEndSub

End Sub

' READY TO TEST
Sub WriteGoToNextRow()

    sSubTitle = "SetNextRow"
    sArgName_1 = "sDummyRow"
    sArgName_2 = "iStepValue"
    
    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle, sArgName_1, sTypeString, sArgName_2, sTypeInteger), 2)

    sStatementToWrite = "If " + sArgName_2 + " < 1 Then " + sArgName_2 + " = 1"
    Call WriteSetterOutput(iSingleIndent, sStatementToWrite, 2)

    Call WriteSetterOutput(iSingleIndent, sSelectOpener + sArgName_1, 2)

    Call WriteSetterOutput(iDoubleIndent, CommentOut(iDoubleIndent, sSourceCodeSetNextRow), 2)

    InitializeAllTableScanners

    Do Until oCbgIiTblMainNameCell = sBlankCell

        MatchPairsUpTo (sTable)

        Call WriteSetterOutput(iDoubleIndent, WriteRowScannerAddStep, 1)

        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        SetNextRowTableScanner (sCbgIiTblRowScanner)


    Loop

    Call WriteSetterOutput(iDoubleIndent, WriteCaseElse(sSubTitle), 2)

    Call WriteSetterOutput(iSingleIndent, sSelectCloser, 2)

    Call WriteSetterOutput(iSingleIndent, "SetTableScanner" + sLeftPar + sVariable + sRightPar, 1)

    WriteEndSub


End Sub

' READY TO TEST
Sub WriteResetTableCells()

    sSubTitle = "ResetTableCells"
    sArgName_1 = "sDummyRow"
    
    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle, sArgName_1, sTypeString), 2)

    InitializeAllTableScanners

    Call WriteSetterOutput(iSingleIndent, sSelectOpener + sArgName_1, 2)

    Do Until oCbgIiTblMainNameCell = sBlankCell

        MatchPairsUpTo (sTable)

        sPrefixInit = sStrPre + GetPrefixInitials(sColumn)
        sStatementToWrite = sCase + sPrefixInit + sRowScanner

        Call WriteSetterOutput(iDoubleIndent, sStatementToWrite, 2)

        Do Until oCbgIiClmnMainNameCell = sBlankCell

            sPrefixInit = GetPrefixInitials(sColumn)

            If oCbgIiClmnTblCell = oCbgIiTblMainNameCell Then

                sLeftHandSide = sObjPre + sPrefixInit + CStr(oCbgIiClmnCodeNameCell) + sCell + sDotValue
                sRightHandSide = "sBlankCell"

                Call WriteSetterOutput(iTripleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)

            End If

            SetNextRowTableScanner (sCbgIiClmnRowScanner)

        Loop

        SetNextRowTableScanner (sCbgSoRowScanner)

        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        InitializeRowScannerAndTable (sCbgIiClmnRowScanner)
        
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

    Call WriteSetterOutput(iDoubleIndent, WriteCaseElse(sSubTitle), 2)

    Call WriteSetterOutput(iSingleIndent, sSelectCloser, 2)

    WriteEndSub


End Sub

' READY TO TEST
Sub WriteTableSorter()

    sSubTitle = "SortTable"
    sArgName_1 = "sDummyWorksheet"
    sArgName_2 = "iDummyColumn"

    InitializeAllTableScanners

    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle, sArgName_1, sTypeString, sArgName_2, sTypeByte), 2)

    ' Initiate Selector Script:
    Call WriteSetterOutput(iSingleIndent, sSelectOpener + sArgName_1, 2)

    ' Propogate Selector Script:
    Do Until oCbgIiTblMainNameCell = sBlankCell

        MatchPairsUpTo (sTable)

        ' Write Code for Selected LHS and RHS And Make Congruent Statements
        ' Prefixes
        sTblPrefixInit = GetPrefixInitials(sTable)
        sElmtPrefixInit = GetPrefixInitials(sColumn)
        
        sStatementToWrite = sCase + sStrPre + sTblPrefixInit + CStr(oCbgIiTblCodeNameCell)
        Call WriteSetterOutput(iDoubleIndent, sStatementToWrite, 2)
        
        ' Selected Worksheet
        sLeftHandSide = sSet + sObjPre + sSelected + sWorksheet
        sRightHandSide = sObjPre + CStr(oCbgIiWbCodeNameCell) + sDotWorksheets + sLeftPar + sStrPre + sWsPrefixInit + CStr(oCbgIiWsCodeNameCell) + sRightPar

        Call WriteSetterOutput(iTripleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)
        
        ' Selected HeaderRow
        sLeftHandSide = sNumPre + sSelected + sHeaderRow
        sRightHandSide = sNumPre + sElmtPrefixInit + sHeaderRow

        Call WriteSetterOutput(iTripleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 2)

        ' Call WriteSetterOutput(iDoubleIndent, "sSelectedTable = " + sTblInit + CStr(oCbgIiTblCodeNameCell), 1)

        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

    ' Terminate Selector Script:
    Call WriteSetterOutput(iDoubleIndent, WriteCaseElse(sSubTitle), 2)
    Call WriteSetterOutput(iSingleIndent, sSelectCloser, 2)

    CreateScriptArray

    Do Until iScriptIndex > iFinalScriptIndex

        Call WriteSetterOutput(iArrIndents(iScriptIndex), sArrStatements(iScriptIndex), CInt(iArrSpaces(iScriptIndex)))
        iScriptIndex = iScriptIndex + 1

    Loop
    
    WriteEndSub


End Sub

' READY TO TEST
Sub WriteSetMasterArrays()

    sSubTitle = "SetMasterArrays"

    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle), 2)

    Call WriteSetterOutput(iSingleIndent, CommentOut(iSingleIndent, "MASTER ARRAYS"), 2)
    
    ' Workbooks

    Do Until iDummyIndex > iVarArrIndex

        SetMasterArray (sArrObjectTypes(iDummyIndex))
        iDummyIndex = iDummyIndex + 1

    Loop

    WriteEndSub

End Sub

' READY TO TEST
Sub WriteSetTableArrays()

    sSubTitle = "SetTableArrays"

    Call WriteSetterOutput(iNoIndent, WriteSub(sSubTitle), 2)

    ' String Assignments As Columns
    InitializeAllTableScanners
    iColumnCounter = iCounterInitialized

    Do Until oCbgIiTblMainNameCell = sBlankCell

        sStatementToWrite = "Index Reassignments For " + CStr(oCbgIiTblMainNameCell) + " Table Columns"
        Call WriteSetterOutput(iSingleIndent, CommentOut(iSingleIndent, sStatementToWrite, 2))
        
        MatchPairsUpTo (sTable)

        Do Until oCbgIiClmnMainNameCell = sBlankCell

            If oCbgIiClmnTblCell = oCbgIiTblMainNameCell Then

                sDelayedIndex = IIf(ColumnCounter = 0, "0", sLaggingIndex + sPlusOne)

                sPrefixInit = sNumPre + GetPrefixInitials(sColumn)
                sCurrentIndex = sPrefixInit + CStr(oCbgIiClmnCodeNameCell) + sIndex
                
                sLeftHandSide = sLocalCon + sPrefixInit + CStr(oCbgIiClmnCodeNameCell) + sIndex + sAs + sTypeByte
                sRightHandSide = sDelayedIndex

                Call WriteSetterOutput(iDoubleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)
                sLaggingIndex = sCurrentIndex

                iColumnCounter = iColumnCounter + 1

            End If

            SetNextRowTableScanner (sCbgIiClmnRowScanner)

        Loop

        SetNextRowTableScanner (sCbgSoRowScanner)

        iColumnCounter = iCounterInitialized
        InitializeRowScannerAndTable (sCbgIiWsRowScanner)
        InitializeRowScannerAndTable (sCbgIiClmnRowScanner)
        
        SetNextRowTableScanner (sCbgIiTblRowScanner)

    Loop

    ' DOUBLE CHECK THESE
    WriteArrayIndexAssigner (sColumn)
    WriteArrayIndexAssigner (sCell)
    WriteArrayIndexAssigner (sHeader)
    WriteArrayIndexAssigner (sHeaderTitle)

    WriteEndSub

End Sub
