Option Explicit

' READY TO TEST
Sub ProduceDeclarations()
      
    SetWorkbookAndWorksheets
    
    InitializeRowScanners (sCbgInputsInterface)
    InitializeRowScanners (sCbgDeclarationsOutput)

    SetTableScanner (sCbgDoRowScanner)
    EraseCurrentDeclarations

    Call WriteDeclarationOutput(iNoIndent, sOptExp, 2)
    
    WriteWorkbookDeclarations
    WriteWorksheetDeclarations
    WriteTableDeclarations

    Call WriteTableElementDeclarations(sColumn, sTypeByte)
    Call WriteTableElementDeclarations(sCell, sTypeObject)
    Call WriteTableElementDeclarations(sHeader, sTypeObject)

    WriteRowScannerDeclarations
    WriteConstantDeclarations
    WriteVariableDeclarations
    WriteArrayDeclarations
    
    ' Test Code
    ' MsgBox ("Check Outputs")

End Sub

' READY TO TEST
Sub ProduceSetters()

    SetWorkbookAndWorksheets
    InitializeRowScanners (sCbgInputsInterface)
    InitializeRowScanners (sCbgSettersOutput)

    SetTableScanner (sCbgSoRowScanner)
    
    EraseCurrentSetters

    Call WriteSetterOutput(iNoIndent, sOptExp, 2)

    WriteWbWsSetters
    WriteSetTablesAndHeaders
    WriteSetTableScanner
    WriteIndividualRSInit
    WriteInitAllRS
    WriteInitializeAllTables
    WriteGoToNextRow
    WriteTableSorter
    WriteResetTableCells
    ' WriteSetMasterArrays
    ' WriteSetTableArrays
    
    ' MsgBox ("Check Outputs")

End Sub

' DO NOT TEST
Sub PopulateAllTables()

    SetWorkbookAndWorksheets
    InitializeRowScanners (True)
    
    ' iArrLoopIndex = 0
    ' Do Until iArrLoopIndex = WorksheetFunction.Max(sArrAllTables)
        
        ' SetTableScanner(sArrRowScanners(iArrLoopIndex))
        
    ' Loop
    
    SetTableScanner (sCbgIiWbRowScanner)
    SetTableScanner (sCbgIiWsRowScanner)
    SetTableScanner (sCbgIiTblRowScanner)
    SetTableScanner (sCbgIiClmnRowScanner)
    SetTableScanner (sCbgIiConstRowScanner)
    SetTableScanner (sCbgIiVarRowScanner)
    
    FillWsTable
    FillTblTable
    FillClmnTable
    CreateDefaultConsts (True)
    CreateDefaultVars (True)
    

End Sub

' READY TO TEST
Sub ClearAllTables()

    SetWorkbookAndWorksheets
    InitializeRowScanners (sCbgInputsInterface)
    
    ' iArrLoopIndex = 0
    ' Do Until iArrLoopIndex = UBound(sArrRowScanners)
        
        ' SetTableScanner(sArrRowScanners(iArrLoopIndex))
        
    ' Loop
    
    SetTableScanner (sCbgIiWbRowScanner)
    SetTableScanner (sCbgIiWsRowScanner)
    SetTableScanner (sCbgIiTblRowScanner)
    SetTableScanner (sCbgIiClmnRowScanner)
    SetTableScanner (sCbgIiConstRowScanner)
    SetTableScanner (sCbgIiVarRowScanner)
    
    ' Can Subs be built into an array?
    ClearWbTable
    ClearWsTable
    ClearTblTable
    ClearClmnTable
    ClearConstTable
    ClearVarTable
    
End Sub
