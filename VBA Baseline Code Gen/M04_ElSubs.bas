Option Explicit

' READY FOR TESTING
Sub CreateScriptArray()

    iScriptIndex = iCounterInitialized

    ' Indents
        iArrIndents(0) = iSingleIndent
        iArrIndents(1) = iSingleIndent
        iArrIndents(2) = iSingleIndent
        iArrIndents(3) = iDoubleIndent
        iArrIndents(4) = iSingleIndent
        iArrIndents(5) = iDoubleIndent
        iArrIndents(6) = iDoubleIndent
        iArrIndents(7) = iSingleIndent
        iArrIndents(8) = iDoubleIndent
        iArrIndents(9) = iDoubleIndent
        iArrIndents(10) = iDoubleIndent
        iArrIndents(11) = iDoubleIndent
        iArrIndents(12) = iSingleIndent

    ' Statements
        sArrStatements(0) = "Set oDummyTable = oSelectedWorksheet.ListObjects(sSelectedTable)"
        sArrStatements(1) = "Set oSortKey = oSelectedWorksheet.Cells(iSelectedHeaderRow, iSelectedColumn)"
        sArrStatements(2) = "oDummyTable.Sort. _"
        sArrStatements(3) = "SortFields.Clear"
        sArrStatements(4) = "oDummyTable.Sort. _"
        sArrStatements(5) = "SortFields.Add2 Key:=oSortKey, SortOn:=xlSortOnValues, _"
        sArrStatements(6) = "Order:=xlAscending, DataOption:=xlSortNormal"
        sArrStatements(7) = "With oDummyTable.Sort"
        sArrStatements(8) = ".Header = xlYes"
        sArrStatements(9) = ".MatchCase = False"
        sArrStatements(10) = ".Orientation = xlTopToBottom"
        sArrStatements(11) = ".SortMethod = xlPinYin"
        sArrStatements(12) = "End With"

    ' Spaces
        iArrSpaces(0) = 1
        iArrSpaces(1) = 2
        iArrSpaces(2) = 1
        iArrSpaces(3) = 2
        iArrSpaces(4) = 1
        iArrSpaces(5) = 1
        iArrSpaces(6) = 2
        iArrSpaces(7) = 2
        iArrSpaces(8) = 1
        iArrSpaces(9) = 1
        iArrSpaces(10) = 1
        iArrSpaces(11) = 2
        iArrSpaces(12) = 1

End Sub

' READY FOR TESTING
Sub FindPairs(oDummyObjectOne As Object, oDummyObjectTwo As Object)

    sDummyRowScanner = GetRowScanner(oDummyObjectOne)

    Do Until oDummyObjectOne = oDummyObjectTwo

        NotFoundTest (oDummyObjectOne)
        SetNextRowTableScanner (sDummyRowScanner)

    Loop

End Sub

' READY FOR TESTING
Sub MatchPairsUpTo(sEndPoint As String)

    Select Case sEndPoint

        Case sWorksheet

            Call FindPairs(oCbgIiWbMainNameCell, oCbgIiWsWbCell)

        Case sTable

            Call FindPairs(oCbgIiWsMainNameCell, oCbgIiTblWsCell)
            Call FindPairs(oCbgIiWbMainNameCell, oCbgIiWsWbCell)

        Case sColumn, sCell, sHeader

            Call FindPairs(oCbgIiTblMainNameCell, oCbgIiClmnTblCell)
            Call FindPairs(oCbgIiWsMainNameCell, oCbgIiTblWsCell)
            Call FindPairs(oCbgIiWbMainNameCell, oCbgIiWsWbCell)

        Case Else: MsgBox (NotSupported(sEndPoint)): End


    End Select

End Sub

' READY FOR TESTING
Sub NotFoundTest(oInputObject As Variant)

    Select Case oInputObject

        Case oCbgIiWbMainNameCell
        
            Set oDummyObject = oCbgInputsInterface.Cells(iCbgIiWbRowScanner, iCbgIiWbMainNameColumn)
            sDummyObjectType = sWorkbook

        Case oCbgIiWsMainNameCell

            Set oDummyObject = oCbgInputsInterface.Cells(iCbgIiWsRowScanner, iCbgIiWsMainNameColumn)
            sDummyObjectType = sWorksheet

        Case oCbgIiTblMainNameCell

            Set oDummyObject = oCbgInputsInterface.Cells(iCbgIiTblRowScanner, iCbgIiTblMainNameColumn)
            sDummyObjectType = sTable

        Case oCbgIiClmnMainNameCell

            Set oDummyObject = oCbgInputsInterface.Cells(iCbgIiClmnRowScanner, iCbgIiClmnMainNameColumn)
            sDummyObjectType = sColumn

        Case oCbgIiVarNameCell

            Set oDummyObject = oCbgInputsInterface.Cells(iCbgIiVarRowScanner, iCbgIiVarNameColumn)
            sDummyObjectType = sVariable

        Case oCbgIiConstNameCell

            Set oDummyObject = oCbgInputsInterface.Cells(iCbgIiConstRowScanner, iCbgIiConstNameColumn)
            sDummyObjectType = sConstant

        Case Else: MsgBox ("Variable " + CStr(oInputObject) + " not found!"): End

    End Select

    If oDummyObject = sBlankCell Then MsgBox (NotSupported(sDummyObjectType)): End

End Sub

' READY FOR TESTING
Sub OptimizedRankSort()

   ' ------SORT-BY-NUMBER------       ------SORT-BY-PARENT------
    SortTable (iCbgIiClmnRankColumn): SortTable (iCbgIiClmnTblColumn)
    SortTable (iCbgIiTblRankColumn):  SortTable (iCbgIiTblWsColumn)
    SortTable (iCbgIiWsRankColumn):   SortTable (iCbgIiWsMainNameColumn)
    SortTable (iCbgIiWbRankColumn)

End Sub

' READY FOR TESTING
Sub PopulateConstantTableRow(sDummyConstName As String, sDummyConstType As String, sDummyConstValue As String)

    oCbgIiConstNameCell.Value = sDummyConstName
    oCbgIiConstTypeCell.Value = sDummyConstType
    oCbgIiConstValueCell.Value = IIf(oCbgIiConstTypeCell = sTypeString, SurroundInQuotes(sDummyConstValue), sDummyConstValue)

    SetNextRowTableScanner (sCbgIiConstRowScanner)

End Sub

' READY FOR TESTING
Sub PopulateVariableTableRow(sDummyVarName As String, sDummyVarType As String)

    oCbgIiVarNameCell.Value = sDummyVarName
    oCbgIiVarTypeCell.Value = sDummyVarType

    SetNextRowTableScanner (sCbgIiVarRowScanner)

End Sub

' READY FOR TESTING, See Additional Comments
Sub SelectCode()

    ' Could be optimized better using the same beginning and endpoint finding algorithms used in the eraser subs.
    Range("B1:B10000").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    ' Turn into string if possible
    MsgBox ("Code Selected and Copied.  Paste Back Into VBA.")

End Sub

' READY FOR TESTING
Sub SortTable(iDummyColumn As Byte)
    
    ' Select Dummy Table (Could be updated later to be selected by worksheet)
    Select Case iDummyColumn

        ' Could use an array for better management
        Case iCbgIiWbMainNameColumn To iCbgIiWbFileNameColumn

            sTableSelected = sCbgIiWorkbooks
            iDummyHeaderRow = iCbgIiWbHeaderRow

        Case iCbgIiWsWbColumn To iCbgIiWsTypeColumn

            sTableSelected = sCbgIiWorksheets
            iDummyHeaderRow = iCbgIiWsHeaderRow

        Case iCbgIiTblWsColumn To iCbgIiTblTypeColumn

            sTableSelected = sCbgIiTables
            iDummyHeaderRow = iCbgIiTblHeaderRow

        Case iCbgIiClmnRankColumn To iCbgIiClmnTypeColumn

            sTableSelected = sCbgIiColumns
            iDummyHeaderRow = iCbgIiClmnHeaderRow

        Case iCbgIiConstNameColumn To iCbgIiConstValueColumn

            sTableSelected = sCbgIiConstants
            iDummyHeaderRow = iCbgIiConstHeaderRow

        Case iCbgIiVarNameColumn To iCbgIiVarTypeColumn

            sTableSelected = sCbgIiVariables
            iDummyHeaderRow = iCbgIiVarHeaderRow

        Case Else: MsgBox (NotSupported(sColumn)): End

    End Select
 
    Set oSortKey = oCbgInputsInterface.Cells(iDummyHeaderRow, iDummyColumn)
    Set oDummyTable = oCbgInputsInterface.ListObjects(sTableSelected)

    oDummyTable.Sort. _
    SortFields.Clear
    
    oDummyTable.Sort. _
        SortFields.Add2 Key:=oSortKey, SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
        
    With oDummyTable.Sort

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With

End Sub

' READY FOR TESTING
Sub WriteDeclarationOutput(iIndentFactor As Byte, sDummyString As String, iSkipFactor As Integer)

    ' Possible to reduce Args through getter statements?
    oCbgDoCodeCell.Value = IndentString(iIndentFactor) + sDummyString
    Call SetNextRowTableScanner(sCbgDoRowScanner, iSkipFactor)

End Sub

' READY FOR TESTING
Sub WriteSetterOutput(iIndentFactor As Byte, sDummyString As String, iSkipFactor As Integer)

    ' Possible to reduce Args through getter statements?
    oCbgSoCodeCell.Value = IndentString(iIndentFactor) + sDummyString
    Call SetNextRowTableScanner(sCbgSoRowScanner, iSkipFactor)

End Sub

' READY FOR TESTING
Sub WriteEndSub()

    SetNextRowTableScanner (sCbgSoRowScanner)
    Call WriteSetterOutput(iNoIndent, sEndSub, 3)

End Sub

' DO NOT USE
Sub WriteArrayIndexAssigner(sTypeObject As String)

    ' Header Titles
    InitializeAllTableScanners
    iColumnCounter = iCounterInitialized

    sSelectedSuffix = MakePlural(sTypeObject)
    sChosenDataType = DetermineDataType(sTypeObject)

    Do Until oCbgIiTblMainNameCell = sBlankCell

        sStatementToWrite = sArrTitlePre + CStr(oCbgIiTblMainNameCell) + sBlankSpace + MakePlural(sHeader)
        Call WriteSetterOutput(iSingleIndent, CommentOut(iSingleIndent, sStatementToWrite), 2)
        
        MatchPairsUpTo (sTable)

        sTblInit = GetPrefixInitials(sColumn)
        sDataInit = GetPrefixInitials(sTypeObject)

        sPrefixInit = CStr(sArrInit + oCbgIiWbInitCell + oCbgIiWsInitCell + oCbgIiTblInitCell)

        Call WriteSetterOutput(iDoubleIndent, sLocalDec + sPrefixInit + sTable + sSelectedSuffix + sLeftPar + sPrefixInit + sTableLength + sRightPar + sAs + sChosenDataType, 2)

        Do Until oCbgIiClmnMainNameCell = sBlankCell

            If oCbgIiClmnTblCell = oCbgIiTblMainNameCell Then

                sLeftHandSide = sPrefixInit + sTable + sSelectedSuffix + sLeftPar + sPrefixInit + CStr(oCbgIiClmnCodeNameCell) + "Index" + sRightPar
                sRightHandSide = sPrefixInit + CStr(oCbgIiClmnCodeNameCell) + sHeader

                Call WriteSetterOutput(iDoubleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)
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

End Sub

' DO NOT USE
Sub SetMasterArray(sTypeObject As String)

    InitializeAllTableScanners

    ' Use Getters To Achieve The Following:
    iArrFinalIndex = GetMaxArrayIndex(sTypeObject)
    sTypeObjectInit = GetObjectInit(sTypeObject)
    sCodeName = GetCodename(sTypeObject)

    If iArrFinalIndex > iCounterInitialized Then 'CALL SOME OTHER SUB

        iArrIndexCounter = iCounterInitialized
        sStatementToWrite = sTypeObject + sBlankSpace + MakePlural(sIndex)

        Call WriteSetterOutput(iSingleIndent, CommentOut(iSingleIndent, sStatementToWrite), 1)
        
        Select Case sTypeObject

            Case sWorkbook, sWorksheet

                Do Until iArrIndexCounter > iArrFinalIndex

                    sLeftHandSide = sTypeObjectInit + sCodeName + sIndex
                    sRightHandSide = CStr(iArrIndexCounter)

                    Call WriteSetterOutput(iDoubleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)
                    
                    SetNextRowTableScanner (sCbgIiWbRowScanner)
                    iArrIndexCounter = iArrIndexCounter + 1

                Loop

                SetNextRowTableScanner (sCbgSoRowScanner)

                InitializeAllTableScanners
                iArrIndexCounter = iCounterInitialized

                sStatementToWrite = sTypeObject + sBlankSpace + sArrays
                Call WriteSetterOutput(iSingleIndent, CommentOut(iSingleIndent, sStatementToWrite), 1)

                Do Until iArrIndexCounter > iArrWbIndexCounter

                    sLeftHandSide = sNumPre + sArr + sTypeObject + CStr(iArrIndexCounter) + sRightPar
                    sRightHandSide = sTypeObjectInit + sCodeName

                    Call WriteSetterOutput(iDoubleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)
                    
                    SetNextRowTableScanner (sCbgIiWbRowScanner)
                    iArrIndexCounter = iArrIndexCounter + 1

                Loop

                SetNextRowTableScanner (sCbgSoRowScanner)

            Case sTable, sNumRowScanner, sStrRowScanner, sHeaderRow, sInitialRow

                Do Until iArrIndexCounter > iArrFinalIndex

                    MatchPairsUpTo (sTable)
                    sPrefixInit = GetPrefixInitials(sTable)

                    sLeftHandSide = sDeclareConst + sPrefixInit + sTblInit + sAs + GetDataType(sTypeObject)
                    sRightHandSide = CStr(iArrIndexCounter)

                    Call WriteSetterOutput(iDoubleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)
                    
                    InitializeRowScannerAndTable (sCbgIiWsRowScanner)

                    SetNextRowTableScanner (sCbgIiTblRowScanner)
                    iArrIndexCounter = iArrIndexCounter + 1

                Loop

                SetNextRowTableScanner (sCbgSoRowScanner)
                
                InitializeAllTableScanners
                iArrIndexCounter = iCounterInitialized

                sStatementToWrite = sTypeObject + sBlankSpace + sArrays
                Call WriteSetterOutput(iSingleIndent, CommentOut(iSingleIndent, sStatementToWrite), 1)

                Do Until iArrIndexCounter > iArrFinalIndex

                    MatchPairsUpTo (sTable)
                    sPrefixInit = GetPrefixInitials(sTable)
                    
                    sLeftHandSide = sFirstInit + sArr + sTypeObject + sLeftPar + sPrefixInit + sTblInit + sIndex
                    sRightHandSide = IIf(sTypeObject = sTable, GetPrefixInitials(sWorksheet) + sTblInit + CStr(oCbgIiTblCodeNameCell), sPrefixInit + sTypeObject)

                    Call WriteSetterOutput(iDoubleIndent, LeftEqualsRight(sLeftHandSide, sRightHandSide), 1)

                    InitializeRowScannerAndTable (sCbgIiWsRowScanner)

                    SetNextRowTableScanner (sCbgIiTblRowScanner)
                    iArrIndexCounter = iArrIndexCounter + 1

                Loop

                SetNextRowTableScanner (sCbgSoRowScanner)

            Case Else: MsgBox (VariableNotFound): End

        End Select

    End If

End Sub