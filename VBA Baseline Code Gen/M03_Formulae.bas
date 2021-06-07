Option Explicit

' I - GENERAL FUNCTIONS
' These functions are general purpose across the board.

' READY FOR TESTING
Function CommentOut(iIndentInput As Byte, sDummyString As String) As String

    ' Formats outputs to have the right amount of Apost, which is ultimately just 1...
    sStringPrefix = IIf(iIndentInput = 0, Chr(iApostMark) + Chr(iApostMark), Chr(iApostMark))

    CommentOut = sStringPrefix + sDummyString

End Function

' READY FOR TESTING
Function ExtendStatement(sInitialStatement As String, sExtendedStatement As String) As String

    ExtendStatement = sInitialStatement + sColon + sFourSpaces + sExtendedStatement

End Function

' READY FOR TESTING
Function IndentString(iIndentInput As Byte) As String

    iIndentCounter = iCounterInitialized
    
    Do Until iIndentCounter = iIndentInput
    
        IndentString = IndentString + sFourSpaces
        iIndentCounter = iIndentCounter + 1

    Loop

End Function

' READY FOR TESTING
Function LeftEqualsRight(sDummyLHS As String, sDummyRHS As String) As String

    LeftEqualsRight = sDummyLHS + sEquals + sDummyRHS

End Function

' READY FOR TESTING
Function MakePlural(sDummyInput As String) As String

    MakePlural = sDummyInput + "s"

End Function

' READY FOR TESTING
Function NotSupported(sDummyInput As String) As String

    NotSupported = sNotFoundPre + sDummyInput + sNotFoundPost

End Function

' READY FOR TESTING
Function SurroundInQuotes(vDummyInput As Variant) As String

    ' Converter
    sDummyInput = CStr(vDummyInput)

    ' Ternary Statement checks if quotation mark already exists:
    SurroundInQuotes = IIf(Left(sDummyInput, 1) = Chr(iQuoteMark), sDummyInput, Chr(iQuoteMark) + sDummyInput + Chr(iQuoteMark))

End Function

' *********************************GETTER**FUNCTIONS*******************************
' READY FOR TESTING

Function GetArrayLength(vArrDummy As Variant) As Byte

    GetArrayLength = UBound(vArrayDummy) - LBound(vArrayDummy) + 1

End Function


Function GetCodename(sObjectSuffix As String) As String

    Select Case sObjectSuffix

        Case sWorkbook

            GetCodename = CStr(oCbgIiWbCodeNameCell)

        Case sWorksheet

            GetCodename = CStr(oCbgIiWsCodeNameCell)

        Case sTable

            GetCodename = CStr(oCbgIiTblCodeNameCell)

        Case sColumn, sCell, sHeader

            GetCodename = CStr(oCbgIiClmnCodeNameCell)

        Case Else: MsgBox (NotSupported(sObjectSuffix)): End

    End Select

End Function

' READY FOR TESTING
Function GetCodenameInit(oDummyCell As Object) As String

    sDummyInput = CStr(oDummyInput)
    ' Splits the Main Name Like Such:
    '                                                       I_0       I_1
    ' sArrMainName = VBA.Split("Example Phrase", " ") = ("Example", "Phrase")
    sArrMainName = VBA.Split(sDummyInput, sBlankSpace)

    If IsArray(sArrMainName) Then

        For iLetterIndex = LBound(sArrMainName) To UBound(sArrMainName)

            If iLetterIndex = LBound(sArrMainName) Then

                ' Makes First Letter Uppercase
                GetCodenameInit = UCase(Left(sArrMainName(iLetterIndex), 1))

            Else

                ' Makes Following Letters Lowercase
                GetCodenameInit = GetCodenameInit + LCase(Left(sArrMainName(iLetterIndex), 1))

            End If

        Next iLetterIndex

    Else

        ' Returns the first two letters of a singular word:
        GetCodenameInit = Left(sDummyString, 2)

    End If

End Function

' ABANDONED
' Function GetEndMsg(sDataType As String) As Variant/Object
' End Sub

' READY FOR TESTING
Function GetDataType(sTypeObject As String) As String

    Select Case sTypeObject

        Case sWorkbook

            GetDataType = sTypeWorkbook

        Case sWorksheet

            GetDataType = sTypeWorksheet

        Case sTable, sObjHeader, sCell
            
            GetDataType = sTypeObject
        
        Case sColumn, sNumRowScanner
        
            GetDataType = sTypeByte

        Case sStrHeader, sStrRowScanner

            GetDataType = sTypeString
        
        Case Else: MsgBox (NotSupported(sTypeObject)): End

    End Select

End Function

' READY FOR TESTING
Function GetDataLetter(sDataType As String) As String

    Select Case sDataType

        Case sTypeObject, sTypeWorkbook, sTypeWorksheet, sTypeTable
            
            GetDataLetter = sObjPre

        Case sTypeByte, sTypeInteger
        
            GetDataLetter = sNumPre

        Case sTypeBoolean
        
            GetDataLetter = sBooPre

        Case sTypeString
            
            GetDataLetter = sStrPre

        Case sTypeVariant

            GetDataLetter = sVarPre

        Case Else: MsgBox (NotSupported(sDataType)): End

    End Select

End Function

' READY FOR TESTING, BUT REVISIT
Function GetMaxArrayIndex(sTypeObject As String) As Byte

    Select Case sTypeObject

        Case sWorkbook

            GetMaxArrayIndex = iArrWbIndexCounter

        Case sWorksheet

            GetMaxArrayIndex = iArrWsIndexCounter

        Case sTable, sColumn, sNumRowScanner, sStrRowScanner, sInitialRow, sHeaderRow, sHeader

            GetMaxArrayIndex = iArrTblIndexCounter

        Case Else: MsgBox (NotSupported(sTypeObject)): End

    End Select

End Function

' READY FOR TESTING
Function GetObjectInit(sTypeObject As String) As String

    Select Case sTypeObject

        Case sWorkbook
            
            GetObjectInit = sWbInit

        Case sWorksheet
        
            GetObjectInit = sWsInit

        Case sTable
        
            GetObjectInit = sTblInit
        
        Case Else: MsgBox (NotSupported(sTypeObject)): End

    End Select

End Function


' READY FOR TESTING
Function GetPrefixInitials(sEndPoint As String) As String

    Select Case sEndPoint

        Case sWorkbook
            
            ' Does not get Object Initials
            GetPrefixInitials = sBlankCell
        
        Case sWorksheet
        
            GetPrefixInitials = CStr(oCbgIiWbInitCell)
        
        Case sTable, sLocalConstant, sLocalVariable
        
            GetPrefixInitials = CStr(oCbgIiWbInitCell + oCbgIiWsInitCell)
        
        Case sColumn, sCell, sHeader, sRowScanner, sInitialRow, sHeaderRow

            GetPrefixInitials = CStr(oCbgIiWbInitCell + oCbgIiWsInitCell + oCbgIiTblInitCell)

        Case sGlobalConstant
        
            GetPrefixInitials = sGblC
        
        Case sGlobalVariable

            GetPrefixInitials = sGblV

        Case Else: MsgBox (NotSupported(sEndPoint)): End


    End Select

End Function

' READY FOR TESTING
Function GetTitle(sTypeObject As String) As String

    Select Case sTypeObject

        Case sColumn
        
            GetTitle = sHeaderColumns

        Case sCell
        
            GetTitle = sHeaderCells

        Case sHeader
        
            GetTitle = sHeaderHeaders

        Case Else: MsgBox (NotSupported(sTypeObject)): End

    End Select

End Function

' READY FOR TESTING
Function GetRowScanner(oDummyObject As Object) As String

    ' May Need Some Conversions To Ensure Data sType Matches

    Select Case oDummyObject

        ' Eventually, I am planning to rewrite this using an Array such as: Case LBound(oArrWorkbookCells) to UBound(oArrWorkbookCells)

        Case oCbgIiWbMainNameCell
        
            GetRowScanner = sCbgIiWbRowScanner

        Case oCbgIiWsMainNameCell
        
            GetRowScanner = sCbgIiWsRowScanner

        Case oCbgIiTblMainNameCell
        
            GetRowScanner = sCbgIiTblRowScanner

        Case oCbgIiClmnMainNameCell
        
            GetRowScanner = sCbgIiClmnRowScanner

        Case oCbgIiConstNameCell
        
            GetRowScanner = sCbgIiConstRowScanner

        Case oCbgIiVarNameCell
        
            GetRowScanner = sCbgIiVarRowScanner

        Case Else: MsgBox (NotSupported(sRowScanner)): End

    End Select

End Function

' READY FOR TESTING
Function GetMainNameObject(sTypeObject As String) As Object

    Select Case sTypeObject

        Case sWorkbook

            Set GetMainNameObject = oCbgInputsInterface.Cells(iCbgIiWbRowScanner, iCbgIiWbMainNameColumn)

        Case sWorksheet

            Set GetMainNameObject = oCbgInputsInterface.Cells(iCbgIiWsRowScanner, iCbgIiWsMainNameColumn)

        Case sTable

            Set GetMainNameObject = oCbgInputsInterface.Cells(iCbgIiTblRowScanner, iCbgIiTblMainNameColumn)

        Case sColumn, sHeader, sCell

            Set GetMainNameObject = oCbgInputsInterface.Cells(iCbgIiTblRowScanner, iCbgIiTblMainNameColumn)

        Case Else: MsgBox (NotSupported(sTypeObject)): End

    End Select

End Function

' READY FOR TESTING
Function GetSuffix(sTypeObject As String) As String

    Select Case sTypeObject

        Case sWorkbook, sWorksheet, sTable, sVariable, sConstant

            ' These types do not have a suffix in their naming schemes.
            GetSuffix = sBlankCell

        Case sColumn

            GetSuffix = sColumn
        
        Case sHeader

            GetSuffix = sHeader

        Case sCell

            GetSuffix = sCell
            
        Case sHeaderRow
        
            GetSuffix = sHeaderRow
            
        Case sInitialRow
        
            GetSuffix = sInitialRow
            
        Case sRowScanner
        
            GetSuffix = sRowScanner

        Case Else: MsgBox (NotSupported(sTypeObject)): End

    End Select

End Function

' READY FOR TESTING
Function GetDeclarationType(sTypeObject As String) As String

    Select Case sTypeObject

        Case sWorkbook

            ' This should be changed when multiple workbooks are supported with selectors.
            GetDeclarationType = sConstant

        Case sWorksheet

            GetDeclarationType = CStr(oCbgIiWsTypeCell)

        Case sTable

            GetDeclarationType = CStr(oCbgIiTblTypeCell)

        Case sColumn, sCell, sHeader

            GetDeclarationType = CStr(oCbgIiClmnTypeCell)

        Case sConstant

            GetDeclarationType = CStr(oCbgIiConstTypeCell)
        
        Case sVariable

            GetDeclarationType = CStr(oCbgIiVarTypeCell)

        Case Else: MsgBox (NotSupported(sTypeObject)): End

    End Select

End Function

' **********************************II**DECLARATION***FUNCTIONS************************************
' COMPLETED

'READY FOR TESTING, ERRORS MAY OCCUR.
Function DeclareConstant(oConstantName As Object, vDataType As Variant, vInputValue As Variant, Optional sTypeObject As String) As String

    ' Converters
    sConstantName = CStr(oConstantName)
    sDataType = CStr(vDataType)
    sInputValue = CStr(vInputValue)

    ' Getters
    sDataInit = GetDataLetter(sDataType)

    ' Determines Left-Hand Side's Full Constant Name
    If IsMissing(sTypeObject) Then

        sConstantFullName = sGblC + sConstantName
    
    Else

        ' Additional Getters
        sPrefixInit = GetPrefixInitials(sTypeObject)
        sSuffix = GetSuffix(sTypeObject)

        sConstantFullName = sPrefixInit + sConstantName + sSuffix

    End If
    
    sLeftHandSide = sDeclareConst + sDataInit + sConstantFullName + sAs + sDataType
    sRightHandSide = IIf(sDataType = sTypeString, SurroundInQuotes(sInputValue), sInputValue)

    '            ---------------LEFTHANDSIDE---------------    RIGHTHANDSIDE
    ' Statement: Public Const dtConstantFullName As DataType = InputValue
    DeclareConstant = LeftEqualsRight(sLeftHandSide, sRightHandSide)

End Function

' READY FOR TESTING
Function DeclareVariable(oVariableName As Object, vDataType As Variant, Optional sTypeObject As String) As String

    ' Converters
    sDataType = CStr(vDataType)
    sVariableName = CStr(oVariableName)

    ' Getters
    sDataInit = GetDataLetter(sDataType)

    If IsMissing(sTypeObject) Then

        ' Statement: Public vGblVVariableName As sDataType
        DeclareVariable = sDeclareVar + sDataInit + sGblV + sVariableName + sAs + sDataType
    
    Else
    
        ' Additional Getters with Object sType
        sPrefixInit = GetPrefixInitials(sTypeObject)
        sSuffix = GetSuffix(sTypeObject)

        ' Statement: Public vPrefInitVariableNameSuffix As sDataType
        DeclareVariable = sDeclareVar + sDataInit + sPrefixInit + sVariableName + sSuffix + sAs + sDataType

    End If


End Function

' READY FOR TESTING, BEWARE OF LAYERS, REVIEW 2nd ARG IN DECLAREVARIABLE()
Function DeclareWorkbook(oWbCodename As Object) As String
 
    ' Converters:
    sDeclType = GetDeclarationType(sWorkbook)
    Set oMainNameCell = GetMainNameObject(sWorkbook)

    ' Statement: Public oWorkbookCodename As Workbook
    DeclareWorkbook = DeclareVariable(oWbCodename, sTypeObject, sWorkbook)

    '            --------------ORIGINAL--------------  ---------------------------EXTENSION---------------------------
    ' Statement: Public oWorkbookCodename As Workbook: Public Const sWorkbookCodename As String = "File Name From Cell"
    If sDeclType = sConstant Then DeclareWorkbook = ExtendStatement(DeclareWorkbook, DeclareConstant(oWbCodename, sTypeString, oMainNameCell, sWorkbook))

End Function

' READY FOR TESTING, BEWARE OF LAYERS
Function DeclareWorksheet(oWsCodename As Object) As String

    ' Converters:
    sDeclType = GetDeclarationType(sWorksheet)
    Set oMainNameCell = GetMainNameObject(sWorksheet)

    ' Statement: Public oWsWorksheetCodename As Worksheet
    DeclareWorksheet = DeclareVariable(oWsCodename, sTypeObject, sWorksheet)

    '            ---------------ORIGINAL---------------  ---------------------------EXTENSION---------------------------
    ' Statement: Public oWsWorksheetCodename As Worksheet: Public Const sWsWorksheetCodename As String = "Main Name"
    If sDeclType = sConstant Then DeclareWorksheet = ExtendStatement(DeclareWorksheet, DeclareConstant(oWsCodename, sTypeString, oMainNameCell, sWorksheet))

End Function

' READY FOR TESTING, BEWARE OF LAYERS
Function DeclareTable(oTblCodename As Object) As String

    ' Converters:
    sDeclType = GetDeclarationType(sTable)
    Set oMainNameCell = GetMainNameObject(sTable)

    ' Statement: Public oWbWsTableCodename As Object          NO TYPE!
    DeclareTable = DeclareVariable(oTblCodename, sTypeObject, sTypeTable)

    '            -------------ORIGINAL-------------  ---------------------------EXTENSION---------------------------
    ' Statement: Public oWbWsTableCodename As Object: Public Const sWbWsTableCodename As String = "Main Name Table"
    If sDeclType = sConstant Then DeclareTable = ExtendStatement(DeclareTable, DeclareConstant(oTblCodename, sTypeString, oMainNameCell, sTable))

End Function

'' ABANDONED
''Function DeclareColumn()
'
'    sLeftHandSide =
'    sRightHandSide =
'
'    DeclareWorkbook = sLeftHandSide + sRightHandSide
'
'     If sDummyVarType = sConstant Then
'
'        sLeftHandSide =
'        sRightHandSide =
'
'        DeclareWorkbook = DeclareWorkbook + sColon + sLeftHandSide + sEquals + sRightHandSide
'
'    End If
'
''End Function
'
'' ABANDONED
''Function DeclareCell()
'
'    sLeftHandSide =
'    sRightHandSide =
'
'    DeclareWorkbook = sLeftHandSide + sRightHandSide
'
'     If sDummyVarType = sConstant Then
'
'        sLeftHandSide =
'        sRightHandSide =
'
'        DeclareWorkbook = DeclareWorkbook + sColon + sLeftHandSide + sEquals + sRightHandSide
'
'    End If
'
''End Function
'
'' ABANDONED
''Function DeclareHeader()
'
'    sLeftHandSide =
'    sRightHandSide =
'
'    DeclareWorkbook = sLeftHandSide + sRightHandSide
'
'     If sDummyVarType = sConstant Then
'
'        sLeftHandSide =
'        sRightHandSide =
'
'        DeclareWorkbook = DeclareWorkbook + sColon + sLeftHandSide + sEquals + sRightHandSide
'
'    End If

'End Function

'*************************************III**WRITER***FUNCTIONS*********************************

' READY FOR TESTING
Function WriteCaseRowScanner() As String

    ' Getter:
    sRowScanInit = GetPrefixInitials(sRowScanner)

    ' Set Up:
    sPreludeClause = sCase + sStrPre + sRowScanInit + sRowScanner

    ' Statement: Case sWbWsTblRowScanner: iWbWsTblRowScanner = iWbWsTblInitialRow
    WriteCaseRowScanner = ExtendStatement(sPreludeClause, WriteInitializeRowScanner) ' See called function for additional comments

End Function

' READY FOR TESTING
Function WriteInitializeRowScanner() As String

    ' Getter
    sRowScanInit = GetPrefixInitials(sRowScanner)
    sFullPre = sNumPre + sRowScanInit

    sLeftHandSide = sFullPre + sRowScanner
    sRightHandSide = sFullPre + sInitialRow

    '            ---LEFTHANDSIDE---   --RIGHTHANDSIDE--
    ' Statement: iWbWsTblRowScanner = iWbWsTblInitialRow
    WriteInitializeRowScanner = LeftEqualsRight(sLeftHandSide, sRightHandSide)

End Function

' READY FOR TESTING
Function WriteSetCell(oClmnCodename As Object, sObjectSuffix As String) As String

    sRowSuffix = IIf(sObjectSuffix = sCell, sRowScanner, sHeaderRow)

    ' Converters
    sClmnCodeName = CStr(oClmnCodename)

    ' Getters
    sWsPrefixInit = GetPrefixInitials(sWorksheet)
    sWsCodename = GetCodename(sWorksheet)
    sClmnPrefixInit = GetPrefixInitials(sColumn)
    

    sLeftHandSide = sSet + sObjPre + sClmnPrefixInit + sClmnCodeName + sObjectSuffix

    ' Right Hand Side Breakdown
    sTypeObjectPrelude = sObjPre + sWsPrefixInit + sWsCodename + sDotCells    'oWbWorksheetCodename.Cells (1)
    sRowStatement = sNumPre + sClmnPrefixInit + sRowSuffix                    'iWbWsHeaderRow (2)
    sColumnStatement = sNumPre + sClmnPrefixInit + sClmnCodeName + sColumn    'iWbWsTblColumnCodenameColumn (3)

    '                     (1)                        (2)                        (3)
    sRightHandSide = sTypeObjectPrelude + sLeftPar + sRowStatement + sComma + sColumnStatement + sRightPar

    '            ---------LEFTHANDSIDE---------   -----------------------------RIGHTHANDSIDE-----------------------------
    ' Statement: Set oWbWsTblColumnCodenameCell = oWbWorksheetCodename.Cells(iWbWsHeaderRow, iWbWsTblColumnCodenameColumn)
    WriteSetCell = LeftEqualsRight(sLeftHandSide, sRightHandSide)

End Function

' READY FOR TESTING
Function WriteSetHeader(oClmnCodename As Object, oDeclType As Object) As String

    ' Converters
    sDeclType = CStr(oDeclType)

    ' Getters
    sWsPrefixInit = GetPrefixInitials(sWorksheet)
    sWsCodename = GetCodename(sWorksheet)
    sTblPrefixInit = GetPrefixInitials(sTable)
    'sTblCodename = GetCodename(sTable)
    sClmnPrefixInit = GetPrefixInitials(sColumn)
    sClmnCodeName = CStr(oClmnCodename)
    
    sPartCodename = sTblPrefixInit + sClmnCodeName
    sFullCodename = sClmnPrefixInit + sClmnCodeName

    ' (Row, Column) Statements
    sRowStatement = sLeftPar + sNumPre + sClmnPrefixInit + sHeaderRow
    sColumnStatement = sNumPre + sFullCodename + sColumn + sRightPar

    '                  ------LEFTHANDSIDE--------   -----------------------RIGHTHANDSIDE-----------------------
    ' Statement Sides: Set oWbWsTblCodenameHeader = oTableCodename.Cells(iWbWsHeaderRow, iWbWsTblCodenameColumn)
    sLeftHandSide = sSet + sObjPre + sFullCodename + sHeader
    sRightHandSide = sObjPre + sWsPrefixInit + sWsCodename + sDotCells + sRowStatement + sComma + sColumnStatement

    WriteSetHeader = LeftEqualsRight(sLeftHandSide, sRightHandSide)

    If sDeclType = sConstant Then

        '                   --------LEFTHANDSIDE-------   ----RIGHTHANDSIDE-----
        ' Statements Sides: oWbWsTblCodenameHeader.Value = sWbWsTblCodenameHeader
        sLeftHandSide = sObjPre + sClmnPrefixInit + sClmnCodeName + sHeader + sDotValue
        sRightHandSide = sStrPre + sClmnPrefixInit + sClmnCodeName + sHeader
        sExtendedStatement = LeftEqualsRight(sLeftHandSide, sRightHandSide)

        '                 --------------------------------------WRITESETHEADER------------------------------------   --------LEFTHANDSIDE-------   ----RIGHTHANDSIDE----
        'Final Statement: Set oWbWsTblCodenameHeader = oTableCodename.Cells(iWbWsHeaderRow, iWbWsTblCodenameColumn): oWbWsTblCodenameHeader.Value = sWbWsTblCodenameHeader
        WriteSetHeader = ExtendStatement(WriteSetHeader, sExtendedStatement)

    End If

End Function

' READY FOR TESTING
Function WriteSetTable(oTblCodename As Object) As String
    
    ' Getters/Converters
    sWsPrefixInit = GetPrefixInitials(sWorksheet)
    sWsCodename = GetCodename(sWorksheet)
    sTblPrefixInit = GetPrefixInitials(sTable)
    sTblCodename = CStr(oTblCodename)

    sLeftHandSide = sSet + sObjPre + sTblPrefixInit + sTblCodename
    sRightHandSide = sObjPre + sWsPrefixInit + sWsCodename + sDotListObjects + sLeftPar + sStrPre + sTblPrefixInit + sTblCodename + sRightPar

    '                  -----LEFTHANDSIDE-----   -------------------RIGHTHANDSIDE-------------------
    ' Statement Sides: Set oWbWsTableCodename = oWbWorksheetCodename.ListObjects(sWbWsTableCodename)
    WriteSetTable = LeftEqualsRight(sLeftHandSide, sRightHandSide)

End Function

' READY FOR TESTING
Function WriteSetWorksheet(oWsCodename As Object) As String

    ' Getters
    sWbCodename = GetCodename(sWorkbook)
    sWsInit = GetPrefixInitials(sWorksheet)
    sWsCodename = CStr(oWsCodename)

    sLeftHandSide = sSet + sObjPre + sWsInit + sWsCodename
    sRightHandSide = sObjPre + sWbCodename + sDotWorksheets + sLeftPar + sStrPre + sWsInit + sWsCodename + sRightPar

    '            ------LEFTHANDSIDE------   ------------------RIGHTHANDSIDE-------------------
    ' Statement: Set oWbWorksheetCodename = oWorkbookCodename.Worksheets(sWbWorksheetCodename)
    WriteSetWorksheet = LeftEqualsRight(sLeftHandSide, sRightHandSide)

End Function

' READY FOR TESTING
Function WriteSetWorkbook(oWbCodename As Object)

    sWbCodename = CStr(oWbCodename)

    sLeftHandSide = sSet + sObjPre + sWbCodename
    sRightHandSide = sThisWorkbook  ' Will need to update later when working with multiple workbooks to use the file name.

    ' ----LEFTHANDSIDE-----   RIGHTHANDSIDE
    ' Set oWorkbookCodename = ThisWorkbook
    WriteSetWorkbook = LeftEqualsRight(sLeftHandSide, sRightHandSide)

End Function

' READY FOR TESTING
Function WriteRowScannerAddStep() As String
    
    sPrefixInitials = GetPrefixInitials(sRowScanner)
    sCodenameRowScanner = sPrefixInitials + sRowScanner

    sPreludeClause = sCase + sStrPre + sCodenameRowScanner
    sLeftHandSide = sNumPre + sCodenameRowScanner
    sRightHandSide = sLeftHandSide + sPlus + sNumPre + sStepValue

    sEqualityStatement = LeftEqualsRight(sLeftHandSide, sRightHandSide)
    
    '            ----PRELUDE CLAUSE-----  ---LEFTHANDSIDE---   ---------RIGHTHANDSIDE---------
    ' Statement: Case sWbWsTblRowScanner: iWbWsTblRowScanner = iWbWsTblRowScanner + iStepValue
    WriteRowScannerAddStep = ExtendStatement(sPreludeClause, sEqualityStatement)

End Function

' READY FOR TESTING
Function WriteSub(sDummyString As String, Optional sDummyArg1 As String, Optional sDummyArgType1 As String, Optional sDummyArg2 As String, Optional sDummyArgType2 As String, Optional sDummyArg3 As String, Optional sDummyArgType3 As String, Optional sDummyArg4 As String, Optional sDummyArgType4 As String, Optional sDummyArg5 As String, Optional sDummyArgType5 As String)

    Dim sSubArgs As String

    Const sArrDummyArg(4) As String = Array(sDummyArg1, sDummyArg2 sDummyArg3, sDummyArg4, sDummyArg5)
    Const sArrDummyArgType(4) As String = Array(sDummyArgType1, sDummyArgType2, sDummyArgType3, sDummyArgType4, sDummyArgType5)

    iDummyIndex = iCounterInitialized

    Do Until sArrDummyArg(iDummyIndex) = sBlankCell

        ' Adds a comma for each argument past the first one
        If iDummyIndex > iCounterInitialized Then sSubArgs = sSubArgs + sComma
        
        sSubArgs = sSubArgs + sArrDummyArg(iDummyIndex) + sAs + sArrDummyArgType(iDummyIndex)
        iDummyIndex = iDummyIndex + 1

    Loop

    ' Writes Final Statement:
    WriteSub = sSub + sDummyString + sLeftPar + sSubArgs + sRightPar

End Function

' READY FOR TESTING, MAY PRODUCE ERRORS
Function WriteTableElement(sDummySuffix As String, sDataType As String) As String

    ' Getters
    sDataInit = GetDataLetter(sDataType)
    sPrefixInit = GetPrefixInitials(sDummySuffix)
    sFullPre = sDataInit + sPrefixInit

    Select Case sDummySuffix

        Case sColumn

            iColumnCounter = iColumnCounter + 1

            If oCbgIiClmnRankCell.Value = iLeadColumn Then sStartingColumn = sFullPre + CStr(oCbgIiClmnCodeNameCell) + sColumn
            
            sFinalColumn = sFullPre + CStr(oCbgIiClmnCodeNameCell) + sColumn

            If oCbgIiClmnTypeCell = sConstant Then

                If oCbgIiClmnRankCell = iLeadColumn Then

                    ' Sets The First Column of The First Table, Otherwise The First Column of the Next Table
                    sRightHandSide = IIf(oCbgIiTblRankCell = iLeadColumn, sFirstColumn, sLaggingColumn + sAddTableSkipFactor)

                Else

                    ' Creates A Nonleading Column Within An Existing Table
                    sRightHandSide = CStr(sLaggingColumn) + sPlusOne
                
                End If
                
                ' Creates The Final Statement As A Constant:
                WriteTableElement = DeclareConstant(oCbgIiClmnCodeNameCell, sTypeByte, sRightHandSide, sDummySuffix)

            Else

                ' Creates Column As A Variable  (Double check inputs with declare variable)
                WriteTableElement = DeclareVariable(oCbgIiClmnCodeNameCell, sTypeByte, sDummySuffix)
                
            End If
            

            ' Creates A Column Variable To Use In Relative Assignment
            sLaggingColumn = sFullPre + CStr(oCbgIiClmnCodeNameCell) + sColumn

        Case sCell

            WriteTableElement = DeclareVariable(oCbgIiClmnCodeNameCell, sTypeObject, sDummySuffix)

        Case sHeader

            WriteTableElement = DeclareVariable(oCbgIiClmnCodeNameCell, sTypeObject, sDummySuffix)

            If oCbgIiClmnTypeCell = sConstant Then
                
                sRightHandSide = CStr(oCbgIiClmnMainNameCell)
                sExtendedStatement = DeclareConstant(oCbgIiClmnCodeNameCell, sTypeString, sRightHandSide, sHeader)

                WriteTableElement = ExtendStatement(WriteTableElement, sExtendedStatement)

            End If

        Case Else: MsgBox (NotSupported(sDummySuffix)): End

    End Select

End Function

''' NEW FUNCTIONS '''

Function WriteIndexReassignmentsHeader(oDummyObject As Object) As String

    sDummyString = CStr(oDummyObject)
    WriteIndexReassignmentsHeader = "Index Reassignments For " + sDummyString + " Table Columns"

End Function

Function WriteConstantLHS(sVarFullName as String, sDataType As String) As String

    WriteConstantLHS = sLocalCon + sVarFullName + sAs + sDataType

End Function

Function WriteSelectCase(sDummyString As String) As String

    WriteSelectCase = sSelectOpener + sDummyString

End Function

Function WriteCaseElse(sDummyString As String) As String

    WriteCaseElse = "Case Else: MsgBox(" + Chr(iQuoteMark) +  sDummyString + " failed to perform properly." + chr(iQuoteMark) + sRightPar

End Function