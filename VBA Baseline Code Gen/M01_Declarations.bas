Option Explicit

' WORKBOOKS
    'Public oREPLACE As Workbook:            Public Const sREPLACE As String = ""
    Public oCodeBitsGen As Workbook:         Public Const sCodeBitsGen As String = "3 - Code Bits Generator"
    
' WORKSHEETS
    'Public oWbREPLACE As Worksheet:            Public Const sWbREPLACE As String = ""
    Public oCbgInputsInterface As Worksheet:    Public Const sCbgInputsInterface As String = "Inputs"
    Public oCbgDeclarationsOutput As Worksheet: Public Const sCbgDeclarationsOutput As String = "Declarations"
    Public oCbgSettersOutput As Worksheet:      Public Const sCbgSettersOutput As String = "Setters"
    
' TABLES
    'Public REPLACETable As Object:          Public Const sWbWsREPLACE As String = ""
    Public oCbgIiWorkbooks As Object:        Public Const sCbgIiWorkbooks As String = "CbgIiWorkbooks"
    Public oCbgIiWorksheets As Object:       Public Const sCbgIiWorksheets As String = "CbgIiWorksheets"
    Public oCbgIiTables As Object:           Public Const sCbgIiTables As String = "CbgIiTables"
    Public oCbgIiColumns As Object:          Public Const sCbgIiColumns As String = "CbgIiColumns"
    Public oCbgIiConstants As Object:        Public Const sCbgIiConstants As String = "CbgIiConstants"
    Public oCbgIiVariables As Object:        Public Const sCbgIiVariables As String = "CbgIiVariables"

' COLUMNS/CELLS/HEADERS
    
    ' COLUMNS
        Public Const iCbgIiFirstColumn As Byte = 2
        Public Const iCbgDoFirstColumn As Byte = 2
        Public Const iCbgSoFirstColumn As Byte = 2

        ' Public Const iCbgIiREPLACEColumn As Byte = iCbgIiREPLACEColumn + 1

        ' Workbooks Table
        Public Const iCbgIiWbMainNameColumn As Byte = iCbgIiFirstColumn
        Public Const iCbgIiWbRankColumn As Byte = iCbgIiWbMainNameColumn + 1
        Public Const iCbgIiWbCodeNameColumn As Byte = iCbgIiWbRankColumn + 1
        Public Const iCbgIiWbInitColumn As Byte = iCbgIiWbCodeNameColumn + 1
        Public Const iCbgIiWbFileNameColumn As Byte = iCbgIiWbInitColumn + 1

        Public Const iCbgIiWbColumnCount As Byte = iCbgIiWbFileNameColumn - iCbgIiWbMainNameColumn + 1
        
        ' Worksheets Table
        Public Const iCbgIiWsWbColumn As Byte = iCbgIiWbFileNameColumn + 2
        Public Const iCbgIiWsRankColumn As Byte = iCbgIiWsWbColumn + 1
        Public Const iCbgIiWsMainNameColumn As Byte = iCbgIiWsRankColumn + 1
        Public Const iCbgIiWsCodeNameColumn As Byte = iCbgIiWsMainNameColumn + 1
        Public Const iCbgIiWsInitColumn As Byte = iCbgIiWsCodeNameColumn + 1
        Public Const iCbgIiWsTypeColumn As Byte = iCbgIiWsInitColumn + 1

        Public Const iCbgIiWsColumnCount As Byte = iCbgIiWsTypeColumn - iCbgIiWsWbColumn + 1
        
        ' Tables Table
        Public Const iCbgIiTblWsColumn As Byte = iCbgIiWsTypeColumn + 2
        Public Const iCbgIiTblMainNameColumn As Byte = iCbgIiTblWsColumn + 1
        Public Const iCbgIiTblCodeNameColumn As Byte = iCbgIiTblMainNameColumn + 1
        Public Const iCbgIiTblRankColumn As Byte = iCbgIiTblCodeNameColumn + 1
        Public Const iCbgIiTblInitColumn As Byte = iCbgIiTblRankColumn + 1
        Public Const iCbgIiTblHeaderRowColumn As Byte = iCbgIiTblInitColumn + 1
        Public Const iCbgIiTblTypeColumn As Byte = iCbgIiTblHeaderRowColumn + 1

        Public Const iCbgIiTblColumnCount As Byte = iCbgIiTblTypeColumn - iCbgIiTblWsColumn + 1
        
        ' Cell / Header / Column Table
        Public Const iCbgIiClmnRankColumn As Byte = iCbgIiTblTypeColumn + 2
        Public Const iCbgIiClmnWsColumn As Byte = iCbgIiClmnRankColumn + 1
        Public Const iCbgIiClmnTblColumn As Byte = iCbgIiClmnWsColumn + 1
        Public Const iCbgIiClmnMainNameColumn As Byte = iCbgIiClmnTblColumn + 1
        Public Const iCbgIiClmnCodeNameColumn As Byte = iCbgIiClmnMainNameColumn + 1
        Public Const iCbgIiClmnTypeColumn As Byte = iCbgIiClmnCodeNameColumn + 1

        Public Const iCbgIiClmnColumnCount As Byte = iCbgIiClmnTypeColumn - iCbgIiClmnRankColumn + 1

        ' Constants Table
        Public Const iCbgIiConstNameColumn As Byte = iCbgIiClmnTypeColumn + 2
        Public Const iCbgIiConstTypeColumn As Byte = iCbgIiConstNameColumn + 1
        Public Const iCbgIiConstValueColumn As Byte = iCbgIiConstTypeColumn + 1

        Public Const iCbgIiConstColumnCount As Byte = iCbgIiConstValueColumn - iCbgIiConstNameColumn + 1

        ' Variables Table
        Public Const iCbgIiVarNameColumn As Byte = iCbgIiConstValueColumn + 2
        Public Const iCbgIiVarTypeColumn As Byte = iCbgIiVarNameColumn + 1

        Public Const iCbgIiVarColumnCount As Byte = iCbgIiVarTypeColumn + iCbgIiVarNameColumn + 1
        

        ' OUTPUTS
        
        ' Declarations Sheet
        Public Const iCbgDoCodeColumn As Byte = iCbgDoFirstColumn

        ' Setters Sheet
        Public Const iCbgSoCodeColumn As Byte = iCbgSoFirstColumn
        
        
    ' CELLS
        ' Public oCbgWsTblREPLACECell As Object

        ' Workbooks Table
        Public oCbgIiWbMainNameCell As Object
        Public oCbgIiWbRankCell As Object
        Public oCbgIiWbCodeNameCell As Object
        Public oCbgIiWbInitCell As Object
        Public oCbgIiWbFileNameCell As Object
        
        ' Worksheets Table
        Public oCbgIiWsWbCell As Object
        Public oCbgIiWsRankCell As Object
        Public oCbgIiWsMainNameCell As Object
        Public oCbgIiWsCodeNameCell As Object
        Public oCbgIiWsInitCell As Object
        Public oCbgIiWsTypeCell As Object
        
        ' Tables Table
        Public oCbgIiTblWsCell As Object
        Public oCbgIiTblMainNameCell As Object
        Public oCbgIiTblCodeNameCell As Object
        Public oCbgIiTblRankCell As Object
        Public oCbgIiTblInitCell As Object
        Public oCbgIiTblHeaderRowCell As Object
        Public oCbgIiTblTypeCell As Object
        
        ' Cell / Header / Column Table
        Public oCbgIiClmnRankCell As Object
        Public oCbgIiClmnWsCell As Object
        Public oCbgIiClmnTblCell As Object
        Public oCbgIiClmnMainNameCell As Object
        Public oCbgIiClmnCodeNameCell As Object
        Public oCbgIiClmnTypeCell As Object

        ' Constants Table
        Public oCbgIiConstNameCell As Object
        Public oCbgIiConstTypeCell As Object
        Public oCbgIiConstValueCell As Object

        ' Variables Table
        Public oCbgIiVarNameCell As Object
        Public oCbgIiVarTypeCell As Object
        
        ' Outputs
        Public oCbgDoCodeCell As Object
        Public oCbgSoCodeCell As Object
    
    ' HEADERS
        ' Public oCbgWsTblREPLACEHeader As Object: Public Const sCbgWsTblREPLACEHeader As String = ""

        ' Workbooks Table
        Public oCbgIiWbMainNameHeader As Object:    Public Const sCbgIiWbMainNameHeader As String = "Workbook"
        Public oCbgIiWbRankHeader As Object:        Public Const sCbgIiWbRankHeader As String = "Rank"
        Public oCbgIiWbCodeNameHeader As Object:    Public Const sCbgIiWbCodeNameHeader As String = "Codename"
        Public oCbgIiWbInitHeader As Object:        Public Const sCbgIiWbInitHeader As String = "Init"
        Public oCbgIiWbFileNameHeader As Object:    Public Const sCbgIiWbFileNameHeader As String = "File Name"
        
        ' Worksheets Table
        Public oCbgIiWsWbHeader As Object:          Public Const sCbgIiWsWbHeader As String = "Workbook"
        Public oCbgIiWsRankHeader As Object:        Public Const sCbgIiWsRankHeader As String = "Rank"
        Public oCbgIiWsMainNameHeader As Object:    Public Const sCbgIiWsMainNameHeader As String = "Worksheet"
        Public oCbgIiWsCodeNameHeader As Object:    Public Const sCbgIiWsCodeNameHeader As String = "Codename"
        Public oCbgIiWsInitHeader As Object:        Public Const sCbgIiWsInitHeader As String = "Init"
        Public oCbgIiWsTypeHeader As Object:        Public Const sCbgIiWsTypeHeader As String = "Type"
        
        ' Tables Table
        Public oCbgIiTblWsHeader As Object:         Public Const sCbgIiTblWsHeader As String = "Worksheet"
        Public oCbgIiTblMainNameHeader As Object:   Public Const sCbgIiTblMainNameHeader As String = "Table"
        Public oCbgIiTblCodeNameHeader As Object:   Public Const sCbgIiTblCodeNameHeader As String = "Codename"
        Public oCbgIiTblRankHeader As Object:       Public Const sCbgIiTblRankHeader As String = "Rank"
        Public oCbgIiTblInitHeader As Object:       Public Const sCbgIiTblInitHeader As String = "Init"
        Public oCbgIiTblHeaderRowHeader As Object:  Public Const sCbgIiTblHeaderRowHeader As String = "HeaderRow"
        Public oCbgIiTblTypeHeader As Object:       Public Const sCbgIiTblTypeHeader As String = "Type"
        
        ' Cell / Header / Column Table
        Public oCbgIiClmnRankHeader As Object:      Public Const sCbgIiClmnRankHeader As String = "Rank"
        Public oCbgIiClmnWsHeader As Object:        Public Const sCbgIiClmnWsHeader As String = "Wkst"
        Public oCbgIiClmnTblHeader As Object:       Public Const sCbgIiClmnTblHeader As String = "Table"
        Public oCbgIiClmnMainNameHeader As Object:  Public Const sCbgIiClmnMainNameHeader As String = "Name"
        Public oCbgIiClmnCodeNameHeader As Object:  Public Const sCbgIiClmnCodeNameHeader As String = "Codename"
        Public oCbgIiClmnTypeHeader As Object:      Public Const sCbgIiClmnTypeHeader As String = "Type"

        ' Constants Table
        Public oCbgIiConstNameHeader As Object:     Public Const sCbgIiConstNameHeader As String = "Const"
        Public oCbgIiConstTypeHeader As Object:     Public Const sCbgIiConstTypeHeader As String = "Data Type"
        Public oCbgIiConstValueHeader As Object:    Public Const sCbgIiConstValueHeader As String = "Value"

        ' Variables Table
        Public oCbgIiVarNameHeader As Object:       Public Const sCbgIiVarNameHeader As String = "Var"
        Public oCbgIiVarTypeHeader As Object:       Public Const sCbgIiVarTypeHeader As String = "Data Type"

        
' ROWS AND ROWSCANNERS
    'Public Const iWbWsTblHeaderRow As Byte = 2
    'Public Const iWbWsTblInitialRow As Byte = iWbWsTblHeaderRow + 1
    'Public iWbWsTblRowScanner As Integer: Public Const sWbWsTblRowScanner As String = "sWbWsTblRowScanner"
    
    ' Workbooks RowScanner
    Public Const iCbgIiWbHeaderRow As Byte = 2
    Public Const iCbgIiWbInitialRow As Byte = iCbgIiWbHeaderRow + 1
    Public iCbgIiWbRowScanner As Integer: Public Const sCbgIiWbRowScanner As String = "sCbgIiWbRowScanner"
    
    ' Worksheets RowScanner
    Public Const iCbgIiWsHeaderRow As Byte = 2
    Public Const iCbgIiWsInitialRow As Byte = iCbgIiWsHeaderRow + 1
    Public iCbgIiWsRowScanner As Integer: Public Const sCbgIiWsRowScanner As String = "sCbgIiWsRowScanner"
    
    ' Tables RowScanner
    Public Const iCbgIiTblHeaderRow As Byte = 2
    Public Const iCbgIiTblInitialRow As Byte = iCbgIiTblHeaderRow + 1
    Public iCbgIiTblRowScanner As Integer: Public Const sCbgIiTblRowScanner As String = "sCbgIiTblRowScanner"
    
    ' Cell / Header / Column RowScanners
    Public Const iCbgIiClmnHeaderRow As Byte = 2
    Public Const iCbgIiClmnInitialRow As Byte = iCbgIiClmnHeaderRow + 1
    Public iCbgIiClmnRowScanner As Integer: Public Const sCbgIiClmnRowScanner As String = "sCbgIiClmnRowScanner"

    ' Constant RowScanners
    Public Const iCbgIiConstHeaderRow As Byte = 2
    Public Const iCbgIiConstInitialRow As Byte = iCbgIiConstHeaderRow + 1
    Public iCbgIiConstRowScanner As Integer: Public Const sCbgIiConstRowScanner As String = "sCbgIiConstRowScanner"

    ' Variables RowScanner
    Public Const iCbgIiVarHeaderRow As Byte = 2
    Public Const iCbgIiVarInitialRow As Byte = iCbgIiVarHeaderRow + 1
    Public iCbgIiVarRowScanner As Integer: Public Const sCbgIiVarRowScanner As String = "sCbgIiVarRowScanner"
    
    
    'OUTPUTS
    ' Declarations Sheet
    Public Const iCbgDoInitialRow As Byte = 1
    Public iCbgDoRowScanner As Integer: Public Const sCbgDoRowScanner As String = "sCbgDoRowScanner"
    
    ' Setters Sheet
    Public Const iCbgSoInitialRow As Byte = 1
    Public iCbgSoRowScanner As Integer: Public Const sCbgSoRowScanner As String = "sCbgSoRowScanner"
        
' CONSTANTS

    ' NUMERIC
        'General
            Public Const iRowStop As Integer = 4
            Public Const iLeadColumn As Byte = 1
            Public Const iCounterInitialized As Byte = 0

            Public Const iArrWbIndex As Byte = 0
            Public Const iArrWsIndex As Byte = iArrWbIndex + 1
            Public Const iArrTblIndex As Byte = iArrWsIndex + 1
            Public Const iArrClmnIndex As Byte = iArrTblIndex + 1
            Public Const iArrConstIndex As Byte = iArrClmnIndex + 1
            Public Const iArrVarIndex As Byte = iArrConstIndex + 1

            Public Const iArrWbIndices As Byte = iCbgIiWbColumnCount
            Public Const iArrWsIndices As Byte = iCbgIiWsColumnCount
            Public Const iArrTblIndices As Byte = iCbgIiTblColumnCount
            Public Const iArrClmnIndices As Byte = iCbgIiClmnColumnCount
            Public Const iArrConstIndices As Byte = iCbgIiConstColumnCount
            Public Const iArrVarIndices As Byte = iCbgIiVarColumnCount

        ' ASCII Constants
            Public Const iQuoteMark As Byte = 34
            Public Const iApostMark As Byte = 39

        ' Indent Constants
            Public Const iNoIndent As Byte = 0
            Public Const iSingleIndent As Byte = iNoIndent + 1
            Public Const iDoubleIndent As Byte = iSingleIndent + 1
            Public Const iTripleIndent As Byte = iDoubleIndent + 1
            Public Const iQuadIndent As Byte = iTripleIndent + 1

        ' Array Lengths
            Public Const iArrLengthInitializer As Integer = -1


    ' STRINGS

        ' Blanks / Spacing
            Public Const sBlankCell As String = ""
            Public Const sBlankSpace As String = " "
            Public Const sFourSpaces As String = "    "

        ' Declaration Initiation
            Public Const sDeclareConst As String = "Public Const "
            Public Const sDeclareVar As String = "Public "

        ' Type Not Found
            Public Const sNotFoundPre As String = "The "
            Public Const sNotFoundPost As String = " you have entered is not listed.  Please check your input and try again."

        ' Data Type Strings (for sDataType)
            Public Const sTypeByte As String = "Byte"
            Public Const sTypeBoolean As String = "Boolean"
            Public Const sTypeInteger As String = "Integer"
            Public Const sTypeObject As String = "Object"
            Public Const sTypeString As String = "String"
            Public Const sTypeTable As String = "Table"
            Public Const sTypeVariant As String = "Variant"
            Public Const sTypeWorksheet As String = "Worksheet"
            Public Const sTypeWorkbook As String = "Workbook"

        ' Data Type Prefixes
            Public Const sStrPre As String = "s"
            Public Const sNumPre As String = "i"
            Public Const sBooPre As String = "b"
            Public Const sObjPre As String = "o"
            Public Const sVarPre As String = "v"

        ' Object Type Suffixes
            Public Const sCell As String = "Cell"
            Public Const sColumn As String = "Column"
            Public Const sHeader As String = "Header"
            Public Const sInitialRow As String = "InitialRow"
            Public Const sNumRowScanner As String = "iRowScanner"
            Public Const sRow As String = "Row"
            Public Const sRowScanner As String = "RowScanner"
            Public Const sStrRowScanner As String = "sRowScanner"
            Public Const sTable As String = "Table"
            Public Const sWorkbook As String = "Workbook"
            Public Const sWorksheet As String = "Worksheet"
            Public Const sHeaderRow As String = "HeaderRow"

    
        ' Variable Declaration Types
            Public Const sVariable As String = "Variable"
            Public Const sConstant As String = "Constant"
            Public Const sLocalConstant As String = "Local Constant"
            Public Const sLocalVariable As String = "Local Variable"
            Public Const sGlobalConstant As String = "Global Constant"
            Public Const sGlobalVariable As String = "Global Variable"
            

        ' HEADERS
            Public Const sHeaderCells As String = "CELLS"
            Public Const sHeaderColumns As String = "COLUMNS"
            Public Const sHeaderConstants As String = "CONSTANTS"
            Public Const sHeaderHeaders As String = "HEADERS"
            Public Const sHeaderRowScanners As String = "ROW SCANNERS"
            Public Const sHeaderTables As String = "TABLES"
            Public Const sHeaderVariables As String = "VARIABLES"
            Public Const sHeaderWorkbooks As String = "WORKBOOKS"
            Public Const sHeaderWorksheets As String = "WORKSHEETS"

        ' SETTER HEADERS
            Public Const sTableSetters As String = "Table Setters"
            Public Const sHeaderSetters As String = "Header Setters"
            Public Const sWorksheetSetters As String = "Worksheet Setters"
            Public Const sWorkbookSetters As String = "Workbook Setters"

        ' Object Prefix Init
            Public Const sArrInit As String = "Arr"
            Public Const sClmnInit As String = "Clmn"
            Public Const sGblC As String = "GblC"
            Public Const sGblV As String = "GblV"
            Public Const sTblInit As String = "Tbl"
            Public Const sWbInit As String = "Wb"
            Public Const sWsInit As String = "Ws"

        ' Expression Phrase Parts
            Public Const sAddTableSkipFactor As String = " + TableSkipFactor"
            Public Const sArr As String = "Arr"
            Public Const sAs As String = " As "
            Public Const sCase As String = "Case "
            Public Const sColon As String = ": "
            Public Const sComma As String = ", "
            Public Const sDotCells As String = ".Cells"
            Public Const sDotListObjects As String = ".ListObjects"
            Public Const sDotValue As String = ".Value "
            Public Const sDotWorksheets As String = ".Worksheets"
            Public Const sEndSub As String = "End Sub"
            Public Const sEquals As String = " = "
            Public Const sFinalIndex As String = "FinalIndex"
            Public Const sFirstColumn As String = "1"
            Public Const sLeftPar As String = "("
            Public Const sMinus As String = " - "
            Public Const sOptExp As String = "Option Explicit"
            Public Const sPairPar As String = "()"
            Public Const sPlus As String = " + "
            Public Const sPlusOne As String = " + 1"
            Public Const sPrefix As String = "s"
            Public Const sRightPar As String = ")"
            Public Const sSet As String = "Set "
            Public Const sSub As String = "Sub "
            Public Const sThisWorkbook As String = "ThisWorkbook"
            Public Const sTableLength As String = "TableLength"
            Public Const sArrTitlePre As String = "Arrays for Table "
            Public Const sLocalDec As String = "Dim "
            Public Const sLocalCon As String = "Const "
            ' Public Const sRowNotFound As String = "Case Else: MsgBox(NotSupported(sRow)): End"
            ' Public Const sTableNotFound As String = "Case Else: MsgBox(NotSupported(sTable)): End"
            ' Public Const sWorksheetNotFound As String = "Case Else: MsgBox(NotSupported(sWorksheet)): End"
            ' Public Const sWorkbookNotFound As String = "Case Else: MsgBox(NotSupported(sWorkbook)): End"
            Public Const sSelectCloser As String = "End Select"
            Public Const sSelectOpener As String = "Select Case "
            Public Const sIndex As String = "Index"
            Public Const sArray As String = "Array"
            Public Const sSelected As String = "Selected"
            Public Const sStepValue As String = "StepValue"

        ' Source Code Phrases
            Public Const sSourceCodeHeader As String = "Set oWbWsTblCOLUMNHeader = oWbWORKSHEET.Cells(iWbWsTblHeaderRow, iWbWsTblCOLUMNColumn):       oWbWsTblCOLUMNHeader.Value = sWbWsTblCOLUMNHeader"
            Public Const sSourceCodeCell As String = "Set oWbWsTblCOLUMNCell = oWbWORKSHEET.Cells(iWbWsTblRowScanner, iWbWsTblCOLUMNColumn)"
            Public Const sSourceCodeTable As String = "Set oWbWsTABLE = oWbWORKSHEET.ListObjects(sWbWsTABLE)"
            Public Const sSourceCodeSetNextRow As String = "Case sWbWsTblRowScanner: iWbWsTblRowScanner = iWbWsTblRowScanner + iStepValue"

' VARIABLES
    ' Boolean
        Public bMultipleElements As Boolean

    ' Numeric
        Public iArrIndexCounter As Byte
        Public iArrTblIndexCounter As Integer
        Public iArrWbIndexCounter As Integer
        Public iArrWsIndexCounter As Integer
        Public iColumnCounter As Byte
        Public iLetterCounter As Integer
        Public iTableCounter As Integer
        Public iRowCounter As Integer
        Public iIndentCounter As Integer
        Public iDummyHeaderRow As Byte
        Public iScriptIndex As Byte

    ' Objects
        Public oDummyObject As Object
        Public oDummyTable As Object
        Public oSortKey As Object
        Public oMainNameCell As Object

    ' Strings
        Public sArrayElements As String
        Public sDataType As String
        Public sDelayedIndex As String
        Public sDummyMessage As String
        Public sDummyInput As String
        Public sFinalColumn As String
        Public sLaggingColumn As String
        Public sLaggingString As String
        Public sNameEnding As String
        Public sLeftHandSide As String
        Public sRightHandSide As String
        Public sSelectedRHS As String
        Public sPrefixInit As String
        Public sStartingColumn As String
        Public sTitlePlaceholder As String
        Public sFullDummyRowName As String
        Public sShortDummyRowName As String
        Public sStringPrefix As String
        Public sTableCodeName As String
        Public sVTblInit As String
        Public sVWsInit As String
        Public sVWbInit As String
        Public sTableSelected As String
        Public sStatementToWrite As String
        Public sFullPre As String
        Public sInputName As String
        Public sVariableName As String
        Public sConstantName As String
        Public sDataInit As String
        Public sSuffix As String
        Public sDummyRowScanner As String
        Public sDummyObjectType As String
        Public sInputValue As String
        Public sWbPrefixInit As String
        Public sWsPrefixInit As String
        Public sTblPrefixInit As String
        Public sClmnPrefixInit As String
        Public sPartCodename As String
        Public sFullCodename As String
        Public sRowStatement As String
        Public sColumnStatement As String
        Public sRowScanInit As String
        Public sPreludeClause As String
        Public sTypeObjectPrelude As String
        Public sPrefixInitials As String
        Public sCodenameRowScanner As String
        Public sEqualityStatement As String
        Public sElmtPrefixInit As String
        Public sArgName_1 As String
        Public sArgName_2 As String
        Public sArgName_3 As String
        Public sArgName_4 As String
        Public sArgName_5 As String
        Public sSubTitle As String
        Public sCallSub As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String
        ' Public As String

    ' Variants
        Public sArrMainName As String
        Public sConstantFullName As String
        Public sDeclType As String
        Public sExtendedStatement As String
        Public sWsCodename As String
        Public sTblCodename As String
        Public sClmnCodeName As String
        Public Const iFinalScriptIndex As Byte = 12

' ARRAYS

    ' Index Declarations

        ' Workbooks
            Public Const iCbgIiWbMainNameIndex As Byte = 0
            Public Const iCbgIiWbRankIndex As Byte = iCbgIiWbMainNameIndex + 1
            Public Const iCbgIiWbCodenameIndex As Byte = iCbgIiWbRankIndex + 1
            Public Const iCbgIiWbInitIndex As Byte = iCbgIiWbCodenameIndex + 1
            Public Const iCbgIiWbFileNameIndex As Byte = iCbgIiWbInitIndex + 1

        ' Worksheets
            Public Const iCbgIiWsWbIndex As Byte = 0
            Public Const iCbgIiWsRankIndex As Byte = iCbgIiWsWbIndex + 1
            Public Const iCbgIiWsMainNameIndex As Byte = iCbgIiWsRankIndex + 1
            Public Const iCbgIiWsCodeNameIndex As Byte = iCbgIiWsMainNameIndex + 1
            Public Const iCbgIiWsInitIndex As Byte = iCbgIiWsCodeNameIndex + 1
            Public Const iCbgIiWsTypeIndex As Byte = iCbgIiWsInitIndex + 1

        ' Tables
            Public Const iCbgIiTblWsIndex As Byte = 0
            Public Const iCbgIiTblMainNameIndex As Byte = iCbgIiTblWsIndex + 1
            Public Const iCbgIiTblCodeNameIndex As Byte = iCbgIiTblMainNameIndex + 1
            Public Const iCbgIiTblRankIndex As Byte = iCbgIiTblCodeNameIndex + 1
            Public Const iCbgIiTblInitIndex As Byte = iCbgIiTblRankIndex + 1
            Public Const iCbgIiTblHeaderRowIndex As Byte = iCbgIiTblInitIndex + 1
            Public Const iCbgIiTblColumnIndex As Byte = iCbgIiTblHeaderRowIndex + 1

        ' Columns
            Public Const iCbgIiClmnRankIndex As Byte = 0
            Public Const iCbgIiClmnWsIndex As Byte = iCbgIiClmnRankIndex + 1
            Public Const iCbgIiClmnTblIndex As Byte = iCbgIiClmnWsIndex + 1
            Public Const iCbgIiClmnMainNameIndex As Byte = iCbgIiClmnTblIndex + 1
            Public Const iCbgIiClmnCodeNameIndex As Byte = iCbgIiClmnMainNameIndex + 1
            Public Const iCbgIiClmnTypeIndex As Byte = iCbgIiClmnCodeNameIndex + 1

        ' Constants
            Public Const iCbgIiConstNameIndex As Byte = 0
            Public Const iCbgIiConstTypeIndex As Byte = iCbgIiConstNameIndex + 1
            Public Const iCbgIiConstValueIndex As Byte = iCbgIiConstTypeIndex + 1

        ' Variables
            Public Const iCbgIiVarNameIndex As Byte = 0
            Public Const iCbgIiVarTypeIndex As Byte = iCbgIiVarNameIndex + 1

        ' Super Array
            Public Const iArrWorkbookIndex As Byte = 0
            Public Const iArrWorksheetIndex As Byte = iArrWorkbookIndex + 1
            Public Const iArrTableIndex As Byte = iArrWorksheetIndex + 1
            Public Const iArrColumnIndex As Byte = iArrTableIndex + 1
            Public Const iArrConstantsIndex As Byte = iArrColumnIndex + 1
            Public Const iArrVariablesIndex As Byte = iArrConstantsIndex + 1
            Public Const iArrDeclarationsIndex As Byte = iArrVariablesIndex + 1
            Public Const iArrSettersIndex As Byte = iArrDeclarationsIndex + 1

    ' Array Declarations
        
        ' All Tables
            Public iArrAllTableColumns(iArrVarIndex) As Byte
            Public oArrAllTableCells(iArrVarIndex) As Object
            Public oArrAllTableHeaders(iArrVarIndex) As Object
            Public sArrAllTableHeaders(iArrVarIndex) As Object
            Public iArrRowScanners(iArrSettersIndex) As Integer
            Public sArrRowScanners(iArrSettersIndex) As String
        
        ' Columns
            Public iArrWorkbookColumns(iArrWbIndices) As Byte
            Public iArrWorksheetColumns(iArrWsIndices) As Byte
            Public iArrTableColumns(iArrTblIndices) As Byte
            Public iArrColumnColumns(iArrClmnIndices) As Byte
            Public iArrConstantsColumns(iArrConstIndices) As Byte
            Public iArrVariablesColumns(iArrVarIndices) As Byte

        ' Cells
            Public oArrWorkbookCells(iArrWbIndices) As Object
            Public oArrWorksheetCells(iArrWsIndices) As Object
            Public oArrTableCells(iArrTblIndices) As Object
            Public oArrColumnCells(iArrClmnIndices) As Object
            Public oArrConstantsCells(iArrConstIndices) As Object
            Public oArrVariablesCells(iArrVarIndices) As Object

        ' Headers (Object)
            Public oArrWorkbookHeaders(iArrWbIndices) As Object
            Public oArrWorksheetHeaders(iArrWsIndices) As Object
            Public oArrTableHeaders(iArrTblIndices) As Object
            Public oArrColumnHeaders(iArrClmnIndices) As Object
            Public oArrConstantsHeaders(iArrConstIndices) As Object
            Public oArrVariablesHeaders(iArrVarIndices) As Object

        ' Headers (String)
            Public sArrWorkbookHeaders(iArrWbIndices) As Object
            Public sArrWorksheetHeaders(iArrWsIndices) As Object
            Public sArrTableHeaders(iArrTblIndices) As Object
            Public sArrColumnHeaders(iArrClmnIndices) As Object
            Public sArrConstantsHeaders(iArrConstIndices) As Object
            Public sArrVariablesHeaders(iArrVarIndices) As Object
            
        ' Script Arrays
            Public iArrIndents(iFinalScriptIndex) As Byte
            Public sArrStatements(iFinalScriptIndex) As String
            Public iArrSpaces(iFinalScriptIndex) As Byte


