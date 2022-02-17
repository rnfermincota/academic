Attribute VB_Name = "EXCEL_VBA_DEPENDENCIES"

'-------------------------------------------------------------------------------------
'See: http://www.cpearson.com/Excel/vbe.aspx
'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------
Public PUB_PROCEDURES_COLLECTION_OBJ As Collection
Public Const PUB_FILES_PATH_STR As String = _
    "\\psf\Home\Desktop\3.4. EUM - ADD-INS\dependencies\modules"
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
'To use these functions you need to set an reference to the VBA Extensibility library.
'In the VBA editor, go the the Tools menu and choose References. In that dialog, scroll down
'to and check the entry for Microsoft Visual Basic For Applications Extensibility 5.3. If you
'do not set this reference, you will receive a User-defined type not defined compiler error.
'-------------------------------------------------------------------------------------

'Import to this workbook all the VBA modules/classes/forms (".bas, .cls, .frm, .frx")
'in the Folder: PUB_FILES_PATH_STR

Sub PROCEDURES_IMPORTING_ALL_FUNC()

Dim FSO_OBJ As Object
Dim FILES_OBJ As Object
Dim FOLDER_OBJ As Object

Set FSO_OBJ = CreateObject("Scripting.FileSystemObject")
Set FOLDER_OBJ = FSO_OBJ.GetFolder(PUB_FILES_PATH_STR)

On Error Resume Next

'Application.ScreenUpdating = False
For Each FILES_OBJ In FOLDER_OBJ.Files
    If Right(FILES_OBJ.name, 3) = "cls" Or Right(FILES_OBJ.name, 3) = "bas" Or Right(FILES_OBJ.name, 3) = "frm" Then
        Application.VBE.ActiveVBProject.VBComponents.Import (PUB_FILES_PATH_STR & FILES_OBJ.name)
    End If
Next FILES_OBJ
'Application.ScreenUpdating = True

End Sub

'Export to this workbook all the VBA modules/classes/forms (".bas, .cls, .frm, .frx")
'in the Folder: PUB_FILES_PATH_STR

Sub PROCEDURES_EXPORTING_ALL_FUNC()

Dim VB_PROJ_OBJ As VBProject
Dim VB_COMP_OBJ As VBComponent

Set VB_PROJ_OBJ = Application.VBE.ActiveVBProject
 
For Each VB_COMP_OBJ In VB_PROJ_OBJ.VBComponents
    If VB_COMP_OBJ.Type = vbext_ct_StdModule Then
        VB_COMP_OBJ.EXPORT PUB_FILES_PATH_STR & VB_COMP_OBJ.name & ".bas"
    ElseIf VB_COMP_OBJ.Type = vbext_ct_ClassModule Then
        VB_COMP_OBJ.EXPORT PUB_FILES_PATH_STR & VB_COMP_OBJ.name & ".cls"
    ElseIf VB_COMP_OBJ.Type = vbext_ct_MSForm Then
        VB_COMP_OBJ.EXPORT PUB_FILES_PATH_STR & VB_COMP_OBJ.name & ".frm"
    Else
        'Debug.Print VB_COMP_OBJ.name
    End If
Next

End Sub


'Identifying dependencies of VB functions and scripts.

Sub TEST_PROCEDURE_DEPENDENTS_FUNC() 'Make sure the libraries are loaded in VBA
Dim i As Long
Dim TEMP_ARR As Variant
Const FUNC_NAME_STR As String = "MATRIX_TRANSPOSE_FUNC"
TEMP_ARR = PROCEDURE_DEPENDENTS_FUNC(FUNC_NAME_STR, True, "_FUNC")
For i = LBound(TEMP_ARR) To UBound(TEMP_ARR)
    Debug.Print TEMP_ARR(i)
Next i
End Sub

Function PROCEDURE_DEPENDENTS_FUNC(ByVal PROCEDURE_NAME_STR As String, _
Optional ByVal CODES_FLAG As Boolean = True, _
Optional ByVal LOOK_STR As String = "_FUNC", _
Optional ByRef VB_PROJ_OBJ As VBIDE.VBProject)

Dim g As Long
Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long

Dim NROWS As Long
Dim LINE_STR As String
Dim pk As vbext_ProcKind

Dim DATA_ARR As Variant
Dim TEMP_ARR As Variant

Dim KEY_STR As String
Dim SRC_MODULE_STR As String

Dim COLLECTION_OBJ As Collection
Dim SRC_COMP_OBJ As VBIDE.VBComponent
Dim CODE_MOD_OBJ As VBIDE.CodeModule

If VB_PROJ_OBJ Is Nothing Then: Set VB_PROJ_OBJ = Application.VBE.ActiveVBProject
Set COLLECTION_OBJ = New Collection

KEY_STR = PROCEDURE_NAME_STR: GoSub MODULE_LINE
If g = 0 Then: GoTo ERROR_LABEL
GoSub EXTRACT_PROCEDURES
n = COLLECTION_OBJ.COUNT
    
If CODES_FLAG = True Then
    Dim DATA_ARR2() As String
    Dim TEMP_ARR2() As String
    For m = 1 To n
        KEY_STR = COLLECTION_OBJ.Item(m): GoSub MODULE_LINE
        If g <> 0 Then: GoSub EXTRACT_CODES
    Next m
    PROCEDURE_DEPENDENTS_FUNC = DATA_ARR2
Else
    ReDim DATA_ARR(1 To n): For m = 1 To n: DATA_ARR(m) = COLLECTION_OBJ.Item(m): Next m
    PROCEDURE_DEPENDENTS_FUNC = DATA_ARR
End If

'----------------------------------------------------------------------------------------
Exit Function
'----------------------------------------------------------------------------------------
MODULE_LINE:
'----------------------------------------------------------------------------------------
If PUB_PROCEDURES_COLLECTION_OBJ Is Nothing Then: Call LOAD_PROCEDURES_PROJECT_FUNC
SRC_MODULE_STR = PUB_PROCEDURES_COLLECTION_OBJ(KEY_STR)
Set SRC_COMP_OBJ = VB_PROJ_OBJ.VBComponents(SRC_MODULE_STR)
Set CODE_MOD_OBJ = SRC_COMP_OBJ.CodeModule
g = CODE_MOD_OBJ.ProcStartLine(KEY_STR, pk)
If g = 0 Then: Return
NROWS = CODE_MOD_OBJ.ProcCountLines(KEY_STR, pk)
'----------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------
ADD_COLLECTION_LINE:
'----------------------------------------------------------------------------------------
    On Error Resume Next
    Call COLLECTION_OBJ.Add(KEY_STR, KEY_STR): If Err.number <> 0 Then Else Err.Clear
    On Error GoTo 0
'----------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------
EXTRACT_PROCEDURES:
'----------------------------------------------------------------------------------------
' Scan through the code module, looking for procedures.
j = g
k = Len(LOOK_STR)
For i = 1 To NROWS
    LINE_STR = CODE_MOD_OBJ.Lines(j, 1)
    If InStr(1, LINE_STR, LOOK_STR) > 0 And REMOVE_EXTRA_SPACES_FUNC(Left(LINE_STR, 1)) <> "'" Then
        l = Len(LINE_STR) - k + 1
        For h = 1 To l
            If Mid(LINE_STR, h, k) <> PROCEDURE_NAME_STR And Mid(LINE_STR, h, k) = LOOK_STR And Mid(LINE_STR, h + k, 1) = "(" Then
                n = h + k
                If n = 0 Then GoTo 1983
                m = h
                Do Until Mid(LINE_STR, m, 1) = " " Or Mid(LINE_STR, m, 1) = "("
                    m = m - 1
                    If m <= 0 Then: GoTo 1983
                Loop
                m = m + 1
                KEY_STR = Mid(LINE_STR, m, n - m): GoSub ADD_COLLECTION_LINE
            End If
1983:
        Next h
    End If
    j = j + 1
Next i
'----------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------
EXTRACT_CODES:
'----------------------------------------------------------------------------------------
    j = g: l = 0: k = 1
    ReDim TEMP_ARR2(1 To NROWS)
    For i = 1 To NROWS
        LINE_STR = CODE_MOD_OBJ.Lines(j, 1)
        If (REMOVE_EXTRA_SPACES_FUNC(LINE_STR) <> "") Then
            TEMP_ARR2(k) = LINE_STR
            k = k + 1
        Else
            l = l + 1
        End If
        j = j + 1
    Next i
    ReDim Preserve TEMP_ARR2(1 To NROWS - l)
    DATA_ARR2 = Split(Join(TEMP_ARR2, Chr(0)) & Chr(0) & Join(DATA_ARR2, Chr(0)), Chr(0))
'----------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------
ERROR_LABEL:
PROCEDURE_DEPENDENTS_FUNC = False
End Function

Private Function LOAD_PROCEDURES_PROJECT_FUNC()
Dim j As Long 'Lines
Dim pk As vbext_ProcKind
Dim PROCEDURE_NAME_STR As String
Dim SRC_COMP_OBJ As VBIDE.VBComponent
Dim CODE_MOD_OBJ As VBIDE.CodeModule
Dim VB_PROJ_OBJ As VBIDE.VBProject

On Error Resume Next

Set PUB_PROCEDURES_COLLECTION_OBJ = New Collection
Set VB_PROJ_OBJ = Application.VBE.ActiveVBProject

For Each SRC_COMP_OBJ In VB_PROJ_OBJ.VBComponents
    Set CODE_MOD_OBJ = SRC_COMP_OBJ.CodeModule
    j = 1
    Do While j < CODE_MOD_OBJ.CountOfLines
        PROCEDURE_NAME_STR = CODE_MOD_OBJ.ProcOfLine(j, pk)
        If PROCEDURE_NAME_STR <> "" Then
            Call PUB_PROCEDURES_COLLECTION_OBJ.Add(SRC_COMP_OBJ.name, PROCEDURE_NAME_STR)
            If Err.number <> 0 Then 'Repeated procedures
                Debug.Print SRC_COMP_OBJ.name & "|" & PROCEDURE_NAME_STR & "|"
                Err.Clear
            End If
            j = j + CODE_MOD_OBJ.ProcCountLines(PROCEDURE_NAME_STR, pk)
        Else
            j = j + 1
        End If
    Loop
    Set CODE_MOD_OBJ = Nothing
    Set SRC_COMP_OBJ = Nothing

Next SRC_COMP_OBJ
LOAD_PROCEDURES_PROJECT_FUNC = True
End Function

'Print the list of all the libraries, subs, and functions in VBA
'Make sure you first load the VBA modules/classes/forms (".bas, .cls, .frm, .frx")
Function PROCEDURES_LISTING_ALL_FUNC()

Dim i As Long
Dim j As Long 'Lines
Dim k As Long
Dim NROWS As Long
Dim TEMP_ARR() As String
Dim sProcName As String
Dim pk As vbext_ProcKind
Dim DST_RNG As Excel.Range

Dim SRC_COMP_OBJ As VBIDE.VBComponent
Dim CODE_MOD_OBJ As VBIDE.CodeModule
Dim VB_PROJ_OBJ As VBIDE.VBProject

On Error GoTo ERROR_LABEL

Set VB_PROJ_OBJ = Application.VBE.ActiveVBProject

NROWS = 1
ReDim TEMP_ARR(1 To NROWS)
For Each SRC_COMP_OBJ In VB_PROJ_OBJ.VBComponents
    ' Find the code module for the project.
    Set CODE_MOD_OBJ = SRC_COMP_OBJ.CodeModule
    ' Scan through the code module, looking for procedures.
    j = 1
    Do While j < CODE_MOD_OBJ.CountOfLines
        sProcName = CODE_MOD_OBJ.ProcOfLine(j, pk)
        If sProcName <> "" Then
            ReDim Preserve TEMP_ARR(1 To NROWS)
            TEMP_ARR(NROWS) = SRC_COMP_OBJ.name & "|" & sProcName & "|"
            NROWS = NROWS + 1
            j = j + CODE_MOD_OBJ.ProcCountLines(sProcName, pk)
        Else ' This line has no procedure, so go to the next line.
            j = j + 1
        End If
    Loop ' clean up
    Set CODE_MOD_OBJ = Nothing
    Set SRC_COMP_OBJ = Nothing
Next SRC_COMP_OBJ
NROWS = NROWS - 1
Call WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC(), ActiveSheet.Parent)
With ActiveSheet
    Set DST_RNG = .Cells(1, 1)
    For k = 1 To NROWS
        i = 1: j = InStr(i, TEMP_ARR(k), "|")
        DST_RNG.Cells(k, 1) = Mid(TEMP_ARR(k), i, j - i)
        i = j + 1: j = InStr(i, TEMP_ARR(k), "|")
        DST_RNG.Cells(k, 2) = Mid(TEMP_ARR(k), i, j - i)
    Next k
End With

PROCEDURES_LISTING_ALL_FUNC = True

Exit Function
ERROR_LABEL:
PROCEDURES_LISTING_ALL_FUNC = False
End Function
