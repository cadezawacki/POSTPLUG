
Attribute VB_Name = "modCmsSheet"
Option Explicit

' =============================================================================
'  modCmsSheet - the curve grid's edit/stage/mark behavior.
'
'  Grid layout (named ranges on the sheet; rows 4 .. [last_row].Row):
'    col A           ticker
'    cms_col         =CURVE($A5,"5Y")                   live 5Y
'    cms_yest_col    =CURVE($A5,"Prev5Y")               T-1 close (read-only)
'    cms_chg_col_2   =O5-P5                             live - close
'    cms_chg_col     =CURVEDIFF($A5,"Prev5Y","5Y")      same, via the store
'
'  Edit semantics (per row; multi-row and multi-area edits handled):
'    edit cms_col     -> stage that level (cell goes red constant), other
'                        three cells recalc off it
'    edit chg col(s)  -> new 5Y = T-1 close + change; staged into cms_col;
'                        the edited cell reverts to its formula
'    edit yest col    -> reverted to its formula (not editable)
'    clear any        -> un-stage the row, restore ALL formulas + formatting
'
'  Wiring:
'    Workbook_SheetChange -> modCmsSheet.CMS_HandleCurveGridChange Sh, Target
'    mark button          -> modCmsSheet.mark_pending_button
'
'  NOTE the yest column uses CURVE(...,"Prev5Y") rather than CMS_PrevQuote:
'  same number, but it subscribes the cell, shows #Pending while the T-1
'  fetch is in flight, and updates itself. Adjust the Formula* functions
'  below if your sheet wants different formulas - row reverts rebuild from
'  these templates.
' =============================================================================

' --- staged-cell style ---
Private Const STAGED_FILL As Long = 14213887     ' RGB(255,222,214) as Long (B*65536+G*256+R)
Private Const STAGED_FONT As Long = 5806         ' RGB(166,22,0)

' =============================================================================
'  FORMULA TEMPLATES (edit these to match your sheet)
' =============================================================================

Private Function FormulaCms(ByVal r As Long) As String
    FormulaCms = "=CURVE($A" & r & ",""5Y"")"
End Function

Private Function FormulaYest(ByVal r As Long) As String
    FormulaYest = "=CURVE($A" & r & ",""Prev5Y"")"
End Function

Private Function FormulaChg2(ByVal Sh As Object, ByVal r As Long) As String
    FormulaChg2 = "=" & Sh.Cells(r, Sh.Range("cms_col").Column).Address(False, False) & _
                  "-" & Sh.Cells(r, Sh.Range("cms_yest_col").Column).Address(False, False)
End Function

Private Function FormulaChg(ByVal r As Long) As String
    FormulaChg = "=CURVEDIFF($A" & r & ",""Prev5Y"",""5Y"")"
End Function

' =============================================================================
'  CHANGE HANDLER ENGINE
'  Call from Workbook_SheetChange:  modCmsSheet.CMS_HandleCurveGridChange Sh, Target
' =============================================================================

Public Sub CMS_HandleCurveGridChange(ByVal Sh As Object, ByVal Target As Range)
    Dim cCms As Long, cYest As Long, cChg2 As Long, cChg As Long
    Dim firstCol As Long, lastCol As Long, lastRow As Long
    Dim zone As Range, hit As Range
    Dim area As Range
    Dim r As Long
    Dim doneRows As Object
    Dim prevEvents As Boolean, prevScreen As Boolean

    ' Cheap bail-outs first: this runs on EVERY sheet change.
    On Error GoTo HardExit
    If Target Is Nothing Then Exit Sub
    If Not GridColumns(Sh, cCms, cYest, cChg2, cChg, lastRow) Then Exit Sub

    firstCol = Application.WorksheetFunction.Min(cCms, cYest, cChg2, cChg)
    lastCol = Application.WorksheetFunction.Max(cCms, cYest, cChg2, cChg)
    Set zone = Sh.Range(Sh.Cells(4, firstCol), Sh.Cells(lastRow, lastCol))
    Set hit = Intersect(Target, zone)
    If hit Is Nothing Then Exit Sub

    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    If hit.Cells.Count > 3 Then Application.ScreenUpdating = False

    On Error GoTo SafeExit
    Set doneRows = CreateObject("Scripting.Dictionary")

    ' Multi-area, multi-row: process each affected ROW exactly once.
    For Each area In hit.Areas
        For r = area.Row To area.Row + area.Rows.Count - 1
            If Not doneRows.Exists(r) Then
                doneRows.Add r, True
                ProcessRow Sh, r, hit, cCms, cYest, cChg2, cChg
            End If
        Next r
    Next area

SafeExit:
    If Err.Number <> 0 Then Debug.Print "CMS grid: " & Err.Description & " (row " & r & ")"
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub
HardExit:
    Application.EnableEvents = True
End Sub

' One row's decision tree. hit = the changed cells (already intersected with
' the grid zone) so we know exactly which columns the user touched.
Private Sub ProcessRow(ByVal Sh As Object, ByVal r As Long, ByVal hit As Range, _
                       ByVal cCms As Long, ByVal cYest As Long, _
                       ByVal cChg2 As Long, ByVal cChg As Long)
    Dim t As String
    Dim eCms As Boolean, eYest As Boolean, eChg2 As Boolean, eChg As Boolean
    Dim vCms As Variant, vChg2 As Variant, vChg As Variant, vYest As Variant
    Dim newLevel As Double
    Dim haveLevel As Boolean

    t = Trim$(CStr(Sh.Cells(r, "A").Value2))
    If Len(t) = 0 Or t = "0" Then Exit Sub

    eCms = Touched(hit, Sh, r, cCms)
    eYest = Touched(hit, Sh, r, cYest)
    eChg2 = Touched(hit, Sh, r, cChg2)
    eChg = Touched(hit, Sh, r, cChg)
    If Not (eCms Or eYest Or eChg2 Or eChg) Then Exit Sub

    vCms = Sh.Cells(r, cCms).Value2
    vYest = Sh.Cells(r, cYest).Value2
    vChg2 = Sh.Cells(r, cChg2).Value2
    vChg = Sh.Cells(r, cChg).Value2

    ' ---- any cleared cell in the row -> un-stage and revert everything ----
    If (eCms And IsBlank(vCms)) Or (eYest And IsBlank(vYest)) _
       Or (eChg2 And IsBlank(vChg2)) Or (eChg And IsBlank(vChg)) Then
        If CMS_IsRegistered(t) Then
            On Error Resume Next
            CMS_ClearPending5y t
            On Error GoTo 0
        End If
        RestoreRow Sh, r
        Exit Sub
    End If

    ' ---- precedence: cms_col beats the change columns ----
    If eCms And IsNumeric(vCms) Then
        newLevel = CDbl(vCms)
        haveLevel = True
    ElseIf (eChg2 And IsNumeric(vChg2)) Or (eChg And IsNumeric(vChg)) Then
        ' new 5Y = T-1 close + change; needs a numeric close in the cache
        If Not IsNumeric(vYest) Then
            Debug.Print "CMS grid row " & r & " (" & t & "): T-1 close not loaded - cannot stage from a change edit"
            RestoreRow Sh, r        ' put the edited chg cell back to its formula
            Exit Sub
        End If
        newLevel = CDbl(vYest) + CDbl(IIf(eChg2 And IsNumeric(vChg2), vChg2, vChg))
        haveLevel = True
    ElseIf eYest Then
        ' read-only column: silently revert
        Sh.Cells(r, cYest).Formula = FormulaYest(r)
        Sh.Cells(r, cYest).Calculate
        Exit Sub
    End If
    If Not haveLevel Then Exit Sub   ' non-numeric garbage typed: leave as-is

    If Not CMS_IsRegistered(t) Then
        Debug.Print "CMS grid row " & r & " (" & t & "): not registered - edit ignored"
        RestoreRow Sh, r
        Exit Sub
    End If

    ' ---- stage ----
    CMS_StagePending5y t, newLevel

    With Sh.Cells(r, cCms)
        .Value2 = newLevel                 ' staged constant in cms_col
        .Interior.Color = STAGED_FILL      ' red = staged, awaiting mark
        .Font.Color = STAGED_FONT
        .Font.Bold = True
    End With
    ' change columns always revert to formulas (they now show the staged move)
    Sh.Cells(r, cYest).Formula = FormulaYest(r)
    Sh.Cells(r, cChg2).Formula = FormulaChg2(Sh, r)
    Sh.Cells(r, cChg).Formula = FormulaChg(r)

    CMS_SubscribeRange t, RowZone(Sh, r, cCms, cYest, cChg2, cChg)
    RowZone(Sh, r, cCms, cYest, cChg2, cChg).Calculate
End Sub

' Restore one row to its pristine state: all four formulas, no styling.
Public Sub RestoreRow(ByVal Sh As Object, ByVal r As Long)
    Dim cCms As Long, cYest As Long, cChg2 As Long, cChg As Long
    Dim lastRow As Long
    If Not GridColumns(Sh, cCms, cYest, cChg2, cChg, lastRow) Then Exit Sub

    Sh.Cells(r, cCms).Formula = FormulaCms(r)
    Sh.Cells(r, cYest).Formula = FormulaYest(r)
    Sh.Cells(r, cChg2).Formula = FormulaChg2(Sh, r)
    Sh.Cells(r, cChg).Formula = FormulaChg(r)

    With RowZone(Sh, r, cCms, cYest, cChg2, cChg)
        .Interior.ColorIndex = xlNone
        .Font.ColorIndex = xlAutomatic
        .Font.Bold = False
        .Calculate
    End With
End Sub

' Rebuild EVERY grid row's formulas/formatting (setup / repair helper).
Public Sub CMS_RebuildCurveGrid()
    Dim Sh As Worksheet
    Dim r As Long
    Dim cCms As Long, cYest As Long, cChg2 As Long, cChg As Long
    Dim lastRow As Long
    Set Sh = GridSheet()
    If Sh Is Nothing Then Exit Sub
    If Not GridColumns(Sh, cCms, cYest, cChg2, cChg, lastRow) Then Exit Sub
    Application.EnableEvents = False
    For r = 4 To lastRow
        If Len(Trim$(CStr(Sh.Cells(r, "A").Value2))) > 0 Then RestoreRow Sh, r
    Next r
    Application.EnableEvents = True
End Sub

' =============================================================================
'  MARK BUTTON + CALLBACKS
' =============================================================================

Public Sub mark_pending_button()
    Dim b As cCmsBatch
    Set b = CMS_MarkPendingAsync("modCmsSheet.OnMarkCurveDone", "modCmsSheet.OnMarkAllDone")
    If b Is Nothing Then
        Application.StatusBar = "CMS: nothing staged"
    Else
        Application.StatusBar = "CMS: publishing " & b.Count & " curve(s)..."
    End If
End Sub

' Per-curve SET confirmation: restore the row (formulas pick up the newly
' confirmed level from the store; red styling clears). Failures stay red so
' the user can see what didn't take.
Public Sub OnMarkCurveDone(curve As cCmsCurve, batch As cCmsBatch)
    Dim Sh As Worksheet
    Dim m As Variant
    On Error Resume Next
    Set Sh = GridSheet()
    If Sh Is Nothing Then Exit Sub

    If curve.IsOk Then
        m = Application.Match(curve.Ticker, Sh.Columns(1), 0)
        If Not IsError(m) Then
            Application.EnableEvents = False
            RestoreRow Sh, CLng(m)
            Application.EnableEvents = True
        End If
        Debug.Print "CMS MARK OK   " & curve.Key & " 5Y=" & CStr(curve.Quote("5Y")) & " (" & curve.ElapsedMs & "ms)"
    Else
        Debug.Print "CMS MARK FAIL " & curve.Key & ": " & curve.ErrorMsg
    End If
    Application.StatusBar = "CMS: " & batch.CompletedCurves & "/" & batch.Count & " confirmed"
    On Error GoTo 0
End Sub

Public Sub OnMarkAllDone(batch As cCmsBatch)
    On Error Resume Next
    If batch.FailedCount > 0 Then
        Application.StatusBar = "CMS: done - " & batch.FailedCount & " of " & batch.Count & _
                                " FAILED (rows left red; see Immediate window)"
    Else
        Application.StatusBar = False
    End If
    On Error GoTo 0
End Sub

' =============================================================================
'  INTERNAL HELPERS
' =============================================================================

' The grid sheet = wherever the cms_col name points.
Private Function GridSheet() As Worksheet
    On Error Resume Next
    Set GridSheet = ThisWorkbook.Names("cms_col").RefersToRange.Worksheet
    On Error GoTo 0
End Function

' Resolve the grid's columns + last row from the sheet's named ranges.
' False (and no error) when this sheet isn't the curve grid.
Private Function GridColumns(ByVal Sh As Object, ByRef cCms As Long, ByRef cYest As Long, _
                             ByRef cChg2 As Long, ByRef cChg As Long, ByRef lastRow As Long) As Boolean
    On Error Resume Next
    cCms = Sh.Range("cms_col").Column
    cYest = Sh.Range("cms_yest_col").Column
    cChg2 = Sh.Range("cms_chg_col_2").Column
    cChg = Sh.Range("cms_chg_col").Column
    lastRow = Sh.Range("last_row").Row
    On Error GoTo 0
    GridColumns = (cCms > 0 And cYest > 0 And cChg2 > 0 And cChg > 0 And lastRow >= 4)
End Function

Private Function RowZone(ByVal Sh As Object, ByVal r As Long, ByVal cCms As Long, _
                         ByVal cYest As Long, ByVal cChg2 As Long, ByVal cChg As Long) As Range
    Set RowZone = Sh.Range( _
        Sh.Cells(r, Application.WorksheetFunction.Min(cCms, cYest, cChg2, cChg)), _
        Sh.Cells(r, Application.WorksheetFunction.Max(cCms, cYest, cChg2, cChg)))
End Function

Private Function Touched(ByVal hit As Range, ByVal Sh As Object, _
                         ByVal r As Long, ByVal c As Long) As Boolean
    Touched = Not Intersect(hit, Sh.Cells(r, c)) Is Nothing
End Function

Private Function IsBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlank = True
    ElseIf VarType(v) = vbString Then
        IsBlank = (Len(Trim$(v)) = 0)
    End If
End Function
