
Attribute VB_Name = "cmsMarker"
Option Explicit

' =============================================================================
'  cmsMarker - workbook-side CMS marking flow, rebuilt on modCms/cCmsBatch.
'
'  Replaces the old winhttpjs.bat + temp-file + cmd.exe pipeline: write_to_cms
'  now scans for changed 5Y marks, computes the new full curve IN MEMORY
'  (ratio front end / parallel long end - see modCms.CMS_ApplyNew5y), and
'  fires one async SET per changed curve through the ExcelBridge HTTP engine.
'  Excel stays responsive; per-curve callbacks recalc each curve row as its
'  SET confirms, and a final callback runs the dependency recalcs once.
'
'  Sheet contract (unchanged from the legacy flow):
'    Main sheet : col A ticker, named ranges cms_col / cms_direct_col /
'                 cms_match_col, HARD_O_STOP
'    Curves     : col B ticker, C full curve label, F ccy, G debt class,
'                 H product, BD:BT the 17-standard quotes (BL = 5Y input)
' =============================================================================

' State shared with the async callbacks (one marking run at a time)
Private gMarkBatch As cCmsBatch
Private gMarkRows As Object        ' fourTuple -> Curves-sheet row
Private gMainSheetName As String
Private gCmsCol As String
Private gCmsDirectCol As String
Private gCmsMatchCol As String
Private gHardStart As Long
Private gHardStop As Long

Public Function ColumnLetter(ColumnNumber As Integer) As String
    Dim n As Long
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    ColumnLetter = s
End Function

Public Sub cms_mark_button()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Publish marks to CMS?", vbQuestion + vbYesNo + vbDefaultButton2)
    If answer = vbYes Then
        Call write_to_cms
    End If
End Sub

Public Sub sync_cms()

    Dim Wb As Workbook
    Set Wb = ThisWorkbook

    Dim ws_main As Worksheet, ws_curves As Worksheet
    Set ws_main = Wb.ActiveSheet
    Set ws_curves = Wb.Sheets("Curves")

    Dim cms_col As String, cms_direct_col As String, cms_match_col As String
    cms_col = ColumnLetter(ws_main.Range("cms_col").Column)
    cms_direct_col = ColumnLetter(ws_main.Range("cms_direct_col").Column)
    cms_match_col = ColumnLetter(ws_main.Range("cms_match_col").Column)

    Dim hard_stop As Long, hard_start As Long
    hard_stop = ws_main.Range("HARD_O_STOP").Row
    hard_start = ws_main.Range("cms_direct_col").Row

    ws_curves.Calculate
    ws_main.Range(cms_direct_col & hard_start & ":" & cms_direct_col & hard_stop).Calculate
    ws_main.Range(cms_match_col & hard_start & ":" & cms_match_col & hard_stop).Calculate
    ws_main.Range(cms_col & hard_start & ":" & cms_col & hard_stop).Calculate

    Wb.Sheets("Main").Range("CurrentTime").Calculate
    Wb.Sheets("Main").Range("SDRData").Calculate

End Sub

Public Sub sync_cms_today_only()

    Dim Wb As Workbook
    Set Wb = ThisWorkbook

    Dim ws_main As Worksheet, ws_curves As Worksheet
    Set ws_main = Wb.ActiveSheet
    Set ws_curves = Wb.Sheets("Curves")

    Dim cms_col As String, cms_direct_col As String, cms_match_col As String
    cms_col = ColumnLetter(ws_main.Range("cms_col").Column)
    cms_direct_col = ColumnLetter(ws_main.Range("cms_direct_col").Column)
    cms_match_col = ColumnLetter(ws_main.Range("cms_match_col").Column)

    Dim hard_stop As Long, hard_start As Long
    hard_stop = ws_main.Range("HARD_O_STOP").Row
    hard_start = ws_main.Range("cms_direct_col").Row

    ws_curves.Range("curves_today").Calculate
    ws_main.Range(cms_direct_col & hard_start & ":" & cms_direct_col & hard_stop).Calculate
    ws_main.Range(cms_match_col & hard_start & ":" & cms_match_col & hard_stop).Calculate
    ws_main.Range(cms_col & hard_start & ":" & cms_col & hard_stop).Calculate

End Sub

Public Sub sync_risk()
    Dim Wb As Workbook
    Set Wb = ThisWorkbook
    Wb.Worksheets("Risk Upload").Range("AllRisk").Dirty
    Wb.Worksheets("Risk Upload").Calculate
    Application.ActiveSheet.Calculate
End Sub

Public Function find_row(key_val As Variant, search_col As String, start_search_row As Integer, sheet_name As String) As Integer
    Dim lr As Long, r As Long
    Dim check_val As Variant
    lr = get_last_row(search_col, start_search_row, sheet_name)
    For r = start_search_row To lr
        check_val = Sheets(sheet_name).Range(search_col & r).Value2
        If check_val = key_val Then
            find_row = r
            Exit Function
        End If
    Next r
    find_row = -1
End Function

Function IsRowVisible(cell As Range)
    On Error Resume Next
    Application.Volatile

    Dim visible As Boolean
    visible = True

    If cell.EntireRow.Hidden = True Then visible = False

    IsRowVisible = visible

End Function

' =============================================================================
'  PUBLISH MARKS (Ctrl+Shift+M)
'  Scan -> compute new curves in memory -> async SET per curve -> callbacks.
'  Returns as soon as the requests are launched; Excel never blocks on CMS.
' =============================================================================

Public Sub write_to_cms()

    Dim state As Object
    Set state = AppState_Optimize()
    On Error GoTo ErrorHandler

    Dim Wb As Workbook
    Set Wb = ThisWorkbook

    Dim ws_main As Worksheet, ws_curves As Worksheet
    Set ws_main = Wb.ActiveSheet
    Set ws_curves = Wb.Sheets("Curves")

    Dim debugCms As Boolean
    On Error Resume Next
    debugCms = CBool([settings.debugCMS])
    On Error GoTo ErrorHandler

    ' Clear any leftover 5Y overrides so BD:BT reflects the unperturbed curve
    Dim lr_curve As Long
    lr_curve = get_last_row("BD", 10, "Curves")
    ws_curves.Range("BL10:BL" & lr_curve).ClearContents

    gMainSheetName = ws_main.Name
    gCmsCol = ColumnLetter(ws_main.Range("cms_col").Column)
    gCmsDirectCol = ColumnLetter(ws_main.Range("cms_direct_col").Column)
    gCmsMatchCol = ColumnLetter(ws_main.Range("cms_match_col").Column)
    gHardStop = ws_main.Range("HARD_O_STOP").Row
    gHardStart = ws_main.Range("cms_direct_col").Row

    ThisWorkbook.ForceFullCalculation = True
    With ws_main
        .Range(gCmsDirectCol & ":" & gCmsDirectCol).Dirty
        .Calculate
    End With

    ' ------------------------------------------------------------------
    ' Scan for changed marks: user-entered value (cms_col) differing from
    ' the live CMS value (cms_direct_col)
    ' ------------------------------------------------------------------
    Dim SR As Integer: SR = 4
    Dim lr As Long
    lr = get_last_row(gCmsMatchCol, SR, ws_main.Name)
    lr = Application.WorksheetFunction.Max(lr, 184) ' Failsafe

    Dim changed As Object
    Set changed = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim orange As Range, drange As Range
    Dim oval As Double, dval As Double
    Dim ticker As String

    For r = SR To lr
        Set orange = ws_main.Range(gCmsCol & r)
        Set drange = ws_main.Range(gCmsDirectCol & r)
        ticker = CStr(ws_main.Range("A" & r).Value2)

        oval = 0: dval = 0
        On Error Resume Next
        oval = CDbl(orange.Value2)
        dval = CDbl(drange.Value2)
        On Error GoTo ErrorHandler

        ' Same filter as the legacy flow: skip zeros, blank tickers, and
        ' formula cells that already match the live CMS value.
        If oval <> 0 And dval <> 0 And ticker <> "" And ticker <> "0" Then
            If Not (orange.HasFormula And (oval = dval)) Then
                If Not changed.Exists(ticker) Then changed.Add ticker, oval
                orange.ClearContents
            End If
        End If
    Next r

    If changed.Count = 0 Then GoTo Finish

    ' ------------------------------------------------------------------
    ' Build the SET batch: new curve = in-memory transform of the existing
    ' curve (BD:BT) onto the new 5Y level
    ' ------------------------------------------------------------------
    Set gMarkRows = CreateObject("Scripting.Dictionary")
    gMarkRows.CompareMode = vbTextCompare

    Dim tuples() As Variant
    Dim quotes() As Variant
    ReDim tuples(0 To changed.Count - 1, 0 To 4)
    ReDim quotes(0 To changed.Count - 1, 0 To 16)

    Dim key As Variant
    Dim chg_row As Long
    Dim n As Long: n = 0
    Dim oldVals As Variant
    Dim old17(0 To 16) As Variant
    Dim new17 As Variant
    Dim i As Long
    Dim fourTuple As String

    For Each key In changed.Keys
        chg_row = find_row(key, "B", 10, "Curves")
        If chg_row >= 10 Then
            oldVals = ws_curves.Range("BD" & chg_row & ":BT" & chg_row).Value2  ' 1x17
            For i = 0 To 16
                If IsNumeric(oldVals(1, i + 1)) And Not IsEmpty(oldVals(1, i + 1)) Then old17(i) = CDbl(oldVals(1, i + 1)) Else old17(i) = Empty
            Next i

            If Not IsEmpty(old17(modCms.CMS_IDX_5Y)) And CDbl(changed(key)) > 0 Then
                If CDbl(old17(modCms.CMS_IDX_5Y)) <> 0 Then
                    new17 = modCms.CMS_ApplyNew5y(old17, CDbl(changed(key)))

                    tuples(n, 0) = ws_curves.Range("B" & chg_row).Value2   ' Ticker
                    tuples(n, 1) = ws_curves.Range("F" & chg_row).Value2   ' Ccy
                    tuples(n, 2) = ws_curves.Range("G" & chg_row).Value2   ' DebtClass
                    tuples(n, 3) = ws_curves.Range("H" & chg_row).Value2   ' Product
                    tuples(n, 4) = "SPREADS"
                    For i = 0 To 16
                        quotes(n, i) = new17(i)
                    Next i

                    fourTuple = modCms.CMS_CompileFourTuple(CStr(tuples(n, 0)), CStr(tuples(n, 1)), _
                                                            CStr(tuples(n, 2)), CStr(tuples(n, 3)))
                    gMarkRows(fourTuple) = chg_row

                    ' Reflect the new 5Y on the Curves sheet (as the legacy
                    ' flow did) so dependent formulas recalc off it
                    ws_curves.Range("BL" & chg_row).Value2 = changed(key)

                    n = n + 1
                End If
            End If
        End If
    Next key

    If n = 0 Then GoTo Finish

    ' Trim to the rows actually populated
    Dim tuplesN() As Variant, quotesN() As Variant
    ReDim tuplesN(0 To n - 1, 0 To 4)
    ReDim quotesN(0 To n - 1, 0 To 16)
    For r = 0 To n - 1
        For i = 0 To 4
            tuplesN(r, i) = tuples(r, i)
        Next i
        For i = 0 To 16
            quotesN(r, i) = quotes(r, i)
        Next i
    Next r

    If debugCms Then
        Debug.Print "DEBUG CMS: would SET " & n & " curve(s) - not sending"
        MsgBox "DEBUG CMS: would send " & n & " curve(s) to CMS", vbOKOnly, "DEBUG CMS"
        GoTo Finish
    End If

    Set gMarkBatch = modCms.CMS_SetCurvesAsync(tuplesN, quotesN, , _
                         "cmsMarker.OnCmsSetCurve", "cmsMarker.OnCmsSetAll")
    Application.StatusBar = "CMS: publishing " & n & " curve(s)..."
    ' Fall through - callbacks take it from here, Excel stays live.

Finish:
ErrorHandler:
    Call AppState_Restore(state)
End Sub

' Fired on the main thread as each curve's SET response lands.
Public Sub OnCmsSetCurve(curve As cCmsCurve, batch As cCmsBatch)
    On Error Resume Next
    If curve.IsOk Then
        If Not gMarkRows Is Nothing Then
            If gMarkRows.Exists(curve.Key) Then
                ThisWorkbook.Sheets("Curves").Range("B" & gMarkRows(curve.Key) & ":BT" & gMarkRows(curve.Key)).Calculate
            End If
        End If
        Debug.Print "CMS SET OK   " & curve.Key & " (" & curve.ElapsedMs & "ms)"
    Else
        Debug.Print "CMS SET FAIL " & curve.Key & ": " & curve.ErrorMsg
    End If
    Application.StatusBar = "CMS: " & (batch.Count - batch.PendingRequests) & "/" & batch.Count & " confirmed"
    On Error GoTo 0
End Sub

' Fired once when every SET in the run has completed.
Public Sub OnCmsSetAll(batch As cCmsBatch)
    On Error Resume Next
    Dim ws_main As Worksheet
    Set ws_main = ThisWorkbook.Sheets(gMainSheetName)

    ThisWorkbook.Sheets("Curves").Range("curveDependencies").Calculate

    ws_main.Range(gCmsDirectCol & gHardStart & ":" & gCmsDirectCol & gHardStop).Calculate
    ws_main.Range(gCmsMatchCol & gHardStart & ":" & gCmsMatchCol & gHardStop).Calculate
    ws_main.Range(gCmsCol & gHardStart & ":" & gCmsCol & gHardStop).Calculate

    ThisWorkbook.Sheets("Main").Range("CurrentTime").Calculate
    ThisWorkbook.Sheets("Main").Range("SDRData").Calculate

    If batch.FailedCount > 0 Then
        Application.StatusBar = "CMS: done, " & batch.FailedCount & " of " & batch.Count & " FAILED (see Immediate window)"
    Else
        Application.StatusBar = False
    End If

    Set gMarkBatch = Nothing
    Set gMarkRows = Nothing
    On Error GoTo 0
End Sub
