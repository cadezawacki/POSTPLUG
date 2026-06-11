Option Explicit
Dim dic As Object

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
    answer = MsgBox("Publish marks to CMS?", vbQuestion + vbYesNo + vbDefaultButton2)
    
    If answer = vbYes Then
        Call write_to_cms
    End If
End Sub

Public Sub sync_cms()

    Dim Wb As Workbook
    Set Wb = ThisWorkbook
    
    Set ws_main = Wb.ActiveSheet
    Set ws_curves = Wb.Sheets("Curves")
    
    Dim cms_col As String, cms_direct_col As String, cms_match_col As String
    cms_col = ColumnLetter(ws_main.Range("cms_col").Column)
    cms_direct_col = ColumnLetter(ws_main.Range("cms_direct_col").Column)
    cms_match_col = ColumnLetter(ws_main.Range("cms_match_col").Column)
    
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
    
    Set ws_main = Wb.ActiveSheet
    Set ws_curves = Wb.Sheets("Curves")
    
    Dim cms_col As String, cms_direct_col As String, cms_match_col As String
    cms_col = ColumnLetter(ws_main.Range("cms_col").Column)
    cms_direct_col = ColumnLetter(ws_main.Range("cms_direct_col").Column)
    cms_match_col = ColumnLetter(ws_main.Range("cms_match_col").Column)
    
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

' SEE HELPERS MODULE FOR NEW IMPLEMENTATION
'Public Function get_last_row(col As String, Optional start_row As Integer = 1, Optional sheet_name As String = "Imports")
'    With Sheets(sheet_name)
'        a = .Range(.Range(col & start_row), .Range(col & start_row).End(xlDown))
'    End With
'    get_last_row = UBound(a) + start_row
'End Function

Public Function find_row(key_val As Variant, search_col As String, start_search_row As Integer, sheet_name As String) As Integer
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


Public Sub write_to_cms()

    Dim full_curve As String, ticker As String, ccy As String, dc As String, prod As String
    Dim v0m As Double, v3m As Double, v6m As Double, v9m As Double
    Dim v1y As Double, v2y As Double, v3y As Double, v4y As Double, v5y As Double, v6y As Double, v7y As Double, v8y As Double, v9y As Double, v10y As Double, v15y As Double, v20y As Double, v30y As Double
            
    ' Keyboard Shortcut: Ctrl+Shift+M
    Dim state As Object
    Set state = AppState_Optimize()
    On Error GoTo ErrorHandler
    
    Dim Wb As Workbook
    Set Wb = ThisWorkbook
    
    Set ws_main = Wb.ActiveSheet
    Set ws_curves = Wb.Sheets("Curves")
    
    lr_curve = get_last_row("BD", 10, "Curves")
    ws_curves.Range("BL" & 10 & ":BL" & lr_curve).ClearContents ' Clear 5yr
    
    Set dic = CreateObject("Scripting.Dictionary")
    Dim cms_col As String, cms_direct_col As String, cms_match_col As String
    cms_col = ColumnLetter(ws_main.Range("cms_col").Column)
    cms_direct_col = ColumnLetter(ws_main.Range("cms_direct_col").Column)
    cms_match_col = ColumnLetter(ws_main.Range("cms_match_col").Column)
    
    Dim hard_stop, hard_start As Integer
    hard_stop = ws_main.Range("HARD_O_STOP").Row
    hard_start = ws_main.Range("cms_direct_col").Row
    
    'ws_curves.Calculate
    ThisWorkbook.ForceFullCalculation = True
    With ws_main
        .Range(cms_direct_col & ":" & cms_direct_col).Dirty
        .Calculate
    End With
    
    Dim SR As Integer: SR = 4
    lr = get_last_row(cms_match_col, SR, ws_main.name)
    lr = Application.WorksheetFunction.Max(lr, 184) ' Failsafe
    
    Dim total_to_ch As Integer: total_to_ch = 0
    
    Dim drange As Range
    Dim orange As Range
    Dim oval As Double
    Dim dval As Double
    
    
    i = 0
    For r = SR To lr
        
        Set orange = ws_main.Range(cms_col & r)
        Set drange = ws_main.Range(cms_direct_col & r)
        ticker = ws_main.Range("A" & r).Value2
        
        If orange Is Nothing Or drange Is Nothing Then
            GoTo ContinueIteration
        End If
        
        oval = CDbl(orange.Value2)
        dval = CDbl(drange.Value2)
        
        If oval = 0 Or dval = 0 Then
            GoTo ContinueIteration
        End If
        
        If orange.HasFormula And (oval = dval) Then
            GoTo ContinueIteration
        End If
    
        If ticker = "" Or ticker = "0" Then
            GoTo ContinueIteration
        End If
        
        total_to_ch = total_to_ch + 1
        dic.Add ticker, oval
        orange.ClearContents
        
ContinueIteration:
    Next r
    
    If total_to_ch = 0 Then
        GoTo ErrorHandler
    End If
    
    Call cHttpBatch.Init
    Dim cmsObject As Object
    Set cmsObject = CreateObject("Scripting.Dictionary")
    Call cHttpBatch.SetCallback(cmsObject, "cmsCallback")
    
    Dim key As Variant
    Dim new_five As Double
    Dim curve_ticker As String
    Dim sXml As String
    
    For Each key In dic.keys
        new_five = dic(key)
        
        On Error Resume Next
        chg_row = find_row(key, "B", 10, "Curves")
        
        If chg_row < 10 Then
            GoTo NextKey
        End If
        
        ws_curves.Range("BL" & chg_row).Value2 = new_five
        With ws_curves
            .Row(chg_row).Dirty
            .Calculate
        End With
        
        With ws_curves
            full_curve = .Range("C" & chg_row).Value2
            curve_ticker = .Range("B" & chg_row).Value2
            ccy = .Range("F" & chg_row).Value2
            dc = .Range("G" & chg_row).Value2
            prod = .Range("H" & chg_row).Value2
            v0m = .Range("BD" & chg_row).Value2
            v3m = .Range("BE" & chg_row).Value2
            v6m = .Range("BF" & chg_row).Value2
            v9m = .Range("BG" & chg_row).Value2
            v1y = .Range("BH" & chg_row).Value2
            v2y = .Range("BI" & chg_row).Value2
            v3y = .Range("BJ" & chg_row).Value2
            v4y = .Range("BK" & chg_row).Value2
            v5y = .Range("BL" & chg_row).Value2
            v6y = .Range("BM" & chg_row).Value2
            v7y = .Range("BN" & chg_row).Value2
            v8y = .Range("BO" & chg_row).Value2
            v9y = .Range("BP" & chg_row).Value2
            v10y = .Range("BQ" & chg_row).Value2
            v15y = .Range("BR" & chg_row).Value2
            v20y = .Range("BS" & chg_row).Value2
            v30y = .Range("BT" & chg_row).Value2
        End With
        
        If v5y > 0 Then
            On Error Resume Next
            sXml = create_cds_xml(full_curve, ticker, dc, ccy, prod, v0m, v3m, v6m, v9m, v1y, v2y, v3y, v4y, v5y, v6y, v7y, v8y, v9y, v10y, v15y, v20y, v30y)
            'Call create_txt(curve_ticker & ".txt", sXml, "c:\temp\")
            Call cHttpBatch.AddPost("http://cms-lxp.lehman.com/CreditMarkingService/", sXml, "Content-Type: text/xml" & vbLf & _
                "charset: UTF-8", 10000)
        End If
NextKey:
    Next key
    
    Application.StatusBar = "Waiting for CMS Response"
    Call cHttpBatch.WaitAll
    Application.StatusBar = False
    
    ws_main.Range(cms_col & r).ClearContents
    
    Dim varKey As Variant
    ws_curves.Range("curveDependencies").Calculate
    x = chg_row
    
    For Each varKey In dic.keys()
        Row = dic(varKey)
        ws_curves.Range("B" & Row & ":" & "BT" & Row).Calculate
    Next
    
    ws_main.Range(cms_direct_col & hard_start & ":" & cms_direct_col & hard_stop).Calculate
    ws_main.Range(cms_match_col & hard_start & ":" & cms_match_col & hard_stop).Calculate
    ws_main.Range(cms_col & hard_start & ":" & cms_col & hard_stop).Calculate
    
    Wb.Sheets("Main").Range("CurrentTime").Calculate
    Wb.Sheets("Main").Range("SDRData").Calculate
    Application.StatusBar = False
    
ErrorHandler:
    Call AppState_Restore(state)
End Sub

Public Sub cmsCallback(requestId As String, r As cHttpResponse)
    Debug.Print (requestId)
    Debug.Print (r)
End Sub


Public Sub test()
    Application.enableEvents = True
End Sub