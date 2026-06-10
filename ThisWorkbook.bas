
Public cached_ticker As String

Sub CreateHotkeys()
    Call Hotkeys_Register(Array(Array("^+m", "write_to_cms"), Array("^c", "pre_copy_capture")))
End Sub

Sub RemoveHotKeys()
    Call Hotkeys_Unregister(Array("^+m", "^c"))
End Sub

Public Sub bridge_start()
    modBridge.StartBridge
    SetupObserver
End Sub

Private Sub Workbook_Open()
    Dim state As Object
    Set state = AppState_Optimize()
    
    On Error GoTo ErrorHandler
    Call CreateHotkeys
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Settings")
    
    Call full_recalc
    Call bridge_start
    
ErrorHandler:
    Call AppState_Restore(state)
End Sub

Public Sub bridge_stop()
    modBridge.StopBridge
    modBridge.StopHeartbeat
    IntersectionObserver.UnregisterAll
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim state As Object
    Set state = AppState_Optimize()
    
    Call bridge_stop
    
    On Error GoTo ErrorHandler
    Call clear_hub(True)
    Call RemoveHotKeys
    ThisWorkbook.Worksheets("MarkSchedule").Range("MarkReminderFlag").value = False
    ThisWorkbook.Worksheets("Settings").Range("HubConnectionAttempt").value = 0
    If am_i_connected() Then
        Call DisconnectFromWs
    End If
ErrorHandler:
    Call AppState_Restore(state)
 
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    
    ' Start clean
    Call AppState_HardFix
    
    If Not NamedRangeExistsIn("copy_mode", ActiveSheet) Then Exit Sub
    ActiveSheet.Range("A1:L4").Calculate
    If (ActiveSheet.Range("copy_mode").Value2 = "RUNZ") Then
        Call enable_runz_button
    Else
        Call enable_ib_button
    End If
    Call clear_filters_button_click
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    Application.enableEvents = True
    Application.screenUpdating = True
    
    If NamedRangeExistsIn("cms_col", Sh) Then
        
        IntersectionObserver.InitializeObserver
        IntersectionObserver.UnregisterAll
    
      ' Link columns A and O together as "ticker_group"
      IntersectionObserver.RegisterZoneWithCallback "$A$4:$A$184", IgnoreIfRangePassIfSingleCell, _
          "sdrFilter.SdrFilterTicker", Sh.name, True, "sdrFilter.SdrClearTicker", True, , _
          "ticker_group"
      
      IntersectionObserver.RegisterZoneWithCallback "$O$4:$O$184", IgnoreIfRangePassIfSingleCell, _
          "sdrFilter.SdrIndirectTicker", Sh.name, False, "sdrFilter.SdrClearTicker", False, , _
          "ticker_group2"
          
        On Error Resume Next
        Call Intersect_Link(Target, ActiveSheet.Columns(axe_strings_address()), , "refresh_axe_strings")
        
        On Error Resume Next
        Call Intersect_Link(Target, ActiveSheet.Columns([axes_side_col].Column), , "side_validate")
        
        'Call clear_hub(True)
      End If
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    IntersectionObserver.CheckIntersection Target
End Sub

Sub SetupObserver()
  
End Sub

Function WorksheetExists(shtName As Variant, Optional Wb As Workbook) As Boolean
    Dim sht As Worksheet
    If Wb Is Nothing Then Set Wb = ThisWorkbook
    On Error Resume Next
    Set sht = Wb.Worksheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Public Sub full_recalc()
    Dim calcMode As XlCalculation
    Dim aborted As Boolean
    aborted = False
    
    With frmProgress
        .show vbModeless
        .UserCancelled = False
        .btnAbort.Enabled = True
        .lblMessage.Caption = "Sheet initializing, one moment..."
    End With
    DoEvents
    
    calcMode = Application.Calculation
    
    With Application
        .enableEvents = False
        .screenUpdating = False
        .Calculation = xlCalculationManual
        .Interactive = False
        .EnableCancelKey = xlDisabled
    End With
    
    On Error GoTo finale
    ThisWorkbook.ForceFullCalculation = True

    Dim sheetNames As Variant
    sheetNames = Array("Settings", "Curves", "Main")
    
    Dim i As Integer
    Dim s As Variant
    For i = 0 To UBound(sheetNames)
        s = sheetNames(i)
        If WorksheetExists(s) = True Then
        
            frmProgress.lblMessage.Caption = "Processing: " & sheetNames(i) & " (" & (i + 1) & "/" & (UBound(sheetNames) + 1) & ")"
            frmProgress.Repaint
                        
            On Error Resume Next
            With ThisWorkbook.Worksheets(sheetNames(i))
                .UsedRange.Dirty
                .Calculate
            End With
            On Error GoTo 0
        End If
    Next i
    
finale:
    With Application
        .Interactive = True
        .EnableCancelKey = xlInterrupt
        .Calculation = calcMode
        .enableEvents = True
        .screenUpdating = True
    End With
    
    ThisWorkbook.ForceFullCalculation = False
    
    On Error Resume Next
    Unload frmProgress
    On Error GoTo 0
    
End Sub

Public Sub aaa()
    Application.enableEvents = True
    Application.screenUpdating = True
    SetupObserver
End Sub

Sub ListActiveReferencesByName()
    Dim ref As Object
    
    Debug.Print "--- Active VBA References ---"
    
    ' Loop through all checked references in the active project
    For Each ref In ThisWorkbook.VBProject.References
        Debug.Print "Name: " & ref.name
        If ref.name = "ExcelBridge" Then
        a = 1
        End If
        Debug.Print "Description: " & ref.Description
        Debug.Print "Path: " & ref.fullPath
        Debug.Print "Broken: " & ref.IsBroken
        Debug.Print "-----------------------------"
    Next ref
End Sub

