
Attribute VB_Name = "IntersectionObserver"
Option Explicit

'Configuration Mode Enumeration
Public Enum IntersectionMode
    IgnoreAll = 0                           ' Do nothing
    PassRange = 1                           ' Pass entire intersecting range to handler
    PassOneAtATime = 2                      ' Pass each cell individually to handler
    IgnoreIfRangePassIfSingleCell = 3       ' Ignore if multi-cell, pass handler if single cell
End Enum

'Zone Registration Structure
Private Type ZoneConfig
    rangeAddress As String
    RangeObject As Range    ' Cached range object
    Mode As IntersectionMode
    callbackId As String    ' Fully-qualified callback ID (module.functionname)
    notInRangeCallbackId As String  ' Optional: callback fired when leaving the zone
    nullCellCallbackId As String  ' Optional: callback fired specifically for empty/null cells
    Enabled As Boolean
    LastCheckTime As Long
    sheetName As String    ' Optional: specific sheet, empty = any sheet
    ignoreEmptyCells As Boolean  ' If True, skip empty cells when passing handlers
    deduplicateValues As Boolean  ' If True, don't fire if cell value same as last fire
    LastFiredValue As String  ' Cached value from last handler execution
    WasInRangeLastCheck As Boolean  ' Tracks if we were in range on last check (for NOT callbacks)
    linkedZoneGroupId As String  ' Optional: group ID for linked zones (share dedup, suppress clear when moving between linked)
End Type

'Module-level state
Private Zones() As ZoneConfig
Private ZoneCount As Long
Private Wb As Workbook
Private CachedRanges As Object ' Dictionary for fast lookups
Private LinkedZoneGroups As Object ' Dictionary mapping LinkedZoneGroupId -> shared LastFiredValue

'===============================================================================
' INITIALIZATION & CONFIGURATION
'===============================================================================

' Idempotent: initializes only if the observer has never been initialized.
' Safe to call from any registration path; preserves registered zones.
Public Sub EnsureInitialized(Optional targetWorkbook As Workbook)
    If Wb Is Nothing Or CachedRanges Is Nothing Then InitializeObserver targetWorkbook
End Sub

Public Sub InitializeObserver(Optional targetWorkbook As Workbook)
    ' Initialize the observer with minimal overhead
    On Error GoTo ErrHandler

    If targetWorkbook Is Nothing Then
        Set Wb = ThisWorkbook
    Else
        Set Wb = targetWorkbook
    End If

    ZoneCount = 0
    ReDim Zones(0 To 49)  ' Pre-allocate for up to 50 zones
    Set CachedRanges = CreateObject("Scripting.Dictionary")
    Set LinkedZoneGroups = CreateObject("Scripting.Dictionary")

    Exit Sub
ErrHandler:
    MsgBox "Failed to initialize Observer: " & Err.Description
End Sub

Public Sub RegisterZone(rangeAddress As String, Mode As IntersectionMode, _
                        handlerName As String, Optional moduleName As String = "ThisWorkbook", _
                        Optional sheetName As String = "", Optional ignoreEmptyCells As Boolean = False)
    ' Legacy method - kept for compatibility
    ' For reliable cross-module handlers, use RegisterZoneWithCallback instead

    ' Try to resolve handler and register with callback
    Dim callbackId As String
    callbackId = moduleName & "." & handlerName

    RegisterZoneWithCallback rangeAddress, Mode, callbackId, sheetName, ignoreEmptyCells
End Sub

Public Sub RegisterZoneWithCallback(rangeAddress As String, Mode As IntersectionMode, _
                                   callbackId As String, Optional sheetName As String = "", _
                                   Optional ignoreEmptyCells As Boolean = False, _
                                   Optional notInRangeCallbackId As String = "", _
                                   Optional deduplicateValues As Boolean = False, _
                                   Optional nullCellCallbackId As String = "", _
                                   Optional linkedZoneGroupId As String = "")
    ' Register zone with callback ID (module.functionname format)
    ' This is the RECOMMENDED way to register zones
    ' Example: RegisterZoneWithCallback "$A$1:$A$150", IgnoreIfRangePassIfSingleCell, "sdrFilter.SdrFilterTicker", "Tickers", True
    '
    ' notInRangeCallbackId: Optional callback fired once when leaving the zone
    ' deduplicateValues: Optional - if True, don't fire handler if cell value matches last fired value
    ' nullCellCallbackId: Optional callback fired specifically for empty/null cells
    ' linkedZoneGroupId: Optional group ID for linked zones
    '   - Zones with same ID share dedup state
    '   - NOT callback won't fire when moving between zones in same group
    On Error GoTo ErrHandler

    If ZoneCount >= UBound(Zones) Then
        ReDim Preserve Zones(0 To UBound(Zones) + 50)
    End If

    With Zones(ZoneCount)
        .rangeAddress = rangeAddress
        .Mode = Mode
        .callbackId = callbackId
        .notInRangeCallbackId = notInRangeCallbackId
        .nullCellCallbackId = nullCellCallbackId
        .Enabled = True
        .LastCheckTime = 0
        .sheetName = sheetName
        .ignoreEmptyCells = ignoreEmptyCells
        .deduplicateValues = deduplicateValues
        .LastFiredValue = ""
        .WasInRangeLastCheck = False
        .linkedZoneGroupId = linkedZoneGroupId

        ' Resolve range: use sheet name if provided
        If sheetName <> "" Then
            Set .RangeObject = Wb.Sheets(sheetName).Range(rangeAddress)
        Else
            Set .RangeObject = Wb.Range(rangeAddress)
        End If

        ' Upsert: re-registering the same zone must not raise a duplicate-key error
        Set CachedRanges(rangeAddress & "|" & sheetName) = .RangeObject

        ' Initialize linked zone group if new
        If linkedZoneGroupId <> "" Then
            If Not LinkedZoneGroups.exists(linkedZoneGroupId) Then
                ' Use sentinel value that won't match any real cell value
                LinkedZoneGroups.Add linkedZoneGroupId, "<NEVER_FIRED>"
            End If
        End If
    End With

    ZoneCount = ZoneCount + 1

    Exit Sub
ErrHandler:
    MsgBox "Failed to register zone " & rangeAddress & ": " & Err.Description
End Sub

Public Sub UnregisterZone(rangeAddress As String)
    ' Remove a zone from observation
    Dim i As Long
    Dim j As Long

    For i = 0 To ZoneCount - 1
        If Zones(i).rangeAddress = rangeAddress Then
            ' Shift remaining zones down
            For j = i To ZoneCount - 2
                Zones(j) = Zones(j + 1)
            Next j
            ZoneCount = ZoneCount - 1
            Exit Sub
        End If
    Next i
End Sub

Public Sub DisableZone(rangeAddress As String)
    Dim i As Long
    For i = 0 To ZoneCount - 1
        If Zones(i).rangeAddress = rangeAddress Then
            Zones(i).Enabled = False
            Exit Sub
        End If
    Next i
End Sub

Public Sub EnableZone(rangeAddress As String)
    Dim i As Long
    For i = 0 To ZoneCount - 1
        If Zones(i).rangeAddress = rangeAddress Then
            Zones(i).Enabled = True
            Exit Sub
        End If
    Next i
End Sub

'===============================================================================
' INTERSECTION CHECKING - CORE PERFORMANCE CRITICAL SECTION
'===============================================================================

Public Function CheckIntersection(sourceRange As Range) As Boolean
    ' Check if sourceRange intersects ANY registered zone
    ' Returns True if intersection(s) found
    On Error GoTo ErrHandler

    If ZoneCount = 0 Then
        CheckIntersection = False
        Exit Function
    End If

    Dim i As Long
    Dim intersect As Range
    Dim zoneRange As Range

    ' Fast-path: iterate through zones and check for intersections
    For i = 0 To ZoneCount - 1
        If Zones(i).Enabled Then
            ' Use cached range object for speed
            Set zoneRange = Zones(i).RangeObject
            Set intersect = Application.intersect(sourceRange, zoneRange)

            If Not intersect Is Nothing Then
                ' Found intersection - execute handler
                ExecuteZoneHandler i, intersect
                CheckIntersection = True
                Set intersect = Nothing
            End If
        End If
    Next i

    ' Check for zones we're leaving (NOT in range callbacks)
    For i = 0 To ZoneCount - 1
        If Zones(i).Enabled And Zones(i).notInRangeCallbackId <> "" Then
            Set zoneRange = Zones(i).RangeObject
            Set intersect = Application.intersect(sourceRange, zoneRange)

            ' We're NOT in this zone
            If intersect Is Nothing Then
                ' Were we in this zone on the last check?
                If Zones(i).WasInRangeLastCheck Then
                    ' Check if we're moving to a linked zone (same group)
                    If Not IsMovingToLinkedZone(i, sourceRange) Then
                        ' Fire the NOT callback once
                        On Error Resume Next
                        Application.Run Zones(i).notInRangeCallbackId, sourceRange
                        On Error GoTo 0
                        ' Clear deduplication cache so same value can fire when we return
                        ClearGroupDedupValue Zones(i).linkedZoneGroupId
                    End If
                    ' Mark that we've fired it so we don't fire again until we re-enter
                    Zones(i).WasInRangeLastCheck = False
                End If
            Else
                ' We're still in this zone - enable NOT callback for next time we leave
                Zones(i).WasInRangeLastCheck = True
                Set intersect = Nothing
            End If
        End If
    Next i

    Exit Function
ErrHandler:
    CheckIntersection = False
End Function

Public Function GetIntersections(sourceRange As Range) As Collection
    ' Return all intersecting zones
    Dim result As Collection
    Set result = New Collection

    On Error GoTo ErrHandler

    If ZoneCount = 0 Then
        Set GetIntersections = result
        Exit Function
    End If

    Dim i As Long
    Dim intersect As Range
    Dim zoneRange As Range
    Dim intersectInfo As Object

    For i = 0 To ZoneCount - 1
        If Zones(i).Enabled Then
            Set zoneRange = Zones(i).RangeObject
            Set intersect = Application.intersect(sourceRange, zoneRange)

            If Not intersect Is Nothing Then
                ' Store intersection info without firing handler
                Set intersectInfo = CreateObject("Scripting.Dictionary")
                intersectInfo.Add "ZoneIndex", i
                intersectInfo.Add "Range", intersect
                intersectInfo.Add "ZoneAddress", Zones(i).rangeAddress
                result.Add intersectInfo
                Set intersectInfo = Nothing
            End If
        End If
    Next i

    Set GetIntersections = result
    Exit Function
ErrHandler:
    Set GetIntersections = result
End Function

'===============================================================================
' HANDLER EXECUTION - OPTIMIZED WITH MINIMAL OVERHEAD
'===============================================================================

Private Sub ExecuteZoneHandler(zoneIndex As Long, intersectRange As Range)
    ' Execute the configured handler for a zone
    ' This is called directly from hot loop - keep it lean
    On Error Resume Next  ' Handler may fail, don't crash observer

    With Zones(zoneIndex)
        Select Case .Mode
            Case IgnoreAll
                ' Do nothing

            Case PassRange
                ' Pass entire range to handler
                If ShouldFireHandler(zoneIndex, intersectRange) Then
                    CallZoneHandler zoneIndex, intersectRange
                    .LastFiredValue = intersectRange.value
                    .WasInRangeLastCheck = True
                End If

            Case PassOneAtATime
                ' Iterate and pass each cell
                Dim cell As Range
                Dim cellCount As Long

                ' Optimization: use Areas collection for discontinuous ranges
                Dim areaIdx As Long
                Dim cellIdx As Long

                For areaIdx = 1 To intersectRange.Areas.count
                    For cellIdx = 1 To intersectRange.Areas(areaIdx).cells.count
                        Set cell = intersectRange.Areas(areaIdx).cells(cellIdx)

                        ' Handle empty cells
                        If cell.value = "" Then
                            ' If null callback is registered, fire it
                            If .nullCellCallbackId <> "" Then
                                On Error Resume Next
                                Application.Run .nullCellCallbackId, cell
                                On Error GoTo 0
                                .WasInRangeLastCheck = True
                            ElseIf Not .ignoreEmptyCells Then
                                ' No null callback, but don't ignore empty cells - fire main handler
                                If ShouldFireHandler(zoneIndex, cell) Then
                                    CallZoneHandler zoneIndex, cell
                                    .LastFiredValue = cell.value
                                    .WasInRangeLastCheck = True
                                End If
                            End If
                        Else
                            ' Non-empty cell - fire normal handler
                            If ShouldFireHandler(zoneIndex, cell) Then
                                CallZoneHandler zoneIndex, cell
                                .LastFiredValue = cell.value
                                .WasInRangeLastCheck = True
                            End If
                        End If
                    Next cellIdx
                Next areaIdx

            Case IgnoreIfRangePassIfSingleCell
                ' Only fire handler if the intersection is a single cell
                If IsSingleCell(intersectRange) Then
                    ' Handle empty cells
                    If intersectRange.value = "" Then
                        ' If null callback is registered, fire it
                        If .nullCellCallbackId <> "" Then
                            On Error Resume Next
                            Application.Run .nullCellCallbackId, intersectRange
                            On Error GoTo 0
                            .WasInRangeLastCheck = True
                        ElseIf Not .ignoreEmptyCells Then
                            ' No null callback, but don't ignore empty cells - fire main handler
                            If ShouldFireHandler(zoneIndex, intersectRange) Then
                                CallZoneHandler zoneIndex, intersectRange
                                .LastFiredValue = intersectRange.value
                                .WasInRangeLastCheck = True
                            End If
                        End If
                    Else
                        ' Non-empty cell - fire normal handler
                        If ShouldFireHandler(zoneIndex, intersectRange) Then
                            CallZoneHandler zoneIndex, intersectRange
                            .LastFiredValue = intersectRange.value
                            .WasInRangeLastCheck = True
                        End If
                    End If
                End If
        End Select
    End With

    On Error GoTo 0
End Sub

Private Function IsSingleCell(rng As Range) As Boolean
    ' Ultra-fast check: is this range exactly one cell?
    ' Single cell ranges have Areas.Count = 1 and Cells.Count = 1
    On Error GoTo ErrHandler
    IsSingleCell = (rng.Areas.count = 1) And (rng.cells.count = 1)
    Exit Function
ErrHandler:
    IsSingleCell = False
End Function

Private Function ShouldFireHandler(zoneIndex As Long, rng As Range) As Boolean
    ' Check if handler should fire (handles deduplication if enabled)
    On Error GoTo ErrHandler

    With Zones(zoneIndex)
        ' Check deduplication
        If .deduplicateValues Then
            Dim lastValue As String

            ' Use shared dedup value if in linked group, otherwise use zone's own
            If .linkedZoneGroupId <> "" Then
                lastValue = LinkedZoneGroups.Item(.linkedZoneGroupId)
            Else
                lastValue = .LastFiredValue
            End If

            ' Skip if value matches last fired value
            If rng.value = lastValue Then
                ShouldFireHandler = False
                Exit Function
            End If
        End If
    End With

    ShouldFireHandler = True
    Exit Function
ErrHandler:
    ShouldFireHandler = True  ' Fire on error to be safe
End Function

Private Sub CallZoneHandler(zoneIndex As Long, rng As Range)
    ' Execute handler for zone using registered callback ID
    On Error Resume Next

    ' Get callback ID from zone
    If Zones(zoneIndex).callbackId <> "" Then
        Application.Run Zones(zoneIndex).callbackId, rng
    End If

    On Error GoTo 0

    ' Auto-update dedup ONLY for non-linked zones
    ' For linked zones, handlers manage their own dedup via UpdateGroupDedup()
    With Zones(zoneIndex)
        If .linkedZoneGroupId = "" Then
            UpdateDedupValue zoneIndex, rng.value
        End If
    End With
End Sub

Private Sub UpdateDedupValue(zoneIndex As Long, cellValue As String)
    ' Update the dedup tracking value (for non-linked zones only)
    With Zones(zoneIndex)
        If .linkedZoneGroupId = "" Then
            ' Update zone's own dedup value
            .LastFiredValue = cellValue
        End If
    End With
End Sub

Private Function IsMovingToLinkedZone(leavingZoneIndex As Long, sourceRange As Range) As Boolean
    ' Check if we're moving to another zone in the same linked group
    On Error GoTo ErrHandler

    Dim leavingGroupId As String
    leavingGroupId = Zones(leavingZoneIndex).linkedZoneGroupId

    ' If not part of a linked group, can't be moving to a linked zone
    If leavingGroupId = "" Then
        IsMovingToLinkedZone = False
        Exit Function
    End If

    ' Check if any other zone in same group has intersection
    Dim i As Long
    For i = 0 To ZoneCount - 1
        If i <> leavingZoneIndex And Zones(i).Enabled Then
            If Zones(i).linkedZoneGroupId = leavingGroupId Then
                Dim intersect As Range
                Set intersect = Application.intersect(sourceRange, Zones(i).RangeObject)
                If Not intersect Is Nothing Then
                    IsMovingToLinkedZone = True
                    Exit Function
                End If
            End If
        End If
    Next i

    IsMovingToLinkedZone = False
    Exit Function
ErrHandler:
    IsMovingToLinkedZone = False
End Function

Private Sub ClearGroupDedupValue(linkedZoneGroupId As String)
    ' Clear the dedup value for a linked group
    If linkedZoneGroupId <> "" Then
        If LinkedZoneGroups.exists(linkedZoneGroupId) Then
            LinkedZoneGroups.Item(linkedZoneGroupId) = ""
        End If
    End If
End Sub

Public Sub UpdateGroupDedup(linkedZoneGroupId As String, dedupValue As String)
    ' PUBLIC: Allow handlers to manually update dedup value for their group
    ' Useful when handler uses indirect/transformed values
    ' Example: Zone B uses value from column A, so handler updates dedup with column A's value
    If linkedZoneGroupId <> "" Then
        If LinkedZoneGroups.exists(linkedZoneGroupId) Then
            LinkedZoneGroups.Item(linkedZoneGroupId) = dedupValue
        End If
    End If
End Sub

'===============================================================================
' UTILITY & STATE MANAGEMENT
'===============================================================================

Public Function GetZoneCount() As Long
    GetZoneCount = ZoneCount
End Function

Public Function GetZoneInfo(zoneIndex As Long) As Object
    ' Return zone configuration as dictionary (for debugging)
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")

    If zoneIndex < 0 Or zoneIndex >= ZoneCount Then
        Exit Function
    End If

    With Zones(zoneIndex)
        info.Add "Address", .rangeAddress
        info.Add "Mode", .Mode
        info.Add "Callback", .callbackId
        info.Add "NotInRangeCallback", .notInRangeCallbackId
        info.Add "NullCellCallback", .nullCellCallbackId
        info.Add "Enabled", .Enabled
        info.Add "Sheet", .sheetName
        info.Add "IgnoreEmptyCells", .ignoreEmptyCells
        info.Add "DeduplicateValues", .deduplicateValues
        info.Add "LinkedZoneGroupId", .linkedZoneGroupId
        info.Add "LastFiredValue", .LastFiredValue
        info.Add "WasInRangeLastCheck", .WasInRangeLastCheck
    End With

    Set GetZoneInfo = info
End Function

Public Sub UnregisterAll()
    ' Remove all registered zones - clean slate
    ZoneCount = 0
    ReDim Zones(0 To 49)
    Set CachedRanges = CreateObject("Scripting.Dictionary")
    Set LinkedZoneGroups = CreateObject("Scripting.Dictionary")
End Sub

Public Sub ClearAllZones()
    ' Alias for UnregisterAll (backwards compatibility)
    UnregisterAll
End Sub

Public Sub RefreshZoneCache()
    ' Force re-cache of all zone ranges (call if ranges change structurally)
    Dim i As Long
    On Error Resume Next

    For i = 0 To ZoneCount - 1
        If Zones(i).sheetName <> "" Then
            Set Zones(i).RangeObject = Wb.Sheets(Zones(i).sheetName).Range(Zones(i).rangeAddress)
        Else
            Set Zones(i).RangeObject = Wb.Range(Zones(i).rangeAddress)
        End If
    Next i

    On Error GoTo 0
End Sub

'===============================================================================
' BATCH OPERATIONS - FOR SCANNING MULTIPLE RANGES EFFICIENTLY
'===============================================================================

Public Sub CheckMultipleRanges(rangeArray() As Range)
    ' Efficiently check multiple ranges against zones
    ' Use for performance-critical scenarios like worksheet change events
    On Error GoTo ErrHandler

    Dim i As Long
    For i = LBound(rangeArray) To UBound(rangeArray)
        CheckIntersection rangeArray(i)
    Next i

    Exit Sub
ErrHandler:
    ' Silently fail to avoid cascading errors
End Sub

Public Sub CheckRangeAddress(rangeAddress As String)
    ' Check intersection for a range specified by address
    On Error GoTo ErrHandler

    Dim rng As Range
    Set rng = Wb.Range(rangeAddress)
    CheckIntersection rng
    Set rng = Nothing

    Exit Sub
ErrHandler:
    ' Invalid range address
End Sub

'===============================================================================
' EXAMPLE USAGE & DOCUMENTATION
'===============================================================================

' Example setup code (RECOMMENDED: Use RegisterZoneWithCallback):
'
' Sub SetupIntersectionObserver()
'     InitializeObserver
'
'     ' Format: RegisterZoneWithCallback(range, Mode, "Module.HandlerName", sheet, ignoreEmpty, "NotCallback", deduplicateValues, "NullCallback")
'
'     ' Zone with deduplication: don't fire if value = last fired value
'     IntersectionObserver.RegisterZoneWithCallback "$A$1:$A$150", IgnoreIfRangePassIfSingleCell, _
'         "sdrFilter.SdrFilterTicker", "Tickers", True, "", True
'
'     ' Zone with null cell handler: separate logic for empty cells
'     IntersectionObserver.RegisterZoneWithCallback "$F$1:$F$100", PassOneAtATime, _
'         "sdrFilter.OnCellEdit", "Tickers", False, , , "sdrFilter.OnEmptyCell"
'
'     ' Zone with NOT callback: fire when leaving the zone
'     IntersectionObserver.RegisterZoneWithCallback "$G$1:$G$50", PassOneAtATime, _
'         "sdrFilter.OnRangeEdit", "Tickers", False, "sdrFilter.OnLeftZone"
'
'     ' All features combined
'     IntersectionObserver.RegisterZoneWithCallback "$H$1:$H$100", PassOneAtATime, _
'         "sdrFilter.OnDataEntry", "Tickers", False, "sdrFilter.OnExitDataEntry", True, "sdrFilter.OnEmptyCell"
'
'     ' Linked zones: Column A and B share dedup, suppress clear when moving between them
'     IntersectionObserver.RegisterZoneWithCallback "$A$1:$A$150", IgnoreIfRangePassIfSingleCell, _
'         "sdrFilter.OnPrimaryTicker", "Tickers", True, "sdrFilter.OnLeaveTickers", True, , "ticker_group"
'     IntersectionObserver.RegisterZoneWithCallback "$B$1:$B$150", IgnoreIfRangePassIfSingleCell, _
'         "sdrFilter.OnSecondaryTicker", "Tickers", False, "sdrFilter.OnLeaveTickers", False, , "ticker_group"
' End Sub
'
' ' In sdrFilter module:
' Sub SdrFilterTicker(cell As Range)
'     ' Only fires if cell.Value <> last value (deduplication enabled)
'     MsgBox "Ticker changed: " & cell.Value
' End Sub
'
' Sub OnCellEdit(cell As Range)
'     Debug.Print "Editing in column F: " & cell.Value
' End Sub
'
' Sub OnEmptyCell(cell As Range)
'     ' Fires specifically for empty cells
'     Debug.Print "Empty cell detected at " & cell.Address
' End Sub
'
' Sub OnLeftZone(triggeringRange As Range)
'     ' Fires once when user leaves the zone
'     Debug.Print "User left zone, now at: " & triggeringRange.Address
' End Sub
'
' Sub OnDataEntry(cell As Range)
'     Debug.Print "Data entry: " & cell.Value
' End Sub
'
' Sub OnExitDataEntry(triggeringRange As Range)
'     Debug.Print "Exiting data entry zone"
' End Sub
'
' Sub OnPrimaryTicker(cell As Range)
'     ' Linked zone: fires on column A selection
'     Debug.Print "Primary ticker: " & cell.Value
' End Sub
'
' Sub OnSecondaryTicker(cell As Range)
'     ' Linked zone: fires on column B selection
'     ' Shares dedup with OnPrimaryTicker - if A fired "ABC", B won't fire "ABC"
'     Debug.Print "Secondary ticker: " & cell.Value
' End Sub
'
' Sub OnLeaveTickers(triggeringRange As Range)
'     ' Linked zone NOT callback: fires when leaving BOTH columns A and B together
'     Debug.Print "Left ticker zone, clearing..."
' End Sub
'
' Use in Worksheet_SelectionChange event (on each monitored sheet):
'
' Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'     IntersectionObserver.CheckIntersection Target
' End Sub



