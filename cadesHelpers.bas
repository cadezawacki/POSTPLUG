
Attribute VB_Name = "cadesHelpers"
Option Explicit

' ==========================================================================================
' Clipboard
' ==========================================================================================
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function RegisterClipboardFormatA Lib "user32" (ByVal lpString As String) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ==========================================================================================
' Utilities
' ==========================================================================================

' Make a dictionary wrapper
Private Function NewDictionary(Optional ByVal caseSensitive As Boolean = True) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If caseSensitive Then
        d.CompareMode = vbBinaryCompare
    Else
        d.CompareMode = vbTextCompare
    End If
    Set NewDictionary = d
End Function

'--- Tests if a Variant is a 1-D dynamic array
Private Function IsDynamicArray(ByVal v As Variant) As Boolean
    On Error GoTo EH
    If IsArray(v) Then
        Dim l As Long, u As Long
        l = LBound(v): u = UBound(v)
        IsDynamicArray = (u - l + 1 >= 0)
    Else
        IsDynamicArray = False
    End If
    Exit Function
EH:
    IsDynamicArray = False
End Function

'--- Returns string value for a Variant (Null-> "", Nothing->"")
Private Function AsString(ByVal v As Variant) As String
    If IsObject(v) Then
        If v Is Nothing Then
            AsString = ""
        Else
            AsString = CStr(v)
        End If
    ElseIf IsNull(v) Then
        AsString = ""
    Else
        AsString = CStr(v)
    End If
End Function

'--- HTML encode minimal entities
Private Function HtmlEncode(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    HtmlEncode = t
End Function

'--- Joins a 1D Variant array to string with delimiter
Private Function Join1D(ByVal arr As Variant, Optional ByVal delimiter As String = vbCrLf) As String
    If Not IsArray(arr) Then
        Join1D = AsString(arr)
        Exit Function
    End If
    Dim l As Long, u As Long, i As Long
    l = LBound(arr): u = UBound(arr)
    Dim parts() As String: ReDim parts(0 To Application.Max(0, u - l))
    For i = l To u
        parts(i - l) = AsString(arr(i))
    Next
    Join1D = VBA.Join(parts, delimiter)
End Function

'--- Ensures we get a Range object from various inputs (Range or Address)
Public Function ResolveRange( _
    ByVal rangeOrAddress As Variant, _
    Optional ByVal Wb As Variant, _
    Optional ByVal ws As Variant) As Range
    Dim twb As Workbook, tws As Worksheet
    If IsObject(Wb) Then
        If TypeName(Wb) = "Workbook" Then Set twb = Wb
    End If
    If twb Is Nothing Then
        If Not Application.ActiveWorkbook Is Nothing Then
            Set twb = Application.ActiveWorkbook
        Else
            Set twb = ThisWorkbook
        End If
    End If
    If IsObject(ws) Then
        If TypeName(ws) = "Worksheet" Then Set tws = ws
    End If
    If tws Is Nothing Then
        If Not twb Is Nothing Then
            If Not twb.ActiveSheet Is Nothing Then
                If TypeName(twb.ActiveSheet) = "Worksheet" Then
                    Set tws = twb.ActiveSheet
                End If
            End If
            If tws Is Nothing Then Set tws = twb.Worksheets(1)
        End If
    End If
    If IsObject(rangeOrAddress) Then
        If TypeName(rangeOrAddress) = "Range" Then
            Set ResolveRange = rangeOrAddress
        Else
            Err.Raise 5, , "ResolveRange: Unsupported object type: " & TypeName(rangeOrAddress)
        End If
    Else
        Set ResolveRange = tws.Range(CStr(rangeOrAddress))
    End If
End Function

' True if v is a 2D array
Private Function Is2DArray(ByVal v As Variant) As Boolean
    On Error GoTo E
    Dim t As Long
    t = LBound(v, 2) ' will error if not 2D
    Is2DArray = True: Exit Function
E:
    Is2DArray = False
End Function

' Returns 1D size
Private Function Size1D(ByVal arr As Variant) As Long
    On Error GoTo E
    Size1D = UBound(arr) - LBound(arr) + 1
    Exit Function
E:
    Size1D = 0
End Function

' Returns (#rows, #cols) for a 2D array
Private Sub Size2D(ByVal arr As Variant, ByRef rows As Long, ByRef cols As Long)
    On Error GoTo E
    rows = UBound(arr, 1) - LBound(arr, 1) + 1
    cols = UBound(arr, 2) - LBound(arr, 2) + 1
    Exit Sub
E:
    rows = 0: cols = 0
End Sub

' Create a composite dictionary key from an array of parts (safe for Null/Empty).
Private Function KeyOfParts(ByVal parts As Variant) As String
    Dim i As Long, s As String
    Dim SEP As String
    SEP = Chr$(30)
    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then s = s & SEP
        s = s & AsString(parts(i))
    Next
    KeyOfParts = s
End Function

' Safely convert a 1D array to a 2D (rows x 1) array for Range writes.
Private Function To2D(ByVal arr1D As Variant) As Variant
    If Not IsArray(arr1D) Then
        Dim t() As Variant: ReDim t(1 To 1, 1 To 1): t(1, 1) = arr1D: To2D = t: Exit Function
    End If
    Dim l As Long, u As Long, i As Long
    l = LBound(arr1D): u = UBound(arr1D)
    Dim t2() As Variant: ReDim t2(1 To u - l + 1, 1 To 1)
    For i = l To u: t2(i - l + 1, 1) = arr1D(i): Next
    To2D = t2
End Function


' ==========================================================================================
' STRING BUILDER functional, delimiter-aware, skip-empty option
' Usage: s = SB_Append(s, "line"); s = SB_Append(s, "next", ", ")
' ==========================================================================================
' Appends a value onto an existing string builder with an optional delimiter (default newline).
' - Start with builder = "".
' - If skipEmpty:=True, empty new values are ignored.
Public Function SB_Append(ByVal builder As String, _
                          ByVal value As Variant, _
                          Optional ByVal delimiter As String = vbCrLf, _
                          Optional ByVal skipEmpty As Boolean = True) As String
    Dim piece As String
    piece = AsString(value)
    If skipEmpty And Len(piece) = 0 Then
        SB_Append = builder
        Exit Function
    End If
    If Len(builder) = 0 Then
        SB_Append = piece
    Else
        SB_Append = builder & delimiter & piece
    End If
End Function


' ==========================================================================================
' ARRAY BUILDER � append to 1D Variant array
' Usage: arr = AB_Append(arr, value)
' ==========================================================================================
' Appends a single value to a 1D dynamic Variant array and returns the new array.
' Start by passing an uninitialized Variant (Dim arr As Variant) or an empty dynamic array.
Public Function AB_Append(ByVal arr As Variant, ByVal value As Variant) As Variant
    Dim l As Long, u As Long
    If Not IsArray(arr) Then
        ReDim arr(0 To 0)
        arr(0) = value
        AB_Append = arr
        Exit Function
    End If
    On Error GoTo EH
    l = LBound(arr): u = UBound(arr)
    ReDim Preserve arr(l To u + 1)
    arr(u + 1) = value
    AB_Append = arr
    Exit Function
EH:
    ' Recover if bounds are invalid or not 1D
    Dim tmp() As Variant
    ReDim tmp(0 To 0)
    tmp(0) = value
    AB_Append = tmp
End Function

Public Function toString(ByVal myInput As Variant) As String
    Dim text As String
    If IsArray(myInput) Then
        text = Join1D(myInput, vbCrLf)
    ElseIf IsObject(myInput) And TypeName(myInput) = "Range" Then
        text = RangeToTSV(myInput)
    Else
        text = AsString(myInput)
    End If
    toString = text
End Function

' ==========================================================================================
' [3] CLIPBOARD � write text or array; optional CF_HTML block like Excel
' Usage:
'   Call Clipboard_Write("Hello")
'   Call Clipboard_Write(arr, True)   ' array becomes newline-delimited, plus HTML Format
' ==========================================================================================
' Writes to clipboard. If input is an array, joins with newline. If includeHtml:=True,
' sets both CF_UNICODETEXT and HTML Format (CF_HTML) using the official header with StartHTML,
' EndHTML, StartFragment, EndFragment (Excel-compatible). If Forms.DataObject is available,
Public Function Clipboard_Write(ByVal myInput As Variant, _
                                Optional ByVal includeHtml As Boolean = False, _
                                Optional ByVal htmlTitle As String = "Clipboard", _
                                Optional ByVal htmlFragment As String = "", _
                                Optional ByVal sourceUrl As String = "") As Boolean
    On Error GoTo EH

    Dim text As String
    If IsArray(myInput) Then
        text = Join1D(myInput, vbCrLf)
    ElseIf IsObject(myInput) And TypeName(myInput) = "Range" Then
        text = RangeToTSV(myInput)
        If includeHtml And Len(htmlFragment) = 0 Then
            htmlFragment = RangeToSimpleHtmlTable(myInput, htmlTitle)
        End If
    Else
        text = AsString(myInput)
    End If

    If Not includeHtml Then
        ' Try MSForms.DataObject for plain text first
        Dim dobj As Object
        On Error Resume Next
        Set dobj = CreateObject("Forms.DataObject")
        On Error GoTo EH
        If Not dobj Is Nothing Then
            dobj.SetText text
            dobj.PutInClipboard
            Clipboard_Write = True
            Exit Function
        End If
        ' Fallback to WinAPI Unicode text
        Clipboard_Write = SetClipboardUnicodeText(text)
        Exit Function
    Else
        If Len(htmlFragment) = 0 Then
            ' Default: simple <pre> fragment
            htmlFragment = "<pre>" & HtmlEncode(text) & "</pre>"
        End If
        Clipboard_Write = SetClipboardTextAndHtml(text, BuildCFHtml(htmlFragment, htmlTitle, sourceUrl))
        Exit Function
    End If

EH:
    Clipboard_Write = False
End Function

'--- Converts a Range to tab-separated
Public Function RangeToTSV(ByVal rng As Range) As String
    Dim r As Long, c As Long, rr As Long, cc As Long
    Dim sb As String, line As String
    For r = 1 To rng.rows.count
        line = ""
        For c = 1 To rng.Columns.count
            If c > 1 Then line = line & vbTab
            line = line & AsString(rng.cells(r, c).text)
        Next
        sb = SB_Append(sb, line, vbCrLf, False)
    Next
    RangeToTSV = sb
End Function

'--- Builds simple HTML <table> for a range
Public Function RangeToSimpleHtmlTable(ByVal rng As Range, Optional ByVal title As String = "Data") As String
    Dim r As Long, c As Long
    Dim html As String
    html = "<table>" & vbCrLf
    For r = 1 To rng.rows.count
        html = html & "<tr>"
        For c = 1 To rng.Columns.count
            html = html & "<td>" & HtmlEncode(AsString(rng.cells(r, c).text)) & "</td>"
        Next
        html = html & "</tr>" & vbCrLf
    Next
    html = html & "</table>"
    RangeToSimpleHtmlTable = "<div><h3>" & HtmlEncode(title) & "</h3>" & html & "</div>"
End Function

'--- Creates a CF_HTML payload per spec with StartHTML/EndHTML/StartFragment/EndFragment.
Private Function BuildCFHtml(ByVal fragmentHtml As String, _
                             Optional ByVal title As String = "Clipboard", _
                             Optional ByVal sourceUrl As String = "") As String
    Dim doc As String, header As String, startFragTag As String, endFragTag As String
    startFragTag = "<!--StartFragment-->"
    endFragTag = "<!--EndFragment-->"
    doc = "<!DOCTYPE html>" & vbCrLf & _
          "<html><head><meta charset=""utf-8""><title>" & HtmlEncode(title) & "</title></head><body>" & _
          startFragTag & fragmentHtml & endFragTag & _
          "</body></html>"

    ' Placeholder header with 10-digit offsets, to be replaced once we know byte positions.
    header = "Version:0.9" & vbCrLf & _
             "StartHTML:0000000000" & vbCrLf & _
             "EndHTML:0000000000" & vbCrLf & _
             "StartFragment:0000000000" & vbCrLf & _
             "EndFragment:0000000000" & vbCrLf & _
             "StartSelection:0000000000" & vbCrLf & _
             "EndSelection:0000000000" & vbCrLf & _
             "SourceURL:" & sourceUrl & vbCrLf

    ' Offsets are in BYTES of the final ANSI (CF_HTML is ANSI). Compute using vbFromUnicode.
    Dim bytesHeader() As Byte, bytesDoc() As Byte
    bytesHeader = StrConv(header, vbFromUnicode)
    bytesDoc = StrConv(doc, vbFromUnicode)

    Dim startHTML As Long, endHTML As Long
    startHTML = UBound(bytesHeader) + 1          ' header byte length
    endHTML = startHTML + (UBound(bytesDoc) + 1) ' header + doc

    ' Find fragment within doc (as bytes)
    Dim idxStart As Long, idxEnd As Long
    idxStart = InStrB(1, doc, startFragTag, vbBinaryCompare) ' returns char index in Unicode string (byte-based in *B)
    idxEnd = InStrB(1, doc, endFragTag, vbBinaryCompare)

    If idxStart = 0 Or idxEnd = 0 Then
        ' Fallback: whole body
        idxStart = InStrB(1, doc, "<body>", vbTextCompare)
        If idxStart = 0 Then idxStart = 1
        idxEnd = InStrB(1, doc, "</body>", vbTextCompare)
        If idxEnd = 0 Then idxEnd = LenB(doc)
    Else
        ' Move to *after* the startFragTag
        idxStart = idxStart + LenB(startFragTag)
    End If

    Dim startFragment As Long, endFragment As Long
    ' Final offsets: header bytes + [fragment offset within doc, as bytes]
    startFragment = startHTML + (idxStart - 1)
    endFragment = startHTML + (idxEnd - 1)

    ' Replace placeholders with fixed-width 10-digit numbers
    header = Replace(header, "StartHTML:0000000000", "StartHTML:" & Right$("0000000000" & CStr(startHTML), 10))
    header = Replace(header, "EndHTML:0000000000", "EndHTML:" & Right$("0000000000" & CStr(endHTML), 10))
    header = Replace(header, "StartFragment:0000000000", "StartFragment:" & Right$("0000000000" & CStr(startFragment), 10))
    header = Replace(header, "EndFragment:0000000000", "EndFragment:" & Right$("0000000000" & CStr(endFragment), 10))
    header = Replace(header, "StartSelection:0000000000", "StartSelection:" & Right$("0000000000" & CStr(startFragment), 10))
    header = Replace(header, "EndSelection:0000000000", "EndSelection:" & Right$("0000000000" & CStr(endFragment), 10))

    BuildCFHtml = header & doc
End Function

'--- Sets only CF_UNICODETEXT using WinAPI (fallback when Forms.DataObject absent)
Private Function SetClipboardUnicodeText(ByVal text As String) As Boolean
    On Error GoTo EH
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = 13

    If OpenClipboard(0) = 0 Then GoTo EH
    If EmptyClipboard() = 0 Then GoTo EH

    Dim byteCount As LongPtr
    byteCount = (Len(text) + 1) * 2 ' wide chars including null
    Dim hMem As LongPtr, pMem As LongPtr
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, byteCount)
    If hMem = 0 Then GoTo EH
    pMem = GlobalLock(hMem)
    If pMem = 0 Then GoTo EH

    CopyMemory pMem, StrPtr(text), byteCount
    GlobalUnlock hMem

    If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then GoTo EH

    CloseClipboard
    SetClipboardUnicodeText = True
    Exit Function
EH:
    On Error Resume Next
    CloseClipboard
    SetClipboardUnicodeText = False
End Function

'--- Sets both CF_UNICODETEXT and CF_HTML
Private Function SetClipboardTextAndHtml(ByVal plainText As String, ByVal cfHtml As String) As Boolean
    On Error GoTo EH
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = 13

    Dim CF_HTML As Long
    CF_HTML = RegisterClipboardFormatA("HTML Format")
    If CF_HTML = 0 Then GoTo EH

    If OpenClipboard(0) = 0 Then GoTo EH
    If EmptyClipboard() = 0 Then GoTo EH

    ' Put Unicode text
    Dim textBytes As LongPtr, hText As LongPtr, pText As LongPtr
    textBytes = (Len(plainText) + 1) * 2
    hText = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, textBytes)
    If hText = 0 Then GoTo EH
    pText = GlobalLock(hText)
    If pText = 0 Then GoTo EH
    CopyMemory pText, StrPtr(plainText), textBytes
    GlobalUnlock hText
    If SetClipboardData(CF_UNICODETEXT, hText) = 0 Then GoTo EH

    ' Put CF_HTML (ANSI)
    Dim htmlBytes() As Byte
    htmlBytes = StrConv(cfHtml, vbFromUnicode)
    Dim hHtml As LongPtr, pHtml As LongPtr
    hHtml = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, UBound(htmlBytes) + 2) ' +1 for ubound offset, +1 for null
    If hHtml = 0 Then GoTo EH
    pHtml = GlobalLock(hHtml)
    If pHtml = 0 Then GoTo EH
    CopyMemory pHtml, VarPtr(htmlBytes(0)), UBound(htmlBytes) + 1
    GlobalUnlock hHtml
    If SetClipboardData(CF_HTML, hHtml) = 0 Then GoTo EH

    CloseClipboard
    SetClipboardTextAndHtml = True
    Exit Function
EH:
    On Error Resume Next
    CloseClipboard
    SetClipboardTextAndHtml = False
End Function


' ==========================================================================================
' JSON -> VBA (Dictionary/Array) � robust recursive parser
' Usage: v = JSON_ToObject(jsonStringOrArrayOfLines)
' Returns: Scripting.Dictionary (for objects), Variant() (for arrays), or scalar
' ==========================================================================================
Public Function JSON_ToObject(ByVal jsonOrText As Variant) As Variant
    Dim s As String
    If IsArray(jsonOrText) Then
        s = Join1D(jsonOrText, vbCrLf)
    Else
        s = AsString(jsonOrText)
    End If
    Dim p As Long: p = 1
    JSON_SkipWS s, p
    Dim result As Variant
    result = JSON_ParseValue(s, p)
    JSON_SkipWS s, p
    If p <= Len(s) Then
        Err.Raise 5, , "JSON parse error: trailing content at position " & p
    End If
    JSON_ToObject = result
End Function

'--- Parse dispatch
Private Function JSON_ParseValue(ByVal s As String, ByRef p As Long) As Variant
    JSON_SkipWS s, p
    If p > Len(s) Then Err.Raise 5, , "JSON parse error: unexpected end"
    Select Case mid$(s, p, 1)
        Case "{": JSON_ParseValue = JSON_ParseObject(s, p): Exit Function
        Case "[": JSON_ParseValue = JSON_ParseArray(s, p): Exit Function
        Case """": JSON_ParseValue = JSON_ParseString(s, p): Exit Function
        Case 2, "f": JSON_ParseValue = JSON_ParseBoolean(s, p): Exit Function
        Case "n": JSON_ParseValue = JSON_ParseNull(s, p): Exit Function
        Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            JSON_ParseValue = JSON_ParseNumber(s, p): Exit Function
        Case Else
            Err.Raise 5, , "JSON parse error: invalid token at position " & p
    End Select
End Function

Private Sub JSON_SkipWS(ByVal s As String, ByRef p As Long)
    Do While p <= Len(s)
        Select Case AscW(mid$(s, p, 1))
            Case 9, 10, 13, 32: p = p + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Function JSON_ParseObject(ByVal s As String, ByRef p As Long) As Variant
    Dim d As Object: Set d = NewDictionary(True)
    p = p + 1  ' skip '{'
    JSON_SkipWS s, p
    If p <= Len(s) And mid$(s, p, 1) = "}" Then
        p = p + 1
        Set JSON_ParseObject = d
        Exit Function
    End If
    Do
        JSON_SkipWS s, p
        If mid$(s, p, 1) <> """" Then Err.Raise 5, , "JSON object expected string key at position " & p
        Dim key As String: key = JSON_ParseString(s, p)
        JSON_SkipWS s, p
        If mid$(s, p, 1) <> ":" Then Err.Raise 5, , "JSON object expected ':' after key at position " & p
        p = p + 1
        Dim val As Variant: val = JSON_ParseValue(s, p)
        d(key) = val
        JSON_SkipWS s, p
        Dim ch As String: ch = mid$(s, p, 1)
        If ch = "}" Then
            p = p + 1: Set JSON_ParseObject = d: Exit Function
        ElseIf ch = "," Then
            p = p + 1
        Else
            Err.Raise 5, , "JSON object expected ',' or '}' at position " & p
        End If
    Loop
End Function

Private Function JSON_ParseArray(ByVal s As String, ByRef p As Long) As Variant
    p = p + 1 ' skip '['
    JSON_SkipWS s, p
    Dim tmp() As Variant, count As Long, cap As Long
    count = 0: cap = 8
    ReDim tmp(0 To cap - 1)
    If p <= Len(s) And mid$(s, p, 1) = "]" Then
        p = p + 1
        ReDim Preserve tmp(0 To -1) ' empty -> return empty array
        JSON_ParseArray = tmp
        Exit Function
    End If
    Do
        If count = cap Then
            cap = cap * 2
            ReDim Preserve tmp(0 To cap - 1)
        End If
        tmp(count) = JSON_ParseValue(s, p)
        count = count + 1
        JSON_SkipWS s, p
        Dim ch As String: ch = mid$(s, p, 1)
        If ch = "]" Then
            p = p + 1
            ReDim Preserve tmp(0 To count - 1)
            JSON_ParseArray = tmp
            Exit Function
        ElseIf ch = "," Then
            p = p + 1
        Else
            Err.Raise 5, , "JSON array expected ',' or ']' at position " & p
        End If
    Loop
End Function

Private Function JSON_ParseString(ByVal s As String, ByRef p As Long) As String
    Dim i As Long: i = p + 1 ' skip opening quote
    Dim sb As String
    Do While i <= Len(s)
        Dim ch As String: ch = mid$(s, i, 1)
        If ch = """" Then
            ' end
            p = i + 1
            JSON_ParseString = sb
            Exit Function
        ElseIf ch <> "\" Then
            sb = sb & ch
            i = i + 1
        Else
            ' escape
            i = i + 1
            If i > Len(s) Then Err.Raise 5, , "JSON string: invalid escape at position " & i
            ch = mid$(s, i, 1)
            Select Case ch
                Case """", "\", "/": sb = sb & ch
                Case "b": sb = sb & vbBack
                Case "f": sb = sb & vbFormFeed
                Case "n": sb = sb & vbLf
                Case "r": sb = sb & vbCr
                Case 2: sb = sb & vbTab
                Case "u"
                    Dim code As Integer
                    code = JSON_ReadHex4(s, i + 1)
                    i = i + 4
                    ' surrogate handling
                    If code >= &HD800 And code <= &HDBFF Then
                        ' high surrogate: expect \uXXXX next
                        If mid$(s, i + 1, 2) = "\u" Then
                            Dim low As Integer
                            low = JSON_ReadHex4(s, i + 3)
                            i = i + 6
                            Dim hi As Long, lo As Long, cp As Long
                            hi = code - &HD800
                            lo = low - &HDC00
                            cp = &H10000 + (hi * &H400) + lo
                            sb = sb & JSON_EncodeUTF16(cp)
                        Else
                            ' standalone surrogate
                            sb = sb & ChrW(code)
                        End If
                    Else
                        sb = sb & ChrW(code)
                    End If
                Case Else
                    Err.Raise 5, , "JSON string: invalid escape '\ " & ch & "' at position " & i
            End Select
            i = i + 1
        End If
    Loop
    Err.Raise 5, , "JSON string: missing closing quote at position " & p
End Function

Private Function JSON_ReadHex4(ByVal s As String, ByVal pos As Long) As Integer
    If pos + 3 > Len(s) Then Err.Raise 5, , "JSON \u escape truncated at position " & pos
    Dim t As String: t = mid$(s, pos, 4)
    Dim v As Long
    v = CLng("&H" & t)
    JSON_ReadHex4 = v
End Function

Private Function JSON_EncodeUTF16(ByVal codePoint As Long) As String
    ' Converts a Unicode code point into UTF-16 surrogate pair (for >= 0x10000).
    If codePoint < &H10000 Then
        JSON_EncodeUTF16 = ChrW(codePoint)
    Else
        Dim cp As Long: cp = codePoint - &H10000
        Dim hi As Integer: hi = &HD800 Or ((cp \ &H400) And &H3FF)
        Dim lo As Integer: lo = &HDC00 Or (cp And &H3FF)
        JSON_EncodeUTF16 = ChrW(hi) & ChrW(lo)
    End If
End Function

Private Function JSON_ParseNumber(ByVal s As String, ByRef p As Long) As Variant
    Dim i As Long: i = p
    Dim ch As String
    Do While i <= Len(s)
        ch = mid$(s, i, 1)
        If InStr(1, "-+0123456789.eE", ch, vbBinaryCompare) = 0 Then Exit Do
        i = i + 1
    Loop
    Dim token As String: token = mid$(s, p, i - p)
    p = i
    If InStr(1, token, ".", vbBinaryCompare) > 0 Or _
       InStr(1, token, "e", vbTextCompare) > 0 Or _
       InStr(1, token, "E", vbBinaryCompare) > 0 Then
        JSON_ParseNumber = CDbl(token)
    Else
        ' Keep as Long if within bounds else as Double
        On Error GoTo Large
        JSON_ParseNumber = CLng(token)
        Exit Function
Large:
        JSON_ParseNumber = CDbl(token)
    End If
End Function

Private Function JSON_ParseBoolean(ByVal s As String, ByRef p As Long) As Variant
    If mid$(s, p, 4) = "true" Then
        p = p + 4: JSON_ParseBoolean = True
    ElseIf mid$(s, p, 5) = "false" Then
        p = p + 5: JSON_ParseBoolean = False
    Else
        Err.Raise 5, , "JSON boolean invalid at position " & p
    End If
End Function

Private Function JSON_ParseNull(ByVal s As String, ByRef p As Long) As Variant
    If mid$(s, p, 4) = "null" Then
        p = p + 4
        JSON_ParseNull = Null
    Else
        Err.Raise 5, , "JSON null invalid at position " & p
    End If
End Function


' ==========================================================================================
' VBA -> JSON serializer (scalars/dates/dictionaries/arrays)
' Usage: s = JSON_Stringify(v, True)  ' pretty:=True
' ==========================================================================================
Public Function JSON_Stringify(ByVal v As Variant, _
                               Optional ByVal pretty As Boolean = False, _
                               Optional ByVal indent As String = "  ") As String
    JSON_Stringify = JSON_WriteValue(v, pretty, indent, 0)
End Function

Private Function JSON_WriteValue(ByVal v As Variant, ByVal pretty As Boolean, ByVal indent As String, ByVal level As Long) As String
    Dim t As String
    If IsObject(v) Then
        If TypeName(v) = "Dictionary" Or TypeName(v) = "Scripting.Dictionary" Then
            JSON_WriteValue = JSON_WriteObject(v, pretty, indent, level)
        ElseIf TypeName(v) = "Collection" Then
            JSON_WriteValue = JSON_WriteArray(CollectionToArray(v), pretty, indent, level)
        ElseIf TypeName(v) = "Range" Then
            JSON_WriteValue = JSON_WriteValue(v.Value2, pretty, indent, level)
        Else
            JSON_WriteValue = """" & JSON_EscapeString(CStr(v)) & """"
        End If
    ElseIf IsArray(v) Then
        JSON_WriteValue = JSON_WriteArray(v, pretty, indent, level)
    ElseIf IsNull(v) Then
        JSON_WriteValue = "null"
    ElseIf VarType(v) = vbDate Then
        ' Serialize Date as ISO 8601 string
        JSON_WriteValue = """" & JSON_EscapeString(Format$(v, "yyyy-mm-dd\Thh:nn:ss")) & """"
    ElseIf VarType(v) = vbBoolean Then
        JSON_WriteValue = IIf(v, "true", "false")
    ElseIf IsNumeric(v) Then
        JSON_WriteValue = CStr(v)
    Else
        JSON_WriteValue = """" & JSON_EscapeString(CStr(v)) & """"
    End If
End Function

Private Function JSON_WriteObject(ByVal d As Object, ByVal pretty As Boolean, ByVal indent As String, ByVal level As Long) As String
    Dim k As Variant
    Dim pieces() As String, i As Long, cnt As Long
    cnt = d.count
    ReDim pieces(0 To Application.Max(0, cnt - 1))
    i = 0
    For Each k In d.keys
        Dim kv As String
        kv = """" & JSON_EscapeString(CStr(k)) & """:" & IIf(pretty, " ", "")
        kv = kv & JSON_WriteValue(d(k), pretty, indent, level + 1)
        pieces(i) = kv
        i = i + 1
    Next
    If pretty Then
        JSON_WriteObject = "{" & vbCrLf & String$((level + 1) * Len(indent), " ") & _
                           VBA.Join(pieces, "," & vbCrLf & String$((level + 1) * Len(indent), " ")) & vbCrLf & _
                           String$(level * Len(indent), " ") & "}"
    Else
        JSON_WriteObject = "{" & VBA.Join(pieces, ",") & "}"
    End If
End Function

Private Function JSON_WriteArray(ByVal arr As Variant, ByVal pretty As Boolean, ByVal indent As String, ByVal level As Long) As String
    Dim l As Long, u As Long, i As Long
    If Not IsArray(arr) Then
        JSON_WriteArray = "[]": Exit Function
    End If
    On Error GoTo EH
    l = LBound(arr): u = UBound(arr)
    Dim pieces() As String
    ReDim pieces(0 To Application.Max(0, u - l))
    For i = l To u
        pieces(i - l) = JSON_WriteValue(arr(i), pretty, indent, level + 1)
    Next
    If pretty Then
        JSON_WriteArray = "[" & vbCrLf & String$((level + 1) * Len(indent), " ") & _
                          VBA.Join(pieces, "," & vbCrLf & String$((level + 1) * Len(indent), " ")) & vbCrLf & _
                          String$(level * Len(indent), " ") & "]"
    Else
        JSON_WriteArray = "[" & VBA.Join(pieces, ",") & "]"
    End If
    Exit Function
EH:
    JSON_WriteArray = "[]"
End Function

Private Function JSON_EscapeString(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCr, "\r")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbTab, "\t")
    JSON_EscapeString = t
End Function

Private Function CollectionToArray(ByVal c As Collection) As Variant
    Dim arr() As Variant, i As Long
    ReDim arr(0 To c.count - 1)
    For i = 1 To c.count
        arr(i - 1) = c(i)
    Next
    CollectionToArray = arr
End Function


' ==========================================================================================
' HOTKEYS simplified OnKey register/unregister
' Usage in Worksheet module:
'   Private Sub Worksheet_Activate(): Hotkeys_Register(Array(Array("^+g","MyMacro"), Array("%h","AnotherMacro"))) : End Sub
'   Private Sub Worksheet_Deactivate(): Hotkeys_Unregister(Array("^+g","%h")) : End Sub
' ==========================================================================================
' Registers a set of hotkeys: mapping = Array(Array("key","MacroName"), Array("key2","Macro2"))
' Keys overwrite application-level keys while active; call Unregister on Deactivate/BeforeClose.
Public Sub Hotkeys_Register(ByVal mapping As Variant)
    Dim i As Long
    For i = LBound(mapping) To UBound(mapping)
        Application.OnKey CStr(mapping(i)(0)), CStr(mapping(i)(1))
    Next
End Sub

' Unregisters keys: pass either same mapping array or a simple array of keys (Array("^+g","%h"))
Public Sub Hotkeys_Unregister(ByVal keysOrMapping As Variant)
    Dim i As Long, k As String
    If IsArray(keysOrMapping) Then
        For i = LBound(keysOrMapping) To UBound(keysOrMapping)
            If IsArray(keysOrMapping(i)) Then
                k = CStr(keysOrMapping(i)(0))
            Else
                k = CStr(keysOrMapping(i))
            End If
            Application.OnKey k
        Next
    End If
End Sub


' ==========================================================================================
' INTERSECTION WATCHER link ranges or trigger callback on edit
' To use: in the Worksheet's code module: call Intersect_Link inside Worksheet_Change(ByVal Target As Range)
' ==========================================================================================
' If Target intersects watchRange:
'   - If destRange provided, calculates destRange
'   - If callbackMacro provided, runs Application.Run callbackMacro, Target, watchRange, destRange
Public Function Intersect_Link(ByVal Target As Range, _
                          ByVal watchRange As Range, _
                          Optional ByVal destRange As Range, _
                          Optional ByVal callbackMacro As String = "") As Boolean
    On Error GoTo EH
    If intersect(Target, watchRange) Is Nothing Then
        Intersect_Link = False
        Exit Function
    End If
    Dim prevEvents As Boolean: prevEvents = Application.enableEvents
    Application.enableEvents = False
    If Not destRange Is Nothing Then destRange.Calculate
    If Len(callbackMacro) > 0 Then
        If Not destRange Is Nothing Then
            Application.Run callbackMacro, Target, watchRange, destRange
        Else
            On Error GoTo simpletry
            Application.Run callbackMacro, Target, watchRange
            GoTo cleanup
simpletry:
            Application.Run callbackMacro, Target
        End If
    End If
cleanup:
    Application.enableEvents = prevEvents
    Intersect_Link = True
    Exit Function
EH:
    Application.enableEvents = True
    Intersect_Link = False
    Err.Raise Err.Number, , "Intersect_Link: " & Err.Description
End Function


' ==========================================================================================
' ARRAY DEDUPLICATION order-preserving
' Usage: unique = Array_Unique(arr, False) ' case-insensitive
' ==========================================================================================
' Removes duplicates from a 1D array. preserveOrder:=True keeps first occurrence.
Public Function Array_Unique(ByVal arr As Variant, _
                             Optional ByVal caseSensitive As Boolean = True, _
                             Optional ByVal preserveOrder As Boolean = True) As Variant
    Dim d As Object: Set d = NewDictionary(caseSensitive)
    Dim out() As Variant, n As Long
    If Not IsArray(arr) Then
        ReDim out(0 To 0): out(0) = arr: Array_Unique = out: Exit Function
    End If
    Dim l As Long, u As Long, i As Long, k As String
    On Error GoTo EH
    l = LBound(arr): u = UBound(arr)
    ReDim out(0 To u - l)
    For i = l To u
        k = CStr(arr(i))
        If Not d.exists(k) Then
            d(k) = True
            out(n) = arr(i)
            n = n + 1
        End If
    Next
    ReDim Preserve out(0 To Application.Max(0, n - 1))
    Array_Unique = out
    Exit Function
EH:
    ' On error, return empty array
    ReDim out(0 To -1)
    Array_Unique = out
End Function


' ==========================================================================================
' RANGE ITERATORS ForEach (map) and Fold (reduce)
' ==========================================================================================
' Invokes a callback for each cell in a range. The callback should be a Public Function.
' - callbackName(cellArg, ParamArray extraArgs) As Variant
' - passCell: "Range" (pass the Range object), "Value" (pass cell.Value2), or "Text" (.Text)
' - includeHidden: include hidden rows/cols
' - includeEmpty: include empty/"" values
' Returns a 1D array of results.
Public Function Range_ForEach(ByVal rng As Range, _
                              ByVal callbackName As String, _
                              Optional ByVal passCell As String = "Value", _
                              Optional ByVal includeHidden As Boolean = False, _
                              Optional ByVal includeEmpty As Boolean = False, _
                              Optional ByVal extraArgs As Variant) As Variant
    Dim c As Range, val As Variant, vis As Boolean
    Dim out() As Variant, n As Long
    ReDim out(0 To rng.cells.count - 1)
    For Each c In rng.cells
        vis = True
        If Not includeHidden Then
            If c.EntireRow.Hidden Or c.EntireColumn.Hidden Then vis = False
        End If
        If vis Then
            Select Case UCase$(passCell)
                Case "RANGE": val = c
                Case "TEXT":  val = c.text
                Case Else:    val = c.Value2
            End Select
            If includeEmpty Or Len(AsString(val)) > 0 Then
                out(n) = Application.Run(callbackName, val, extraArgs)
                n = n + 1
            End If
        End If
    Next
    ReDim Preserve out(0 To Application.Max(0, n - 1))
    Range_ForEach = out
End Function

Public Function Range_Fold(ByVal rng As Range, _
                           ByVal folderName As String, _
                           ByVal initialState As Variant, _
                           Optional ByVal passCell As String = "Value", _
                           Optional ByVal includeHidden As Boolean = False, _
                           Optional ByVal includeEmpty As Boolean = False, _
                           Optional ByVal extraArgs As Variant) As Variant
    Dim state As Variant: state = initialState
    Dim c As Range, val As Variant, vis As Boolean
    For Each c In rng.cells
        vis = True
        If Not includeHidden Then
            If c.EntireRow.Hidden Or c.EntireColumn.Hidden Then vis = False
        End If
        If vis Then
            Select Case UCase$(passCell)
                Case "RANGE": val = c
                Case "TEXT":  val = c.text
                Case Else:    val = c.Value2
            End Select
            If includeEmpty Or Len(AsString(val)) > 0 Then
                state = Application.Run(folderName, val, state, extraArgs)
            End If
        End If
    Next
    Range_Fold = state
End Function


' ==========================================================================================
' Miscellaneous
' ==========================================================================================

Public Function AppState_Optimize(Optional ByVal calcMode As XlCalculation = xlCalculationManual, _
                                  Optional ByVal screenUpdating As Boolean = False, _
                                  Optional ByVal enableEvents As Boolean = False, _
                                  Optional ByVal displayAlerts As Boolean = False) As Object
    Dim prev As Object: Set prev = NewDictionary(True)
    prev("Calc") = Application.Calculation
    prev("Screen") = Application.screenUpdating
    prev("Events") = Application.enableEvents
    prev("Alerts") = Application.displayAlerts

    Application.Calculation = calcMode
    Application.screenUpdating = screenUpdating
    Application.enableEvents = enableEvents
    Application.displayAlerts = displayAlerts

    Set AppState_Optimize = prev
End Function

Public Sub AppState_Restore(ByVal state As Object)
    On Error Resume Next
    Application.Calculation = state("Calc")
    Application.screenUpdating = state("Screen")
    Application.enableEvents = state("Events")
    Application.displayAlerts = state("Alerts")
End Sub

Public Sub AppState_HardFix()
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Application.screenUpdating = True
    Application.enableEvents = True
    Application.displayAlerts = True
End Sub


'--- Range/Array conversions
Public Function Range_ToArray1D(ByVal rng As Range) As Variant
    Dim r As Long, c As Long, n As Long
    Dim arr() As Variant
    ReDim arr(0 To rng.cells.count - 1)
    n = 0
    For r = 1 To rng.rows.count
        For c = 1 To rng.Columns.count
            arr(n) = rng.cells(r, c).Value2
            n = n + 1
        Next
    Next
    Range_ToArray1D = arr
End Function

'--- CSV/TSV builders
Public Function Array_ToCSV(ByVal arr As Variant, Optional ByVal delimiter As String = ",") As String
    Dim l As Long, u As Long, i As Long
    If Not IsArray(arr) Then Array_ToCSV = AsString(arr): Exit Function
    l = LBound(arr): u = UBound(arr)
    Dim parts() As String: ReDim parts(0 To u - l)
    For i = l To u
        parts(i - l) = CSV_Escape(AsString(arr(i)), delimiter)
    Next
    Array_ToCSV = VBA.Join(parts, delimiter)
End Function

Private Function CSV_Escape(ByVal s As String, ByVal delim As String) As String
    Dim needs As Boolean
    needs = (InStr(1, s, """") > 0) Or (InStr(1, s, vbCr) > 0) Or (InStr(1, s, vbLf) > 0) Or (InStr(1, s, delim) > 0)
    If needs Then
        CSV_Escape = """" & Replace(s, """", """""") & """"
    Else
        CSV_Escape = s
    End If
End Function

'--- Simple wait
Public Sub DoEventsWait(ByVal milliseconds As Long)
    Dim t As Double: t = Timer + milliseconds / 1000#
    Do While Timer < t
        DoEvents
    Loop
End Sub

'--- StableSortArray
' Stable sort for 1D or 2D arrays. For 2D, provide keyColumn (1-based). Optional custom comparer macro.
' comparerMacro signature: CompareFn(a As Variant, b As Variant) As Long (-1/0/1); return <0 if a<b.
Public Function StableSortArray(ByVal arr As Variant, _
                                Optional ByVal keyColumn As Long = 0, _
                                Optional ByVal descending As Boolean = False, _
                                Optional ByVal caseSensitive As Boolean = True, _
                                Optional ByVal nullsLast As Boolean = True, _
                                Optional ByVal comparerMacro As String = "") As Variant
    If Not IsArray(arr) Then StableSortArray = arr: Exit Function

    Dim is2D As Boolean: is2D = Is2DArray(arr)
    Dim n As Long, i As Long
    If is2D Then
        Dim r As Long, c As Long
        Size2D arr, r, c
        If r <= 1 Then StableSortArray = arr: Exit Function
        If keyColumn < 1 Or keyColumn > c Then Err.Raise 5, , "StableSortArray: invalid keyColumn."
        n = r
    Else
        n = Size1D(arr)
        If n <= 1 Then StableSortArray = arr: Exit Function
    End If

    ' Build index array
    Dim idx() As Long: ReDim idx(1 To n)
    For i = 1 To n: idx(i) = i: Next

    ' Merge sort indices by comparator
    Call MergeSort_Indices(arr, idx, 1, n, is2D, keyColumn, descending, caseSensitive, nullsLast, comparerMacro)

    ' Rebuild sorted array
    If is2D Then
        Dim cols As Long: Size2D arr, n, cols ' reuse n as rows
        Dim baseR As Long, baseC As Long
        baseR = LBound(arr, 1): baseC = LBound(arr, 2)
        Dim out() As Variant: ReDim out(baseR To baseR + n - 1, baseC To baseC + cols - 1)
        Dim rr As Long, cc As Long
        For i = 1 To n
            rr = idx(i) + baseR - 1
            For cc = baseC To baseC + cols - 1
                out(baseR + i - 1, cc) = arr(rr, cc)
            Next
        Next
        StableSortArray = out
    Else
        Dim lo As Long, hi As Long: lo = LBound(arr): hi = UBound(arr)
        Dim out1() As Variant: ReDim out1(lo To lo + n - 1)
        For i = 1 To n
            out1(lo + i - 1) = arr(lo + idx(i) - 1)
        Next
        StableSortArray = out1
    End If
End Function

' --- Stable merge sort of indices using a comparator
Private Sub MergeSort_Indices(ByVal arr As Variant, ByRef idx() As Long, ByVal lo As Long, ByVal hi As Long, _
                              ByVal is2D As Boolean, ByVal keyCol As Long, ByVal desc As Boolean, _
                              ByVal caseSensitive As Boolean, ByVal nullsLast As Boolean, ByVal cmpMacro As String)
    If lo >= hi Then Exit Sub
    Dim mid As Long: mid = (lo + hi) \ 2
    MergeSort_Indices arr, idx, lo, mid, is2D, keyCol, desc, caseSensitive, nullsLast, cmpMacro
    MergeSort_Indices arr, idx, mid + 1, hi, is2D, keyCol, desc, caseSensitive, nullsLast, cmpMacro

    Dim tmp() As Long: ReDim tmp(lo To hi)
    Dim i As Long, j As Long, k As Long
    i = lo: j = mid + 1: k = lo

    Do While i <= mid And j <= hi
        If CompareItems(arr, idx(i), idx(j), is2D, keyCol, desc, caseSensitive, nullsLast, cmpMacro) <= 0 Then
            tmp(k) = idx(i): i = i + 1
        Else
            tmp(k) = idx(j): j = j + 1
        End If
        k = k + 1
    Loop
    Do While i <= mid: tmp(k) = idx(i): i = i + 1: k = k + 1: Loop
    Do While j <= hi: tmp(k) = idx(j): j = j + 1: k = k + 1: Loop

    For k = lo To hi: idx(k) = tmp(k): Next
End Sub

Private Function CompareItems(ByVal arr As Variant, ByVal i As Long, ByVal j As Long, _
                              ByVal is2D As Boolean, ByVal keyCol As Long, ByVal desc As Boolean, _
                              ByVal caseSensitive As Boolean, ByVal nullsLast As Boolean, _
                              ByVal cmpMacro As String) As Long
    Dim a As Variant, b As Variant
    If is2D Then
        Dim br As Long: br = LBound(arr, 1): Dim bc As Long: bc = LBound(arr, 2)
        a = arr(br + i - 1, bc + keyCol - 1)
        b = arr(br + j - 1, bc + keyCol - 1)
    Else
        Dim lo As Long: lo = LBound(arr)
        a = arr(lo + i - 1): b = arr(lo + j - 1)
    End If

    Dim r As Long
    If Len(cmpMacro) > 0 Then
        r = CLng(Application.Run(cmpMacro, a, b))
    Else
        r = DefaultCompare(a, b, caseSensitive, nullsLast)
    End If
    If desc Then r = -r
    CompareItems = r
End Function

' Default value comparator (numbers vs text; Null/Empty order; case sensitivity).
Private Function DefaultCompare(ByVal a As Variant, ByVal b As Variant, ByVal caseSensitive As Boolean, ByVal nullsLast As Boolean) As Long
    Dim aEmpty As Boolean, bEmpty As Boolean
    aEmpty = (IsNull(a) Or IsEmpty(a) Or AsString(a) = "")
    bEmpty = (IsNull(b) Or IsEmpty(b) Or AsString(b) = "")
    If aEmpty Or bEmpty Then
        If aEmpty And bEmpty Then DefaultCompare = 0: Exit Function
        If nullsLast Then DefaultCompare = IIf(aEmpty, 1, -1) Else DefaultCompare = IIf(aEmpty, -1, 1)
        Exit Function
    End If

    If IsNumeric(a) And IsNumeric(b) Then
        If CDbl(a) < CDbl(b) Then
            DefaultCompare = -1
        ElseIf CDbl(a) > CDbl(b) Then
            DefaultCompare = 1
        Else
            DefaultCompare = 0
        End If
    Else
        Dim sa As String, sb As String
        sa = CStr(a): sb = CStr(b)
        If Not caseSensitive Then sa = LCase$(sa): sb = LCase$(sb)
        If sa < sb Then
            DefaultCompare = -1
        ElseIf sa > sb Then
            DefaultCompare = 1
        Else
            DefaultCompare = 0
        End If
    End If
End Function

'--- MultiKeySort2D
' Stable multi-key sort by columns. keys: array of 1-based column indices; orders: "ASC"/"DESC" or Boolean False/True for descending.
Public Function MultiKeySort2D(ByVal arr2D As Variant, _
                               ByVal keys As Variant, _
                               Optional ByVal orders As Variant = Empty, _
                               Optional ByVal caseSensitive As Boolean = True, _
                               Optional ByVal nullsLast As Boolean = True) As Variant
    If Not Is2DArray(arr2D) Then Err.Raise 5, , "MultiKeySort2D: input must be 2D array."
    Dim i As Long, k As Long, desc As Boolean
    Dim last As Long: last = UBound(keys)
    For i = last To LBound(keys) Step -1
        k = CLng(keys(i))
        If IsEmpty(orders) Then
            desc = False
        ElseIf IsArray(orders) Then
            desc = IsDesc(orders, i)
        ElseIf VarType(orders) = vbString Then
            desc = (UCase$(orders) = "DESC")
        Else
            desc = CBool(orders)
        End If
        arr2D = StableSortArray(arr2D, keyColumn:=k, descending:=desc, caseSensitive:=caseSensitive, nullsLast:=nullsLast)
    Next
    MultiKeySort2D = arr2D
End Function

Private Function IsDesc(ByVal orders As Variant, ByVal idx As Long) As Boolean
    On Error GoTo E
    If VarType(orders(idx)) = vbString Then
        IsDesc = (UCase$(orders(idx)) = "DESC")
    Else
        IsDesc = CBool(orders(idx))
    End If
    Exit Function
E:
    IsDesc = False
End Function

'--- GroupByAggregate2D
' Groups by key columns and computes aggregates over value columns.
' aggSpecs: Array of Array(columnIndex, op) where op in {"sum","avg","min","max","count"}.
' firstRowHasHeaders controls header handling; includeHeaders toggles header row in output.
Public Function GroupByAggregate2D(ByVal arr2D As Variant, _
                                   ByVal keyCols As Variant, _
                                   ByVal aggSpecs As Variant, _
                                   Optional ByVal firstRowHasHeaders As Boolean = True, _
                                   Optional ByVal includeHeaders As Boolean = True) As Variant
    If Not Is2DArray(arr2D) Then Err.Raise 5, , "GroupByAggregate2D: input must be 2D."
    Dim rows As Long, cols As Long: Size2D arr2D, rows, cols
    If rows = 0 Then GroupByAggregate2D = arr2D: Exit Function

    Dim startRow As Long: startRow = LBound(arr2D, 1) + IIf(firstRowHasHeaders, 1, 0)
    Dim baseR As Long: baseR = LBound(arr2D, 1)
    Dim baseC As Long: baseC = LBound(arr2D, 2)

    ' Build groups
    Dim g As Object: Set g = NewDictionary(True)
    Dim r As Long, i As Long, keyParts() As Variant, k As String
    For r = startRow To baseR + rows - 1
        ReDim keyParts(LBound(keyCols) To UBound(keyCols))
        For i = LBound(keyCols) To UBound(keyCols)
            keyParts(i) = arr2D(r, baseC + CLng(keyCols(i)) - 1)
        Next
        k = KeyOfParts(keyParts)
        If Not g.exists(k) Then
            Dim state As Object: Set state = NewDictionary(True)
            state("keys") = keyParts
            Dim aIdx As Long
            For aIdx = LBound(aggSpecs) To UBound(aggSpecs)
                Dim colIdx As Long: colIdx = CLng(aggSpecs(aIdx)(0))
                Dim op As String: op = LCase$(CStr(aggSpecs(aIdx)(1)))
                Dim s As Object: Set s = NewDictionary(True)
                s("op") = op
                s("sum") = 0#
                s("count") = 0&
                s("minSet") = False
                s("min") = 0#
                s("maxSet") = False
                s("max") = 0#
                s("col") = colIdx
                state(CStr(aIdx)) = s
            Next
            Set g(k) = state
        End If
        Dim st As Object: Set st = g(k)
        For i = LBound(aggSpecs) To UBound(aggSpecs)
            Dim col As Long: col = CLng(aggSpecs(i)(0))
            Dim v As Variant: v = arr2D(r, baseC + col - 1)
            Dim ss As Object: Set ss = st(CStr(i))
            Dim isNum As Boolean: isNum = IsNumeric(v)
            If LCase$(ss("op")) = "count" Then
                If Len(AsString(v)) > 0 Then ss("count") = ss("count") + 1
            ElseIf isNum Then
                ss("sum") = ss("sum") + CDbl(v)
                ss("count") = ss("count") + 1
                If Not ss("minSet") Or CDbl(v) < ss("min") Then ss("min") = CDbl(v): ss("minSet") = True
                If Not ss("maxSet") Or CDbl(v) > ss("max") Then ss("max") = CDbl(v): ss("maxSet") = True
            End If
        Next
    Next

    Dim outCols As Long: outCols = (UBound(keyCols) - LBound(keyCols) + 1) + (UBound(aggSpecs) - LBound(aggSpecs) + 1)
    Dim outRows As Long: outRows = g.count + IIf(includeHeaders, 1, 0)
    Dim out() As Variant: ReDim out(1 To outRows, 1 To outCols)

    Dim rowPtr As Long: rowPtr = 1
    If includeHeaders Then
        ' headers
        Dim hc As Long: hc = 1
        If firstRowHasHeaders Then
            For i = LBound(keyCols) To UBound(keyCols)
                out(1, hc) = arr2D(baseR, baseC + CLng(keyCols(i)) - 1): hc = hc + 1
            Next
            For i = LBound(aggSpecs) To UBound(aggSpecs)
                Dim opName As String: opName = LCase$(CStr(aggSpecs(i)(1)))
                Dim srcCol As Long: srcCol = CLng(aggSpecs(i)(0))
                out(1, hc) = opName & "(" & AsString(arr2D(baseR, baseC + srcCol - 1)) & ")"
                hc = hc + 1
            Next
        Else
            For i = LBound(keyCols) To UBound(keyCols)
                out(1, hc) = "Key" & i: hc = hc + 1
            Next
            For i = LBound(aggSpecs) To UBound(aggSpecs)
                out(1, hc) = LCase$(CStr(aggSpecs(i)(1))) & "_C" & CStr(aggSpecs(i)(0))
                hc = hc + 1
            Next
        End If
        rowPtr = 2
    End If

    Dim kKey As Variant, colPtr As Long
    For Each kKey In g.keys
        Dim st2 As Object: Set st2 = g(kKey)
        colPtr = 1
        Dim kp As Variant: kp = st2("keys")
        For i = LBound(kp) To UBound(kp)
            out(rowPtr, colPtr) = kp(i): colPtr = colPtr + 1
        Next
        Dim agg As Object
        For i = LBound(aggSpecs) To UBound(aggSpecs)
            Set agg = st2(CStr(i))
            Select Case LCase$(agg("op"))
                Case "sum": out(rowPtr, colPtr) = agg("sum")
                Case "avg": If agg("count") > 0 Then out(rowPtr, colPtr) = agg("sum") / agg("count") Else out(rowPtr, colPtr) = Null
                Case "min": If agg("minSet") Then out(rowPtr, colPtr) = agg("min") Else out(rowPtr, colPtr) = Null
                Case "max": If agg("maxSet") Then out(rowPtr, colPtr) = agg("max") Else out(rowPtr, colPtr) = Null
                Case "count": out(rowPtr, colPtr) = agg("count")
                Case Else: out(rowPtr, colPtr) = Null
            End Select
            colPtr = colPtr + 1
        Next
        rowPtr = rowPtr + 1
    Next

    GroupByAggregate2D = out
End Function

'--- Pivot2D
' Build a cross-tab (pivot) from normalized rows.
' rowKeyCols: array of 1-based row key columns; colKeyCol: column key index; valueCol: numeric value column.
' aggregator: "sum","count","min","max","avg"; missingValue used where no data.
Public Function Pivot2D(ByVal arr2D As Variant, _
                        ByVal rowKeyCols As Variant, _
                        ByVal colKeyCol As Long, _
                        ByVal valueCol As Long, _
                        Optional ByVal aggregator As String = "sum", _
                        Optional ByVal firstRowHasHeaders As Boolean = True, _
                        Optional ByVal includeHeaders As Boolean = True, _
                        Optional ByVal missingValue As Variant = 0) As Variant
    If Not Is2DArray(arr2D) Then Err.Raise 5, , "Pivot2D: input must be 2D."
    Dim rows As Long, cols As Long: Size2D arr2D, rows, cols
    Dim baseR As Long: baseR = LBound(arr2D, 1)
    Dim baseC As Long: baseC = LBound(arr2D, 2)
    Dim startRow As Long: startRow = baseR + IIf(firstRowHasHeaders, 1, 0)

    Dim rowMap As Object: Set rowMap = NewDictionary(True)
    Dim colMap As Object: Set colMap = NewDictionary(True)
    Dim data As Object: Set data = NewDictionary(True)

    Dim r As Long, i As Long, rowKeyParts() As Variant, rowKey As String, colKey As String
    For r = startRow To baseR + rows - 1
        ReDim rowKeyParts(LBound(rowKeyCols) To UBound(rowKeyCols))
        For i = LBound(rowKeyCols) To UBound(rowKeyCols)
            rowKeyParts(i) = arr2D(r, baseC + CLng(rowKeyCols(i)) - 1)
        Next
        rowKey = KeyOfParts(rowKeyParts)
        colKey = AsString(arr2D(r, baseC + colKeyCol - 1))

        If Not rowMap.exists(rowKey) Then rowMap(rowKey) = rowMap.count + 1
        If Not colMap.exists(colKey) Then colMap(colKey) = colMap.count + 1

        Dim key As String: key = rowKey & "|" & colKey
        Dim state As Object
        If data.exists(key) Then
            Set state = data(key)
        Else
            Set state = NewDictionary(True)
            state("sum") = 0#: state("count") = 0&: state("minSet") = False: state("min") = 0#: state("maxSet") = False: state("max") = 0#
            data(key) = state
        End If
        Dim v As Variant: v = arr2D(r, baseC + valueCol - 1)
        If LCase$(aggregator) = "count" Then
            If Len(AsString(v)) > 0 Then state("count") = state("count") + 1
        ElseIf IsNumeric(v) Then
            Dim d As Double: d = CDbl(v)
            state("sum") = state("sum") + d
            state("count") = state("count") + 1
            If Not state("minSet") Or d < state("min") Then state("min") = d: state("minSet") = True
            If Not state("maxSet") Or d > state("max") Then state("max") = d: state("maxSet") = True
        End If
    Next

    ' Output dimensions
    Dim outRows As Long: outRows = rowMap.count + IIf(includeHeaders, 1, 0)
    Dim outCols As Long: outCols = (UBound(rowKeyCols) - LBound(rowKeyCols) + 1) + colMap.count
    Dim out() As Variant: ReDim out(1 To outRows, 1 To outCols)

    Dim rowKeys() As String: rowKeys = rowMap.keys
    Dim colKeys() As String: colKeys = colMap.keys

    ' Headers
    Dim c As Long
    If includeHeaders Then
        If firstRowHasHeaders Then
            For i = LBound(rowKeyCols) To UBound(rowKeyCols)
                out(1, i - LBound(rowKeyCols) + 1) = arr2D(baseR, baseC + CLng(rowKeyCols(i)) - 1)
            Next
            For c = 1 To colMap.count
                out(1, (UBound(rowKeyCols) - LBound(rowKeyCols) + 1) + c) = colKeys(c - 1)
            Next
        Else
            For i = LBound(rowKeyCols) To UBound(rowKeyCols)
                out(1, i - LBound(rowKeyCols) + 1) = "Key" & i
            Next
            For c = 1 To colMap.count
                out(1, (UBound(rowKeyCols) - LBound(rowKeyCols) + 1) + c) = "Col" & c
            Next
        End If
    End If

    ' Body
    Dim ri As Long, ci As Long, outRow As Long, outCol As Long
    For ri = 0 To rowMap.count - 1
        Dim rowKeyName As String: rowKeyName = rowKeys(ri)
        Dim rowKeyVals() As String: rowKeyVals = Split(rowKeyName, Chr$(30))
        outRow = ri + 1 + IIf(includeHeaders, 1, 0)
        For i = LBound(rowKeyVals) To UBound(rowKeyVals)
            out(outRow, i + 1) = rowKeyVals(i)
        Next
        For ci = 0 To colMap.count - 1
            outCol = (UBound(rowKeyCols) - LBound(rowKeyCols) + 2) + ci - 1 + 1
            Dim k2 As String: k2 = rowKeyName & "|" & colKeys(ci)
            If data.exists(k2) Then
                Dim st As Object: Set st = data(k2)
                Select Case LCase$(aggregator)
                    Case "sum": out(outRow, outCol) = st("sum")
                    Case "avg": If st("count") > 0 Then out(outRow, outCol) = st("sum") / st("count") Else out(outRow, outCol) = missingValue
                    Case "min": If st("minSet") Then out(outRow, outCol) = st("min") Else out(outRow, outCol) = missingValue
                    Case "max": If st("maxSet") Then out(outRow, outCol) = st("max") Else out(outRow, outCol) = missingValue
                    Case "count": out(outRow, outCol) = st("count")
                    Case Else: out(outRow, outCol) = missingValue
                End Select
            Else
                out(outRow, outCol) = missingValue
            End If
        Next
    Next

    Pivot2D = out
End Function

'--- Unpivot2D
' Unpivots a cross-tab: first idColsCount columns are identifiers; remaining columns become (Variable, Value).
Public Function Unpivot2D(ByVal arr2D As Variant, _
                          ByVal idColsCount As Long, _
                          Optional ByVal firstRowHasHeaders As Boolean = True, _
                          Optional ByVal variableHeader As String = "Variable", _
                          Optional ByVal valueHeader As String = "Value") As Variant
    If Not Is2DArray(arr2D) Then Err.Raise 5, , "Unpivot2D: input must be 2D."
    Dim rows As Long, cols As Long: Size2D arr2D, rows, cols
    Dim baseR As Long: baseR = LBound(arr2D, 1)
    Dim baseC As Long: baseC = LBound(arr2D, 2)
    Dim startRow As Long: startRow = baseR + IIf(firstRowHasHeaders, 1, 0)

    Dim valueCols As Long: valueCols = cols - idColsCount
    If valueCols <= 0 Then Err.Raise 5, , "Unpivot2D: idColsCount >= number of columns."

    Dim outRows As Long: outRows = (rows - IIf(firstRowHasHeaders, 1, 0)) * valueCols + IIf(firstRowHasHeaders, 1, 0)
    Dim outCols As Long: outCols = idColsCount + 2
    Dim out() As Variant: ReDim out(1 To outRows, 1 To outCols)

    Dim r As Long, c As Long, orow As Long: orow = 1

    If firstRowHasHeaders Then
        For c = 1 To idColsCount
            out(1, c) = arr2D(baseR, baseC + c - 1)
        Next
        out(1, idColsCount + 1) = variableHeader
        out(1, idColsCount + 2) = valueHeader
        orow = 2
    End If

    For r = startRow To baseR + rows - 1
        For c = idColsCount + 1 To cols
            Dim oc As Long
            For oc = 1 To idColsCount
                out(orow, oc) = arr2D(r, baseC + oc - 1)
            Next
            out(orow, idColsCount + 1) = IIf(firstRowHasHeaders, arr2D(baseR, baseC + c - 1), "Col" & c)
            out(orow, idColsCount + 2) = arr2D(r, baseC + c - 1)
            orow = orow + 1
        Next
    Next

    Unpivot2D = out
End Function

'--- LeftJoin2D / InnerJoin2D
Public Function LeftJoin2D(ByVal leftArr As Variant, ByVal rightArr As Variant, _
                           ByVal leftKeyCols As Variant, ByVal rightKeyCols As Variant, _
                           Optional ByVal firstRowHasHeaders As Boolean = True, _
                           Optional ByVal includeHeaders As Boolean = True, _
                           Optional ByVal caseSensitive As Boolean = False) As Variant
    LeftJoin2D = Join2D(leftArr, rightArr, leftKeyCols, rightKeyCols, "LEFT", firstRowHasHeaders, includeHeaders, caseSensitive)
End Function

Public Function InnerJoin2D(ByVal leftArr As Variant, ByVal rightArr As Variant, _
                            ByVal leftKeyCols As Variant, ByVal rightKeyCols As Variant, _
                            Optional ByVal firstRowHasHeaders As Boolean = True, _
                            Optional ByVal includeHeaders As Boolean = True, _
                            Optional ByVal caseSensitive As Boolean = False) As Variant
    InnerJoin2D = Join2D(leftArr, rightArr, leftKeyCols, rightKeyCols, "INNER", firstRowHasHeaders, includeHeaders, caseSensitive)
End Function

Private Function Join2D(ByVal leftArr As Variant, ByVal rightArr As Variant, _
                        ByVal leftKeyCols As Variant, ByVal rightKeyCols As Variant, _
                        ByVal joinType As String, _
                        ByVal firstRowHasHeaders As Boolean, ByVal includeHeaders As Boolean, _
                        ByVal caseSensitive As Boolean) As Variant
    If Not Is2DArray(leftArr) Or Not Is2DArray(rightArr) Then Err.Raise 5, , "Join2D: inputs must be 2D."

    Dim lr As Long, lc As Long, rr As Long, rc As Long
    Size2D leftArr, lr, lc: Size2D rightArr, rr, rc
    Dim lbr As Long, lbc As Long, rbr As Long, rbc As Long
    lbr = LBound(leftArr, 1): lbc = LBound(leftArr, 2)
    rbr = LBound(rightArr, 1): rbc = LBound(rightArr, 2)

    Dim startL As Long: startL = lbr + IIf(firstRowHasHeaders, 1, 0)
    Dim startR As Long: startR = rbr + IIf(firstRowHasHeaders, 1, 0)

    ' Build right lookup (key -> array of row indexes)
    Dim dict As Object: Set dict = NewDictionary(caseSensitive)
    Dim r As Long
    For r = startR To rbr + rr - 1
        Dim k As String: k = JoinKey(rightArr, r, rbc, rightKeyCols)
        If dict.exists(k) Then
            dict(k) = AB_Append(dict(k), r)
        Else
            Dim v As Variant
            v = Empty
            v = AB_Append(v, r)
            dict(k) = v
        End If
    Next

    ' Estimate rows
    Dim maxRows As Long
    If UCase$(joinType) = "INNER" Then
        maxRows = lr ' upper bound
    Else
        maxRows = lr + dict.count ' rough upper bound
    End If

    Dim out() As Variant, rowPtr As Long, outCols As Long
    outCols = lc + (rc - (UBound(rightKeyCols) - LBound(rightKeyCols) + 1))
    ReDim out(1 To IIf(includeHeaders, 1, 0) + maxRows, 1 To outCols)

    ' Headers
    rowPtr = 1
    Dim c As Long, oc As Long
    If includeHeaders Then
        ' left headers
        If firstRowHasHeaders Then
            For c = 1 To lc: out(1, c) = leftArr(lbr, lbc + c - 1): Next
            ' right headers (excluding key columns)
            oc = lc + 1
            For c = 1 To rc
                If Not InArray(c, rightKeyCols) Then
                    Dim nameR As String: nameR = AsString(rightArr(rbr, rbc + c - 1))
                    If Len(nameR) = 0 Then nameR = "R" & c
                    out(1, oc) = nameR
                    oc = oc + 1
                End If
            Next
        Else
            For c = 1 To outCols: out(1, c) = "C" & c: Next
        End If
        rowPtr = 2
    End If

    ' Rows
    Dim l As Long
    For l = startL To lbr + lr - 1
        Dim lk As String: lk = JoinKey(leftArr, l, lbc, leftKeyCols)
        If dict.exists(lk) Then
            Dim matches As Variant: matches = dict(lk)
            Dim mi As Long
            For mi = LBound(matches) To UBound(matches)
                If rowPtr > UBound(out, 1) Then ReDim Preserve out(1 To rowPtr * 2, 1 To outCols)
                ' copy left row
                For c = 1 To lc: out(rowPtr, c) = leftArr(l, lbc + c - 1): Next
                ' copy right non-key cols
                oc = lc + 1
                For c = 1 To rc
                    If Not InArray(c, rightKeyCols) Then
                        out(rowPtr, oc) = rightArr(matches(mi), rbc + c - 1)
                        oc = oc + 1
                    End If
                Next
                rowPtr = rowPtr + 1
            Next
        Else
            If UCase$(joinType) = "LEFT" Then
                If rowPtr > UBound(out, 1) Then ReDim Preserve out(1 To rowPtr * 2, 1 To outCols)
                For c = 1 To lc: out(rowPtr, c) = leftArr(l, lbc + c - 1): Next
                oc = lc + 1
                Do While oc <= outCols: out(rowPtr, oc) = Empty: oc = oc + 1: Loop
                rowPtr = rowPtr + 1
            End If
        End If
    Next

    ' Trim rows
    If rowPtr = 1 Then
        ReDim out(1 To IIf(includeHeaders, 1, 0), 1 To outCols)
    Else
        ReDim Preserve out(1 To rowPtr - 1, 1 To outCols)
    End If
    Join2D = out
End Function

Private Function JoinKey(ByVal arr2D As Variant, ByVal r As Long, ByVal baseC As Long, ByVal colsIdx As Variant) As String
    Dim parts() As Variant, i As Long
    ReDim parts(LBound(colsIdx) To UBound(colsIdx))
    For i = LBound(colsIdx) To UBound(colsIdx)
        parts(i) = arr2D(r, baseC + CLng(colsIdx(i)) - 1)
    Next
    JoinKey = KeyOfParts(parts)
End Function

Private Function InArray(ByVal val As Variant, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then InArray = True: Exit Function
    Next
    InArray = False
End Function

'--- FlattenArray
' Flattens nested arrays/collections into a single 1D Variant array. maxDepth: -1 = infinite.
Public Function Array_Flatten(ByVal myInput As Variant, Optional ByVal maxDepth As Long = -1) As Variant
    Dim out As Variant: out = Empty
    Call Flatten_Recur(myInput, out, maxDepth, 0)
    Array_Flatten = out
End Function

Private Sub Flatten_Recur(ByVal v As Variant, ByRef out As Variant, ByVal maxDepth As Long, ByVal depth As Long)
    If maxDepth >= 0 And depth > maxDepth Then
        out = AB_Append(out, v)
        Exit Sub
    End If
    If IsArray(v) Then
        Dim i As Long
        For i = LBound(v) To UBound(v)
            Flatten_Recur v(i), out, maxDepth, depth + 1
        Next
    ElseIf IsObject(v) And TypeName(v) = "Collection" Then
        Dim j As Long
        For j = 1 To v.count
            Flatten_Recur v(j), out, maxDepth, depth + 1
        Next
    Else
        out = AB_Append(out, v)
    End If
End Sub

'--- ZipArrays
' Zips multiple 1D arrays element-wise into a 2D matrix. align: "shortest" or "longest"; fillValue for padding.
Public Function Array_Zip(ByVal arrays As Variant, _
                          Optional ByVal align As String = "shortest", _
                          Optional ByVal fillValue As Variant = Empty) As Variant
    If Not IsArray(arrays) Then Err.Raise 5, , "Array_Zip: arrays must be an array of arrays."
    Dim i As Long, n As Long: n = UBound(arrays) - LBound(arrays) + 1
    Dim lens() As Long: ReDim lens(1 To n)
    For i = LBound(arrays) To UBound(arrays)
        If Not IsArray(arrays(i)) Then Err.Raise 5, , "Array_Zip: element " & i & " is not array."
        lens(i - LBound(arrays) + 1) = Size1D(arrays(i))
    Next
    Dim rows As Long
    If LCase$(align) = "longest" Then
        rows = MaxLong(lens)
    Else
        rows = MinLong(lens)
    End If
    If rows < 0 Then rows = 0

    Dim out() As Variant: ReDim out(1 To rows, 1 To n)
    Dim r As Long, c As Long
    For r = 1 To rows
        For c = 1 To n
            Dim a As Variant: a = arrays(LBound(arrays) + c - 1)
            If r - 1 <= UBound(a) - LBound(a) Then
                out(r, c) = a(LBound(a) + r - 1)
            Else
                out(r, c) = fillValue
            End If
        Next
    Next
    Array_Zip = out
End Function

Private Function MaxLong(ByVal arr As Variant) As Long
    Dim i As Long, m As Long: m = -2147483648#
    For i = LBound(arr) To UBound(arr): If arr(i) > m Then m = arr(i)
    Next: MaxLong = m
End Function
Private Function MinLong(ByVal arr As Variant) As Long
    Dim i As Long, m As Long: m = 2147483647
    For i = LBound(arr) To UBound(arr): If arr(i) < m Then m = arr(i)
    Next: MinLong = m
End Function

'--- RegexMatch
' Returns a 2D table: [FullMatch, Group1, Group2, ...]; sets columns to max group count found; empty if no matches.
Public Function Regex_Matches(ByVal text As String, ByVal pattern As String, _
                              Optional ByVal caseSensitive As Boolean = True, _
                              Optional ByVal multiLine As Boolean = False) As Variant
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.Global = True
    re.multiLine = multiLine
    re.IgnoreCase = Not caseSensitive

    Dim matches As Object: Set matches = re.Execute(text)
    If matches.count = 0 Then
        Dim emptyArr() As Variant: ReDim emptyArr(1 To 0, 1 To 0): Regex_Matches = emptyArr: Exit Function
    End If

    Dim m As Object, maxGroups As Long: maxGroups = 0
    For Each m In matches
        If m.SubMatches.count > maxGroups Then maxGroups = m.SubMatches.count
    Next

    Dim out() As Variant: ReDim out(1 To matches.count, 1 To 1 + maxGroups)
    Dim i As Long: i = 1
    For Each m In matches
        out(i, 1) = m.value
        Dim g As Long
        For g = 0 To maxGroups - 1
            If g <= m.SubMatches.count - 1 Then out(i, 2 + g) = m.SubMatches(g) Else out(i, 2 + g) = Empty
        Next
        i = i + 1
    Next
    Regex_Matches = out
End Function

'--- RegexReplace
' Simple replace with pattern/replacement (supports $1, $2, ...).
Public Function Regex_Replace(ByVal text As String, ByVal pattern As String, ByVal replacement As String, _
                              Optional ByVal caseSensitive As Boolean = True, _
                              Optional ByVal multiLine As Boolean = False, _
                              Optional ByVal globalReplace As Boolean = True) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = Not caseSensitive
    re.multiLine = multiLine
    re.Global = globalReplace
    Regex_Replace = re.Replace(text, replacement)
End Function

' Replacement via callback macro. evaluatorMacro(cellText, fullMatch As String, groups() As Variant, matchIndex As Long) -> String
Public Function Regex_ReplaceEval(ByVal text As String, ByVal pattern As String, ByVal evaluatorMacro As String, _
                                  Optional ByVal caseSensitive As Boolean = True, _
                                  Optional ByVal multiLine As Boolean = False) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = Not caseSensitive
    re.multiLine = multiLine
    re.Global = True

    Dim matches As Object: Set matches = re.Execute(text)
    If matches.count = 0 Then Regex_ReplaceEval = text: Exit Function

    Dim sb As String, pos As Long, m As Object, idx As Long
    pos = 1: idx = 0
    For Each m In matches
        sb = sb & mid$(text, pos, m.FirstIndex - pos + 1)
        Dim groups() As Variant: ReDim groups(0 To m.SubMatches.count - 1)
        Dim i As Long
        For i = 0 To m.SubMatches.count - 1: groups(i) = m.SubMatches(i): Next
        Dim rep As String: rep = CStr(Application.Run(evaluatorMacro, text, m.value, groups, idx))
        sb = sb & rep
        pos = m.FirstIndex + m.Length + 1
        idx = idx + 1
    Next
    sb = sb & mid$(text, pos)
    Regex_ReplaceEval = sb
End Function

'--- HeaderIndexMap
' Maps requested header names to column indices (1-based). Supports exact/normalized/fuzzy.
' headersToFind: array of strings. aliases: optional Dictionary alias->canonical.
Public Function Header_IndexMap(ByVal arr2D As Variant, ByVal headersToFind As Variant, _
                                Optional ByVal firstRow As Long = 1, _
                                Optional ByVal caseSensitive As Boolean = False, _
                                Optional ByVal allowFuzzy As Boolean = True, _
                                Optional ByVal maxDistance As Long = 2, _
                                Optional ByVal aliases As Variant) As Object
    If Not Is2DArray(arr2D) Then Err.Raise 5, , "Header_IndexMap: input must be 2D."
    Dim cols As Long, rows As Long: Size2D arr2D, rows, cols
    Dim baseC As Long: baseC = LBound(arr2D, 2)
    Dim rowIdx As Long: rowIdx = LBound(arr2D, 1) + firstRow - 1

    Dim normMap As Object: Set normMap = NewDictionary(False)
    Dim c As Long
    For c = 1 To cols
        Dim raw As String: raw = AsString(arr2D(rowIdx, baseC + c - 1))
        If Len(raw) > 0 Then
            Dim nm As String: nm = NormalizeHeader(raw, caseSensitive:=caseSensitive)
            If Not normMap.exists(nm) Then normMap(nm) = c
        End If
    Next

    Dim aliasMap As Object: Set aliasMap = NewDictionary(False)
    If IsObject(aliases) Then
        Dim k As Variant
        For Each k In aliases.keys
            aliasMap(NormalizeHeader(CStr(k), caseSensitive:=caseSensitive)) = NormalizeHeader(CStr(aliases(k)), caseSensitive:=caseSensitive)
        Next
    End If

    Dim out As Object: Set out = NewDictionary(True)
    Dim i As Long
    For i = LBound(headersToFind) To UBound(headersToFind)
        Dim want As String: want = CStr(headersToFind(i))
        Dim wantNorm As String: wantNorm = NormalizeHeader(want, caseSensitive:=caseSensitive)

        If aliasMap.exists(wantNorm) Then wantNorm = aliasMap(wantNorm)

        If normMap.exists(wantNorm) Then
            out(want) = normMap(wantNorm)
        ElseIf allowFuzzy Then
            Dim best As Long: best = 0
            Dim bestDist As Long: bestDist = 32767
            Dim cand As Variant
            For Each cand In normMap.keys
                Dim d As Long: d = LevDistance(wantNorm, CStr(cand))
                If d < bestDist Then bestDist = d: best = normMap(cand)
            Next
            If bestDist <= maxDistance Then out(want) = best
        End If
        ' if not found, omit or set to 0
        If Not out.exists(want) Then out(want) = 0
    Next
    Set Header_IndexMap = out
End Function

Private Function NormalizeHeader(ByVal s As String, ByVal caseSensitive As Boolean) As String
    Dim t As String: t = Trim$(s)
    If Not caseSensitive Then t = LCase$(t)
    Dim ch As String, i As Long, r As String
    For i = 1 To Len(t)
        ch = mid$(t, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            r = r & ch
        End If
    Next
    NormalizeHeader = r
End Function

Private Function LevDistance(ByVal a As String, ByVal b As String) As Long
    Dim la As Long: la = Len(a)
    Dim lb As Long: lb = Len(b)
    Dim d() As Long: ReDim d(0 To la, 0 To lb)
    Dim i As Long, j As Long
    For i = 0 To la: d(i, 0) = i: Next
    For j = 0 To lb: d(0, j) = j: Next
    For i = 1 To la
        For j = 1 To lb
            Dim cost As Long: cost = IIf(mid$(a, i, 1) = mid$(b, j, 1), 0, 1)
            d(i, j) = Application.Min(Application.Min(d(i - 1, j) + 1, d(i, j - 1) + 1), d(i - 1, j - 1) + cost)
        Next
    Next
    LevDistance = d(la, lb)
End Function

'--- SelectColumns2D
' Projects a 2D array to a new 2D with selected columns (by index or header names)
Public Function Array_SelectColumns(ByVal arr2D As Variant, ByVal selectors As Variant, _
                                    Optional ByVal firstRowHasHeaders As Boolean = True, _
                                    Optional ByVal caseSensitive As Boolean = False, _
                                    Optional ByVal includeHeaders As Boolean = True) As Variant
    If Not Is2DArray(arr2D) Then Err.Raise 5, , "Array_SelectColumns: input must be 2D."
    Dim rows As Long, cols As Long: Size2D arr2D, rows, cols
    Dim baseR As Long: baseR = LBound(arr2D, 1)
    Dim baseC As Long: baseC = LBound(arr2D, 2)

    Dim selIdx() As Long
    If IsArray(selectors) And VarType(selectors(LBound(selectors))) <> vbString Then
        ' numeric indices
        Dim i As Long: ReDim selIdx(LBound(selectors) To UBound(selectors))
        For i = LBound(selectors) To UBound(selectors)
            selIdx(i) = CLng(selectors(i))
        Next
    Else
        ' header names
        Dim map As Object: Set map = Header_IndexMap(arr2D, selectors, 1, caseSensitive, True, 2)
        Dim s As Variant: ReDim selIdx(LBound(selectors) To UBound(selectors))
        For i = LBound(selectors) To UBound(selectors)
            s = selectors(i)
            If map.exists(s) And map(s) > 0 Then selIdx(i) = map(s) Else selIdx(i) = 0
        Next
    End If

    Dim outRows As Long: outRows = rows
    Dim outCols As Long: outCols = UBound(selIdx) - LBound(selIdx) + 1
    Dim out() As Variant: ReDim out(1 To IIf(includeHeaders, 1, 0) + (rows - IIf(firstRowHasHeaders, 1, 0)), 1 To outCols)

    Dim r As Long, c As Long, oc As Long, orow As Long
    orow = 1
    If includeHeaders And firstRowHasHeaders Then
        oc = 1
        For c = LBound(selIdx) To UBound(selIdx)
            If selIdx(c) > 0 Then
                out(1, oc) = arr2D(baseR, baseC + selIdx(c) - 1)
            Else
                out(1, oc) = ""
            End If
            oc = oc + 1
        Next
        orow = 2
    End If

    For r = baseR + IIf(firstRowHasHeaders, 1, 0) To baseR + rows - 1
        oc = 1
        For c = LBound(selIdx) To UBound(selIdx)
            If selIdx(c) > 0 Then out(orow, oc) = arr2D(r, baseC + selIdx(c) - 1) Else out(orow, oc) = Empty
            oc = oc + 1
        Next
        orow = orow + 1
    Next

    Array_SelectColumns = out
End Function

'--- TryRun
' Executes a callback safely and returns a record: Success, Result, ErrorNumber, ErrorDescription, DurationMs.
Public Function TryRun(ByVal callbackMacro As String, ParamArray args() As Variant) As Object
    Dim t0 As Double: t0 = Timer
    Dim r As Object: Set r = NewDictionary(True)
    On Error GoTo EH
    Dim res As Variant
    res = Application.Run(callbackMacro, args)
    r("Success") = True
    r("Result") = res
    r("ErrorNumber") = 0
    r("ErrorDescription") = ""
    r("DurationMs") = CLng((Timer - t0) * 1000)
    Set TryRun = r
    Exit Function
EH:
    r("Success") = False
    r("Result") = Empty
    r("ErrorNumber") = Err.Number
    r("ErrorDescription") = Err.Description
    r("DurationMs") = CLng((Timer - t0) * 1000)
    Set TryRun = r
End Function

'--- RetryWithBackoff
' Retries a callback with exponential backoff/jitter. shouldRetryMacro can inspect attempt/result/error.
' Returns: Success, Attempts, Result, ErrorNumber, ErrorDescription, TotalDurationMs.
Public Function Retry_WithBackoff(ByVal callbackMacro As String, _
                                  Optional ByVal attempts As Long = 3, _
                                  Optional ByVal initialDelayMs As Long = 250, _
                                  Optional ByVal maxDelayMs As Long = 5000, _
                                  Optional ByVal jitterPct As Double = 0.2, _
                                  Optional ByVal shouldRetryMacro As String = "", _
                                  Optional ByRef args As Variant) As Object
    Dim t0 As Double: t0 = Timer
    Dim res As Variant, att As Long, delay As Long
    Dim lastErrNum As Long, lastErrDesc As String
    For att = 1 To Application.Max(1, attempts)
        On Error GoTo TryEH
        res = Application.Run(callbackMacro, args)
        If Len(shouldRetryMacro) > 0 Then
            Dim doRetry As Boolean
            doRetry = CBool(Application.Run(shouldRetryMacro, True, 0&, "", res, att))
            If Not doRetry Then Exit For
        Else
            Exit For
        End If
TryNext:
        If att < attempts Then
            delay = ComputeDelay(initialDelayMs, maxDelayMs, jitterPct, att)
            Sleep delay
        End If
        On Error GoTo 0
    Next

    Dim r As Object: Set r = NewDictionary(True)
    r("Success") = (lastErrNum = 0)
    r("Attempts") = att
    r("Result") = res
    r("ErrorNumber") = lastErrNum
    r("ErrorDescription") = lastErrDesc
    r("TotalDurationMs") = CLng((Timer - t0) * 1000)
    Set Retry_WithBackoff = r
    Exit Function

TryEH:
    lastErrNum = Err.Number: lastErrDesc = Err.Description
    Dim retry As Boolean: retry = (att < attempts)
    If Len(shouldRetryMacro) > 0 Then
        On Error Resume Next
        retry = CBool(Application.Run(shouldRetryMacro, False, Err.Number, Err.Description, Empty, att))
        On Error GoTo 0
    End If
    If retry Then
        Resume TryNext
    Else
        Resume Next
    End If
End Function

Private Function ComputeDelay(ByVal initialMs As Long, ByVal maxMs As Long, ByVal jitter As Double, ByVal attempt As Long) As Long
    Dim base As Double: base = CDbl(initialMs) * (2 ^ (attempt - 1))
    If base > maxMs Then base = maxMs
    Dim rndJ As Double: rndJ = (Rnd() - 0.5) * 2 * jitter
    ComputeDelay = Application.Max(0, CLng(base * (1# + rndJ)))
End Function

'--- ThrottleDoEvents + Debounce
' ThrottleDoEvents: Call DoEvents only if minIntervalMs elapsed since lastYieldTick
Public Sub Throttle_DoEvents(ByRef lastYieldTick As Double, ByVal minIntervalMs As Long)
    Dim nowT As Double: nowT = Timer
    If lastYieldTick = 0# Or ElapsedMs(lastYieldTick, nowT) >= minIntervalMs Then
        DoEvents
        lastYieldTick = nowT
    End If
End Sub

Private Function ElapsedMs(ByVal tStart As Double, ByVal tNow As Double) As Double
    If tNow >= tStart Then
        ElapsedMs = (tNow - tStart) * 1000#
    Else
        ElapsedMs = ((86400# - tStart) + tNow) * 1000#
    End If
End Function


'--- StatusBarScope
' Push a temporary status message; pass returned token to StatusBar_Pop to restore.
Public Function StatusBar_Push(ByVal Message As String) As Variant
    Dim token As Variant: token = Application.StatusBar
    Application.StatusBar = Message
    StatusBar_Push = token
End Function

Public Sub StatusBar_Pop(ByVal token As Variant)
    Application.StatusBar = token
End Sub

'--- BatchRangeWrite
' Writes a 1D/2D array to a destination range with optional resize, clipping and clearing extra area.
Public Sub Range_BatchWrite(ByVal dest As Range, ByVal data As Variant, _
                            Optional ByVal autoResize As Boolean = True, _
                            Optional ByVal allowClip As Boolean = True, _
                            Optional ByVal clearExtra As Boolean = False)
    Dim rng As Range: Set rng = dest
    Dim arr2D As Variant
    If Is2DArray(data) Then
        arr2D = data
    Else
        arr2D = To2D(data)
    End If
    Dim rows As Long, cols As Long: Size2D arr2D, rows, cols

    Dim stCalc As XlCalculation: stCalc = Application.Calculation
    Dim stScreen As Boolean: stScreen = Application.screenUpdating
    Dim stEvents As Boolean: stEvents = Application.enableEvents
    Application.Calculation = xlCalculationManual
    Application.screenUpdating = False
    Application.enableEvents = False

    On Error GoTo Finally
    If autoResize Then
        Set rng = rng.Resize(rows, cols)
    ElseIf Not allowClip Then
        If rng.rows.count <> rows Or rng.Columns.count <> cols Then Err.Raise 5, , "Range_BatchWrite: size mismatch."
    Else
        rows = Application.Min(rows, rng.rows.count)
        cols = Application.Min(cols, rng.Columns.count)
    End If

    rng.Value2 = arr2D

    If clearExtra And (rng.rows.count > rows Or rng.Columns.count > cols) Then
        Dim extra As Range
        If rng.rows.count > rows Then
            Set extra = rng.Offset(rows, 0).Resize(rng.rows.count - rows, rng.Columns.count)
            extra.ClearContents
        End If
        If rng.Columns.count > cols Then
            Set extra = rng.Offset(0, cols).Resize(rng.rows.count, rng.Columns.count - cols)
            extra.ClearContents
        End If
    End If

Finally:
    Application.Calculation = stCalc
    Application.screenUpdating = stScreen
    Application.enableEvents = stEvents
End Sub

'--- TextFileIO
' Reads a text file
Public Function Text_ReadFile(ByVal filePath As String, Optional ByVal defaultEncoding As String = "utf-8") As String
    If Len(Dir$(filePath, vbNormal)) = 0 Then Err.Raise 53, , "Text_ReadFile: file not found."
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile filePath
    Dim bytes() As Byte
    bytes = stream.Read
    stream.Close

    Text_ReadFile = DecodeBytesToString(bytes, defaultEncoding)
End Function

' Writes text with specified encoding
Public Sub Text_WriteFile(ByVal filePath As String, ByVal content As String, _
                          Optional ByVal encoding As String = "utf-8", _
                          Optional ByVal withBOM As Boolean = True, _
                          Optional ByVal createDirs As Boolean = True)
    If createDirs Then EnsureFolderForFile filePath
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = NormalizeEncodingName(encoding)
    st.Open
    st.WriteText content, 0
    If LCase$(NormalizeEncodingName(encoding)) = "utf-8" And Not withBOM Then
        st.Position = 0
        st.Type = 1
        st.Position = 3
        st.SaveToFile filePath, 2
        st.Close
    Else
        st.SaveToFile filePath, 2
        st.Close
    End If
End Sub

' Appends text to a file (creates if not exists) using encoding.
Public Sub Text_AppendFile(ByVal filePath As String, ByVal content As String, _
                           Optional ByVal encoding As String = "utf-8", _
                           Optional ByVal withBOM As Boolean = True)
    Dim exists As Boolean: exists = (Len(Dir$(filePath, vbNormal)) > 0)
    If Not exists Then
        Text_WriteFile filePath, content, encoding, withBOM, True
    Else
        Dim current As String: current = Text_ReadFile(filePath, encoding)
        Text_WriteFile filePath, current & content, encoding, withBOM, False
    End If
End Sub

' --- helpers for Text I/O
Private Function DecodeBytesToString(ByRef bytes() As Byte, ByVal defaultEncoding As String) As String
    If UBound(bytes) < 0 Then DecodeBytesToString = "": Exit Function
    Dim hasUTF8BOM As Boolean, hasUTF16LE As Boolean, hasUTF16BE As Boolean
    hasUTF8BOM = (UBound(bytes) >= 2 And bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF)
    hasUTF16LE = (UBound(bytes) >= 1 And bytes(0) = &HFF And bytes(1) = &HFE)
    hasUTF16BE = (UBound(bytes) >= 1 And bytes(0) = &HFE And bytes(1) = &HFF)

    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    Dim sb As Object: Set sb = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open
    st.Write bytes
    st.Position = 0

    sb.Type = 2
    If hasUTF8BOM Then
        sb.Charset = "utf-8"
        st.Position = 3
    ElseIf hasUTF16LE Then
        sb.Charset = "utf-16le"
        st.Position = 2
    ElseIf hasUTF16BE Then
        sb.Charset = "utf-16be"
        st.Position = 2
    Else
        sb.Charset = NormalizeEncodingName(defaultEncoding)
        st.Position = 0
    End If

    sb.Open
    st.CopyTo sb
    sb.Position = 0
    DecodeBytesToString = sb.ReadText
    sb.Close: st.Close
End Function

Private Function NormalizeEncodingName(ByVal enc As String) As String
    Dim E As String: E = LCase$(Trim$(enc))
    Select Case E
        Case "utf8": NormalizeEncodingName = "utf-8"
        Case "utf16": NormalizeEncodingName = "utf-16le"
        Case Else: NormalizeEncodingName = E
    End Select
End Function

Private Sub EnsureFolderForFile(ByVal filePath As String)
    Dim p As Long: p = InStrRev(filePath, "\")
    If p > 0 Then EnsureFolder Left$(filePath, p - 1)
End Sub

'--- ConditionalFormatBuilder

' Clear all conditional formats from a range.
Public Sub CF_Clear(ByVal rng As Range)
    ResolveRange(rng).FormatConditions.Delete
End Sub

' Add a 3-color scale; pass xlConditionValue* constants for types and raw values.
Public Function CF_AddColorScale3(ByVal rng As Range, _
                                  Optional ByVal minType As Long = xlConditionValueLowestValue, _
                                  Optional ByVal midType As Long = xlConditionValuePercentile, _
                                  Optional ByVal maxType As Long = xlConditionValueHighestValue, _
                                  Optional ByVal minValue As Variant = Empty, _
                                  Optional ByVal midValue As Variant = 50, _
                                  Optional ByVal maxValue As Variant = Empty) As ColorScale
    Dim cs As ColorScale
    Set cs = ResolveRange(rng).FormatConditions.AddColorScale(ColorScaleType:=3)
    With cs.ColorScaleCriteria(1)
        .Type = minType
        If Not IsEmpty(minValue) Then .value = minValue
    End With
    With cs.ColorScaleCriteria(2)
        .Type = midType
        If Not IsEmpty(midValue) Then .value = midValue
    End With
    With cs.ColorScaleCriteria(3)
        .Type = maxType
        If Not IsEmpty(maxValue) Then .value = maxValue
    End With
    Set CF_AddColorScale3 = cs
End Function

' Add an icon set condition (e.g., xl3TrafficLights1). Reverse order toggles semantics.
Public Function CF_AddIconSet(ByVal rng As Range, ByVal iconSet As XlIconSet, _
                              Optional ByVal reverseOrder As Boolean = False, _
                              Optional ByVal showIconOnly As Boolean = False) As IconSetCondition
    Dim isc As IconSetCondition
    Set isc = ResolveRange(rng).FormatConditions.AddIconSetCondition
    With isc
        .iconSet = iconSet
        .reverseOrder = reverseOrder
        .showIconOnly = showIconOnly
    End With
    Set CF_AddIconSet = isc
End Function

' Add a cell-value based rule (>, <, between, etc.). Returns the FormatCondition for styling.
Public Function CF_AddRule_CellValue(ByVal rng As Range, ByVal op As XlFormatConditionOperator, _
                                     ByVal formula1 As Variant, Optional ByVal formula2 As Variant, _
                                     Optional ByVal stopIfTrue As Boolean = False) As FormatCondition
    Dim fc As FormatCondition
    If IsMissing(formula2) Then
        Set fc = ResolveRange(rng).FormatConditions.Add(Type:=xlCellValue, Operator:=op, formula1:=formula1)
    Else
        Set fc = ResolveRange(rng).FormatConditions.Add(Type:=xlCellValue, Operator:=op, formula1:=formula1, formula2:=formula2)
    End If
    fc.stopIfTrue = stopIfTrue
    Set CF_AddRule_CellValue = fc
End Function

' Add an expression-based rule (xlExpression).
Public Function CF_AddRule_Expression(ByVal rng As Range, ByVal formula As String, _
                                      Optional ByVal stopIfTrue As Boolean = False) As FormatCondition
    Dim fc As FormatCondition
    Set fc = ResolveRange(rng).FormatConditions.Add(Type:=xlExpression, formula1:="=" & formula)
    fc.stopIfTrue = stopIfTrue
    Set CF_AddRule_Expression = fc
End Function

' Apply simple formatting to a FormatCondition (font color/bold, fill color, number format).
Public Sub CF_SetFormat(ByVal fc As FormatCondition, _
                        Optional ByVal fontColor As Variant, _
                        Optional ByVal bold As Variant, _
                        Optional ByVal interiorColor As Variant, _
                        Optional ByVal numberFormat As Variant)
    If Not IsMissing(fontColor) Then fc.Font.Color = CLng(fontColor)
    If Not IsMissing(bold) Then fc.Font.bold = CBool(bold)
    If Not IsMissing(interiorColor) Then fc.Interior.Color = CLng(interiorColor)
    If Not IsMissing(numberFormat) Then fc.numberFormat = CStr(numberFormat)
End Sub

'--- ClipboardRead
' Reads text (Unicode) and optionally the CF_HTML payload (and fragment).
' Returns Dictionary: Success, Text, Html, Fragment, SourceURL
Public Function Clipboard_Read(Optional ByVal includeHtml As Boolean = True) As Object
    Dim d As Object: Set d = NewDictionary(True)
    d("Success") = False: d("Text") = "": d("Html") = "": d("Fragment") = "": d("SourceURL") = ""

    On Error Resume Next
    Dim dobj As Object: Set dobj = CreateObject("Forms.DataObject")
    On Error GoTo 0
    If Not dobj Is Nothing Then
        On Error Resume Next
        dobj.GetFromClipboard
        d("Text") = dobj.GetText
        On Error GoTo 0
    Else
        d("Text") = GetClipboardUnicodeText()
    End If

    If includeHtml Then
        Dim html As String: html = GetClipboardCFHtml()
        If Len(html) > 0 Then
            d("Html") = html
            Dim meta As Object: Set meta = ParseCFHtml(html)
            If meta Is Nothing Then
                d("Fragment") = ""
            Else
                d("Fragment") = mid$(html, meta("StartFragment") + 1, meta("EndFragment") - meta("StartFragment"))
                d("SourceURL") = meta("SourceURL")
            End If
        End If
    End If

    d("Success") = (Len(d("Text")) > 0) Or (Len(d("Html")) > 0)
    Set Clipboard_Read = d
End Function

Private Function GetClipboardUnicodeText() As String
    Const CF_UNICODETEXT As Long = 13
    Dim s As String: s = ""
    If OpenClipboard(0) = 0 Then GetClipboardUnicodeText = "": Exit Function
    If IsClipboardFormatAvailable(CF_UNICODETEXT) <> 0 Then
        Dim h As LongPtr: h = GetClipboardData(CF_UNICODETEXT)
        If h <> 0 Then
            Dim p As LongPtr: p = GlobalLock(h)
            If p <> 0 Then
                Dim sizeB As LongPtr: sizeB = GlobalSize(h)
                Dim b() As Byte: ReDim b(0 To CLng(sizeB - 1))
                CopyMemory VarPtr(b(0)), p, sizeB
                Dim i As Long
                ' find terminating null
                For i = 0 To UBound(b) - 1 Step 2
                    If b(i) = 0 And b(i + 1) = 0 Then Exit For
                Next
                If i > 0 Then
                    ReDim Preserve b(0 To i + 1)
                    s = bToStr(b)
                End If
                GlobalUnlock h
            End If
        End If
    End If
    CloseClipboard
    GetClipboardUnicodeText = s
End Function

Private Function GetClipboardCFHtml() As String
    Dim CF_HTML As Long: CF_HTML = RegisterClipboardFormatA("HTML Format")
    Dim s As String: s = ""
    If CF_HTML = 0 Then GetClipboardCFHtml = "": Exit Function
    If OpenClipboard(0) = 0 Then GetClipboardCFHtml = "": Exit Function
    If IsClipboardFormatAvailable(CF_HTML) <> 0 Then
        Dim h As LongPtr: h = GetClipboardData(CF_HTML)
        If h <> 0 Then
            Dim p As LongPtr: p = GlobalLock(h)
            If p <> 0 Then
                Dim sizeB As LongPtr: sizeB = GlobalSize(h)
                Dim b() As Byte: ReDim b(0 To CLng(sizeB - 1))
                CopyMemory VarPtr(b(0)), p, sizeB
                Dim t As String: t = StrConv(b, vbUnicode)
                Dim z As Long: z = InStr(1, t, vbNullChar, vbBinaryCompare)
                If z > 0 Then t = Left$(t, z - 1)
                s = t
                GlobalUnlock h
            End If
        End If
    End If
    CloseClipboard
    GetClipboardCFHtml = s
End Function

' Parse CF_HTML header to offsets dictionary.
Private Function ParseCFHtml(ByVal html As String) As Object
    On Error GoTo E
    Dim d As Object: Set d = NewDictionary(True)
    Dim lines() As String: lines = Split(html, vbCrLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        If InStr(1, lines(i), "StartHTML:", vbTextCompare) = 1 Then d("StartHTML") = CLng(mid$(lines(i), 11))
        If InStr(1, lines(i), "EndHTML:", vbTextCompare) = 1 Then d("EndHTML") = CLng(mid$(lines(i), 9))
        If InStr(1, lines(i), "StartFragment:", vbTextCompare) = 1 Then d("StartFragment") = CLng(mid$(lines(i), 15))
        If InStr(1, lines(i), "EndFragment:", vbTextCompare) = 1 Then d("EndFragment") = CLng(mid$(lines(i), 13))
        If InStr(1, lines(i), "SourceURL:", vbTextCompare) = 1 Then d("SourceURL") = mid$(lines(i), 11)
        If Len(lines(i)) = 0 Then Exit For ' header done
    Next
    Set ParseCFHtml = d
    Exit Function
E:
    Set ParseCFHtml = Nothing
End Function

Private Function bToStr(ByRef b() As Byte) As String
    Dim s As String
    s = String$(UBound(b) \ 2, vbNullChar)
    CopyMemory StrPtr(s), VarPtr(b(0)), (UBound(b) + 1)
    bToStr = s
End Function

'--- USER INFO HELPERS

' Get current username
Public Function User_GetUserName() As String
    On Error Resume Next
    Dim wn As Object: Set wn = CreateObject("WScript.Network")
    If Not wn Is Nothing Then
        User_GetUserName = CStr(wn.username)
    Else
        User_GetUserName = Environ$("USERNAME")
    End If
End Function

' Get current user domain (AD or machine).
Public Function User_GetDomain() As String
    On Error Resume Next
    Dim wn As Object: Set wn = CreateObject("WScript.Network")
    If Not wn Is Nothing Then
        User_GetDomain = CStr(wn.UserDomain)
    Else
        User_GetDomain = Environ$("USERDOMAIN")
    End If
End Function

' Get machine/computer name.
Public Function User_GetMachineName() As String
    On Error Resume Next
    Dim wn As Object: Set wn = CreateObject("WScript.Network")
    If Not wn Is Nothing Then
        User_GetMachineName = CStr(wn.ComputerName)
    Else
        User_GetMachineName = Environ$("COMPUTERNAME")
    End If
End Function

' Get Excel/Office info: Version, Build, Bitness.
Public Function User_GetExcelInfo() As Object
    Dim d As Object: Set d = NewDictionary(True)
    d("ExcelVersion") = Application.Version
    d("Build") = Application.Build
#If Win64 Then
    d("Bitness") = "64-bit"
#Else
    d("Bitness") = "32-bit"
#End If
#If VBA7 Then
    d("VBA7") = True
#Else
    d("VBA7") = False
#End If
    Set User_GetExcelInfo = d
End Function

' Get OS info via WMI (Caption, Version, BuildNumber, OSArchitecture).
Public Function User_GetOSInfo() As Object
    Dim d As Object: Set d = NewDictionary(True)
    On Error GoTo E
    Dim svc As Object: Set svc = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Dim col As Object: Set col = svc.ExecQuery("SELECT Caption, Version, BuildNumber, OSArchitecture FROM Win32_OperatingSystem")
    Dim it As Object
    For Each it In col
        d("Caption") = it.Caption
        d("Version") = it.Version
        d("BuildNumber") = it.BuildNumber
        d("OSArchitecture") = it.OSArchitecture
        Exit For
    Next
    Set User_GetOSInfo = d
    Exit Function
E:
    d("Caption") = Application.OperatingSystem
    Set User_GetOSInfo = d
End Function

' Get local IP addresses (IPv4/IPv6).
Public Function User_GetIPAddresses() As Variant
    Dim out As Variant: out = Empty
    'On Error GoTo E
    Dim svc As Object: Set svc = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Dim col As Object: Set col = svc.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    Dim it As Object, ips As Variant, i As Long
    For Each it In col
        If Not IsNull(it.IPAddress) Then
            On Error GoTo check_1
            User_GetIPAddresses = it.IPAddress(0)
            Exit Function
check_1:
            User_GetIPAddresses = it.IPAddress(1)
            Exit Function
        End If
    Next
E:
    User_GetIPAddresses = ""
End Function

' Get an environment variable
Public Function User_GetEnv(ByVal name As String) As String
    User_GetEnv = Environ$(name)
End Function

' Get all environment variables into a Dictionary.
Public Function User_GetAllEnv() As Object
    Dim d As Object: Set d = NewDictionary(True)
    Dim i As Long, env As String
    For i = 1 To 1000
        env = Environ$(i)
        If Len(env) = 0 Then Exit For
        Dim p As Long: p = InStr(1, env, "=", vbBinaryCompare)
        If p > 0 Then d(Left$(env, p - 1)) = mid$(env, p + 1)
    Next
    Set User_GetAllEnv = d
End Function

' Special folder path
Public Function User_GetSpecialFolder(ByVal name As String) As String
    On Error Resume Next
    Dim ws As Object: Set ws = CreateObject("WScript.Shell")
    Dim sf As Object: Set sf = ws.SpecialFolders
    Dim p As String: p = ""
    If Not sf Is Nothing Then
        p = sf.Item(CStr(name))
    End If
    If Len(p) = 0 Then
        Select Case LCase$(name)
            Case "desktop": p = Environ$("USERPROFILE") & "\Desktop"
            Case "mydocuments", "documents": p = Environ$("USERPROFILE") & "\Documents"
            Case "appdata": p = Environ$("APPDATA")
            Case "localappdata": p = Environ$("LOCALAPPDATA")
            Case Else: p = ""
        End Select
    End If
    User_GetSpecialFolder = p
End Function

' File/folder existence & access
Public Function File_Exists(ByVal path As String) As Boolean
    File_Exists = (Len(Dir$(path, vbNormal)) > 0)
End Function

Public Function Folder_Exists(ByVal path As String) As Boolean
    On Error Resume Next
    Folder_Exists = (GetAttr(path) And vbDirectory) = vbDirectory
End Function

' Ensure folder exists (creates as needed).
Public Sub EnsureFolder(ByVal folderPath As String)
    Dim p As Long, cur As String
    If Len(folderPath) = 0 Then Exit Sub
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    For p = 1 To Len(folderPath)
        If mid$(folderPath, p, 1) = "\" Then
            cur = Left$(folderPath, p - 1)
            If Len(cur) > 0 And Not Folder_Exists(cur) Then MkDir cur
        End If
    Next
    If Not Folder_Exists(folderPath) Then MkDir folderPath
End Sub

' Is folder writable? (create+delete a temp file).
Public Function Folder_IsWritable(ByVal folderPath As String) As Boolean
    On Error GoTo E
    If Not Folder_Exists(folderPath) Then Folder_IsWritable = False: Exit Function
    Dim tmp As String: tmp = folderPath & IIf(Right$(folderPath, 1) = "\", "", "\") & ".__write_test__" & Format$(Timer, "000000") & ".tmp"
    Dim f As Integer: f = FreeFile
    Open tmp For Output As #f
    Print #f, "test"
    Close #f
    Kill tmp
    Folder_IsWritable = True
    Exit Function
E:
    On Error Resume Next
    Close #f
    Folder_IsWritable = False
End Function


'-- Reads a registry value; returns defaultValue if missing. For the (Default) value, pass valueName:=""
Public Function Registry_Read(ByVal rootKey As String, _
                              ByVal subKey As String, _
                              Optional ByVal valueName As String = "", _
                              Optional ByVal defaultValue As Variant = Empty, _
                              Optional ByVal expandEnvStrings As Boolean = True) As Variant
    On Error GoTo Missing
    Dim ws As Object: Set ws = CreateObject("WScript.Shell")
    Dim fullPath As String: fullPath = BuildRegPath(rootKey, subKey, valueName)
    Dim v As Variant: v = ws.RegRead(fullPath)
    If expandEnvStrings Then
        If VarType(v) = vbString Then
            If InStr(1, v, "%", vbBinaryCompare) > 0 Then v = ws.ExpandEnvironmentStrings(CStr(v))
        End If
    End If
    Registry_Read = v
    Exit Function
Missing:
    Registry_Read = defaultValue
End Function

'-- Writes a registry value (creates the key path when needed).
' valueKind: "REG_SZ","REG_EXPAND_SZ","REG_DWORD","REG_BINARY","REG_MULTI_SZ","REG_QWORD"
Public Function Registry_Write(ByVal rootKey As String, _
                               ByVal subKey As String, _
                               Optional ByVal valueName As String = "", _
                               Optional ByVal valueData As Variant, _
                               Optional ByVal valueKind As String) As Boolean
    On Error GoTo EH
    Dim ws As Object: Set ws = CreateObject("WScript.Shell")
    Dim fullPath As String: fullPath = BuildRegPath(rootKey, subKey, valueName)
    Dim kind As String: kind = UCase$(Trim$(valueKind))

    Select Case kind
        Case "REG_SZ", "REG_EXPAND_SZ"
            ws.RegWrite fullPath, CStr(valueData), kind
        Case "REG_DWORD"
            ws.RegWrite fullPath, CLng(valueData), "REG_DWORD"
        Case "REG_BINARY"
            Dim b() As Byte
            If IsArray(valueData) Then
                b = valueData
            Else
                b = HexStringToBytes(CStr(valueData)) ' supports "01 0A FF" or "010AFF"
            End If
            ws.RegWrite fullPath, b, "REG_BINARY"
        Case "REG_MULTI_SZ"
            If IsArray(valueData) Then
                ws.RegWrite fullPath, valueData, "REG_MULTI_SZ"  ' array of strings
            Else
                Dim parts() As String: parts = Split(CStr(valueData), vbCrLf)
                ws.RegWrite fullPath, parts, "REG_MULTI_SZ"
            End If
        Case "REG_QWORD"
            ' WScript.Shell lacks REG_QWORD directly; write as REG_BINARY (8 bytes, little-endian).
            Dim q As Currency ' 64-bit container, but we will write via helper
            Dim ll As Double  ' accept large numeric; we will pack
            ll = CDbl(valueData)
            ws.RegWrite fullPath, QWordToBytes(ll), "REG_BINARY"
        Case Else
            Err.Raise 5, , "Registry_Write: Unsupported kind '" & valueKind & "'."
    End Select
    Registry_Write = True
    Exit Function
EH:
    Registry_Write = False
End Function

'-- True if a key exists (checks by attempting to read its default value)
Public Function Registry_KeyExists(ByVal rootKey As String, ByVal subKey As String) As Boolean
    On Error GoTo NotFound
    Dim ws As Object: Set ws = CreateObject("WScript.Shell")
    Dim p As String: p = BuildRegPath(rootKey, subKey, "")
    Dim v As Variant: v = ws.RegRead(p) ' default value
    Registry_KeyExists = True
    Exit Function
NotFound:
    Registry_KeyExists = False
End Function

'-- True if a named value exists under the key
Public Function Registry_ValueExists(ByVal rootKey As String, ByVal subKey As String, ByVal valueName As String) As Boolean
    On Error GoTo NotFound
    Dim ws As Object: Set ws = CreateObject("WScript.Shell")
    Dim p As String: p = BuildRegPath(rootKey, subKey, valueName)
    Dim v As Variant: v = ws.RegRead(p)
    Registry_ValueExists = True
    Exit Function
NotFound:
    Registry_ValueExists = False
End Function

'--- helpers (private) ---
Private Function BuildRegPath(ByVal rootKey As String, ByVal subKey As String, ByVal valueName As String) As String
    Dim root As String: root = NormalizeRootKey(rootKey)
    Dim sk As String: sk = subKey
    If Left$(sk, 1) = "\" Then sk = mid$(sk, 2)
    If Right$(sk, 1) = "\" Then sk = Left$(sk, Len(sk) - 1)
    If Len(valueName) = 0 Then
        BuildRegPath = root & "\" & sk & "\"      ' default value requires trailing "\"
    Else
        BuildRegPath = root & "\" & sk & "\" & valueName
    End If
End Function

Private Function NormalizeRootKey(ByVal k As String) As String
    Dim t As String: t = UCase$(Trim$(k))
    Select Case t
        Case "HKCU", "HKEY_CURRENT_USER": NormalizeRootKey = "HKCU"
        Case "HKLM", "HKEY_LOCAL_MACHINE": NormalizeRootKey = "HKLM"
        Case "HKCR", "HKEY_CLASSES_ROOT": NormalizeRootKey = "HKCR"
        Case "HKU", "HKEY_USERS": NormalizeRootKey = "HKU"
        Case "HKCC", "HKEY_CURRENT_CONFIG": NormalizeRootKey = "HKCC"
        Case Else: Err.Raise 5, , "NormalizeRootKey: Unknown root '" & k & "'."
    End Select
End Function

Private Function HexStringToBytes(ByVal hexText As String) As Byte()
    Dim s As String: s = Replace$(Replace$(hexText, " ", ""), vbTab, "")
    Dim n As Long: n = Len(s) \ 2
    Dim b() As Byte: ReDim b(0 To Application.Max(0, n - 1))
    Dim i As Long
    For i = 0 To n - 1
        b(i) = CByte("&H" & mid$(s, 2 * i + 1, 2))
    Next
    HexStringToBytes = b
End Function

Private Function QWordToBytes(ByVal num As Double) As Byte()
    Dim v As Currency
    Dim b() As Byte: ReDim b(0 To 7)
    Dim x As Double: x = num
    Dim i As Long, n As Double
    For i = 0 To 7
        n = x Mod 256#
        If n < 0 Then n = n + 256#
        b(i) = CByte(n)
        x = Fix((x - n) / 256#)
    Next
    QWordToBytes = b
End Function


'--- COLUMN LETTER / NUMBER CONVERSION + ADDRESS A1 <-> (Row, Col)
'-- Converts a 1-based column number to letters ("A", "Z", "AA"). If strictExcelLimit:=True, enforces 1..16384.
Public Function Column_NumberToLetter(ByVal colNumber As Long, Optional ByVal strictExcelLimit As Boolean = True) As String
    If colNumber < 1 Then Err.Raise 5, , "Column_NumberToLetter: colNumber must be >= 1."
    If strictExcelLimit And colNumber > 16384 Then Err.Raise 5, , "Column_NumberToLetter: exceeds Excel max (16384)."
    Dim n As Long: n = colNumber
    Dim s As String: s = ""
    Do While n > 0
        Dim r As Long: r = (n - 1) Mod 26
        s = Chr$(65 + r) & s
        n = (n - 1) \ 26
    Loop
    Column_NumberToLetter = s
End Function

'-- Converts column letters to a 1-based column number (e.g., "XFD" -> 16384). Accepts $ and whitespace.
Public Function Column_LetterToNumber(ByVal colLetters As String, Optional ByVal strictExcelLimit As Boolean = True) As Long
    Dim t As String: t = UCase$(Replace$(Replace$(Trim$(colLetters), "$", ""), " ", ""))
    If Len(t) = 0 Then Err.Raise 5, , "Column_LetterToNumber: empty input."
    Dim i As Long, n As Long: n = 0
    For i = 1 To Len(t)
        Dim ch As String: ch = mid$(t, i, 1)
        If ch < "A" Or ch > "Z" Then Err.Raise 5, , "Column_LetterToNumber: invalid char '" & ch & "'."
        n = n * 26 + (Asc(ch) - 64)
    Next
    If strictExcelLimit And n > 16384 Then Err.Raise 5, , "Column_LetterToNumber: exceeds Excel max (16384)."
    Column_LetterToNumber = n
End Function

'-- Builds an A1 address from row/col, with optional absolute row/col and optional sheet name prefix.
Public Function Address_FromRC(ByVal rowNumber As Long, _
                               ByVal colNumber As Long, _
                               Optional ByVal absoluteRow As Boolean = False, _
                               Optional ByVal absoluteCol As Boolean = False, _
                               Optional ByVal sheetName As String = "") As String
    If rowNumber < 1 Or colNumber < 1 Then Err.Raise 5, , "Address_FromRC: row/col must be >= 1."
    Dim colTxt As String: colTxt = Column_NumberToLetter(colNumber, True)
    Dim s As String
    s = IIf(absoluteCol, "$", "") & colTxt & IIf(absoluteRow, "$", "") & CStr(rowNumber)
    If Len(sheetName) > 0 Then
        s = QuoteSheetName(sheetName) & "!" & s
    End If
    Address_FromRC = s
End Function

'-- Parses an A1 address and returns row/col and absolute flags (ignores workbook/sheet prefix). Returns True if ok.
Public Function Address_A1ToRC(ByVal addressText As String, _
                               ByRef rowOut As Long, _
                               ByRef colOut As Long, _
                               Optional ByRef absoluteRowOut As Boolean, _
                               Optional ByRef absoluteColOut As Boolean) As Boolean
    On Error GoTo Bad
    Dim a As String: a = Trim$(addressText)
    ' Strip workbook/sheet prefix if present (everything up to last '!')
    Dim p As Long: p = InStrRev(a, "!")
    If p > 0 Then a = mid$(a, p + 1)

    Dim i As Long: i = 1
    Dim absC As Boolean: If i <= Len(a) And mid$(a, i, 1) = "$" Then absC = True: i = i + 1
    Dim colStr As String: Do While i <= Len(a)
        Dim ch As String: ch = mid$(a, i, 1)
        If ch >= "A" And ch <= "Z" Or ch >= "a" And ch <= "z" Then
            colStr = colStr & ch: i = i + 1
        Else
            Exit Do
        End If
    Loop
    If Len(colStr) = 0 Then GoTo Bad
    Dim absR As Boolean: If i <= Len(a) And mid$(a, i, 1) = "$" Then absR = True: i = i + 1
    Dim rowStr As String: rowStr = mid$(a, i)
    If Len(rowStr) = 0 Or Not IsNumeric(rowStr) Then GoTo Bad

    colOut = Column_LetterToNumber(colStr, True)
    rowOut = CLng(rowStr)
    absoluteRowOut = absR
    absoluteColOut = absC
    Address_A1ToRC = True
    Exit Function
Bad:
    Address_A1ToRC = False
End Function

Private Function QuoteSheetName(ByVal s As String) As String
    If InStr(1, s, " ", vbBinaryCompare) > 0 Or InStr(1, s, "'", vbBinaryCompare) > 0 Or InStr(1, s, "!", vbBinaryCompare) > 0 Or InStr(1, s, "]", vbBinaryCompare) > 0 Or InStr(1, s, "[", vbBinaryCompare) > 0 Then
        QuoteSheetName = "'" & Replace$(s, "'", "''") & "'"
    Else
        QuoteSheetName = s
    End If
End Function


'--- POP-UP WRAPPERS (error / warning / confirm)

'Shows a critical error dialog; returns vbOK.
Public Function UI_MsgError(ByVal messageText As String, Optional ByVal titleText As String = "Error") As VbMsgBoxResult
    UI_MsgError = MsgBox(messageText, vbOKOnly Or vbCritical Or vbApplicationModal, titleText)
End Function

'Shows a warning dialog; returns vbOK.
Public Function UI_MsgWarning(ByVal messageText As String, Optional ByVal titleText As String = "Warning") As VbMsgBoxResult
    UI_MsgWarning = MsgBox(messageText, vbOKOnly Or vbExclamation Or vbApplicationModal, titleText)
End Function

'Shows a confirmation dialog; set defaultYes:=True for Yes default; set allowCancel:=True for Yes/No/Cancel; returns the button clicked.
Public Function UI_MsgConfirm(ByVal questionText As String, _
                              Optional ByVal titleText As String = "Confirm", _
                              Optional ByVal defaultYes As Boolean = True, _
                              Optional ByVal allowCancel As Boolean = False) As VbMsgBoxResult
    Dim buttons As VbMsgBoxStyle
    buttons = IIf(allowCancel, vbYesNoCancel, vbYesNo)
    buttons = buttons Or IIf(defaultYes, vbDefaultButton1, vbDefaultButton2)
    UI_MsgConfirm = MsgBox(questionText, buttons Or vbQuestion Or vbApplicationModal, titleText)
End Function


'-- C-STYLE STRING FORMATTING
'-- Formats a string with C-like specifiers: %s, %d, %i, %u, %f, %x, %X, %c, and %%.
'   Supports flags 0 (zero-pad), - (left align), + (force sign); width; precision for %f.
'   Pass arguments in an array, e.g., Format_C("Hello %s %02d %.2f%%", Array("Cade", 7, 3.5))
Public Function Format_C(ByVal formatText As String, ByVal args As Variant) As String
    Dim out As String: out = ""
    Dim i As Long: i = 1
    Dim argIdx As Long: argIdx = 0
    Dim argc As Long: argc = IIf(IsArray(args), (UBound(args) - LBound(args) + 1), 0)
    Do While i <= Len(formatText)
        Dim ch As String: ch = mid$(formatText, i, 1)
        If ch <> "%" Then
            out = out & ch: i = i + 1
        Else
            ' Handle %%
            If i < Len(formatText) And mid$(formatText, i + 1, 1) = "%" Then
                out = out & "%": i = i + 2: GoTo NextLoop
            End If
            ' Parse flags
            Dim zeroPad As Boolean, leftAlign As Boolean, forceSign As Boolean
            Dim j As Long: j = i + 1
            Dim f As String
            Do While j <= Len(formatText)
                f = mid$(formatText, j, 1)
                If f = "0" Then
                    zeroPad = True
                ElseIf f = "-" Then
                    leftAlign = True
                ElseIf f = "+" Then
                    forceSign = True
                Else
                    Exit Do
                End If
                j = j + 1
            Loop
            ' Width
            Dim width As Long: width = 0
            Do While j <= Len(formatText) And mid$(formatText, j, 1) Like "[0-9]"
                width = width * 10 + (Asc(mid$(formatText, j, 1)) - 48)
                j = j + 1
            Loop
            ' Precision
            Dim precision As Long: precision = -1
            If j <= Len(formatText) And mid$(formatText, j, 1) = "." Then
                j = j + 1: precision = 0
                Do While j <= Len(formatText) And mid$(formatText, j, 1) Like "[0-9]"
                    precision = precision * 10 + (Asc(mid$(formatText, j, 1)) - 48)
                    j = j + 1
                Loop
            End If
            ' Specifier
            If j > Len(formatText) Then Err.Raise 5, , "Format_C: incomplete format specifier."
            Dim sp As String: sp = mid$(formatText, j, 1)
            j = j + 1

            Dim val As Variant
            If argIdx >= argc Then Err.Raise 5, , "Format_C: not enough arguments for specifiers."
            If IsArray(args) Then
                val = args(LBound(args) + argIdx)
            Else
                val = Empty
            End If
            argIdx = argIdx + 1

            Dim chunk As String
            Select Case sp
                Case "s"
                    chunk = CStr(val)
                    If precision >= 0 Then chunk = Left$(chunk, precision)
                    chunk = PadFormatted(chunk, width, leftAlign, zeroPad:=False, forceSign:=False)
                Case "d", "i"
                    chunk = FormatSignedInteger(val, width, leftAlign, zeroPad, forceSign)
                Case "u"
                    chunk = FormatUnsignedInteger(val, width, leftAlign, zeroPad)
                Case "f", "F"
                    chunk = FormatFloat(val, width, leftAlign, zeroPad, precision)
                Case "x", "X"
                    chunk = FormatHex(val, width, leftAlign, zeroPad, UCase$(sp) = "X")
                Case "c"
                    chunk = Chr$(CLng(val))
                Case Else
                    Err.Raise 5, , "Format_C: unsupported specifier '%" & sp & "'."
            End Select
            out = out & chunk
            i = j
        End If
NextLoop:
    Loop
    Format_C = out
End Function

'--- Format helpers (private) ---
Private Function PadFormatted(ByVal s As String, ByVal width As Long, ByVal leftAlign As Boolean, ByVal zeroPad As Boolean, ByVal forceSign As Boolean) As String
    Dim padChar As String: padChar = IIf(zeroPad And Not leftAlign, "0", " ")
    Dim sign As String: sign = ""
    If forceSign And Len(s) > 0 Then
        If Left$(s, 1) <> "-" And Left$(s, 1) <> "+" Then sign = "+"
    End If
    Dim core As String: core = sign & s
    If width <= 0 Or Len(core) >= width Then
        PadFormatted = core
    Else
        Dim padLen As Long: padLen = width - Len(core)
        If leftAlign Then
            PadFormatted = core & String$(padLen, " ")
        Else
            If padChar = "0" And sign = "+" Then
                PadFormatted = "+" & String$(padLen, "0") & s
            Else
                PadFormatted = String$(padLen, padChar) & core
            End If
        End If
    End If
End Function

Private Function FormatSignedInteger(ByVal v As Variant, ByVal width As Long, ByVal leftAlign As Boolean, ByVal zeroPad As Boolean, ByVal forceSign As Boolean) As String
    Dim n As Long
    n = CLng(v)
    Dim s As String: s = CStr(n)
    FormatSignedInteger = PadFormatted(s, width, leftAlign, zeroPad, forceSign)
End Function

Private Function FormatUnsignedInteger(ByVal v As Variant, ByVal width As Long, ByVal leftAlign As Boolean, ByVal zeroPad As Boolean) As String
    Dim d As Double: d = CDbl(v)
    If d < 0 Then d = d + 2 ^ 32 ' wraplike
    Dim s As String: s = CStr(Fix(d))
    FormatUnsignedInteger = PadFormatted(s, width, leftAlign, zeroPad, False)
End Function

Private Function FormatFloat(ByVal v As Variant, ByVal width As Long, ByVal leftAlign As Boolean, ByVal zeroPad As Boolean, ByVal precision As Long) As String
    Dim p As Long: p = IIf(precision >= 0, precision, 6)
    Dim fmt As String: fmt = "0"
    If p > 0 Then fmt = fmt & "." & String$(p, "0")
    Dim s As String: s = Format$(CDbl(v), fmt)
    FormatFloat = PadFormatted(s, width, leftAlign, zeroPad, False)
End Function

Private Function FormatHex(ByVal v As Variant, ByVal width As Long, ByVal leftAlign As Boolean, ByVal zeroPad As Boolean, ByVal upper As Boolean) As String
    Dim n As Currency: n = CDec(v)
    If n < 0 Then n = n + 2 ^ 32 ' simple wrap
    Dim s As String: s = Hex$(CLng(n))
    If Not upper Then s = LCase$(s)
    FormatHex = PadFormatted(s, width, leftAlign, zeroPad, False)
End Function

'-- FIND VALUE IN ROW OR COLUMN (exact or fuzzy)
' Returns the first matching cell or Nothing
'-- Finds in a single-row range; fuzzy uses Levenshtein distance. Set caseSensitive for string equality.
Public Function Row_FindFirst(ByVal rowRange As Range, _
                              ByVal seekValue As Variant, _
                              Optional ByVal fuzzy As Boolean = False, _
                              Optional ByVal maxDistance As Long = 2, _
                              Optional ByVal caseSensitive As Boolean = False) As Range
    Dim r As Range: Set r = ResolveRange(rowRange)
    If r.rows.count <> 1 Then Err.Raise 5, , "Row_FindFirst: rowRange must be a single row."
    Set Row_FindFirst = Vector_FindFirst(r, seekValue, fuzzy, maxDistance, caseSensitive)
End Function

'-- Finds in a single-column range; fuzzy uses Levenshtein distance. Set caseSensitive for string equality.
Public Function Column_FindFirst(ByVal colRange As Range, _
                                 ByVal seekValue As Variant, _
                                 Optional ByVal fuzzy As Boolean = False, _
                                 Optional ByVal maxDistance As Long = 2, _
                                 Optional ByVal caseSensitive As Boolean = False) As Range
    Dim r As Range: Set r = ResolveRange(colRange)
    If r.Columns.count <> 1 Then Err.Raise 5, , "Column_FindFirst: colRange must be a single column."
    Set Column_FindFirst = Vector_FindFirst(r, seekValue, fuzzy, maxDistance, caseSensitive)
End Function

'--- internal: works for a 1D vector range
Private Function Vector_FindFirst(ByVal vec As Range, _
                                  ByVal seekValue As Variant, _
                                  ByVal fuzzy As Boolean, _
                                  ByVal maxDistance As Long, _
                                  ByVal caseSensitive As Boolean) As Range
    Dim c As Range
    If Not fuzzy Then
        ' Exact match (case-sensitive by manual loop)
        For Each c In vec.cells
            If IsEqualityMatch(c.Value2, seekValue, caseSensitive) Then
                Set Vector_FindFirst = c
                Exit Function
            End If
        Next
    Else
        ' Fuzzy: pick the cell with smallest distance to seekValue (as string)
        Dim Target As String: Target = CStr(seekValue)
        If Not caseSensitive Then Target = LCase$(Target)
        Dim bestCell As Range, best As Long: best = 32767
        For Each c In vec.cells
            Dim s As String: s = CStr(c.Value2)
            If Not caseSensitive Then s = LCase$(s)
            Dim d As Long: d = LevDistance(Target, s) ' from earlier helper set
            If d < best Then Set bestCell = c: best = d
            If d = 0 Then Exit For
        Next
        If best <= maxDistance Then Set Vector_FindFirst = bestCell
    End If
End Function

Private Function IsEqualityMatch(ByVal a As Variant, ByVal b As Variant, ByVal caseSensitive As Boolean) As Boolean
    If IsNumeric(a) And IsNumeric(b) Then
        IsEqualityMatch = (CDbl(a) = CDbl(b))
    Else
        Dim sa As String: sa = CStr(a)
        Dim sb As String: sb = CStr(b)
        If Not caseSensitive Then sa = LCase$(sa): sb = LCase$(sb)
        IsEqualityMatch = (sa = sb)
    End If
End Function


'-- SIMPLE SWITCH-BASED CELL FORMATTING
'Applies formatting to a cell based on its value using an array of rules.
'   rules: Array of Array(matchKey, [interiorColor], [fontColor], [bold As Boolean], [numberFormat])
'   If useLikePattern:=True, matchKey is treated as VBA Like pattern (e.g., "ERR*").
Public Function Cell_FormatSwitch(ByVal targetCell As Range, _
                                  ByVal rules As Variant, _
                                  Optional ByVal caseSensitive As Boolean = False, _
                                  Optional ByVal useLikePattern As Boolean = False, _
                                  Optional ByVal defaultInteriorColor As Variant, _
                                  Optional ByVal defaultFontColor As Variant, _
                                  Optional ByVal defaultBold As Variant, _
                                  Optional ByVal defaultNumberFormat As Variant) As Boolean
    Dim cellObj As Range: Set cellObj = ResolveRange(targetCell)
    Dim v As Variant: v = cellObj.Value2
    Dim valText As String: valText = CStr(v)
    Dim valCmp As String: valCmp = IIf(caseSensitive, valText, LCase$(valText))

    Dim applied As Boolean: applied = False
    Dim i As Long
    If IsArray(rules) Then
        For i = LBound(rules) To UBound(rules)
            Dim rowRule As Variant: rowRule = rules(i)
            Dim key As String: key = CStr(rowRule(0))
            Dim keyCmp As String: keyCmp = IIf(caseSensitive, key, LCase$(key))

            Dim isMatch As Boolean
            If useLikePattern Then
                isMatch = (valCmp Like keyCmp)
            Else
                isMatch = (valCmp = keyCmp)
            End If

            If isMatch Then
                ApplyCellFormat cellObj, rowRule
                applied = True
                Exit For
            End If
        Next
    End If

    If Not applied Then
        Dim dRule(0 To 4) As Variant
        dRule(0) = ""
        dRule(1) = IIf(IsMissing(defaultInteriorColor), Empty, defaultInteriorColor)
        dRule(2) = IIf(IsMissing(defaultFontColor), Empty, defaultFontColor)
        dRule(3) = IIf(IsMissing(defaultBold), Empty, defaultBold)
        dRule(4) = IIf(IsMissing(defaultNumberFormat), Empty, defaultNumberFormat)
        ApplyCellFormat cellObj, dRule
    End If
    Cell_FormatSwitch = True
End Function

Private Sub ApplyCellFormat(ByVal cellObj As Range, ByVal ruleRow As Variant)
    On Error Resume Next
    If UBound(ruleRow) >= 1 Then If Not IsEmpty(ruleRow(1)) Then cellObj.Interior.Color = CLng(ruleRow(1))
    If UBound(ruleRow) >= 2 Then If Not IsEmpty(ruleRow(2)) Then cellObj.Font.Color = CLng(ruleRow(2))
    If UBound(ruleRow) >= 3 Then If Not IsEmpty(ruleRow(3)) Then cellObj.Font.bold = CBool(ruleRow(3))
    If UBound(ruleRow) >= 4 Then If Not IsEmpty(ruleRow(4)) Then cellObj.numberFormat = CStr(ruleRow(4))
    On Error GoTo 0
End Sub





