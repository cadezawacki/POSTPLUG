
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSoapEnvelope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
'  cSoapEnvelope � lightweight SOAP envelope builder with slot-based templating
'
'  Construct envelopes from registered templates or build ad-hoc. Slots are
'  named placeholders ({{SlotName}}) replaced at render time. Supports both
'  SOAP 1.1 and 1.2, custom namespaces, headers, and nested element builders.
' =============================================================================

Private mSoapVersion    As String       ' "1.1" or "1.2"
Private mAction         As String       ' SOAPAction / ws-addressing Action
Private mNamespaces     As Object       ' dict: prefix -> uri
Private mHeaderBlocks   As Collection   ' ordered XML fragments for <Header>
Private mBodyXml        As String       ' raw body XML (between <Body>...</Body>)
Private mSlots          As Object       ' dict: slotName -> value (for template rendering)
Private mRawSlots       As Object       ' dict: slotName -> value (NO xml-escaping)

' --- Lifecycle ---

Private Sub Class_Initialize()
    mSoapVersion = "1.1"
    mAction = ""
    Set mNamespaces = CreateObject("Scripting.Dictionary")
    Set mHeaderBlocks = New Collection
    mBodyXml = ""
    Set mSlots = CreateObject("Scripting.Dictionary")
    Set mRawSlots = CreateObject("Scripting.Dictionary")
    mSlots.CompareMode = vbTextCompare
    mRawSlots.CompareMode = vbTextCompare
    mNamespaces.CompareMode = vbTextCompare
End Sub

' --- Configuration ---

Public Property Let soapVersion(v As String)
    If v = "1.2" Or v = "1.1" Then
        mSoapVersion = v
    Else
        Err.Raise 5, "cSoapEnvelope", "SoapVersion must be ""1.1"" or ""1.2"""
    End If
End Property
Public Property Get soapVersion() As String
    soapVersion = mSoapVersion
End Property

Public Property Let action(v As String)
    mAction = v
End Property
Public Property Get action() As String
    action = mAction
End Property

Public Sub AddNamespace(prefix As String, uri As String)
    mNamespaces(prefix) = uri
End Sub

Public Sub AddHeader(xmlFragment As String)
    mHeaderBlocks.Add xmlFragment
End Sub

Public Sub ClearHeaders()
    Set mHeaderBlocks = New Collection
End Sub

' --- Body building ---

Public Property Let bodyXml(v As String)
    mBodyXml = v
End Property
Public Property Get bodyXml() As String
    bodyXml = mBodyXml
End Property

' --- Slot system ---

Public Sub SetSlot(slotName As String, value As String)
    mSlots(slotName) = value
End Sub

' Set a raw slot � value is inserted WITHOUT xml-escaping.
' Use for embedding pre-built XML fragments.
Public Sub SetRawSlot(slotName As String, value As String)
    mRawSlots(slotName) = value
End Sub

Public Sub SetSlots(ParamArray pairs() As Variant)
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs) - 1 Step 2
        mSlots(CStr(pairs(i))) = CStr(pairs(i + 1))
    Next i
End Sub

Public Sub ClearSlots()
    mSlots.RemoveAll
    mRawSlots.RemoveAll
End Sub

Public Function GetSlot(slotName As String) As String
    If mSlots.exists(slotName) Then
        GetSlot = mSlots(slotName)
    ElseIf mRawSlots.exists(slotName) Then
        GetSlot = mRawSlots(slotName)
    Else
        GetSlot = ""
    End If
End Function

' --- Render ---

Public Function Render() As String
    Dim sb As String
    Dim envNs As String

    If mSoapVersion = "1.2" Then
        envNs = "http://www.w3.org/2003/05/soap-envelope"
    Else
        envNs = "http://schemas.xmlsoap.org/soap/envelope/"
    End If

    sb = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbLf
    sb = sb & "<soap:Envelope xmlns:soap=""" & envNs & """"

    Dim k As Variant
    For Each k In mNamespaces.keys
        sb = sb & " xmlns:" & CStr(k) & "=""" & XmlEsc(CStr(mNamespaces(k))) & """"
    Next k
    sb = sb & ">" & vbLf

    ' Header
    If mHeaderBlocks.count > 0 Then
        sb = sb & "  <soap:Header>" & vbLf
        Dim hdr As Variant
        For Each hdr In mHeaderBlocks
            sb = sb & "    " & CStr(hdr) & vbLf
        Next hdr
        sb = sb & "  </soap:Header>" & vbLf
    End If

    ' Body
    sb = sb & "  <soap:Body>" & vbLf
    sb = sb & "    " & mBodyXml & vbLf
    sb = sb & "  </soap:Body>" & vbLf
    sb = sb & "</soap:Envelope>"

    ' Resolve slots
    Render = ResolveSlots(sb)
End Function

Public Function ContentType() As String
    If mSoapVersion = "1.2" Then
        ContentType = "application/soap+xml; charset=utf-8; action=""" & mAction & """"
    Else
        ContentType = "text/xml; charset=utf-8"
    End If
End Function

Public Function SoapActionHeader() As String
    If mSoapVersion = "1.1" Then
        SoapActionHeader = "SOAPAction: """ & mAction & """"
    Else
        SoapActionHeader = ""
    End If
End Function

' --- Internal ---

Private Function ResolveSlots(ByVal xml As String) As String
    Dim k As Variant
    ' Pass 1: raw slots � triple brace {{{Key}}} and double brace {{Key}} both resolve raw
    For Each k In mRawSlots.keys
        xml = Replace(xml, "{{{" & CStr(k) & "}}}", CStr(mRawSlots(k)))
        xml = Replace(xml, "{{" & CStr(k) & "}}", CStr(mRawSlots(k)))
    Next k
    ' Pass 2: escaped slots � triple brace {{{Key}}} resolves raw, double brace {{Key}} escapes
    For Each k In mSlots.keys
        xml = Replace(xml, "{{{" & CStr(k) & "}}}", CStr(mSlots(k)))
        xml = Replace(xml, "{{" & CStr(k) & "}}", XmlEsc(CStr(mSlots(k))))
    Next k
    ResolveSlots = xml
End Function

Private Function XmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    XmlEsc = s
End Function



