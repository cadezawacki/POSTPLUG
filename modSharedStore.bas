
Attribute VB_Name = "modSharedStore"
Option Explicit

' =============================================================================
'  modSharedStore - generic namespaced key-value store shared across modules,
'  sheet event handlers, async callbacks, userforms, etc.
'
'  Each namespace is its own Scripting.Dictionary (text-compare keys), created
'  on first use. Values can be anything: numbers, strings, arrays, objects.
'
'      Shared_Set "amends", "AEP", 125.5
'      If Shared_Has("amends", "AEP") Then lvl = Shared_Get("amends", "AEP")
'      Shared_Remove "amends", "AEP"
'      For Each k In Shared_Keys("amends"): ... : Next
'      Set d = Shared_Drain("amends")     ' atomically take everything + clear
'
'  CAVEAT (same as every module-level cache in VBA): pressing End on an error
'  dialog, Stop/Reset in the VBE, or editing code in break mode wipes the lot.
'  Treat it as a session cache, not durable state.
' =============================================================================

Private gStores As Object   ' namespace -> Scripting.Dictionary

' The namespace dictionary itself (created on demand) - use directly when you
' want the full Dictionary API.
Public Function Shared_Namespace(ByVal Ns As String) As Object
    If gStores Is Nothing Then
        Set gStores = CreateObject("Scripting.Dictionary")
        gStores.CompareMode = vbTextCompare
    End If
    If Not gStores.Exists(Ns) Then
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d.CompareMode = vbTextCompare
        gStores.Add Ns, d
    End If
    Set Shared_Namespace = gStores(Ns)
End Function

' Store a value or object under a key (overwrites).
Public Sub Shared_Set(ByVal Ns As String, ByVal Key As String, ByVal Value As Variant)
    Dim d As Object
    Set d = Shared_Namespace(Ns)
    If IsObject(Value) Then
        Set d(Key) = Value
    Else
        d(Key) = Value
    End If
End Sub

' Read a value (Empty if missing). Works for objects too:
'   Set o = Shared_Get("ns", "key")   /   v = Shared_Get("ns", "key")
Public Function Shared_Get(ByVal Ns As String, ByVal Key As String) As Variant
    Dim d As Object
    Set d = Shared_Namespace(Ns)
    If d.Exists(Key) Then
        If IsObject(d(Key)) Then
            Set Shared_Get = d(Key)
        Else
            Shared_Get = d(Key)
        End If
    End If
End Function

Public Function Shared_GetOrDefault(ByVal Ns As String, ByVal Key As String, _
                                    ByVal Default As Variant) As Variant
    If Shared_Has(Ns, Key) Then
        Shared_GetOrDefault = Shared_Get(Ns, Key)
    ElseIf IsObject(Default) Then
        Set Shared_GetOrDefault = Default
    Else
        Shared_GetOrDefault = Default
    End If
End Function

Public Function Shared_Has(ByVal Ns As String, ByVal Key As String) As Boolean
    Shared_Has = Shared_Namespace(Ns).Exists(Key)
End Function

Public Sub Shared_Remove(ByVal Ns As String, ByVal Key As String)
    Dim d As Object
    Set d = Shared_Namespace(Ns)
    If d.Exists(Key) Then d.Remove Key
End Sub

Public Sub Shared_Clear(ByVal Ns As String)
    Shared_Namespace(Ns).RemoveAll
End Sub

Public Function Shared_Keys(ByVal Ns As String) As Variant
    Shared_Keys = Shared_Namespace(Ns).Keys
End Function

Public Function Shared_Count(ByVal Ns As String) As Long
    Shared_Count = Shared_Namespace(Ns).Count
End Function

' Atomic take-all: returns the namespace's dictionary and replaces it with a
' fresh empty one. Ideal for "drain the staged amends on mark-button press":
' you keep a stable snapshot to iterate while new entries land in the new dict.
Public Function Shared_Drain(ByVal Ns As String) As Object
    Dim old As Object
    Set old = Shared_Namespace(Ns)
    Dim fresh As Object
    Set fresh = CreateObject("Scripting.Dictionary")
    fresh.CompareMode = vbTextCompare
    Set gStores(Ns) = fresh
    Set Shared_Drain = old
End Function

' List the namespaces currently alive (diagnostics).
Public Function Shared_Namespaces() As Variant
    If gStores Is Nothing Then
        Shared_Namespaces = Array()
    Else
        Shared_Namespaces = gStores.Keys
    End If
End Function
