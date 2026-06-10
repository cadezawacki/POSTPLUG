
Attribute VB_Name = "modTrades"

Option Explicit

' ============================================================
'  Trade record handlers registered via modSocketRouter.
'  Extend this module with your business logic.
' ============================================================

Public Sub HandleTradeRecord(jsonPayload As String)
    Debug.Print toString(jsonPayload)
    Dim ticker As String: ticker = modSocketRouter.ExtractJsonValue(jsonPayload, "ticker")
    Dim action As String: action = modSocketRouter.ExtractJsonValue(jsonPayload, "action")
    Dim notional As String: notional = modSocketRouter.ExtractJsonValue(jsonPayload, "newNotionalStr")
    Debug.Print "Trade: " & ticker & " " & action & " " & notional
End Sub


