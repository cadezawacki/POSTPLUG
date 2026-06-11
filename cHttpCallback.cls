
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHttpCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Handle(requestId As String, r As cHttpResponse)
    Debug.Print "[cb] " & requestId & " " & r.Status & " " & r.ElapsedMs & "ms"
End Sub



