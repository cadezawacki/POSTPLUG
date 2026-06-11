
Attribute VB_Name = "modPublicUtils"
Option Explicit

' VA 2009-04-24: http://www.cpearson.com/Excel/Scope.aspx
' Option Private Module


' VA 2011-03-30: Extended list (STEA,STEJ,BLCDS,SUKU) + "SPECREF"
' VA 2009-06-18: We now test STANDARD_CONTRACT/NON_STANDARD_CONTRACT curves
' SNAC & STEC at the moment; might be extended in the future
' VA 2009-09-01: Adding STEM curves.
' VA 2009-09-09: Adding STAS & STAJ
Public Function CMS_GetSupportedContractStandards() As Variant
    CMS_GetSupportedContractStandards = Array("BLCDS", "SNAC", "STAJ", "STAS", "STEA", "STEC", "STEJ", "STEM", "SUKU")
End Function


Public Function CMS_GetSupportedTenorStandards() As Variant
    CMS_GetSupportedTenorStandards = Array(11, 12, 13, 17, "<Ad Hoc 1D Tenor Array>")
End Function

' -----------------------------------------------------

Public Function CMS_GetServerUrl() As String
    CMS_GetServerUrl = ThisWorkbook.CMS_WEB_SERVICE_URL
End Function

' VA 2009-04-21: For backward comaptibility
Public Function GetCMSServerMode() As String
    GetCMSServerMode = CMS_GetServerMode
End Function

Public Function CMS_GetServerMode() As String
        
    Debug.Print "CMS_GetServerMode: " & ThisWorkbook.CMS_WEB_SERVICE_URL
    
    Select Case (ThisWorkbook.CMS_WEB_SERVICE_URL)
        Case "":
            ThisWorkbook.CMS_SERVER_MODE = "INVALID"
        Case CMS_WSDL_DEV:
            ThisWorkbook.CMS_SERVER_MODE = "DEV"
        Case CMS_WSDL_STAGE:
            ThisWorkbook.CMS_SERVER_MODE = "STAGE"
        Case CMS_WSDL_PROD:
            ThisWorkbook.CMS_SERVER_MODE = "PROD"
        Case Else:
            ThisWorkbook.CMS_SERVER_MODE = "CUSTOM"
    End Select
    CMS_GetServerMode = ThisWorkbook.CMS_SERVER_MODE
End Function


'-----------------------------------------
' UTILS
'-----------------------------------------

Public Function CMS_GetVersion() As String
    CMS_GetVersion = ADDIN_VERSION
End Function

Public Function CMS_GetApplicationId() As String
    CMS_GetApplicationId = APP_ID
End Function



' http://www.devx.com/vb2themax/Tip/19162
' Encode an string so that it can be displayed correctly
' inside the browser.
'
' Same effect as the Server.HTMLEncode method in ASP

Function HTMLEncode(ByVal Text As String) As String
    Dim i As Integer
    Dim acode As Integer
    Dim repl As String

    HTMLEncode = Text

    For i = Len(HTMLEncode) To 1 Step -1
        acode = Asc(Mid$(HTMLEncode, i, 1))
        Select Case acode
            Case 32
                repl = " "
            Case 34
                repl = "&quot;"
            Case 38
                repl = "&amp;"
            Case 60
                repl = "&lt;"
            Case 62
                repl = "&gt;"
            Case 32 To 127
                ' don't touch alphanumeric chars
            Case Else
                repl = "&#" & CStr(acode) & ";"
        End Select
        If Len(repl) Then
            HTMLEncode = Left$(HTMLEncode, i - 1) & repl & Mid$(HTMLEncode, _
                i + 1)
            repl = ""
        End If
    Next
End Function



' ------------------------------------------------------
' http://www.mvps.org/access/modules/mdl0044.htm
' Determining the number of dimensions for an array
' ------------------------------------------------------
'This code was originally written by Lyle Fairfield
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Lyle Fairfield
'
 Function DimensionCount(b As Variant) As Long
    Dim V As Variant, z As Long
    For Each V In b
        z = z + 1
    Next V
    Do
        DimensionCount = DimensionCount + 1
        z = z / (UBound(b, DimensionCount) - LBound(b, DimensionCount) + 1)
    Loop Until z = 1
 End Function

'Sub testArray()
' Dim a(3 To 9, 4 To 7, 0, 1 To 12) As Variant, b As Variant
' Dim varReturn As Long
'    b = a
'    varReturn = fDummy(b)
'    MsgBox varReturn
' End Sub
'
' Function fDummy(b As Variant) As Long
'    fDummy = ElementCount(b)
' End Function

' VA 2008-12-12: Meant to be LastMarkedAs, but for the time being...
 ' SV 2012-07-31: Updated to use latestCurveMarkedType attribute in response
  Public Function GetQuoteType(ByVal Response As MSXML2.DOMDocument) As String
  
    Dim conventionTag As String
    Dim point As MSXML2.IXMLDOMNode

    On Error GoTo GetQuoteType_Err
    ' latestCurveMarkedType
    Set point = Response.selectSingleNode("//genericKeys/nameValuePair[@name='latestCurveMarkedType']")
    If (Not point Is Nothing) Then
        Select Case point.Text
            Case "QUOTED"
                conventionTag = "QUOTED_SPREADS"
            Case "SPREAD"
                conventionTag = "SPREADS"
            Case Else
                conventionTag = point.Text
        End Select
        GetQuoteType = conventionTag
    Else
        GetQuoteType = "UNKOWN"
    'Else
     '   If (QuoteConvention <> CURVE_CONVENTION_QUOTED) Then
      '    GetQuoteType = QuoteConvention
       ' Else
        '    GetQuoteType = CURVE_CONVENTION_UNKNOWN
         '   conventionTag = "creditUpfront"
          '  Set point = Response.selectSingleNode("//point[creditUpfront[periodMultiplier=5][period='Y']]")
           ' If (Not point Is Nothing) Then
            '    GetQuoteType = CURVE_CONVENTION_UPFRONT
            'Else
             '   Set point = Response.selectSingleNode("//point[creditSpread[periodMultiplier=5][period='Y']]")
              '  If (Not point Is Nothing) Then
               '     GetQuoteType = CURVE_CONVENTION_SPREADS
                'End If
            'End If
        'End If
    End If
GetQuoteType_Err:
   If (Err.Number <> 0) Then
   RaiseError ("Failed to recognize Curve Type: " & Err.Description)
   End If
  End Function
  
  
  Public Sub RaiseError(ByVal errDesc As String)
    Call Err.Raise(-1, APP_ID, errDesc)
  End Sub






   








