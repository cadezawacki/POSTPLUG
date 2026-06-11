Public Sub RunMarkerCmd(ticker As String, Optional path As String = "")
    Dim oShell As Object
    Dim oExec As Object
    Set oShell = VBA.CreateObject("Wscript.Shell")
    Dim sCmd As String
    
    winhttp_path = path & "winhttpjs.bat"
    ticker_path = path & ticker
    header_path = path & "headers.txt"
    
    sCmd = "cmd /k " & winhttp_path & " " & Chr(34) & "http://cms-lxp.lehman.com/CreditMarkingService/" & Chr(34) & " -body-file " & ticker_path & ".txt -method POST -header " & header_path & " && del /f " & ticker_path & ".txt && exit"
    'sCmd = "cmd /k " & " " & Chr(34) & "timeout 5 && exit"
    
    ' Debug.Print sCmd
    If Not [settings.debugCMS] Then
        Call oShell.Run(sCmd, 0, False)
    Else
        MsgBox "SENDING " & ticker & " TO CMS", vbOKOnly, "DEBUG CMS"
    End If
    a = 1
End Sub

Public Sub RunMarkerCmdWait(ticker As String, Optional path As String = "")
    Dim sCmd As String
    Dim oShell As Object
    Dim oExec As Object
    Set oShell = VBA.CreateObject("Wscript.Shell")
    winhttp_path = path & "winhttpjs.bat"
    ticker_path = path & ticker
    header_path = path & "headers.txt"
    sCmd = "cmd /k " & winhttp_path & " " & Chr(34) & "http://cms-lxp.lehman.com/CreditMarkingService/" & Chr(34) & " -body-file " & ticker_path & ".txt -method POST -header " & header_path & " && del /f " & ticker_path & ".txt && exit"
    'sCmd = "cmd /k " & " " & Chr(34) & "timeout 5 && exit"
    ' Debug.Print sCmd
    If Not [settings.debugCMS] Then
        Call oShell.Run(sCmd, 0, True)
    Else
        MsgBox "SENDING " & ticker & "TO CMS", vbOKOnly, "DEBUG CMS"
    End If
    a = 1
End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      ' First remove readonly attribute, if set
      SetAttr FileToDelete, vbNormal
      ' Then delete the file
      Kill FileToDelete
   End If
End Sub

Function FileThere(FileName As String) As Boolean
     FileThere = (Dir(FileName) > "")
End Function

Public Sub create_winhttp_files()

    path = Application.ActiveWorkbook.path
    
    If Not FileThere("c:\Temp\winhttpjs.bat") Then
        FileCopy path & "\winhttpjs.bat", "C:\Temp\winhttpjs.bat"
    End If
    
    If Not FileThere("c:\\Temp\\headers.txt") Then
        FileCopy path & "\headers.txt", "c:\Temp\headers.txt"
    End If
    
    If Not FileThere("c:\Temp\winhttpjs.bat") Then
        MsgBox "Error writing to TEMP drive!"
    End If

End Sub

Public Sub create_txt(name As String, value As String, Optional path As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = FSO.CreateTextFile(path & name)
    oFile.WriteLine value
    oFile.Close
    Set FSO = Nothing
    Set oFile = Nothing
End Sub

Public Function runProcess(cmd As String) As String

    Dim oShell As Object
    Dim oExec As Object, oOutput As Object
    Dim s As String, sLine As String

    Set oShell = VBA.CreateObject("Wscript.Shell")

    Set oExec = oShell.Exec(cmd)
    Set oOutput = oExec.StdOut

    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbNewLine
    Wend
    
    Set oOutput = Nothing: Set oExec = Nothing
    Set oShell = Nothing
    
    ' Debug.Print s
    runProcess = s

End Function

Public Sub compName()
    sFilePath = VBA.Environ$("username")
    x = 1
End Sub

Public Function create_cds_xml(full_curve_name As String, _
    issuer_ticker As String, _
    debt_class As String, _
    ccy As String, _
    curve_type As String, _
    v0m As Double, _
    v3m As Double, _
    v6m As Double, _
    v9m As Double, _
    v1y As Double, _
    v2y As Double, _
    v3y As Double, _
    v4y As Double, _
    v5y As Double, _
    v6y As Double, _
    v7y As Double, _
    v8y As Double, _
    v9y As Double, _
    v10y As Double, _
    v15y As Double, _
    v20y As Double, _
    v30y As Double) As String
    
    result = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "no" & Chr(34) & "?><SOAP-ENV:Envelope xmlns:SOAPSDK1=" & Chr(34) & "http://www.w3.org/2001/XMLSchema" & Chr(34) & " xmlns:SOAPSDK2=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & " xmlns:SOAPSDK3=" & Chr(34) & "http://schemas.xmlsoap.org/soap/encoding/" & Chr(34) & " xmlns:SOAP-ENV=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34) & "><SOAP-ENV:Body SOAP-ENV:encodingStyle=" & Chr(34) & "http://schemas.xmlsoap.org/soap/encoding/" & Chr(34) & "><SOAPSDK4:setMarketData xmlns:SOAPSDK4=" & Chr(34) & "http://www.lehman.com/fta/CreditMarkingService" & Chr(34) & ">"
    result = result & create_xml_user_info()
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 0, "M", v0m)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 3, "M", v3m)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 6, "M", v6m)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 9, "M", v9m)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 1, "Y", v1y)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 2, "Y", v2y)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 3, "Y", v3y)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 4, "Y", v4y)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 5, "Y", v5y)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 6, "Y", v6y)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 7, "Y", v7y)
    ' result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 8, "Y", v8y)
    ' result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 9, "Y", v9y)
    result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 10, "Y", v10y)
    ' result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 15, "Y", v15y)
    ' result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 20, "Y", v20y)
    ' result = result & create_xml_point(full_curve_name, issuer_ticker, debt_class, ccy, curve_type, 30, "Y", v30y)
    result = result & "/marketSet&gt;&lt;/marketData&gt;&lt;/market&gt;&lt;/Rosetta&gt;</string0></SOAPSDK4:setMarketData></SOAP-ENV:Body></SOAP-ENV:Envelope>"
    
    create_cds_xml = result
    
    
    End Function
    
Private Function create_xml_user_info() As String
    Dim username As String: username = LCase(VBA.Environ$("username"))
    Dim sHostName As String: sHostName = VBA.Environ$("computername")
    Dim sFilePath As String: sFilePath = Application.ActiveWorkbook.FullName
    Dim sDate As String: sDate = Format(Now(), "yyyymmdd")
    Dim sLocation As String: sLocation = "NYC"
    
    result = "<string>&lt;user-metadata&gt;&lt;user-details&gt;&lt;user-id&gt;"
    result = result & username
    result = result & "&lt;/user-id&gt;&lt;user-domain&gt;INTRANET&lt;/user-domain&gt;&lt;application-id&gt;CMS - Excel Addin&lt;/application-id&gt;&lt;application-version&gt;4.9.4&lt;/application-version&gt;&lt;application-batch-id&gt;"
    result = result & sFilePath
    result = result & "&lt;/application-batch-id&gt;&lt;id&gt;"
    result = result & username
    result = result & "&lt;/id&gt;&lt;host-id&gt;"
    result = result & sHostName
    result = result & "&lt;/host-id&gt;&lt;/user-details&gt;&lt;/user-metadata&gt;</string>"
    result = result & "<string0>&lt;Rosetta version=&quot;5.0.15&quot;&gt;&lt;market&gt;&lt;marketData&gt;&lt;action&gt;SET&lt;/action&gt;&lt;date&gt;"
    result = result & sDate
    result = result & "&lt;/date&gt;&lt;marketSet&gt;&lt;location&gt;"
    result = result & sLocation
    result = result & "&lt;/location&gt;&lt;logicalTime&gt;LIVE&lt;/logicalTime&gt;&lt;"
    
    create_xml_user_info = result

End Function

Private Function create_xml_point(full_curve_name As String, _
    issuer_ticker As String, _
    debt_class As String, _
    ccy As String, _
    curve_type As String, _
    tenor As Integer, _
    tenorPeriod As String, _
    value As Double)
    
    result = "point&gt;&lt;label&gt;"
    result = result & full_curve_name
    result = result & "&lt;/label&gt;&lt;creditSpread&gt;&lt;issuerTicker&gt;"
    result = result & issuer_ticker
    result = result & "&lt;/issuerTicker&gt;&lt;debtClass&gt;"
    result = result & debt_class
    result = result & "&lt;/debtClass&gt;&lt;currency&gt;"
    result = result & ccy
    result = result & "&lt;/currency&gt;&lt;creditCurveType&gt;"
    result = result & curve_type
    result = result & "&lt;/creditCurveType&gt;&lt;periodMultiplier&gt;"
    result = result & tenor
    result = result & "&lt;/periodMultiplier&gt;&lt;period&gt;"
    result = result & tenorPeriod
    result = result & "&lt;/period&gt;&lt;contractualSpread&gt;-777&lt;/contractualSpread&gt;&lt;/creditSpread&gt;&lt;value&gt;"
    result = result & value
    result = result & "&lt;/value&gt;&lt;/point&gt;&lt;"
    
    create_xml_point = result
    
    End Function