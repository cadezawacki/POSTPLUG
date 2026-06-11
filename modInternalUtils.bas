
Attribute VB_Name = "modInternalUtils"
Option Explicit

' VA 2009-04-24: http://www.cpearson.com/Excel/Scope.aspx
Option Private Module

' ------------------------------------------------------------------------------
' VA 2011-03-30: To implement CMS_GetFiveTupleArray(), we need to know
' how to decompose DebtClass into Tier/RefId and DocClause
' based on concatChar ("~" for SPECREF Products, "_" for the rest.
' ------------------------------------------------------------------------------

'''' http://www.freevbcode.com/ShowCode.Asp?ID=2856
'''Public Function StringStartsWith(ByVal strValue As String, _
'''  CheckFor As String, Optional CompareType As VbCompareMethod _
'''   = vbBinaryCompare) As Boolean
'''
''''Determines if a string starts with the same characters as
''''CheckFor string
'''
''''True if starts with CheckFor, false otherwise
''''Case sensitive by default.  If you want non-case sensitive, set
''''last parameter to vbTextCompare
'''
'''    'Examples:
'''    'MsgBox StringStartsWith("Test", "TE") 'false
'''    'MsgBox StringStartsWith("Test", "TE", vbTextCompare) 'True
'''
'''  Dim sCompare As String
'''  Dim lLen As Long
'''
'''  lLen = Len(CheckFor)
'''  If lLen > Len(strValue) Then Exit Function
'''  sCompare = Left(strValue, lLen)
'''  StringStartsWith = StrComp(sCompare, CheckFor, CompareType) = 0
'''
'''End Function
'''
'''' http://www.freevbcode.com/ShowCode.Asp?ID=2856
'''Public Function StringEndsWith(ByVal strValue As String, _
'''   CheckFor As String, Optional CompareType As VbCompareMethod _
'''   = vbBinaryCompare) As Boolean
''' 'Determines if a string ends with the same characters as
''' 'CheckFor string
'''
''' 'True if end with CheckFor, false otherwise
'''
''' 'Case sensitive by default.  If you want non-case sensitive, set
''' 'last parameter to vbTextCompare
'''
'''  'Examples
'''  'MsgBox StringEndsWith("Test", "ST") 'False
'''  'MsgBox StringEndsWith("Test", "ST", vbTextCompare) 'True
'''
'''  Dim sCompare As String
'''  Dim lLen As Long
'''
'''  lLen = Len(CheckFor)
'''  If lLen > Len(strValue) Then Exit Function
'''  sCompare = Right(strValue, lLen)
'''  StringEndsWith = StrComp(sCompare, CheckFor, CompareType) = 0
'''
'''End Function
'''

' http://msdn.microsoft.com/en-us/library/aa189091(v=office.10).aspx
Function CountOccurrences(strText As String, _
                          strFind As String, _
                          Optional lngCompare As VbCompareMethod) As Long

   ' Count occurrences of a particular character or characters.
   ' If lngCompare argument is omitted, procedure performs binary comparison.
   
   Dim lngPos       As Long
   Dim lngTemp      As Long
   Dim lngCount     As Long
   
   ' Specify a starting position. We don't need it the first
   ' time through the loop, but we'll need it on subsequent passes.
   lngPos = 1
   ' Execute the loop at least once.
   Do
      ' Store position at which strFind first occurs.
      lngPos = InStr(lngPos, strText, strFind, lngCompare)
      ' Store position in a temporary variable.
      lngTemp = lngPos
      ' Check that strFind has been found.
      If lngPos > 0 Then
         ' Increment counter variable.
         lngCount = lngCount + 1
         ' Define a new starting position.
         lngPos = lngPos + Len(strFind)
      End If
   ' Loop until last occurrence has been found.
   Loop Until lngPos = 0
   ' Return the number of occurrences found.
   CountOccurrences = lngCount
End Function




' ------------------------------------------------------------------------------
' VA 2009-07-23: As suggested by Gavin Shanks & Merwyn Dsouza,
' let's try a singleton CMS SOAP connection
' ------------------------------------------------------------------------------

Function GetCmsConnection() As MSOSOAPLib30.SoapClient30
   
   
   If (ThisWorkbook.CmsClient Is Nothing) Then
        If (ThisWorkbook.CMS_WEB_SERVICE_URL = "") Then
         'Debug.Print "GetCmsConnection: WS URl is nothing: " & ThisWorkbook.CMS_WEB_SERVICE_URL
         Call RaiseError("CMS Environment got corrupted. Please reconnect to CMS using the menu bar")
        End If
        
        Set ThisWorkbook.CmsClient = New MSOSOAPLib30.SoapClient30
        Call ThisWorkbook.CmsClient.MSSoapInit(ThisWorkbook.CMS_WEB_SERVICE_URL)
        
        ' VAS 2009-07-16: Restored
        ' VA 2009-06-23: This setting fails in STAGE!!!
        ' VA 2009-04-08 D-Day: Timeout issue
         ThisWorkbook.CmsClient.ClientProperty("ServerHTTPRequest") = True
         ThisWorkbook.CmsClient.ConnectorProperty("Timeout") = CMS_SERVER_GET_TIMEOUT
    End If
    Set GetCmsConnection = ThisWorkbook.CmsClient

End Function

Private Sub RaiseError(errDesc As String)
    Call Err.Raise(-1, APP_ID, errDesc)
End Sub



' ------------------------------------------------------------------------------
' VA 2009-06-18: Now we handle CONTRACT STANDARDS
' ASSUMPTION: Product starts with it (e.g., SNAC100, STEC750))
' For supported Product "prefixes" see CMS_GetSupportedContractStandards()
' ------------------------------------------------------------------------------

Public Function IsStandardContractProduct(ByVal Product As String) As Boolean
    Dim SupportedContractStandards() As Variant
    Dim supportedContractStandard As Variant
    Dim standardName As String
    
    ' IsContractStandardProduct = False
    
    SupportedContractStandards = CMS_GetSupportedContractStandards
    
    For Each supportedContractStandard In CMS_GetSupportedContractStandards
        standardName = supportedContractStandard
        If StringStartsWith(UCase(Product), standardName) Then
            IsStandardContractProduct = True
            Exit Function
        End If
    Next
End Function






' =================================================================
' TEXT UTILITIES
' http://www.freevbcode.com/ShowCode.Asp?ID=2856
' =================================================================
Public Function StringStartsWith(ByVal strValue As String, _
  CheckFor As String, Optional CompareType As VbCompareMethod _
   = vbBinaryCompare) As Boolean
   
'Determines if a string starts with the same characters as
'CheckFor string

'True if starts with CheckFor, false otherwise
'Case sensitive by default.  If you want non-case sensitive, set
'last parameter to vbTextCompare
    
    'Examples:
    'MsgBox StringStartsWith("Test", "TE") 'false
    'MsgBox StringStartsWith("Test", "TE", vbTextCompare) 'True
    
  Dim sCompare As String
  Dim lLen As Long
   
  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Left(strValue, lLen)
  StringStartsWith = StrComp(sCompare, CheckFor, CompareType) = 0

End Function

Public Function StringEndsWith(ByVal strValue As String, _
   CheckFor As String, Optional CompareType As VbCompareMethod _
   = vbBinaryCompare) As Boolean
 'Determines if a string ends with the same characters as
 'CheckFor string
 
 'True if end with CheckFor, false otherwise

 'Case sensitive by default.  If you want non-case sensitive, set
 'last parameter to vbTextCompare
 
  'Examples
  'MsgBox StringEndsWith("Test", "ST") 'False
  'MsgBox StringEndsWith("Test", "ST", vbTextCompare) 'True

  Dim sCompare As String
  Dim lLen As Long
   
  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Right(strValue, lLen)
  StringEndsWith = StrComp(sCompare, CheckFor, CompareType) = 0

End Function




' =================================================================
' REGISTRY UTILITIES
' =================================================================

Public Sub GetCmsRegistrySettings()

    ' Compile the Proc & Common Section names
    Dim procCmsSectionName As String
    Dim cmnCmsSectionName As String
    ' procCmsSectionName = REG_CMS_SECTION_NAME & ".P" & GetProcessId
    procCmsSectionName = GetCurrentProcessId
    cmnCmsSectionName = "Common"
    
    ' ------------------------------------------------------
    ' Current User environment
    ' ------------------------------------------------------
    ThisWorkbook.CMS_USER_LOGIN_NAME = GetSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_USER_LOGIN_NAME")
    ThisWorkbook.CMS_USER_MACHINE_NAME = GetSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_USER_MACHINE_NAME")
    ThisWorkbook.CMS_USER_DOMAIN_NAME = GetSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_USER_DOMAIN_NAME")
    
    ' ------------------------------------------------------
    ' SET ENVIRONMENT CORRECTLY
    ' AND UPDATE THE ADD-IN TITLE ACCORDINGLY!!!
    ' ------------------------------------------------------
    'Debug.Print "GetCmsRegistrySettings: Initializing from registry: " & ThisWorkbook.CMS_WEB_SERVICE_URL
    ThisWorkbook.CMS_WEB_SERVICE_URL = GetSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_WEB_SERVICE_URL")
    ThisWorkbook.CMS_SERVER_MODE = GetSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_SERVER_MODE")
    'Debug.Print "GetCmsRegistrySettings: After initializing from registry: " & ThisWorkbook.CMS_WEB_SERVICE_URL & " " & ThisWorkbook.CMS_SERVER_MODE
    
    ' This guy comes from the Common Registry section
    ThisWorkbook.CMS_LAST_CUSTOM_WEB_SERVICE_URL = GetSetting(REG_CMS_APP_NAME, cmnCmsSectionName, "CMS_LAST_CUSTOM_WEB_SERVICE_URL")
    
    ' ==========================================================================================
    ' Impersonated User Name (for testing, delegation, etc.)
    ' ==========================================================================================
    ThisWorkbook.IMPERSONATED_USER_NAME = GetSetting(REG_CMS_APP_NAME, procCmsSectionName, "IMPERSONATED_USER_NAME")
    
    ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN = GetSetting(REG_CMS_APP_NAME, procCmsSectionName, "IS_ALLOWED_TO_USE_ADDIN")
    ThisWorkbook.IS_ALLOWED_TO_CONFIG = GetSetting(REG_CMS_APP_NAME, procCmsSectionName, "IS_ALLOWED_TO_CONFIG")
    ThisWorkbook.IS_ALLOWED_TO_IMPERSONATE = GetSetting(REG_CMS_APP_NAME, procCmsSectionName, "IS_ALLOWED_TO_IMPERSONATE")

End Sub

' VA 2009-04-30: Settings go into Process-specific Registry slot...
' ... but for CMS_LAST_CUSTOM_WEB_SERVICE_URL that goes to Common one
Public Sub SaveCmsRegistrySettings()

     ' Compile the Proc & Common Section names
    Dim procCmsSectionName As String
    Dim cmnCmsSectionName As String
    procCmsSectionName = GetCurrentProcessId
    ' cmnCmsSectionName = "Common"
    
    ' ------------------------------------------------------
    ' Current User environment
    ' ------------------------------------------------------
    
    ' VA 2009-05-04: Timestamp
    ThisWorkbook.CMS_ENVIRONMENT_TIMESTAMP = Now()
    Call SaveSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_ENVIRONMENT_TIMESTAMP", ThisWorkbook.CMS_ENVIRONMENT_TIMESTAMP)
    
    ' VA 2009-05-05: Let's keep the Process IDs in Processes for easy cleanup!
    Call SaveSetting(REG_CMS_APP_NAME, REG_CMS_SESSION_SECTION_NAME, procCmsSectionName, ThisWorkbook.CMS_ENVIRONMENT_TIMESTAMP)
    
    ' VA 2009-04-29: Just in case... :-)
    ThisWorkbook.CMS_USER_LOGIN_NAME = CMS_GetUserLoginName()
    ThisWorkbook.CMS_USER_MACHINE_NAME = CMS_GetUserMachineName()
    ThisWorkbook.CMS_USER_DOMAIN_NAME = CMS_GetUserDomainName()
    
    Call SaveSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_USER_LOGIN_NAME", ThisWorkbook.CMS_USER_LOGIN_NAME)
    Call SaveSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_USER_MACHINE_NAME", ThisWorkbook.CMS_USER_MACHINE_NAME)
    Call SaveSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_USER_DOMAIN_NAME", ThisWorkbook.CMS_USER_DOMAIN_NAME)
    
    ' ------------------------------------------------------
    ' SET ENVIRONMENT CORRECTLY
    ' AND UPDATE THE ADD-IN TITLE ACCORDINGLY!!!
    ' ------------------------------------------------------
    'Debug.Print "SaveCmsRegistrySettings: saving url to registries: " & ThisWorkbook.CMS_WEB_SERVICE_URL
    Call SaveSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_WEB_SERVICE_URL", ThisWorkbook.CMS_WEB_SERVICE_URL)
    Call SaveSetting(REG_CMS_APP_NAME, procCmsSectionName, "CMS_SERVER_MODE", ThisWorkbook.CMS_SERVER_MODE)
    
    ' This guys goes into the Common section - if it DOES exist...
    If (Len(ThisWorkbook.CMS_LAST_CUSTOM_WEB_SERVICE_URL) > 0) Then
        Call SaveSetting(REG_CMS_APP_NAME, REG_CMS_COMMON_SECTION_NAME, "CMS_LAST_CUSTOM_WEB_SERVICE_URL", ThisWorkbook.CMS_LAST_CUSTOM_WEB_SERVICE_URL)
    Else
        ' VA 2009-05-07: Don't do it! We'll cleanup next time add-in opens up
        ' Call DeleteSetting(REG_CMS_APP_NAME, cmnCmsSectionName, "CMS_LAST_CUSTOM_WEB_SERVICE_URL")
    End If
    
    ' ==========================================================================================
    ' Impersonated User Name (for testing, delegation, etc.)
    ' ==========================================================================================
    Call SaveSetting(REG_CMS_APP_NAME, procCmsSectionName, "IMPERSONATED_USER_NAME", ThisWorkbook.IMPERSONATED_USER_NAME)
    
    Call SaveSetting(REG_CMS_APP_NAME, procCmsSectionName, "IS_ALLOWED_TO_USE_ADDIN", ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN)
    Call SaveSetting(REG_CMS_APP_NAME, procCmsSectionName, "IS_ALLOWED_TO_CONFIG", ThisWorkbook.IS_ALLOWED_TO_CONFIG)
    Call SaveSetting(REG_CMS_APP_NAME, procCmsSectionName, "IS_ALLOWED_TO_IMPERSONATE", ThisWorkbook.IS_ALLOWED_TO_IMPERSONATE)

End Sub

' Delete current Process-specific settings
Public Sub DeleteCmsCurrentProcRegistrySettings()
    On Error Resume Next
    Call DeleteSetting(REG_CMS_APP_NAME, GetCurrentProcessId)
End Sub

' VA 2009-05-06: Cleanup the Regsitry
' If the Process ID is not current, remove section from Registry.
' DD 2022-23-03:
' Here "Computer\HKEY_CURRENT_USER\Software\VB and VBA Program Settings\CMS.DataAccessAddIn" PID of Excel is saved.
' This function checks if the current Excel PID is the same as that saved in the regitry and if not registry setting is deleted.
Public Sub CleanupCmsRegistrySettings()

    Dim procInfo As modWindowsAPI.ProcessInfo
    Dim cmsAllProcessArray As Variant
    Dim rowIndex As Integer
    Dim curentPID As Long
    Dim arrayHelper As New CArrayHelper
    
    ' Read registry Computer\HKEY_CURRENT_USER\Software\VB and VBA Program Settings\CMS.DataAccessAddIn
    cmsAllProcessArray = GetAllSettings(REG_CMS_APP_NAME, REG_CMS_SESSION_SECTION_NAME)
    
    ' Check for missing or empty directory
    If (arrayHelper.IsArrayEmpty(cmsAllProcessArray)) Then
        Exit Sub
    End If
    
    ' Iterate through all session sections (a 2-column array; 1st column - PID)
    For rowIndex = 0 To UBound(cmsAllProcessArray)
        If (IsNumeric(cmsAllProcessArray(rowIndex, 0))) Then
            curentPID = cmsAllProcessArray(rowIndex, 0)
            Call GetProcessInfo(curentPID, procInfo)
            If (Not (procInfo.processExists And LCase(procInfo.processExeName) = "excel.exe")) Then
                    Call DeleteSetting(REG_CMS_APP_NAME, REG_CMS_SESSION_SECTION_NAME, curentPID)
                    Call DeleteSetting(REG_CMS_APP_NAME, curentPID)
            End If
        End If
    Next
  
End Sub


' VA 2009-05-11: "Generic" method to calc Tenor shifts
' ASSUMPTION: TenorStandard is already checked and is vaild
Function GetTenorColumnShift(Optional ByVal TenorStandard As Variant = TENORS_11_STANDARD) As Integer
    
    Dim rngAdHocTenors As Range
    Dim adHocTenorArray() As Variant
    
    GetTenorColumnShift = 0
    
    If (IsObject(TenorStandard) And Not IsNumeric(TenorStandard)) Then
        
        ' VA 2009-05-12: It better be a Range... :-)
        ' For performance reasons we are NOT validating the ad hoc Tenor range every time
        ' Blood in, blood out... :-)
        
        Set rngAdHocTenors = TenorStandard
        ' adHocTenorArray

        ' Transpose to a "Landscape" if necessary - and sorry, I am lazy... :-)
        If (rngAdHocTenors.Rows.Count > rngAdHocTenors.Columns.Count) Then
            adHocTenorArray = WorksheetFunction.Transpose(rngAdHocTenors)
        Else
            adHocTenorArray = WorksheetFunction.Transpose(WorksheetFunction.Transpose(rngAdHocTenors))
        End If
        
        GetTenorColumnShift = UBound(adHocTenorArray) - TENORS_11_STANDARD
        Exit Function
    End If
    Select Case (TenorStandard)
        ' "ALL_TENORS"
        Case (TENORS_ALL_STANDARD_TAG):
            GetTenorColumnShift = TENORS_ALL_STANDARD_VALUE - TENORS_11_STANDARD
        
        ' VA 2009-06-22: Adding 12 Tenors
        ' "Regular" Tenor Standards
        Case TENORS_17_STANDARD, TENORS_13_STANDARD, TENORS_12_STANDARD, TENORS_11_STANDARD:
            GetTenorColumnShift = TenorStandard - TENORS_11_STANDARD
    End Select
End Function


' VA 2009-05-11: Nice to have...
Function ValidateAdHocTenors(AdHocTenors() As Variant, ValidationErrors() As String) As Boolean

End Function


' VA 2009-05-12: Assume we have only one-character on the right side ('M' or 'Y')
' "3M" --> "3:M"; "20Y" --> "20:Y"
Function ConvertTermsToTermStructure(TermNames() As Variant) As Variant()
    Dim arrayHelper As New CArrayHelper
    Dim TermStructure() As Variant
    Dim itemIndex As Integer
    Dim termName As String
    
    If (IsArray(TermNames) And arrayHelper.NumberOfArrayDimensions(TermNames) = 1 And Not arrayHelper.IsArrayEmpty(TermNames)) Then
        ReDim TermStructure(LBound(TermNames) To UBound(TermNames))
        For itemIndex = LBound(TermNames) To UBound(TermNames)
            termName = UCase(TermNames(itemIndex))
            TermStructure(itemIndex) = Left(termName, Len(termName) - 1) & ":" & Right(termName, 1)
        Next
    End If
    ConvertTermsToTermStructure = TermStructure
End Function

Function GetAdHocTenorsFromRange(AdHocTenorRange As Range) As Variant()

    Dim arrayHelper As New CArrayHelper
    Dim wrkArray() As Variant
    Dim tenorArray() As Variant
    Dim itemIndex As Integer
   
    ' VA 2009-05-12: Ad Hoc Tenor Range is passed
    If (AdHocTenorRange.Cells.Count <> 0) Then
        wrkArray = arrayHelper.GetArrayFromRange(AdHocTenorRange, 0, True)
        Call arrayHelper.GetRow(wrkArray, tenorArray, 0)
        
        ' Convert to upper case
        For itemIndex = 0 To UBound(tenorArray)
            tenorArray(itemIndex) = UCase(tenorArray(itemIndex))
        Next
    End If
    
    GetAdHocTenorsFromRange = tenorArray

End Function

Sub SendEmailToCms()
    'Call SendEmailMessage(CMS_MAIL_RECIPIENTS, CMS_MAIL_CC_LIST, CMS_MAIL_SUBJECT)
End Sub

'' VA 2009-05-15: Send Mails to CMS
'Public Const CMS_MAIL_RECIPIENTS = "CMS Client Support"
'Public Const CMS_MAIL_SUBJECT = "CMS Data Access Add-in issues: "
'Public Const CMS_MAIL_BODY = "CMS Client Support," & vbNewLine & vbNewLine




'' -----------------------------------------------------------------------------------------------------------
'' VA 2011-10-04: Temporarily disabled due to issues with Outlook 2003/2007 references.
'' --------------------------------------------------------
'' VA 2009-05-15: SendEmailMessage()
'' --------------------------------------------------------
'' VBA Send Outlook Email
'' http://snipplr.com/view/10041/vba-send-outlook-email/
'' --------------------------------------------------------
'' Insert Outlook Signature in mail
'' http://www.rondebruin.nl/mail/folder3/signature.htm
'' --------------------------------------------------------
'' http://www.eggheadcafe.com/forumarchives/officedeveloperoutlookVisualBasica/Sep2005/post23600061.asp
'' -----------------------------------------------------------------------------------------------------------
'Sub SendEmailMessage(MailRecipients As String, MailCCList As String, MailSubject As String, Optional MailHTMLBody, Optional DisplayMsg = True, Optional AttachmentPath)
'    Dim objOutlook As Outlook.Application
'    Dim objOutlookMsg As Outlook.MailItem
'    Dim objOutlookRecip As Outlook.Recipient
'    Dim objOutlookAttach As Outlook.Attachment
'    Dim compiledCCAdresses As String
'
'    ' Create the Outlook session.
'    Set objOutlook = CreateObject("Outlook.Application")
'
'    ' Create the message.
'    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
'
'    With objOutlookMsg
'        ' Add the To recipient(s) to the message.
'        Set objOutlookRecip = .Recipients.Add(MailRecipients)
'        objOutlookRecip.Type = olTo
'
'        ' Add the CC recipient(s) to the message.
'        compiledCCAdresses = objOutlookMsg.Session.CurrentUser
'        If (Len(MailCCList) <> 0) Then
'            compiledCCAdresses = compiledCCAdresses & ";" & MailCCList
'        End If
'        Set objOutlookRecip = .Recipients.Add(compiledCCAdresses)
'        objOutlookRecip.Type = olCC
'
''        ' Add the BCC recipient(s) to the message.
''        Set objOutlookRecip = .Recipients.Add("XXX")
''        objOutlookRecip.Type = olBCC
'
'       ' Set the Subject, Body, and Importance of the message.
'       .Subject = MailSubject
'       If Not IsMissing(MailHTMLBody) Then
'            .HTMLBody = MailHTMLBody & vbNewLine & vbNewLine & .HTMLBody
'       End If
'       .Importance = olImportanceNormal ' Normal
'
'       ' Add attachments to the message.
'       If Not IsMissing(AttachmentPath) Then
'           Set objOutlookAttach = .Attachments.Add(AttachmentPath)
'       End If
'
'       ' Resolve each Recipient's name.
'       For Each objOutlookRecip In .Recipients
'           objOutlookRecip.Resolve
'       Next
'
'       ' Should we display the message before sending?
'       If DisplayMsg Then
'           .Display
'       Else
'           .Save
'           .send
'       End If
'    End With
'    Set objOutlook = Nothing
'End Sub

Sub NavigateToWebPage(URL As String)
    Dim ieControl As InternetExplorer
    Set ieControl = New InternetExplorer
    ieControl.Visible = True
    ieControl.Navigate2 (URL)
End Sub

