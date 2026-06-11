
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWebServiceHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Started migration to class-base implementation...

Option Explicit


' VA 2009-05-05: CMS throws an exception if the user is "unknown".
Friend Function GetUserAuthorization() As Boolean
    
    'MSXML Dom object for processing the response
    Dim domUserAuthorization As MSXML2.DOMDocument
    Dim success As Boolean
    
    Dim nodestatus As MSXML2.IXMLDOMNode
    Dim LastMarkedOn As MSXML2.IXMLDOMNode
    Dim LastMarkedBy As MSXML2.IXMLDOMNode
    Dim Owner        As MSXML2.IXMLDOMNode
    Dim IsParent     As MSXML2.IXMLDOMNode
    
    Dim nodeCmsAddinAuthorized As MSXML2.IXMLDOMNode
    Dim nodeCmsConfigAuthorized As MSXML2.IXMLDOMNode
    Dim nodeImpersonationAuthorized As MSXML2.IXMLDOMNode
 
'    Dim cmsAddinAuthorized As Boolean
'    Dim cmsConfigAuthorized As Boolean
'    Dim impersonationAuthorized As Boolean
    
    Dim rosettaUserMetaData As String
    Dim rosettaXmlRequest As String
    Dim cmsResponse As String
    
    
    On Error GoTo ERROR_HANDLING
    
    ' VA 2009-05-05: By Paresh's request, no access by default
    ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN = False
    ThisWorkbook.IS_ALLOWED_TO_CONFIG = False
    ThisWorkbook.IS_ALLOWED_TO_IMPERSONATE = False
    
    '
    rosettaUserMetaData = CompileUserMetaData
    
'    ' Debug
'    Debug.Print "GetCMSUsersByQuery: User Metadata: "
'    Debug.Print rosettaUserMetaData
    
    rosettaXmlRequest = CompileRosettaGetUsersByQueryRequest
    
'    Debug.Print "GetCMSUsersByQuery: Rosetta XML Request: "
'    Debug.Print rosettaXmlRequest
    
    ' VA 2009-05-04: Authorization is a 'special" call that should always go to PROD.
    cmsResponse = GetCmsXmlResponse(rosettaUserMetaData, rosettaXmlRequest, CMS_WS_GET_MARKET_DATA_GetCMSUsersByQuery_CALL, CMS_WSDL_PROD)
    
'    ' Debug
'    Debug.Print "GetCMSUsersByQuery: Rosetta XML Response: "
'    Debug.Print cmsResponse
'
'    ' VA 2009-05-01: Temporarily, before Paresh fixes the Response
'    ' cmsResponse = "<cmsusers><cmsuser id=""1"" userid=""abramovv"" firstname=""Victor"" lastname=""Abramovich"" status=""ACTIVE"" email=""victor.abramovich@barclayscapital.com"" region=""New York"" cmsaddin=""true"" impersonation=""true"" cmsurl=""http://cms-lxp.lehman.com""/></cmsusers>"
'    Debug.Print "Rosetta XML Response:"
'    Debug.Print cmsResponse
    
    ' See GetMarketData() - XML DOM, etc.
    Set domUserAuthorization = New MSXML2.DOMDocument
    domUserAuthorization.async = False
    domUserAuthorization.validateOnParse = False
    success = domUserAuthorization.LoadXML(cmsResponse)
    
    ' Add-in usage authorization
    Set nodeCmsAddinAuthorized = domUserAuthorization.selectSingleNode("//cmsusers/cmsuser/@cmsaddin")
    If (Not nodeCmsAddinAuthorized Is Nothing) Then
        ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN = CBool(nodeCmsAddinAuthorized.Text)
    End If

    ' Add-in configuration authorization
    Set nodeCmsConfigAuthorized = domUserAuthorization.selectSingleNode("//cmsusers/cmsuser/@cmsenvironment")
    If (Not nodeCmsAddinAuthorized Is Nothing) Then
        ThisWorkbook.IS_ALLOWED_TO_CONFIG = CBool(nodeCmsAddinAuthorized.Text)
    End If
    
    ' Impersonation authorization
    Set nodeImpersonationAuthorized = domUserAuthorization.selectSingleNode("//cmsusers/cmsuser/@impersonation")
    If (Not nodeImpersonationAuthorized Is Nothing) Then
        ThisWorkbook.IS_ALLOWED_TO_IMPERSONATE = CBool(nodeImpersonationAuthorized.Text)
    End If
    
    ThisWorkbook.CMS_AUTHORIZATION_SUCCESS = True
    GetUserAuthorization = ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN
    Exit Function
        
ERROR_HANDLING:
    ThisWorkbook.CMS_AUTHORIZATION_SUCCESS = False
    ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN = False
    ThisWorkbook.IS_ALLOWED_TO_CONFIG = False
    ThisWorkbook.IS_ALLOWED_TO_IMPERSONATE = False
    GetUserAuthorization = ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN
End Function


Friend Function CompileUserMetaData(Optional ByVal ImpersonatedUserName As String) As String

    ' Get UserName - might be impersonated
    If (ImpersonatedUserName = "") Then
        ImpersonatedUserName = CMS_GetUserLoginName()
    End If
    
    ' Encode the ActiveWorkbook.FullName - might break XML parsing on CMS side!
    ' VA 2009-01-28: Rolled back...
    ' VA 2009-04-14: ... back in business
    ' VA 2009-07-13: added <application-version>
    Dim workbookFullNameEncoded As String
    If (Not ActiveWorkbook Is Nothing) Then
        workbookFullNameEncoded = HTMLEncode(ActiveWorkbook.FullName)
    Else
        workbookFullNameEncoded = CMS_GetApplicationId()
    End If
    
    CompileUserMetaData = "<user-metadata><user-details>" _
                        + "<user-id>" + ImpersonatedUserName + "</user-id>" _
                        + "<user-domain>" + CMS_GetUserDomainName() + "</user-domain>" _
                        + "<application-id>" + APP_ID + "</application-id>" _
                        + "<application-version>" + ADDIN_VERSION + "</application-version>" _
                        + "<application-batch-id>" + workbookFullNameEncoded + "</application-batch-id>" _
                        + "<id>" + CMS_GetUserLoginName() + "</id>" _
                        + "<host-id>" + CMS_GetUserMachineName() + "</host-id>" _
                        + "</user-details></user-metadata>"
'    Debug.Print CompileGetUserMetaData
End Function



Friend Function CompileRosettaGetUsersByQueryRequest(Optional ByVal ImpersonatedUserName As String) As String
    
    ' Get UserName - might be impersonated
    If (ImpersonatedUserName = "") Then
        ImpersonatedUserName = CMS_GetUserLoginName()
    End If
    
    CompileRosettaGetUsersByQueryRequest = "<cmsuser-queries><cmsuser-query userid=""" & ImpersonatedUserName & """/></cmsuser-queries>"
End Function

' ============================================================================================================================
' ============================================================================================================================
' ============================================================================================================================


' getCMSUsersByQuery

' ====================================================================
' CMS WEB Service: See HY_CDX & LCDX Blotters by Victor Abramovich
' ====================================================================
Friend Function GetCmsXmlResponse(rosettaUserMetaData As String, rosettaXmlRequest As String, cmsWSCallName As String, Optional cmsWebServiceUrl As String = "") As String

   'CMS connection object
    Dim cms As MSOSOAPLib30.SoapClient30
    
    On Error GoTo ERROR_HANDLING
    
    If (ThisWorkbook.CMS_WEB_SERVICE_URL = "") Then
      Debug.Print "GetCmsXmlResponse: url is empty"
      Call RaiseError("Environment got corrupted. Please reconnect to CMS using the menu bar")
    End If
    
    ' Recover cmsWebServiceUrl
    If (Len(cmsWebServiceUrl) = 0) Then
        cmsWebServiceUrl = ThisWorkbook.CMS_WEB_SERVICE_URL
    End If
    
    ' Just in case...
        If (Len(cmsWebServiceUrl) = 0) Then
        cmsWebServiceUrl = CMS_WSDL_PROD
    End If

    Set cms = New MSOSOAPLib30.SoapClient30
       
   ' VA 2009-04-08 D-Day: Timeout issue
   'cms.ClientProperty("ServerHTTPRequest") = True
   
   'Debug.Print "GetCmsXmlResponse: before initializing soap: " & ThisWorkbook.CMS_WEB_SERVICE_URL & " soap url is " & cmsWebServiceUrl
   Call cms.MSSoapInit(cmsWebServiceUrl)
   
   cms.ConnectorProperty("Timeout") = CMS_SERVER_GET_TIMEOUT
   
    Select Case (cmsWSCallName):
   
        Case CMS_WS_GET_MARKET_DATA_GetMarketData_CALL:
            GetCmsXmlResponse = cms.GetMarketData(rosettaUserMetaData, rosettaXmlRequest)
            Exit Function
        Case CMS_WS_GET_MARKET_DATA_SetMarketData_CALL:
            GetCmsXmlResponse = cms.SetMarketData(rosettaUserMetaData, rosettaXmlRequest)
            Exit Function
            
        ' VA 2009-05-04: Special Case - moved to CWebServiceHelper
        Case CMS_WS_GET_MARKET_DATA_GetCMSUsersByQuery_CALL:
            ' VA 2009-05-04: To avoid "catch-22", we always get authorization data from PROD...
            ' ... yet for now (development) it's DEV
            GetCmsXmlResponse = cms.GetCMSUsersByQuery(rosettaUserMetaData, rosettaXmlRequest)
            Exit Function
default:
         Call RaiseError("Add-in CMS WS: Unknown Function call '" & "cmsWSCallName" & ".")
   End Select
   

'    ' DEBUG
'     Debug.Print "Rosetta XML Response: "
'     Debug.Print GetCmsXmlResponse
'    MsgBox GetCmsXmlResponse

ERROR_HANDLING:
         Call RaiseError("GetCmsXmlResponse(): CMS WEB Service call Failed:" & vbNewLine & vbNewLine & Err.Description)
End Function

Friend Sub RaiseError(errDesc As String)
    Call Err.Raise(-1, APP_ID, errDesc)
End Sub


