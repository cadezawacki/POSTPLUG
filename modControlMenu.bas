
Attribute VB_Name = "modControlMenu"
' -----------------------------------------------------------------------------------
' VA 2009-04-14
' THIS CODE IS MOSTLY BASED ON CDAMenu.xla from CDAToolkit
' written by Richard Sharp (8773 9265)
' -----------------------------------------------------------------------------------

Option Explicit

' VA 2009-04-24: http://www.cpearson.com/Excel/Scope.aspx
Option Private Module

Private xlaPath As String

Public Sub SetDefaultCmsSettings()
    
    Dim webServiceHelper As New CWebServiceHelper
    
    'Debug.Print "SetDefaultCmsSettings: before setting url: " & ThisWorkbook.CMS_WEB_SERVICE_URL
    ThisWorkbook.CMS_WEB_SERVICE_URL = CMS_WSDL_PROD
    'Debug.Print "SetDefaultCmsSettings: after setting url: " & ThisWorkbook.CMS_WEB_SERVICE_URL
    
    ' No impersonation by default
    ThisWorkbook.IMPERSONATED_USER_NAME = ""
    
    ' The only thing that might be recovered from Registry
    ' is CMS_LAST_CUSTOM_WEB_SERVICE_URL
    ThisWorkbook.CMS_LAST_CUSTOM_WEB_SERVICE_URL = GetSetting(REG_CMS_APP_NAME, REG_CMS_COMMON_SECTION_NAME, "CMS_LAST_CUSTOM_WEB_SERVICE_URL")
    
    ' Call CMS WEB Service here
    ' to define whether user is allowed to config CMS & Impersonate
    Call webServiceHelper.GetUserAuthorization
    
    If (Not ThisWorkbook.CMS_AUTHORIZATION_SUCCESS) Then
        ' Set CMS Server to INVALID
        'Debug.Print "Authorization has failed. Setting server to invalid. " & ThisWorkbook.CMS_WEB_SERVICE_URL
        Set ThisWorkbook.CmsClient = Nothing
        ThisWorkbook.CMS_WEB_SERVICE_URL = ""
        ThisWorkbook.CMS_SERVER_MODE = "INVALID"
        
        ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN = False
        ThisWorkbook.IS_ALLOWED_TO_CONFIG = False
        ThisWorkbook.IS_ALLOWED_TO_IMPERSONATE = False
        
        MsgBox "CMS Authorization procedure failed." & vbNewLine & "Please try again using ""Reconnect to CMS"" Menu option." & _
                vbNewLine & vbNewLine & "If the error persists, contact CMS Client Support.", vbExclamation, "CMS AUTHORIZATION FAILED"
    
'''        ' -----------------------------------------------------------------
'''        ' VA 2009-05-21: PROD CLEANUP
'''        ' DEBUG 2009-05-06 - Comment it out for Production!
'''        MsgBox "VA 05/05/06: For Debugging, reset to SuperUser" & vbNewLine & vbNewLine & "FOR PRODUCTION, REMOVE IT FROM modControlMenu.SetDefaultCmsSettings()!!!", _
'''        vbExclamation, "DEBUG: CMS AUTHORIZATION RESET TO SUPERUSER"
'''        ThisWorkbook.CMS_WEB_SERVICE_URL = CMS_WSDL_PROD
'''        ThisWorkbook.CMS_SERVER_MODE = "PROD"
'''        ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN = True
'''        ThisWorkbook.IS_ALLOWED_TO_CONFIG = True
'''        ThisWorkbook.IS_ALLOWED_TO_IMPERSONATE = True
'''        ' -----------------------------------------------------------------
    
    ElseIf (ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN) Then
        ' Set CMS Server to PROD
        'Debug.Print "SetDefaultCmsSettings: before setting url to prod: " & ThisWorkbook.CMS_WEB_SERVICE_URL
        ThisWorkbook.CMS_WEB_SERVICE_URL = CMS_WSDL_PROD
        'Debug.Print "SetDefaultCmsSettings: after setting url to prod: " & ThisWorkbook.CMS_WEB_SERVICE_URL
        ThisWorkbook.CMS_SERVER_MODE = "PROD"
    Else
        ' Set CMS Server to INVALID
        'Debug.Print "Authorization hasn't failed but not allowed to use add-in. Setting server to invalid. " & ThisWorkbook.CMS_WEB_SERVICE_URL
        ThisWorkbook.CMS_WEB_SERVICE_URL = ""
        ThisWorkbook.CMS_SERVER_MODE = "INVALID"
        ThisWorkbook.IS_ALLOWED_TO_CONFIG = False
        ThisWorkbook.IS_ALLOWED_TO_IMPERSONATE = False
        
        MsgBox "Sorry, you are not authorized to use CMSDataAccess Add-in." & vbNewLine & vbNewLine & _
        "Please contact CMS Client Support.", vbExclamation

    End If
    
    ' Store current settings in Registry
    Call SaveCmsRegistrySettings
    
    Call frmAddinEnvironment.InitializeAddinEnvironmentForm
    
End Sub


Public Sub Auto_Open()
    'cmsxla_setupMenu
End Sub

Public Sub cmsxla_about()


End Sub




Public Sub cmsxla_displayHelp()
    Dim TheBrowser As Object
    Set TheBrowser = CreateObject("InternetExplorer.Application")
    TheBrowser.Visible = True
    Dim DocumentationPath As String
    DocumentationPath = ThisWorkbook.Path & "\" & CMS_HELP_FILE_NAME
    TheBrowser.Navigate URL:="file://" & DocumentationPath
End Sub

Public Sub cmsxla_setupMenu()

    Dim helpMenuI As Integer
    Dim thisFolder As Folder
    Dim menuPath As String
    Dim fs As Scripting.FileSystemObject
    Dim mainMenu As CommandBar, cdaMainMenu As CommandBarControl, submenu As CommandBarControl
    
    ' Delete command bar if it exists already
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls(CMS_MENU_TITLE).Delete

    On Error GoTo ErrorHandler
    Set fs = CreateObject("Scripting.FileSystemObject")
    xlaPath = ThisWorkbook.Path & "\"
'    menuPath = xlaPath & "menu\"
'    Set thisFolder = fs.GetFolder(menuPath)
    Set mainMenu = Application.CommandBars("Worksheet Menu Bar")
    helpMenuI = mainMenu.Controls("Help").index
    Set cdaMainMenu = mainMenu.Controls.Add(Type:=msoControlPopup, Before:=helpMenuI, temporary:=True)
    With cdaMainMenu
        .Caption = CMS_MENU_TITLE
    End With
    
    With cdaMainMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
        .Caption = "Check/Set CMS Environment"
        .OnAction = "cmsxla_SetCmsEnvironment"
    End With
    
    With cdaMainMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
        .Caption = "Reconnect to CMS"
        .OnAction = "ReconnectToCMS"
    End With
    
    
    Set submenu = cdaMainMenu.Controls.Add(Type:=msoControlPopup, temporary:=True)
    With submenu
        .Caption = "Help"
    End With
    With submenu.Controls.Add(Type:=msoControlButton, temporary:=True)
        .Caption = "CMSDataAccess Help"
        .OnAction = "cmsxla_displayHelp"
    End With
'    With submenu.Controls.Add(Type:=msoControlButton, temporary:=True)
'        .Caption = "About"
'        .OnAction = "cmsxla_about"
'
'    End With
    
    Exit Sub
ErrorHandler:
    MsgBox "Error occured creating CMS menu" & vbNewLine & vbNewLine & Err.Description, vbOKOnly, "Error:"
End Sub

Sub cmsxla_SetCmsEnvironment()
    ' VA 2009-04-28: If Excel has a "loss of state", it cleans up global variables!
    ' http://www.mrexcel.com/forum/showthread.php?p=1662217
    

    Call GetCmsRegistrySettings
    
    Call frmAddinEnvironment.InitializeAddinEnvironmentForm
    
    If (Not ThisWorkbook.IS_ALLOWED_TO_USE_ADDIN) Then
        MsgBox "Sorry, you are not authorized to use CMSDataAccess Add-in." & vbNewLine & vbNewLine & _
               "Please contact CMS Client Support.", vbExclamation
        ' Exit Sub
    End If
    frmAddinEnvironment.Show
End Sub

Sub cmsxla_refreshMenu()
    Call cmsxla_deleteMenu
    Call cmsxla_setupMenu
End Sub

Sub cmsxla_deleteMenu()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls(CMS_MENU_TITLE).Delete
    On Error GoTo 0
End Sub


Public Sub SetControlMenu()
 Call modControlMenu.SetDefaultCmsSettings
 Call modControlMenu.cmsxla_refreshMenu
End Sub

Public Sub ReconnectToCMS()
    ' VA 2009-07-16: Found catch-22...
    'Debug.Print "ReconnectToCMS: before setting url to prod: " & ThisWorkbook.CMS_WEB_SERVICE_URL
    ThisWorkbook.CMS_WEB_SERVICE_URL = CMS_WSDL_PROD
    'Debug.Print "ReconnectToCMS: after setting url to prod: " & ThisWorkbook.CMS_WEB_SERVICE_URL
    
    Call modControlMenu.SetDefaultCmsSettings
    frmAddinEnvironment.Show
End Sub








