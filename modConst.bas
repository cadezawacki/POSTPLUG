
Attribute VB_Name = "modConst"
' ------------------------------------------------------
' SET ENVIRONMENT CORRECTLY
' AND UPDATE THE ADD-IN TITLE ACCORDINGLY!!!
' ------------------------------------------------------

Option Explicit

' VA 2009-04-24: http://www.cpearson.com/Excel/Scope.aspx
Option Private Module

' VA 2008-11-21
' Application ID
Public Const APP_ID = "CMS - Excel Addin"

' VA 2008-12-23 - Set it Properly in file properties!
' ThisWorkbook.IsAddin = False/True
Public Const ADDIN_VERSION = "5.0.1"

' -------------------------------------------------------
' VA 2009-06-23: DEBUG SETTINGS
' VA 2010-07-26: Make it configurable
' -------------------------------------------------------
Public DEBUG_PRINT_CMS_USER_METADATA As Boolean
Public DEBUG_PRINT_CMS_ROSETTA_XML As Boolean

' VA 2011-03-30: Specific Reference Curve Support
Public Const SPEC_REF_CURVE_PRODUCT = "SPECREF"
Public Const SPEC_REF_DEBT_CLASS_CONCAT_CHAR = "~"
Public Const REG_DEBT_CLASS_CONCAT_CHAR = "_"

' VA 2011-08-11: Issues with different versions of Outlook installed:
' hardcoded reference to 2007 in VBA project might be broken!
Public Const OFFICE_11_OUTLOOK_FILE_PATH = "C:\Program Files\Microsoft Office\Office11\MSOUTL.OLB"
Public Const OFFICE_12_OUTLOOK_FILE_PATH = "C:\Program Files\Microsoft Office\Office12\MSOUTL.OLB"

' TODO - clean up the obsolete ones!
' Enum
Public Enum ImpliedQuoteTypes
    QuoteType_Unknown = 0
    QuoteType_NoQuoteData
    QuoteType_100Spreads
    QuoteType_100Upfronts
    QuoteType_100Both
    QuoteType_500Spreads
    QuoteType_500Upfronts
    QuoteType_500Both
    
    QuoteType_ParSpreads
    QuoteType_ParUpfronts
    QuoteType_ContractualSpread
    QuoteType_RunningSpread
    
    ' VA 2009-06-18: Updated for STANDARD
    QuoteType_StandardContractSpreads
    QuoteType_StandardContractUpfronts
    QuoteType_StandardContractBoth
End Enum

Public Type TermQuoteData
    Term As String
    TermCreditSpread As Variant
    TermSpreadContractual As Variant
    TermCreditUpfront As Variant
    TermUpfrontContractual As Variant
End Type

' VA 2009-05-18: CMS Help
Public Const CMS_MENU_TITLE = "C&MS"
Public Const CMS_HELP_FILE_NAME = "CMSDataAccessHelp.mht"


' VA 2009-05-15: URLs to navigate
Public Const BARCAP_URL = "http://home.barcapint.com/"
Public Const CMS_URL = "http://my.lehman.com/CRD/cms/jsp/CurveMainPage.jsp"


' VA 2009-05-15: Send Mails to CMS
Public Const CMS_MAIL_RECIPIENTS = "CMS Client Support"
Public Const CMS_MAIL_CC_LIST = ""
Public Const CMS_MAIL_SUBJECT = "CMS Data Access Add-in Issues: "
' Public Const CMS_MAIL_HTML_BODY = ""


' VA 2009-05-01: Let's make CMS WS calls generic!
Public Const CMS_WS_GET_MARKET_DATA_GetMarketData_CALL = "GetMarketData"
Public Const CMS_WS_GET_MARKET_DATA_SetMarketData_CALL = "SetMarketData"
Public Const CMS_WS_GET_MARKET_DATA_GetCMSUsersByQuery_CALL = "GetCMSUsersByQuery"

' VA 2009-04-29: Now we store session variables in Registry
Public Const REG_CMS_APP_NAME = "CMS.DataAccessAddIn"
Public Const REG_CMS_SESSION_SECTION_NAME = "Sessions"
Public Const REG_CMS_COMMON_SECTION_NAME = "Common"


' VA 2010-08-09: 5Tuplets For new Bulk GET requests
Public Const BULK_REQUEST_TICKER_INDEX = 1
Public Const BULK_REQUEST_CCY_INDEX = 2
Public Const BULK_REQUEST_DEBT_CLASS_INDEX = 3
Public Const BULK_REQUEST_PRODUCT_INDEX = 4
Public Const BULK_REQUEST_QUOTE_CONVENTION_INDEX = 5

' VA 2010-08-16: "Too many errors" CMS message
Public Const TOO_MANY_ERRORS_CMS_MESSAGE = "Too many errors while processing request. Please check the request."

' ==========================================================================================
' CMS Web Service
' ==========================================================================================
'Using DEV WSDL for development
Public Const CMS_WSDL_DEV = "http://cms-lxd.lehman.com/CreditMarkingService/?WSDL"

'Using STAGE WSDL for development
Public Const CMS_WSDL_STAGE = "http://cms-lxs.lehman.com/CreditMarkingService/?WSDL"

' Using PROD WSDL for production
Public Const CMS_WSDL_PROD = "http://cms-lxp.lehman.com/CreditMarkingService/?WSDL"



' Standard Java Error -  to remove from Error messages
Public Const STANDARD_JAVA_ERROR = "Exception during processing: javax.ejb.EJBException: Unexpected exception: javax.ejb.EJBException:"

' ==========================================================================================

' 2009-02-20: Tenor Standards
Public Const TENORS_11_STANDARD = 11
Public Const TENORS_12_STANDARD = 12
Public Const TENORS_13_STANDARD = 13
Public Const TENORS_17_STANDARD = 17


' VA 2009-06-22: Adding 12 Standard
' VA 2009-05-08: Extended tenor list + dynamic request
' Replace with next extended standard!
Public Const TENORS_ALL_STANDARD_TAG = "ALL_TENORS"
Public Const TENORS_ALL_STANDARD_VALUE = TENORS_17_STANDARD

Public Const DEFAULT_TENOR_STANDARD = TENORS_11_STANDARD



'---------------------------------------------------------
' Credit Curve Convention
'---------------------------------------------------------
Public Const CURVE_CONVENTION_UNKNOWN = "UNKNOWN"
Public Const CURVE_CONVENTION_SPREADS = "SPREADS"
Public Const CURVE_CONVENTION_UPFRONT = "UPFRONT"
Public Const CURVE_CONVENTION_QUOTED = "QUOTED"
' 2009-02-20: SNAC
Public Const CURVE_CONVENTION_QUOTEDSPREADS = "QUOTED_SPREADS"

Public Const CURVE_CONVENTION_RUNNING = "RUNNING"
' "EXTENDED" = "UPFRONT" + "RUNNING"
Public Const CURVE_CONVENTION_EXTENDED = "EXTENDED"

'---------------------------------------------------------
' Name-Value pair tags
'---------------------------------------------------------
Public Const NAME_VALUE_OWNER_TAG = "owner"
Public Const NAME_VALUE_LAST_MARKED_ON_TAG = "marked"
Public Const NAME_VALUE_LAST_MARKED_BY_TAG = "lastmarkedby"
Public Const NAME_VALUE_IS_PARENT_TAG = "isparent"
Public Const NAME_VALUE_INTERPOLATION_TYPE_TAG = "interpolationtype"

'---------------------------------------------------------
' Status
'---------------------------------------------------------
Public Const REQUEST_STATUS_FAILED = "FAILED"
Public Const REQUEST_STATUS_OK = "OK"

'---------------------------------------------------------
' Column Structure
'---------------------------------------------------------

' VA 2009-07-14:
Public Const CMS_SCLR_COL_COUNT = 8

' VA 2009-03-08: Indicies for "scalar" array
Public Const CMS_SCLR_RECOVERY = 0
Public Const CMS_SCLR_LAST_MARKED_ON = 1
Public Const CMS_SCLR_LAST_MARKED_BY = 2
Public Const CMS_SCLR_OWNER = 3
Public Const CMS_SCLR_IS_PARENT = 4
Public Const CMS_SCLR_CURVE_TYPE = 5
Public Const CMS_SCLR_STATUS = 6
Public Const CMS_SCLR_ERROR = 7

' VA 2009-06-22: Adding 12 Tenors
' VA 2009-02-23: These indices are for "default" Tenor standard (11)
Public Const CMS_TENOR_12_COL_SHIFT = 1
Public Const CMS_TENOR_13_COL_SHIFT = 2
Public Const CMS_COL_3M = 0
Public Const CMS_COL_6M = 1
Public Const CMS_COL_1Y = 2
Public Const CMS_COL_2Y = 3
Public Const CMS_COL_3Y = 4
Public Const CMS_COL_4Y = 5
Public Const CMS_COL_5Y = 6
Public Const CMS_COL_7Y = 7
Public Const CMS_COL_10Y = 8
Public Const CMS_COL_20Y = 9
Public Const CMS_COL_30Y = 10
Public Const CMS_COL_RECOVERY = 11
Public Const CMS_COL_LAST_MARKED_ON = 12
Public Const CMS_COL_LAST_MARKED_BY = 13
Public Const CMS_COL_OWNER = 14
Public Const CMS_COL_IS_PARENT = 15
Public Const CMS_COL_CURVE_TYPE = 16
Public Const CMS_COL_STATUS = 17
Public Const CMS_COL_ERROR = 18

Public Const CMS_COL_FLAT = 6
Public Const CMS_TERM_COUNT = 11
Public Const CMS_COL_TERM_START = 0
Public Const CMS_COL_COUNT = 19


' Term & Risk Data Columns
Public Const CMS_TERM_DATA_COL_COUNT = 9
Public Const CMS_RISK_DATA_COL_COUNT = 8
Public Const CMS_TERM_COL_QUOTE = 0
Public Const CMS_TERM_COL_RECOVERY = 1
Public Const CMS_TERM_COL_LAST_MARKED_ON = 2
Public Const CMS_TERM_COL_LAST_MARKED_BY = 3
Public Const CMS_TERM_COL_OWNER = 4
Public Const CMS_TERM_COL_IS_PARENT = 5
Public Const CMS_TERM_COL_CURVE_TYPE = 6
Public Const CMS_TERM_COL_STATUS = 7
Public Const CMS_TERM_COL_ERROR = 8

'---------------------------------------------------------
' Data Types
'---------------------------------------------------------
Public Const CREDIT_SPREAD_DATA_TYPE = "creditSpread"
Public Const CREDIT_UPFRONT_DATA_TYPE = "creditUpfront"
Public Const ISSUER_RECOVERY_DATA_TYPE = "issuerRecovery"

' VA 2009-06-18: Now we care about STANDART_CONTRACT & NON_STANDART_CONTRACT curves


' VA 2009-02-27
Public Const SNAC_100_PRODUCT = "SNAC100"
Public Const SNAC_500_PRODUCT = "SNAC500"


Public Const CONTRACTUAL_SPREAD_TAG = "contractualSpread"
Public Const CONTRACTUAL_SPREAD_100 = 0.01
Public Const CONTRACTUAL_SPREAD_500 = 0.05

' VA 2009-06-19: for SET requests, we don't care about the CONTRACTUAL_SPREAD value anymore...
Public Const CONTRACTUAL_SPREAD_DUMMY = -777

'---------------------------------------------------------
' Default Settings
'---------------------------------------------------------
Public Const DEFAULT_DELIMITER = "."
Public Const DEFAULT_LOCATION = "NYC"
Public Const DEFAULT_CCY = "USD"
Public Const DEFAULT_PRODUCT = "DERIV"
Public Const DEFAULT_CURVE_CONVENTION = "SPREADS"
Public Const DEFAULT_DATA_SOURCE = "LIVE"
Public Const DEFAULT_DATA_TYPE = "creditSpread"
Public Const DEFAULT_LOGICAL_TIME = "LIVE"

'---------------------------------------------------------
' Misc
'---------------------------------------------------------
Public Const DUMMY_DOUBLE_VALUE = -77777#


' VA 2009-04-08 D-Day: Timeout issue
Public Const CMS_SERVER_GET_TIMEOUT = 180000
Public Const CMS_SERVER_SET_TIMEOUT = 180000

' VA 2009-04-15: Set CMS Environment Form
Public Const CMS_ENV_FORM_HEADER1 = "CMS Data Access Add-in"
Public Const CMS_ENV_FORM_HEADER2 = "ENVIRONMENT SETTINGS"



