
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMarketData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' VA 2009-03-08: A class to process CMS XML responses
'------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------
' Private Members
'------------------------------------------------------------------

' An instance of CArrayHelper
Private arrayHelper As New CArrayHelper

' Consts
Private Const SCALAR_DATA_COUNT = 8


' Tenor Standard (we'll default it to 11)
Public TenorStandard As Variant

' 4Typle et al.
Public fourTuple As String
Public Ticker As String
Public Ccy As String
Public debtClass As String
Public Product As String

' VA 2009-06-18: Now we are going to handle STANDARD_CONTRACT curves (SNAC & STEC at the moment),
' and all the others (just DERIV for now).
Public IsStandardContractCurve As Boolean

'----------------------------------------------------
' Arrays (have to be private)
'----------------------------------------------------
Private m_strFourTupleArray() As Variant
Private m_vaTermStructure() As Variant
Private m_vaCreditSpreads() As Variant
Private m_vaSpreadContractuals() As Variant
Private m_vaCreditUpfronts() As Variant
Private m_vaUpfrontContractuals() As Variant

Private m_va100RunningSpreads() As Variant
Private m_va500RunningSpreads() As Variant

Private m_vaImpliedQuoteTypeNames() As Variant

' All the "scalar" data - Recovery, Last Marked on, etc.
Private m_vaScalarData() As Variant

' "Ultimate" array, containing the quote array, "scalar" data, status, etc.
Private m_vaMarketDataQuotes() As Variant

' "Ultimate" array, containing the Spreads array, "scalar" data, status, etc.
Private m_vaMarketDataSpreads() As Variant

' "Ultimate" array, containing the Upfronts array, "scalar" data, status, etc.
Private m_vaMarketDataUpfronts() As Variant

' "Ultimate" array, containing the Contractuals/500 Running & "scalar" data, status, etc.
Private m_vaMarketDataContractuals() As Variant

'----------------------------------------------
' Booleans
'----------------------------------------------
'Private m_bHasCreditSpreads As Boolean
'Private m_bHasSpreadContractuals As Boolean
'Private m_bHasCreditUpfronts As Boolean
'Private m_bHasUpfrontContractuals As Variant

Public HasCreditSpreads As Boolean
Public HasSpreadContractuals As Boolean
Public HasCreditUpfronts As Boolean
Public HasUpfrontContractuals As Variant

' Quote Type
Public ImpliedQuoteType As ImpliedQuoteTypes
Public ImpliedQuoteTypeName As String

'----------------------------------------------
' "Scalar" Data: Recovery    Last Marked On  Last Marked By  Owner   Is Parent   Curve Type  Status  Error Message
'----------------------------------------------
'Private m_dRecoveryRate As Double
'Private m_dtLastMarkedOn As Date
'Private m_strLastMarkedBy As String
'Private m_strOwner As String
'Private m_bIsParent As Boolean
'Private m_strCurveType As Boolean
'Private m_strStatus As String
'Private m_strErrorMessage As String

Public IssuerRecoveryRate As Double
Public LastMarkedOn As Date
Public LastMarkedBy As String
Public Owner As String
Public IsParent As Boolean
Public status As String
Public errorMessage As String

'------------------------------------------------------
' Public Properties
' http://www.ozgrid.com/forum/showthread.php?t=30333
'------------------------------------------------------

' FourTupleArray
Public Property Get fourTupleArray(index As Integer) As Variant
    fourTupleArray(index) = m_strFourTupleArray(index)
End Property
'Public Property Let FourTupleArray(Index As Integer, Value As Variant)
'    m_strFourTupleArray(Index) = Value
'End Property

' TermStructure
Public Property Get TermStructure(index As Integer) As Variant
    TermStructure(index) = m_vaTermStructure(index)
End Property
'Public Property Let TermStructure(Index As Integer, Value As Variant)
'    m_vaTermStructure(Index) = Value
'End Property

' Private m_vaCreditSpreads() As Variant
Public Property Get CreditSpreads(index As Integer) As Variant
    CreditSpreads(index) = m_vaCreditSpreads(index)
End Property
Public Property Let CreditSpreads(index As Integer, Value As Variant)
    m_vaCreditSpreads(index) = Value
End Property

' Private m_vaSpreadContractuals() As Variant
Public Property Get SpreadContractuals(index As Integer) As Variant
    SpreadContractuals(index) = m_vaSpreadContractuals(index)
End Property
Public Property Let SpreadContractuals(index As Integer, Value As Variant)
    m_vaSpreadContractuals(index) = Value
End Property

' Private m_vaCreditUpfronts() As Variant
Public Property Get CreditUpfronts(index As Integer) As Variant
    CreditUpfronts(index) = m_vaCreditUpfronts(index)
End Property
Public Property Let CreditUpfronts(index As Integer, Value As Variant)
    m_vaCreditUpfronts(index) = Value
End Property

' Private m_vaUpfrontContractuals() As Variant
Public Property Get UpfrontContractuals(index As Integer) As Variant
    UpfrontContractuals(index) = m_vaUpfrontContractuals(index)
End Property
Public Property Let UpfrontContractuals(index As Integer, Value As Variant)
    m_vaUpfrontContractuals = Value
End Property

' Private m_vaScalarData() As Variant
Public Property Get MarketScalarData() As Variant
    MarketScalarData = m_vaScalarData
End Property

' "Ultimate" array, containing the quote arrays, "scalar" data, status, etc.
' I.e., it's either Spreads or Upfronts depending on the request.
' Private m_vaMarketDataQuotes() As Variant
Public Property Get MarketDataQuotes() As Variant
    MarketDataQuotes = m_vaMarketDataQuotes
End Property

' "Ultimate" array, containing the Spreads array, "scalar" data, status, etc.
' Private m_vaMarketDataSpreads() As Variant
Public Property Get MarketDataSpreads() As Variant
    MarketDataSpreads = m_vaMarketDataSpreads
End Property

' "Ultimate" array, containing the Upfronts array, "scalar" data, status, etc.
' Private m_vaMarketDataUpfronts() As Variant
Public Property Get MarketDataUpfronts() As Variant
    MarketDataUpfronts = m_vaMarketDataUpfronts
End Property

' "Ultimate" array, containing the Contractuals/500 Running & "scalar" data, status, etc.
' Private m_vaMarketDataContractuals() As Variant
Public Property Get MarketDataContractuals() As Variant
    MarketDataContractuals = m_vaMarketDataContractuals
End Property

' ImpliedQuoteTypeNames
Public Property Get ImpliedQuoteTypeNames() As Variant
    ImpliedQuoteTypeNames = m_vaImpliedQuoteTypeNames
End Property


'------------------------------------------------------------------------------
' "Default" Ctor - defaulted to DEFAULT_TENOR_STANDARD = TENORS_11_STANDARD
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Call Initialize
End Sub

' VA 2009-05-12: Now we process ad hoc Tenors!
' Initialize arrays, etc. - TENORS_11_STANDARD as default
Public Sub Initialize(Optional InputTenorStandard As Variant = DEFAULT_TENOR_STANDARD)
    Dim index As Integer
    Dim arrayLength As Long
    
    ' Ad hoc Tenors - InputTenorStandard is a Range
    ' Again, we are NOT validating the Range here!
    ' VA 2009-05-12: TODO: Move it to reusable Procedure!
    ' NOTE: For some reason passed Range(1, N) it transformed to Variant array(1,N) in InputTenorStandard...
    If IsArray(InputTenorStandard) Then
        
        If (UBound(InputTenorStandard, 1) > UBound(InputTenorStandard, 2)) Then
            ' This is a "column"
            Call arrayHelper.GetColumn(InputTenorStandard, m_vaTermStructure, 1)
        Else
            ' It's a "row"
            Call arrayHelper.GetRow(InputTenorStandard, m_vaTermStructure, 1)
        End If
        
        ' Resize from [1:N] to [0:N-1]
        Me.TenorStandard = UBound(m_vaTermStructure)
        Call arrayHelper.ChangeBoundsOfArray(m_vaTermStructure, 0, TenorStandard - 1)
        m_vaTermStructure = ConvertTermsToTermStructure(m_vaTermStructure)
        
    Else
        ' "Regular" scalar cases: 11, 13, and 17 as of 2009-09-12.
        Me.TenorStandard = InputTenorStandard
        Select Case (TenorStandard)
        Case TENORS_17_STANDARD:
            m_vaTermStructure = Array("0:M", "3:M", "6:M", "9:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "6:Y", "7:Y", "8:Y", "9:Y", "10:Y", "15:Y", "20:Y", "30:Y")
        Case TENORS_13_STANDARD:
            m_vaTermStructure = Array("0:M", "3:M", "6:M", "9:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
        Case TENORS_12_STANDARD:
            m_vaTermStructure = Array("0:M", "3:M", "6:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
        Case TENORS_11_STANDARD:
default:
            m_vaTermStructure = Array("3:M", "6:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
    End Select

End If
    
'    ' Implied Quote Type Names
'    QuoteType_Unknown = 0
'    QuoteType_NoQuoteData
'    QuoteType_100Spreads
'    QuoteType_100Upfronts
'    QuoteType_100Both
'    QuoteType_500Spreads
'    QuoteType_500Upfronts
'    QuoteType_500Both
'    QuoteType_ParSpreads
'    QuoteType_ParUpfronts
'    QuoteType_ContractualSpread
'    QuoteType_RunningSpread
'    ' VA 2009-06-18: Updated for STANDARD
'    QuoteType_StandardContractSpreads
'    QuoteType_StandardContractUpfronts
'    QuoteType_StandardContractBoth

    m_vaImpliedQuoteTypeNames = Array("UNKNOWN", "NO_QUOTE_DATA", "SNAC100_SPREADS", "SNAC100_UPFRONTS", "SNAC100_BOTH", _
                                      "SNAC500_SPREADS", "SNAC500_UPFRONTS", "SPREADS", "UPFRONTS", "SNAC500_BOTH", _
                                      "RUNNING_SPREADS", "RUNNING_SPREADS", "QUOTED_SPREADS", "UPFRONT", "BOTH")
    
    ReDim m_vaCreditSpreads(0 To TenorStandard - 1)
    ReDim m_vaSpreadContractuals(0 To TenorStandard - 1)
    ReDim m_vaCreditUpfronts(0 To TenorStandard - 1)
    ReDim m_vaUpfrontContractuals(0 To TenorStandard - 1)
    
    ReDim m_vaScalarData(0 To SCALAR_DATA_COUNT - 1)
    
    ' "Output" arrays
    ReDim m_vaMarketDataQuotes(0 To TenorStandard + SCALAR_DATA_COUNT - 1)
    ReDim m_vaMarketDataSpreads(0 To TenorStandard + SCALAR_DATA_COUNT - 1)
    ReDim m_vaMarketDataUpfronts(0 To TenorStandard + SCALAR_DATA_COUNT - 1)
    ReDim m_vaMarketDataContractuals(0 To TenorStandard + SCALAR_DATA_COUNT - 1)
    
    ' Initialize m_va100RunningSpreads & m_va500RunningSpreads
    ReDim m_va100RunningSpreads(0 To TenorStandard - 1)
    ReDim m_va500RunningSpreads(0 To TenorStandard - 1)
    For index = 0 To TenorStandard - 1
           m_va100RunningSpreads(index) = 100#
           m_va500RunningSpreads(index) = 500#
    Next
    
    
    ' Set Booleans
    Me.HasCreditSpreads = False
    Me.HasSpreadContractuals = False
    Me.HasCreditUpfronts = False
    Me.HasUpfrontContractuals = False

    ' Set Quote Type
    Me.ImpliedQuoteType = QuoteType_Unknown
    Me.ImpliedQuoteTypeName = ImpliedQuoteTypeNames(QuoteType_Unknown)
    
    ' VA 2010-08-12
    Me.errorMessage = ""
    Me.status = REQUEST_STATUS_OK

End Sub

Public Sub GetMarketDataFromResponse(ByVal CmsResponseXml As String, IsSetResponse As Boolean)
    
    '--------------------------------------------------------
    ' XML Variables
    '--------------------------------------------------------
    ' MSXML Dom object for processing the response
    Dim docMarketDataResponse   As MSXML2.DOMDocument
    Dim ndLabel                 As MSXML2.IXMLDOMNode
    Dim ndStatus                As MSXML2.IXMLDOMNode
    Dim ndIssuerRecoveryRate    As MSXML2.IXMLDOMNode
    Dim ndLastMarkedOn          As MSXML2.IXMLDOMNode
    Dim ndLastMarkedBy          As MSXML2.IXMLDOMNode
    Dim ndOwner                 As MSXML2.IXMLDOMNode
    Dim ndIsParent              As MSXML2.IXMLDOMNode
    Dim ndErrorMsg              As MSXML2.IXMLDOMNode
    
    Dim loadSuccess             As Boolean
    Dim Term                    As Variant
    Dim currentFourTuppleArray() As String
        
    On Error GoTo GetMarketDataFromResponse_Err
    
    ' VA 2010-08-12: Clean up from potential previous errors
    Me.errorMessage = ""
    Me.status = REQUEST_STATUS_OK
    
    ' Load the XML into DOMDocument
    Set docMarketDataResponse = New MSXML2.DOMDocument
    docMarketDataResponse.async = False
    docMarketDataResponse.validateOnParse = False
    loadSuccess = docMarketDataResponse.LoadXML(CmsResponseXml)
    If (Not loadSuccess) Then
       RaiseError ("Failed to load CMS response in DOMDocument")
    End If
    
    '-------------------------------------
    ' Process Status
    '-------------------------------------
    Set ndStatus = docMarketDataResponse.selectSingleNode("//systemMessages/statusMessage/statusCode")
    If ndStatus Is Nothing Then
       RaiseError ("Failed to retrieve Status from CMS Response")
    End If
    
    If StrComp(ndStatus.Text, "OK") <> 0 Then
        Me.status = "FAILED"
         Set ndErrorMsg = docMarketDataResponse.selectSingleNode("//systemErrors/errorMessage/errorMessageContent")
         If ndErrorMsg Is Nothing Then
             Me.errorMessage = "Unknown Error Occured while retrieving spreads from CMS"
         Else
             Me.errorMessage = Replace(ndErrorMsg.Text, STANDARD_JAVA_ERROR, "")
         End If
    
    Else
     Me.status = "OK"
     Me.errorMessage = ""
    End If
   
    ' -------------------------------------------------
    ' VA 2009-03-06: Kinjal guarantees that
    ' Label = 4Tuple! :-)
    ' -------------------------------------------------
     ' VA 2009-03-09 !!! For SNAC SET
     Set ndLabel = docMarketDataResponse.selectSingleNode("//marketSet/label")
     If (ndLabel Is Nothing) Then
        Set ndLabel = docMarketDataResponse.selectSingleNode("//marketSet/point/label")
     End If
     
     If (Not ndLabel Is Nothing) Then
        Me.fourTuple = ndLabel.Text
        currentFourTuppleArray = Split(Me.fourTuple, ".")
        Me.Ticker = currentFourTuppleArray(0)
        If (UBound(currentFourTuppleArray) = 3) Then
            Me.Ccy = currentFourTuppleArray(1)
            Me.debtClass = currentFourTuppleArray(2)
            Me.Product = currentFourTuppleArray(3)
        End If
     End If
     m_strFourTupleArray = Array(Ticker, Ccy, debtClass, Product)
     
'''    ' -------------------------------------------------------------
'''    ' VA 2009-06-18: Now we handle CONTRACT STANDARDS
'''    ' ASSUMPTION: Product starts with it (e.g., SNAC100, STEC750))
'''    ' -------------------------------------------------------------

    Me.IsStandardContractCurve = IsStandardContractProduct(Me.Product)
    
    ' -------------------------------------------------
    ' Extended set of attributes - "scalar" data
    ' -------------------------------------------------
     
    ' Issuer Recovery Rate
    Set ndIssuerRecoveryRate = docMarketDataResponse.selectSingleNode("//point[issuerRecovery]")
    If Not ndIssuerRecoveryRate Is Nothing Then
        Me.IssuerRecoveryRate = (ndIssuerRecoveryRate.LastChild.Text)
    End If
     
     ' LastMarkedOn
     ' VA 2008-12-15: New format from Kinjal: "Mon Dec 15 12:22:10 EST 2008"
     Set ndLastMarkedOn = docMarketDataResponse.selectSingleNode("//genericKeys/nameValuePair[@name='" & NAME_VALUE_LAST_MARKED_ON_TAG & "']")
     If (Not ndLastMarkedOn Is Nothing) Then
        Me.LastMarkedOn = ConvertToDate(ndLastMarkedOn.Text)
     End If
     ' LastMarkedBy
     Set ndLastMarkedBy = docMarketDataResponse.selectSingleNode("//genericKeys/nameValuePair[@name='" & NAME_VALUE_LAST_MARKED_BY_TAG & "']")
     If (Not ndLastMarkedBy Is Nothing) Then
        Me.LastMarkedBy = ndLastMarkedBy.Text
     End If
     ' Owner
     Set ndOwner = docMarketDataResponse.selectSingleNode("//genericKeys/nameValuePair[@name='" & NAME_VALUE_OWNER_TAG & "']")
     If (Not ndOwner Is Nothing) Then
        Me.Owner = ndOwner.Text
     End If
     ' Is Parent
     Set ndIsParent = docMarketDataResponse.selectSingleNode("//genericKeys/nameValuePair[@name='" & NAME_VALUE_IS_PARENT_TAG & "']")
     If (Not ndIsParent Is Nothing) Then
        Me.IsParent = ndIsParent.Text
     End If
     
    
    '-----------------------------------------------------------------
    ' Get Current Quote data - and update HAS flags by the way
    '-----------------------------------------------------------------
    Dim currentTermQuoteData As TermQuoteData
    Dim TermInt As Integer
    TermInt = 0
    For TermInt = 0 To UBound(m_vaTermStructure)
       Term = m_vaTermStructure(TermInt)
       currentTermQuoteData = GetTermQuoteData(Term, docMarketDataResponse, IsSetResponse)
       
       ' Check for all the Has flags
       If (Not currentTermQuoteData.TermCreditSpread = Empty) Then
            Let Me.CreditSpreads(TermInt) = currentTermQuoteData.TermCreditSpread
            Me.HasCreditSpreads = True
       End If
       If (Not currentTermQuoteData.TermCreditUpfront = Empty) Then
            Me.CreditUpfronts(TermInt) = currentTermQuoteData.TermCreditUpfront
            Me.HasCreditUpfronts = True
       End If
       If (Not currentTermQuoteData.TermSpreadContractual = Empty) Then
            Me.SpreadContractuals(TermInt) = currentTermQuoteData.TermSpreadContractual
            Me.HasSpreadContractuals = True
       End If
       If (Not currentTermQuoteData.TermUpfrontContractual = Empty) Then
            m_vaUpfrontContractuals(TermInt) = currentTermQuoteData.TermUpfrontContractual
            Me.HasUpfrontContractuals = True
       End If
    Next TermInt
    
    ' Now, based on the HAS flags, we could determine the Quote type
    Dim strHasFlags As String
    
    ' SET Response contains numbers in the same scale as Request -
    ' Don't need to multiply
    Dim doCalcFlag As Boolean
    doCalcFlag = True
    
    Dim bCopySuccess As Boolean
    strHasFlags = UCase(Me.HasCreditSpreads & "|" & Me.HasSpreadContractuals & "|" & Me.HasCreditUpfronts & "|" & Me.HasUpfrontContractuals)
    Select Case (strHasFlags)
                   
        ' "Pure" PAR Spreads
        Case "TRUE|FALSE|FALSE|FALSE":
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataSpreads, m_vaCreditSpreads)
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataQuotes, m_vaCreditSpreads)
            ' No Contractuals!
        
        ' SNAC Spreads + Contractuals
        Case "TRUE|TRUE|FALSE|FALSE":
            
            ' VA 2009-06-18
'''            If (Me.Product = SNAC_100_PRODUCT) Then
'''                Me.ImpliedQuoteType = QuoteType_100Spreads
'''            ' Just 2 cases for now...
'''            Else
'''                Me.ImpliedQuoteType = QuoteType_500Spreads
'''            End If

            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataSpreads, m_vaCreditSpreads)
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataQuotes, m_vaCreditSpreads)
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataContractuals, m_vaSpreadContractuals)
        
        ' "Pure" PAR Upfronts
        Case "FALSE|FALSE|TRUE|FALSE":
            
            ' There is no contracutals; yet we put m_va500RunningSpreads in m_vaMarketDataContractuals
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataUpfronts, m_vaCreditUpfronts, False)
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataQuotes, m_vaCreditUpfronts, False)
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataContractuals, m_va500RunningSpreads)
        
        ' SNAC Upfronts + Contractuals
        Case "FALSE|FALSE|TRUE|TRUE":
'''            If (Me.Product = SNAC_100_PRODUCT) Then
'''                Me.ImpliedQuoteType = QuoteType_100Upfronts
'''            ' Just 2 cases for now...
'''            Else
'''                Me.ImpliedQuoteType = QuoteType_500Upfronts
'''            End If

            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataUpfronts, m_vaCreditUpfronts, False)
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataQuotes, m_vaCreditUpfronts)
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataContractuals, m_vaUpfrontContractuals)
            
        ' SNAC Spreads + Upfronts + 2 Contractuals - MUST be the SET Response!
        Case "TRUE|TRUE|TRUE|TRUE":
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataSpreads, m_vaCreditSpreads)
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataUpfronts, m_vaCreditUpfronts, False)
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataQuotes, m_vaCreditUpfronts)
            bCopySuccess = arrayHelper.CopyArray(m_vaMarketDataContractuals, m_vaUpfrontContractuals)
                    
    End Select
    
    ' ImpliedQuoteTypeName
    ImpliedQuoteTypeName = GetQuoteType(docMarketDataResponse)
    
    ' Initialize the "Scalar" array
    m_vaScalarData = Array(Me.IssuerRecoveryRate, Me.LastMarkedOn, Me.LastMarkedBy, Me.Owner, Me.IsParent, Me.ImpliedQuoteTypeName, Me.status, Me.errorMessage)

    ' Apppend the "Scalar" data to MarketDataQuotes & m_vaMarketDataContractuals
    Dim startCopyIndex As Integer
    startCopyIndex = TenorStandard
    
    bCopySuccess = arrayHelper.CopyArraySubSetToArray(m_vaScalarData, m_vaMarketDataQuotes, 0, UBound(m_vaScalarData), Me.TenorStandard)
    bCopySuccess = arrayHelper.CopyArraySubSetToArray(m_vaScalarData, m_vaMarketDataSpreads, 0, UBound(m_vaScalarData), Me.TenorStandard)
    bCopySuccess = arrayHelper.CopyArraySubSetToArray(m_vaScalarData, m_vaMarketDataUpfronts, 0, UBound(m_vaScalarData), Me.TenorStandard)
            
    m_vaScalarData(CMS_SCLR_CURVE_TYPE) = ImpliedQuoteTypeName
    
    bCopySuccess = arrayHelper.CopyArraySubSetToArray(m_vaScalarData, m_vaMarketDataContractuals, 0, UBound(m_vaScalarData), Me.TenorStandard)
   
   Exit Sub
GetMarketDataFromResponse_Err:
     Me.status = REQUEST_STATUS_FAILED
     Me.errorMessage = "Failed to retrieve Market Data from CMS XML Response"
     If (Err.Description <> "") Then
        Me.errorMessage = Me.errorMessage & ": " & Err.Description
     Else
        
     End If
     m_vaMarketDataQuotes(Me.TenorStandard + CMS_SCLR_STATUS) = Me.status
     m_vaMarketDataContractuals(Me.TenorStandard + CMS_SCLR_STATUS) = Me.status
     m_vaMarketDataQuotes(Me.TenorStandard + CMS_SCLR_ERROR) = Me.errorMessage
     m_vaMarketDataContractuals(Me.TenorStandard + CMS_SCLR_ERROR) = Me.errorMessage
End Sub

' Get ALL the Term data available in CMS Response
 Private Function GetTermQuoteData(ByVal Term As String, ByVal Response As MSXML2.DOMDocument, _
                                   ByVal IsSetResponse As Boolean) _
                                   As TermQuoteData
        
        Dim ndPoint As MSXML2.IXMLDOMNode
        Dim ndContractualSpread As MSXML2.IXMLDOMNode
        Dim termSplit() As String
        Dim pointContractualSpread As MSXML2.DOMDocument
        Dim currentTermQuoteData As TermQuoteData
        Dim xsltTag As String
        
        termSplit = Strings.Split(Term, ":")
        currentTermQuoteData.Term = Term

        
        ' Get creditSpread & spreadContractual
        xsltTag = "//point[" & CREDIT_SPREAD_DATA_TYPE & "[periodMultiplier=" & termSplit(0) & "][period='" & termSplit(1) & "']]"
        
        ' Get creditSpread & spreadContractual
        Set ndPoint = Response.selectSingleNode("//point[" & CREDIT_SPREAD_DATA_TYPE & "[periodMultiplier=" & termSplit(0) & "][period='" & termSplit(1) & "']]")
        If (Not ndPoint Is Nothing) Then
            ' Spreads in Basis Points - but not in SET Response!
            currentTermQuoteData.TermCreditSpread = ndPoint.LastChild.Text
            If (Not IsSetResponse) Then
                currentTermQuoteData.TermCreditSpread = currentTermQuoteData.TermCreditSpread * 10000#
            End If

            ' Get spreadContractual
            Set ndContractualSpread = ndPoint.selectSingleNode("creditSpread/contractualSpread")
            If (Not ndContractualSpread Is Nothing) Then
                currentTermQuoteData.TermSpreadContractual = ndPoint.LastChild.Text
                If (Not IsSetResponse) Then
                    currentTermQuoteData.TermSpreadContractual = ndContractualSpread.LastChild.Text * 10000#
                End If
                
            End If
        End If
        
        ' Get creditUpfront & upfrontContractual
        Set ndPoint = Response.selectSingleNode("//point[" & CREDIT_UPFRONT_DATA_TYPE & "[periodMultiplier=" & termSplit(0) & "][period='" & termSplit(1) & "']]")
        If (Not ndPoint Is Nothing) Then
            ' Upfronts in percents - but not in SET Response!
            currentTermQuoteData.TermCreditUpfront = ndPoint.LastChild.Text
            If (Not IsSetResponse) Then
                currentTermQuoteData.TermCreditUpfront = currentTermQuoteData.TermCreditUpfront * 100#
            End If
            
            ' Get upfrontContractual
            Set ndContractualSpread = ndPoint.selectSingleNode("creditUpfront/contractualSpread")
            If (Not ndContractualSpread Is Nothing) Then
                currentTermQuoteData.TermUpfrontContractual = ndContractualSpread.LastChild.Text
                If (Not IsSetResponse) Then
                    currentTermQuoteData.TermUpfrontContractual = currentTermQuoteData.TermUpfrontContractual * 10000#
                End If
            End If
        End If

        GetTermQuoteData = currentTermQuoteData
        
End Function

Private Function GetIssuerRecovery(ByVal Response As MSXML2.DOMDocument) As Double
                            
        Dim point As MSXML2.IXMLDOMNode
        Set point = Response.selectSingleNode("//point[issuerRecovery]")
                                
        If Not point Is Nothing Then
            GetIssuerRecovery = (point.LastChild.Text)
        End If
        
End Function



'-------------------------------------------------------------------
' UTILITIES
'-------------------------------------------------------------------

Private Sub RaiseError(errDesc As String)
    Call Err.Raise(-1, APP_ID, errDesc)
End Sub

' VA 2008-12-15: Due to the error on CMS server we have to process 2 foramts:
' "2008-12-12 17:04:41.137" & "Mon Dec 15 12:22:10 EST 2008"
Private Function ConvertToDate(DateString As String) As Date
    Dim tmpDate As String
    Dim tmpTime As String
    
    If (InStr(1, DateString, "-") > 0) Then
        ' "2008-12-12 17:04:41.137" Format
        tmpDate = Replace(Left(DateString, 10), "-", "/")
        tmpTime = Mid(DateString, 12, 8)
    Else
        ' "Mon Dec 15 12:22:10 EST 2008" Format
        tmpDate = Mid(DateString, 9, 2) & "-" & Mid(DateString, 5, 3) & "-" & Mid(DateString, 25, 4)
        tmpTime = Mid(DateString, 12, 8)
    End If
    
    ' Convert to "reasonable" Excel Date-Time format
    ConvertToDate = CDate(tmpDate & " " & tmpTime)

End Function

' 2009-01-27: CMS fails to parse XML if it's coming from WEB query
' http://www.devx.com/vb2themax/Tip/19162
' Encode an string so that it can be displayed correctly
' inside the browser.
'
' Same effect as the Server.HTMLEncode method in ASP
Private Function HTMLEncode(ByVal Text As String) As String
    Dim i As Integer
    Dim acode As Integer
    Dim repl As String

    HTMLEncode = Text

    For i = Len(HTMLEncode) To 1 Step -1
        acode = Asc(Mid$(HTMLEncode, i, 1))
        Select Case acode
            Case 32
                repl = "&nbsp;"
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

Private Function GetTenorColumnShift(Optional ByVal TenorStandard = TENORS_11_STANDARD) As Integer
    Select Case (TenorStandard)
        Case TENORS_13_STANDARD:
            GetTenorColumnShift = 2
        Case TENORS_12_STANDARD:
            GetTenorColumnShift = 1
        Case TENORS_11_STANDARD:
default:
            GetTenorColumnShift = 0
    End Select
End Function







