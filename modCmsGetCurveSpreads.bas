
Attribute VB_Name = "modCmsGetCurveSpreads"
'----------------------------------------------------------------------------------
' Revision 2009-03-26.13-00 (SNAC Marking; 11 Tenors (placeholder for 13)
'----------------------------------------------------------------------------------
Option Explicit

'----------------------------------------------------------------------------------
' VA 2008-11-21: XML functionality originally provided by Chaitanya Penubarthi
'----------------------------------------------------------------------------------
Private TERM_STRUCTURE() As Variant
Private FULL_TERM_STRUCTURE_NAMES() As Variant
Private CURVE_QUOTE_DATA_HEADER() As Variant
Private CURVE_QUOTES_HEADER() As Variant
Private TERM_QUOTE_DATA_HEADER() As Variant
Private RECOVERY_RATE_HEADER() As Variant
Private IMPERSONATED_USER As String


'----------------------------------------------------------------------------------
' VA 2009-01-16: Header functionality originally induced by Ozgur Kaya
'----------------------------------------------------------------------------------

'' VA 2009-06-18: We now test STANDARD_CONTRACT/NON_STANDARD_CONTRACT curves
'' SNAC & STEC at the moment; might be extended in the future
'Public Function CMS_GetSupportedContractStandards() As Variant
'    CMS_GetSupportedContractStandards = Array("SNAC", "STEC")
'End Function
'
'
'Public Function CMS_GetSupportedTenorStandards() As Variant
'    CMS_GetSupportedTenorStandards = Array(11, 13, 17, "<Ad Hoc 1D Tenor Array>")
'End Function


Public Function CMS_GetDataSizingInfo(Optional TenorStandard = TENORS_11_STANDARD) As String
         
         Dim tenorCount As Integer
         Dim curveDetailsCount As Integer
         Dim colCount As Integer
         Dim adHoc As String
         
         ' Both are 0-based arrays
         tenorCount = UBound(CMS_GetCurveQuotesHeader(TenorStandard)) + 1
         curveDetailsCount = UBound(CMS_GetCurveDetailsHeader()) + 1
         colCount = tenorCount + curveDetailsCount
         If (tenorCount <> 11 And tenorCount <> 12 And tenorCount <> 13 And tenorCount <> 17) Then
            adHoc = " AD HOC"
         End If
         
         ' VA 2009-05-13: Now we support Ad Hoc Tenor arrays as well
                  CMS_GetDataSizingInfo = _
         "'----------------------------------------------------" & vbNewLine & _
         "' CmsDataAccess Add-in Sizing Info (" & tenorCount & adHoc & " Tenors)" & vbNewLine & _
         "'----------------------------------------------------" & vbNewLine & _
         "Supported TENOR STDs - 11, 12, 13, 17, and AD HOC ARRAY" & vbNewLine & _
         "FourTupleArray       - VariantArray(1x4)" & vbNewLine & _
         "TermStructure        - VariantArray(1x" & tenorCount & ")" & vbNewLine & _
         "CurveQuoteData       - VariantArray(1x" & colCount & ")" & vbNewLine & _
         "                       (All modes but 'EXTENDED')" & vbNewLine & _
         "                       VariantArray(2x" & colCount & ")" & vbNewLine & _
         "                       ('EXTENDED' mode)" & vbNewLine & _
         "CurveQuotes          - VariantArray(1x" & tenorCount & ")" & vbNewLine & _
         "TermQuoteData        - VariantArray(1x9)" & vbNewLine & _
         "TermQuote            - Variant" & vbNewLine & _
         "RecoveryRate         - Variant"
         
         Exit Function
         
         Select Case (TenorStandard)
         
         ' VA 2009-07-08
         Case TENORS_17_STANDARD:
         CMS_GetDataSizingInfo = _
         "'------------------------------------------------" & vbNewLine & _
         "' CmsDataAccess Add-in Sizing Info (17 Tenors)" & vbNewLine & _
         "'------------------------------------------------" & vbNewLine & _
         "Supported TENOR STDs - 11, 12, 13, 17, and AD HOC ARRAY" & vbNewLine & _
         "FourTupleArray       - VariantArray(1x4)" & vbNewLine & _
         "TermStructure        - VariantArray(1x17)" & vbNewLine & _
         "CurveQuoteData       - VariantArray(1x25)" & vbNewLine & _
         "                       (All modes but 'EXTENDED')" & vbNewLine & _
         "                       VariantArray(2x25)" & vbNewLine & _
         "                       ('EXTENDED' mode)" & vbNewLine & _
         "CurveQuotes          - VariantArray(1x17)" & vbNewLine & _
         "TermQuoteData        - VariantArray(1x9)" & vbNewLine & _
         "TermQuote            - Variant" & vbNewLine & _
         "RecoveryRate         - Variant"
    
         Case TENORS_13_STANDARD:
         CMS_GetDataSizingInfo = _
         "'------------------------------------------------" & vbNewLine & _
         "' CmsDataAccess Add-in Sizing Info (13 Tenors)" & vbNewLine & _
         "'------------------------------------------------" & vbNewLine & _
         "Supported TENOR STDs - 11, 12, 13, 17, and AD HOC ARRAY" & vbNewLine & _
         "FourTupleArray       - VariantArray(1x4)" & vbNewLine & _
         "TermStructure        - VariantArray(1x13)" & vbNewLine & _
         "CurveQuoteData       - VariantArray(1x21)" & vbNewLine & _
         "                       (All modes but 'EXTENDED')" & vbNewLine & _
         "                       VariantArray(2x21)" & vbNewLine & _
         "                       ('EXTENDED' mode)" & vbNewLine & _
         "CurveQuotes          - VariantArray(1x13)" & vbNewLine & _
         "TermQuoteData        - VariantArray(1x9)" & vbNewLine & _
         "TermQuote            - Variant" & vbNewLine & _
         "RecoveryRate         - Variant"
         
         Case TENORS_12_STANDARD:
         CMS_GetDataSizingInfo = _
         "'------------------------------------------------" & vbNewLine & _
         "' CmsDataAccess Add-in Sizing Info (13 Tenors)" & vbNewLine & _
         "'------------------------------------------------" & vbNewLine & _
         "Supported TENOR STDs - 11, 12, 13, 17, and AD HOC ARRAY" & vbNewLine & _
         "FourTupleArray       - VariantArray(1x4)" & vbNewLine & _
         "TermStructure        - VariantArray(1x12)" & vbNewLine & _
         "CurveQuoteData       - VariantArray(1x20)" & vbNewLine & _
         "                       (All modes but 'EXTENDED')" & vbNewLine & _
         "                       VariantArray(2x20)" & vbNewLine & _
         "                       ('EXTENDED' mode)" & vbNewLine & _
         "CurveQuotes          - VariantArray(1x12)" & vbNewLine & _
         "TermQuoteData        - VariantArray(1x9)" & vbNewLine & _
         "TermQuote            - Variant" & vbNewLine & _
         "RecoveryRate         - Variant"
         
         
         
         Case TENORS_11_STANDARD:
default:
         CMS_GetDataSizingInfo = _
         "'------------------------------------------------" & vbNewLine & _
         "' CmsDataAccess Add-in Sizing Info (11 Tenors)" & vbNewLine & _
         "'------------------------------------------------" & vbNewLine & _
         "Supported TENOR STDs - 11, 12, 13, 17, and AD HOC ARRAY" & vbNewLine & _
         "FourTupleArray       - VariantArray(1x4)" & vbNewLine & _
         "TermStructure        - VariantArray(1x11)" & vbNewLine & _
         "CurveQuoteData       - VariantArray(1x19)" & vbNewLine & _
         "                       (All modes but 'EXTENDED')" & vbNewLine & _
         "                       VariantArray(2x19)" & vbNewLine & _
         "                       ('EXTENDED' mode)" & vbNewLine & _
         "CurveQuotes          - VariantArray(1x11)" & vbNewLine & _
         "TermQuoteData        - VariantArray(1x9)" & vbNewLine & _
         "TermQuote            - Variant" & vbNewLine & _
         "RecoveryRate         - Variant"
         End Select
         
         
End Function

Public Function CMS_GetCurveDetailsHeader() As Variant
    CMS_GetCurveDetailsHeader = Array("Recovery", "Last Marked On", "Last Marked By", "Owner", "Is Parent", "Curve Type", "Status", "Error Message")
End Function


Public Function CMS_GetCurveQuoteDataHeader(Optional TenorStandard As Variant = TENORS_11_STANDARD) As Variant
    Dim arrayHelper As New CArrayHelper
    Dim curveQuotesHeader() As Variant
    Dim curveDetailsHeader() As Variant
    
    curveQuotesHeader = CMS_GetCurveQuotesHeader(TenorStandard)
    curveDetailsHeader = CMS_GetCurveDetailsHeader()
    CURVE_QUOTE_DATA_HEADER = arrayHelper.Merge1DArrays(curveQuotesHeader, curveDetailsHeader, 0)
    
    CMS_GetCurveQuoteDataHeader = CURVE_QUOTE_DATA_HEADER
End Function


Public Function CMS_GetCurveQuotesHeader(Optional TenorStandard As Variant = TENORS_11_STANDARD) As Variant()
    
    Dim itemIndex As Integer
    Dim wrkArray() As Variant
    Dim arrayHelper As New CArrayHelper
    Dim rngAdHocTenors As Range

    ' VA 2009-05-12: Ad Hoc Tenor Range is passed
    If (IsObject(TenorStandard) And Not IsNumeric(TenorStandard)) Then
        Set rngAdHocTenors = TenorStandard
        wrkArray = arrayHelper.GetArrayFromRange(rngAdHocTenors, 0, True)
        Call arrayHelper.GetRow(wrkArray, CURVE_QUOTES_HEADER, 0)
        
        ' Convert to upper case
        For itemIndex = 0 To UBound(CURVE_QUOTES_HEADER)
            CURVE_QUOTES_HEADER(itemIndex) = UCase(CURVE_QUOTES_HEADER(itemIndex))
        Next
        
    ' VA 2009-06-22: Adding 12 Tenors
    ' "Regular" Tenors (11, 13, & 17 as of 2009-05-12)
    Else
        Select Case (TenorStandard)
            Case TENORS_17_STANDARD:
                CURVE_QUOTES_HEADER = Array("0M", "3M", "6M", "9M", "1Y", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "8Y", "9Y", "10Y", "15Y", "20Y", "30Y")
            Case TENORS_13_STANDARD:
                CURVE_QUOTES_HEADER = Array("0M", "3M", "6M", "9M", "1Y", "2Y", "3Y", "4Y", "5Y", "7Y", "10Y", "20Y", "30Y")
            Case TENORS_12_STANDARD:
                CURVE_QUOTES_HEADER = Array("0M", "3M", "6M", "1Y", "2Y", "3Y", "4Y", "5Y", "7Y", "10Y", "20Y", "30Y")
            Case TENORS_11_STANDARD:
                CURVE_QUOTES_HEADER = Array("3M", "6M", "1Y", "2Y", "3Y", "4Y", "5Y", "7Y", "10Y", "20Y", "30Y")
            Case Else:
                CURVE_QUOTES_HEADER = Array("ERROR: Only 11, 12, 13 & 17 are supported as TENOR_STANDARD value")
        End Select
    End If
    CMS_GetCurveQuotesHeader = CURVE_QUOTES_HEADER
End Function

Public Function CMS_GetTermQuoteDataHeader() As Variant()
    TERM_QUOTE_DATA_HEADER() = Array("Term Quote", "Recovery", "Last Marked On", "Last Marked By", "Owner", "Is Parent", "Curve Type", "Status", "Error Message")
    CMS_GetTermQuoteDataHeader = TERM_QUOTE_DATA_HEADER()
End Function

Public Function CMS_GetRecoveryRateDataHeader() As Variant()
    RECOVERY_RATE_HEADER() = Array("Recovery", "Last Marked On", "Last Marked By", "Owner", "Is Parent", "Curve Type", "Status", "Error Message")
    CMS_GetRecoveryRateDataHeader = RECOVERY_RATE_HEADER()
End Function


' VA 2009-05-08: Need delimited Term Structure for Tenor quote processing
Private Function GetDelimitedTermStructure(Optional TenorStandard As Variant = TENORS_11_STANDARD) As Variant()
    Dim delimitedTermStructure() As Variant
    
    ' VA 2009-05-12: Ad Hoc Tenor Range is passed
    If (IsObject(TenorStandard) And Not IsNumeric(TenorStandard)) Then
        
        Dim rngAdHocTenors As Range
        Set rngAdHocTenors = TenorStandard
        GetDelimitedTermStructure = ConvertTermsToTermStructure(GetAdHocTenorsFromRange(rngAdHocTenors))
        Exit Function
    End If
    
    ' "Regular" Tenors (11, 13, 17)
    Select Case (TenorStandard)
        Case TENORS_17_STANDARD:
            delimitedTermStructure = Array("0:M", "3:M", "6:M", "9:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "6:Y", "7:Y", "8:Y", "9:Y", "10:Y", "15:Y", "20:Y", "30:Y")
        Case TENORS_13_STANDARD:
            delimitedTermStructure = Array("0:M", "3:M", "6:M", "9:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
        Case TENORS_12_STANDARD:
            delimitedTermStructure = Array("0:M", "3:M", "6:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
        Case TENORS_11_STANDARD:
default:
        delimitedTermStructure = Array("3:M", "6:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
    End Select
        GetDelimitedTermStructure = delimitedTermStructure
End Function



' VA 2009-05-18: Nope, we'd rather use Ad Hoc tenors...
' VA 2008-12-08: Implement it dynamic and curve-specific, taking 4-tuple as parameters!
Public Function CMS_GetTermStructure(Optional TenorStandard As Variant = TENORS_11_STANDARD) As Variant()
    FULL_TERM_STRUCTURE_NAMES = CMS_GetCurveQuotesHeader(TenorStandard)
    CMS_GetTermStructure = FULL_TERM_STRUCTURE_NAMES
End Function

Public Function CMS_GetFourTupleArray(ByVal fourTuple As String, Optional ByVal Delimiter = DEFAULT_DELIMITER) As String()
Attribute CMS_GetFourTupleArray.VB_Description = "Retrieves Ticker, Currency, DebtClass, and Product from 4Touple (VariantArray(1x4))."
    Dim FourTupleItems() As String
    ReDim Preserve FourTupleItems(0 To 3)
    If (Len(fourTuple) > 0) Then
        fourTuple = UCase(fourTuple)
        FourTupleItems = Split(fourTuple, Delimiter)
    End If
    CMS_GetFourTupleArray = FourTupleItems
End Function

Public Function CMS_CompileFourTuple(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, ByVal Product As String, _
                                   Optional ByVal Delimiter = DEFAULT_DELIMITER) As String
Attribute CMS_CompileFourTuple.VB_Description = "Compiles 4Touple from Ticker, Currency, DebtClass, and Product (String). Default Delimiter = "".""."
    CMS_CompileFourTuple = UCase(Ticker) & UCase(Delimiter) & UCase(Ccy) & Delimiter & UCase(debtClass) & Delimiter & UCase(Product)
End Function

Public Function CMS_GetCurveQuoteData(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = DEFAULT_PRODUCT, _
                                     Optional ByVal Tag = "LIVE", Optional ByVal QuoteDate As Date, _
                                     Optional ByVal QuoteConvention = DEFAULT_CURVE_CONVENTION, _
                                     Optional ByVal TenorStandard As Variant = TENORS_11_STANDARD) As Variant
                                     
CMS_GetCurveQuoteData = CMS_GetCurveQuoteDataImpl(Ticker, Ccy, debtClass, Product, Tag, _
                                                  QuoteDate, QuoteConvention, TenorStandard)

End Function


' VA 2010-08-11: Extended with optional Response XML parameter to be used in Bulk request processing
' and made it private not to be visible in Excel
Private Function CMS_GetCurveQuoteDataImpl(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = DEFAULT_PRODUCT, _
                                     Optional ByVal Tag = "LIVE", Optional ByVal QuoteDate As Date, _
                                     Optional ByVal QuoteConvention = DEFAULT_CURVE_CONVENTION, _
                                     Optional ByVal TenorStandard As Variant = TENORS_11_STANDARD, _
                                     Optional CmsCurveDataResponse As String = "") As Variant

    Dim arrayHelper As New CArrayHelper
    Dim IsExtendedRequest As Boolean
    Dim quoteResults() As Variant
    Dim tenorColumnShift As Integer

    Dim spreadMarks() As Variant
    Dim runningMarks() As Variant
    Dim rosettaUserMetaData As String
    Dim rosettaXmlRequest As String
    Dim cmsResponse As String
    Dim index As Integer
        
    ' VA 2010-08-12
    
    
    On Error GoTo CMS_GetCurveQuoteData_Err
    
    ' Get the column shift for the current Tenor Standard
    tenorColumnShift = GetTenorColumnShift(TenorStandard)
    
    ' Redim - cleanup
    ReDim quoteResults(0 To 1, 0 To CMS_COL_COUNT + tenorColumnShift - 1)
    ReDim spreadMarks(0 To CMS_COL_COUNT + tenorColumnShift - 1)
    
    ' Check the QuoteConvention
    If (QuoteConvention = CURVE_CONVENTION_EXTENDED) Then
        QuoteConvention = CURVE_CONVENTION_UPFRONT
        IsExtendedRequest = True
    Else
    End If
    
    
    ' VA 2010-08-10: Skip getting CMSResponse if it's already there...
    If (CmsCurveDataResponse = "") Then
        rosettaUserMetaData = CompileGetUserMetaData(CMS_GetUserLoginName())
        
        ' ' Debug
        If (DEBUG_PRINT_CMS_USER_METADATA) Then
             Debug.Print "User Metadata: "
             Debug.Print rosettaUserMetaData
            ' MsgBox rosettaUserMetaData
        End If
        
        rosettaXmlRequest = CompileRosettaGetXmlRequest(Ticker, Ccy, debtClass, Product, Tag, QuoteDate, QuoteConvention)
        
        ' Debug
        If (DEBUG_PRINT_CMS_ROSETTA_XML) Then
            Debug.Print "Rosetta XML Request: "
            Debug.Print rosettaXmlRequest
            'MsgBox rosettaXmlRequest
        End If
        
        ' Get the CMS Response
        cmsResponse = GetCmsXmlResponse(rosettaUserMetaData, rosettaXmlRequest)
        ' Debug
        If (DEBUG_PRINT_CMS_ROSETTA_XML) Then
        Debug.Print "Rosetta XML Response: "
        Debug.Print cmsResponse
        'MsgBox cmsResponse
    End If
    Else
        cmsResponse = CmsCurveDataResponse
    End If
    
    ' =============================================================
    ' VA 2008-03-09: Using CMarketData Class now
    ' Sorry, no time to clean it up...
    ' =============================================================
    Dim currentMarketData As CMarketData
    Set currentMarketData = New CMarketData
    currentMarketData.Initialize (TenorStandard)
    
    Call currentMarketData.GetMarketDataFromResponse(cmsResponse, True)
    
'''    Select Case (currentMarketData.Product)
'''        Case "SNAC100", "SNAC500":
            
    ' VA 2009-06-18
    If (currentMarketData.IsStandardContractCurve) Then
            
            Call currentMarketData.GetMarketDataFromResponse(cmsResponse, False)
            spreadMarks = currentMarketData.MarketDataQuotes
            runningMarks = currentMarketData.MarketDataContractuals
            
            If (IsExtendedRequest) Then
                runningMarks = currentMarketData.MarketDataContractuals
                For index = 0 To CMS_COL_COUNT + tenorColumnShift - 1
                    quoteResults(0, index) = spreadMarks(index)
                    'If (Index < (CMS_COL_RECOVERY + tenorColumnShift) Or Index > (CMS_COL_IS_PARENT + tenorColumnShift)) Then
                    quoteResults(1, index) = runningMarks(index)
                    ' End If
                Next
                CMS_GetCurveQuoteDataImpl = quoteResults
            ElseIf QuoteConvention = CURVE_CONVENTION_RUNNING Then
                CMS_GetCurveQuoteDataImpl = runningMarks
            Else
              CMS_GetCurveQuoteDataImpl = spreadMarks
            End If
        
'''        Case Else:
            
            ' VA 2009-06-18
            Else
          ' =================================================================================
          ' OLD STUFF
          ' =================================================================================
          spreadMarks = GetMarketData(cmsResponse, QuoteConvention, TenorStandard)
          
   
          
         If (IsExtendedRequest) Then
              For index = 0 To CMS_COL_COUNT + tenorColumnShift - 1
                  quoteResults(0, index) = spreadMarks(index)
                      ' VA 2009-03-10
                      ' VA 2009-06-18 !!!!
                      ' quoteResults(1, index) = runningMarks(index)
                      If (index < currentMarketData.TenorStandard And spreadMarks(index) <> 0) Then
                        quoteResults(1, index) = 500
                      End If
              Next
              CMS_GetCurveQuoteDataImpl = quoteResults
          Else
              CMS_GetCurveQuoteDataImpl = spreadMarks
          End If
          
   ' VA 2011-05-27: This is a hack to make CURVE_CONVENTION_RUNNING
    ' work for SPECREF product
    Dim Term As Variant
    Dim TermInt As Integer
    If currentMarketData.Product = SPEC_REF_CURVE_PRODUCT And QuoteConvention = CURVE_CONVENTION_RUNNING Then
        For Each Term In TERM_STRUCTURE
                If Not IsEmpty(currentMarketData.MarketDataSpreads(TermInt)) Then
                    spreadMarks(TermInt) = 500
                End If
            TermInt = TermInt + 1
        Next Term
        CMS_GetCurveQuoteDataImpl = spreadMarks
    End If
          
    
'''    End Select

    End If
 
    
    Exit Function
    
CMS_GetCurveQuoteData_Err:
    If (Err.Number <> 0) Then
        If (IsExtendedRequest) Then
            quoteResults(0, CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_FAILED
            quoteResults(0, CMS_COL_ERROR + tenorColumnShift) = Err.Description
            CMS_GetCurveQuoteDataImpl = quoteResults
        Else
            spreadMarks(CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_FAILED
            spreadMarks(CMS_COL_ERROR + tenorColumnShift) = Err.Description
            CMS_GetCurveQuoteDataImpl = spreadMarks
        End If
    End If
    
    ' VA 2009-06-23: !!! This excludes #VALUE!'s from the sheet
    Resume CMS_GetCurveQuoteData_Exit
    
CMS_GetCurveQuoteData_Exit:
    
    
End Function

Public Function CMS_GetCurveQuotes(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = "DERIV", _
                                     Optional ByVal Tag = "LIVE", Optional ByVal QuoteDate As Date, Optional ByVal QuoteConvention = "SPREADS", _
                                     Optional ByVal TenorStandard As Variant = TENORS_11_STANDARD) As Variant

    Dim tenorColumnShift As Integer
    Dim spreadMarks() As Variant
    
    tenorColumnShift = GetTenorColumnShift(TenorStandard)
    spreadMarks = CMS_GetCurveQuoteData(Ticker, Ccy, debtClass, Product, Tag, QuoteDate, QuoteConvention, TenorStandard)
    
    If (spreadMarks(CMS_COL_STATUS + tenorColumnShift) <> REQUEST_STATUS_OK) Then
        spreadMarks(CMS_TERM_COUNT + tenorColumnShift - 1) = spreadMarks(CMS_COL_ERROR + tenorColumnShift)
    End If
    ReDim Preserve spreadMarks(0 To CMS_TERM_COUNT + tenorColumnShift - 1)
    CMS_GetCurveQuotes = spreadMarks
End Function

Public Function CMS_GetTermQuoteData(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = "DERIV", _
                                     Optional ByVal Tag = "LIVE", Optional ByVal QuoteDate As Date, Optional ByVal QuoteConvention = "SPREADS", Optional Term = "5Y") As Variant
Attribute CMS_GetTermQuoteData.VB_Description = "Returns the  data releated to specific Term (defaulted to 5Y): TermQuote, Recovery, LastMarkedOn, LastMarkedBy, Owner, IsParent, Status, and ErrorMessage."

    Dim spreadMarks() As Variant
    Dim termMarks() As Variant
    Dim TermNames() As Variant
    Dim colDelta As Integer
    Dim index As Integer

    ReDim spreadMarks(0 To CMS_COL_COUNT - 1)
    ReDim TermNames(0 To CMS_TERM_COUNT - 1)
    ReDim termMarks(0 To CMS_TERM_DATA_COL_COUNT - 1)
    
    colDelta = CMS_COL_COUNT - CMS_TERM_DATA_COL_COUNT
    
    ' Term Names
    TermNames = CMS_GetTermStructure()

    ' Get all quotes
    spreadMarks = CMS_GetCurveQuoteData(Ticker, Ccy, debtClass, Product, Tag, QuoteDate, QuoteConvention)
    
    ' Get the Term Spread/Upfront
    termMarks(0) = 0
    For index = 0 To CMS_TERM_COUNT - 1
      If UCase(Term) = TermNames(index) Then
          termMarks(0) = spreadMarks(index)
      End If
    Next
    
    ' Now, all the "lookups"
    For index = CMS_TERM_COUNT To CMS_COL_COUNT - 1
        termMarks(index - colDelta) = spreadMarks(index)
    Next
    
    CMS_GetTermQuoteData = termMarks

End Function

Public Function CMS_GetTermQuote(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = "DERIV", _
                                     Optional ByVal Tag = "LIVE", Optional ByVal QuoteDate As Date, Optional ByVal QuoteConvention = "SPREADS", Optional Term = "5Y") As Variant
Attribute CMS_GetTermQuote.VB_Description = "Returns Quote for specified Term (defaulted to 5Y). In case of error contains ErrorMessage.\r\n"
    Dim termMarks() As Variant

    ReDim termMarks(0 To CMS_TERM_DATA_COL_COUNT - 1)

    ' Get Term Data
    termMarks = CMS_GetTermQuoteData(Ticker, Ccy, debtClass, Product, Tag, QuoteDate, QuoteConvention, Term)
    
    If (termMarks(CMS_TERM_COL_STATUS) <> REQUEST_STATUS_OK) Then
        CMS_GetTermQuote = termMarks(CMS_TERM_COL_ERROR)
    Else
        CMS_GetTermQuote = termMarks(CMS_TERM_COL_QUOTE)
    End If
End Function

'--------------------------------------------------------
' VA 2011-10-04: Extra "single item" methods
'--------------------------------------------------------

' VA 2011-10-04
Public Function CMS_GetCurveQuotedType(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = "DERIV", _
                                       Optional ByVal Tag = "LIVE", Optional ByVal QuoteDate As Date) As Variant

    Dim termMarks() As Variant

    ReDim termMarks(0 To TENORS_11_STANDARD - 1)

    ' Get Curve Quote Data
    termMarks = CMS_GetCurveQuoteData(Ticker, Ccy, debtClass, Product, Tag, QuoteDate, CURVE_CONVENTION_QUOTED, TENORS_11_STANDARD)
    
    If (termMarks(CMS_COL_STATUS) <> REQUEST_STATUS_OK) Then
        CMS_GetCurveQuotedType = termMarks(CMS_COL_ERROR)
    Else
        CMS_GetCurveQuotedType = termMarks(CMS_COL_CURVE_TYPE)
    End If
End Function

' VA 2011-10-04: LIVE Curves only
Public Function CMS_GetCurveLastMarkedOn(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = "DERIV") As Variant

    Dim termMarks() As Variant

    ReDim termMarks(0 To CMS_TERM_DATA_COL_COUNT - 1)

    ' Get Term Data
    termMarks = CMS_GetTermQuoteData(Ticker, Ccy, debtClass, Product)
    
    If (termMarks(CMS_TERM_COL_STATUS) <> REQUEST_STATUS_OK) Then
        CMS_GetCurveLastMarkedOn = termMarks(CMS_TERM_COL_ERROR)
    Else
        CMS_GetCurveLastMarkedOn = termMarks(CMS_TERM_COL_LAST_MARKED_ON)
    End If
End Function

' VA 2011-10-04: LIVE Curves only
Public Function CMS_GetCurveLastMarkedBy(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = "DERIV") As Variant

    Dim termMarks() As Variant

    ReDim termMarks(0 To CMS_TERM_DATA_COL_COUNT - 1)

    ' Get Term Data
    termMarks = CMS_GetTermQuoteData(Ticker, Ccy, debtClass, Product)
    
    If (termMarks(CMS_TERM_COL_STATUS) <> REQUEST_STATUS_OK) Then
        CMS_GetCurveLastMarkedBy = termMarks(CMS_TERM_COL_ERROR)
    Else
        CMS_GetCurveLastMarkedBy = termMarks(CMS_TERM_COL_LAST_MARKED_BY)
    End If
End Function


' VA 2011-10-04: LIVE Curves only
Public Function CMS_GetCurveOwner(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = "DERIV") As Variant

    Dim termMarks() As Variant

    ReDim termMarks(0 To CMS_TERM_DATA_COL_COUNT - 1)

    ' Get Term Data
    termMarks = CMS_GetTermQuoteData(Ticker, Ccy, debtClass, Product)
    
    If (termMarks(CMS_TERM_COL_STATUS) <> REQUEST_STATUS_OK) Then
        CMS_GetCurveOwner = termMarks(CMS_TERM_COL_ERROR)
    Else
        CMS_GetCurveOwner = termMarks(CMS_TERM_COL_OWNER)
    End If
End Function


' VA 2011-10-04: LIVE Curves only
Public Function CMS_GetCurveIsParent(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = "DERIV") As Variant

    Dim termMarks() As Variant

    ReDim termMarks(0 To CMS_TERM_DATA_COL_COUNT - 1)

    ' Get Term Data
    termMarks = CMS_GetTermQuoteData(Ticker, Ccy, debtClass, Product)
    
    If (termMarks(CMS_TERM_COL_STATUS) <> REQUEST_STATUS_OK) Then
        CMS_GetCurveIsParent = termMarks(CMS_TERM_COL_ERROR)
    Else
        CMS_GetCurveIsParent = UCase(termMarks(CMS_TERM_COL_IS_PARENT))
    End If
End Function


Public Function CMS_GetRecoveryRate(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = "DERIV", _
                                     Optional ByVal Tag = "LIVE", Optional ByVal QuoteDate As Date) As Variant
Attribute CMS_GetRecoveryRate.VB_Description = "Returns RecoveryRate (Variant).\r\nIn case of error contains ErrorMessage.\r\n"

    Dim termMarks() As Variant

    ReDim termMarks(0 To CMS_TERM_DATA_COL_COUNT - 1)

    ' Get Term Data
    termMarks = CMS_GetTermQuoteData(Ticker, Ccy, debtClass, Product, Tag, QuoteDate)
    
    If (termMarks(CMS_TERM_COL_STATUS) <> REQUEST_STATUS_OK) Then
        CMS_GetRecoveryRate = termMarks(CMS_TERM_COL_ERROR)
    Else
        CMS_GetRecoveryRate = termMarks(CMS_TERM_COL_RECOVERY)
    End If
End Function

' ====================================================================
' CMS WEB Service: See HY_CDX & LCDX Blotters by Victor Abramovich
' ====================================================================
Private Function GetCmsXmlResponse(rosettaUserMetaData As String, rosettaXmlRequest As String) As String

    Dim webServiceURL As String
    Dim cmsSoapClient As MSOSOAPLib30.SoapClient30
   
   On Error GoTo ERROR_HANDLING
   
'''   ' VA 2009-07-22: As suggested by Gavin Shanks & Merwyn Dsouza,
'''   ' let's try a singleton connection
'''   If (ThisWorkbook.CmsClient Is Nothing) Then
'''        If (ThisWorkbook.CMS_WEB_SERVICE_URL = "") Then
'''         Call RaiseError("CMS Environment got corrupted. Please reconnect to CMS using the menu bar")
'''        End If
'''
'''        Set ThisWorkbook.CmsClient = New SoapClient30
'''        Call ThisWorkbook.CmsClient.MSSoapInit(ThisWorkbook.CMS_WEB_SERVICE_URL)
'''
'''        ' VAS 2009-07-16: Restored
'''        ' VA 2009-06-23: This setting fails in STAGE!!!
'''        ' VA 2009-04-08 D-Day: Timeout issue
'''         ThisWorkbook.CmsClient.ClientProperty("ServerHTTPRequest") = True
'''         ThisWorkbook.CmsClient.ConnectorProperty("Timeout") = CMS_SERVER_GET_TIMEOUT
'''    End If
   
   Set cmsSoapClient = GetCmsConnection()
   GetCmsXmlResponse = ThisWorkbook.CmsClient.GetMarketData(rosettaUserMetaData, rosettaXmlRequest)

'    ' DEBUG
'    If (DEBUG_PRINT_CMS_USER_METADATA) Then
'         Debug.Print "Rosetta XML Response: "
'         Debug.Print GetCmsXmlResponse
'        MsgBox GetCmsXmlResponse
'    End If
    
    Exit Function

ERROR_HANDLING:
    Call RaiseError("Failed to connect to CMS " & CMS_GetServerMode() & " Server:" + Err.Description)
End Function

' ==========================================================================================
' modUtils
' ==========================================================================================

Private Function CMS_ImpersonateUserAs(ImpersonatedUser As String) As String
    If (Len(ImpersonatedUser) = 0) Then
        IMPERSONATED_USER = CMS_GetUserLoginName()
    Else
    IMPERSONATED_USER = ImpersonatedUser
    End If
    CMS_ImpersonateUserAs = IMPERSONATED_USER
End Function

Private Function CMS_GetImpersonatedAs() As String
    If (Len(IMPERSONATED_USER) = 0) Then
        IMPERSONATED_USER = CMS_GetUserLoginName()
    End If
    CMS_GetImpersonatedAs = IMPERSONATED_USER
End Function

Private Function CompileGetUserMetaData(Optional ByVal ImpersonatedUserName As String) As String
    ' Get UserName
    ' Dim userName As String
    If (ImpersonatedUserName = "") Then
        ImpersonatedUserName = CMS_GetUserLoginName()
    End If
    
    ' Encode the ActiveWorkbook.FullName - might break XML parsing on CMS side!
    ' VA 2009-01-28: Rolled back...
    ' VA 2009-04-14: ... back in business
    ' VA 2009-07-13: Added <application-version>
    Dim workbookFullNameEncoded As String
    workbookFullNameEncoded = HTMLEncode(ActiveWorkbook.FullName)
    
    CompileGetUserMetaData = "<user-metadata><user-details>" _
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





Private Function CompileRosettaGetXmlRequest(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, ByVal Product As String, _
                                         ByVal Tag As String, ByVal QuoteDate As Date, ByVal QuoteConvention As String)

    Dim fourTuple As String
    Dim quoteConventionTag As String
    Dim timeStamp As String
    
    ' 4Tuple & upper cases
    fourTuple = CMS_CompileFourTuple(Ticker, Ccy, debtClass, Product)
    Tag = UCase(Tag)
    QuoteConvention = UCase(QuoteConvention)
    
    If (QuoteConvention = CURVE_CONVENTION_RUNNING) Then
        quoteConventionTag = CURVE_CONVENTION_UPFRONT
    Else
        quoteConventionTag = QuoteConvention
    End If
    
    ' Check the QuoteDate - it could be "default" (12/30/1899)
    If (QuoteDate = 0) Then
        QuoteDate = Now()
    End If
    timeStamp = Format(QuoteDate, "yyyyMMdd")
     
    CompileRosettaGetXmlRequest = "<Rosetta version=""5.0.15""><market><marketData><action>GET</action>" _
                             & "<date>" & timeStamp & "</date>" _
                             & CompileRosettaGetMarketSet(, , Ccy) _
                             & "<label>" & fourTuple & "</label>" _
                             & "<genericKeys>" _
                                & "<nameValuePair name=""tag"">" & Tag & "</nameValuePair>" _
                                & "<nameValuePair name=""returnCreditCurveType"">" & quoteConventionTag & "</nameValuePair>" _
                             & "</genericKeys>" _
                             & "</marketSet></marketData></market></Rosetta>"
End Function



Private Function CompileRosettaGetMarketSet(Optional ByVal Location = DEFAULT_LOCATION, Optional ByVal Tag = DEFAULT_DATA_SOURCE, Optional ByVal Ccy = DEFAULT_CCY) As String
    CompileRosettaGetMarketSet = "<marketSet><location>" & Location & "</location>" _
                            & "<currency>" & Ccy & "</currency>" _
                            & "<dataSource>" & Tag & "</dataSource><version>CLOSE</version><type>creditSpread</type>"
End Function


Private Function GetMarketData(ByVal CmsResponseXml As String, ByVal QuoteConvention As String, ByVal TenorStandard As Variant) As Variant

 Dim spreadMarks() As Variant
 
 ' VA 2009-02-27 Now we have contraactual spreads
 Dim contractualSpreads() As Variant
    
 'MSXML Dom object for processing the response
 Dim getMarketDataResponse As MSXML2.DOMDocument
 Dim success As Boolean
 Dim tenorColumnShift As Integer
 Dim ContractualSpread As Variant
 
 ' DEBUG
' Debug.Print "CmsResponseXml"
' Debug.Print CmsResponseXml
 
     ' VA 2009-05-08: Now we handle 3 Tenor Standards (11, 13, & 17)
     ' So, get the corresponding Column Shift & Term Structure
     TERM_STRUCTURE = GetDelimitedTermStructure(TenorStandard)
     tenorColumnShift = GetTenorColumnShift(TenorStandard)
     

 
    ReDim spreadMarks(0 To CMS_COL_COUNT + tenorColumnShift - 1)

    ReDim contractualSpreads(0 To CMS_COL_COUNT + tenorColumnShift - 1)



Set getMarketDataResponse = New MSXML2.DOMDocument
 getMarketDataResponse.async = False
 getMarketDataResponse.validateOnParse = False

 success = getMarketDataResponse.LoadXML(CmsResponseXml)
 
 If (Not success) Then
    RaiseError ("Failed to load CMS response in DOMDocument")
 End If
  
 
 Dim status As MSXML2.IXMLDOMNode
 
 ' VA 2008-12-12
 Dim LastMarkedOn As MSXML2.IXMLDOMNode
 Dim LastMarkedBy As MSXML2.IXMLDOMNode
 Dim Owner        As MSXML2.IXMLDOMNode
 Dim IsParent     As MSXML2.IXMLDOMNode
 
 Dim ErrorMsg As MSXML2.IXMLDOMNode

 Set status = getMarketDataResponse.selectSingleNode("//systemMessages/statusMessage/statusCode")
 If status Is Nothing Then
    RaiseError ("Failed to retrieve Status from CMS Response")
 End If
 
 If StrComp(status.Text, "OK") <> 0 Then
     spreadMarks(CMS_COL_STATUS + tenorColumnShift) = "FAILED"
     Set ErrorMsg = getMarketDataResponse.selectSingleNode("//systemErrors/errorMessage/errorMessageContent")
     If ErrorMsg Is Nothing Then
         spreadMarks(CMS_COL_ERROR + tenorColumnShift) = "Unknown Error Occured while retrieving spreads from CMS"
         ' Debug.Print "Error occured while retrieving spreads "; "Unknown Error Occured"
     Else
         spreadMarks(CMS_COL_ERROR + tenorColumnShift) = Replace(ErrorMsg.Text, STANDARD_JAVA_ERROR, "")
         ' Debug.Print "Error occured while retrieving spreads "; ErrorMsg.Text
     End If
     
 Else
     spreadMarks(CMS_COL_STATUS + tenorColumnShift) = "OK"
     spreadMarks(CMS_COL_ERROR + tenorColumnShift) = ""
     
 
 ' VA 2010-08-12: For Bulk requests we should ignore batch-level error and proceed with processing
 End If
 
     ' -------------------------------------------------
     ' VA 2008-12-12: Extended set of attributes
     ' -------------------------------------------------
   
     ' LastMarkedOn
     ' VA 2008-12-15: New format from Kinjal: "Mon Dec 15 12:22:10 EST 2008"
     Set LastMarkedOn = getMarketDataResponse.selectSingleNode("//genericKeys/nameValuePair[@name='" & NAME_VALUE_LAST_MARKED_ON_TAG & "']")
     If (Not LastMarkedOn Is Nothing) Then
        spreadMarks(CMS_COL_LAST_MARKED_ON + tenorColumnShift) = ConvertToDate(LastMarkedOn.Text)
     End If
     ' LastMarkedBy
     Set LastMarkedBy = getMarketDataResponse.selectSingleNode("//genericKeys/nameValuePair[@name='" & NAME_VALUE_LAST_MARKED_BY_TAG & "']")
     If (Not LastMarkedBy Is Nothing) Then
        spreadMarks(CMS_COL_LAST_MARKED_BY + tenorColumnShift) = LastMarkedBy.Text
     End If
     ' Owner
     Set Owner = getMarketDataResponse.selectSingleNode("//genericKeys/nameValuePair[@name='" & NAME_VALUE_OWNER_TAG & "']")
     If (Not Owner Is Nothing) Then
        spreadMarks(CMS_COL_OWNER + tenorColumnShift) = Owner.Text
     End If
     ' Is Parent
     Set IsParent = getMarketDataResponse.selectSingleNode("//genericKeys/nameValuePair[@name='" & NAME_VALUE_IS_PARENT_TAG & "']")
     If (Not IsParent Is Nothing) Then
        spreadMarks(CMS_COL_IS_PARENT + tenorColumnShift) = IsParent.Text
     End If
     
     ' Quote Type
     spreadMarks(CMS_COL_CURVE_TYPE + tenorColumnShift) = GetQuoteType(getMarketDataResponse)
     
     spreadMarks(CMS_COL_RECOVERY + tenorColumnShift) = GetIssuerRecovery(getMarketDataResponse)
  
    'Process the spreads and write them to the specific cells

    Dim spread As String
    Dim Term As Variant
    
    Dim TermInt As Integer
    TermInt = CMS_COL_TERM_START
    


    ' =============================================================
    
    For Each Term In TERM_STRUCTURE
       spreadMarks(TermInt) = GetTermQuote(Term, getMarketDataResponse, QuoteConvention, ContractualSpread)
       TermInt = TermInt + 1
    Next Term
 
  
  GetMarketData = spreadMarks
  
End Function
    
  Private Function GetTermQuote(ByVal Term As String, ByVal Response As MSXML2.DOMDocument, ByVal QuoteConvention As String, ByRef ContractualSpread As Variant) As Variant
        
        
        Dim termSplit() As String
        Dim pointContractualSpread As MSXML2.DOMDocument
        
        termSplit = Strings.Split(Term, ":")
        
        ' VA 2009-02-27: Now we retrieve contractualSpread as well :-)
        ContractualSpread = Empty
        
        ' VA 2008-12-8: Now we retrieve CreditUpfront as well
        Dim conventionTag As String
        
        If (QuoteConvention = CURVE_CONVENTION_SPREADS Or QuoteConvention = CURVE_CONVENTION_QUOTEDSPREADS) Then
            conventionTag = CREDIT_SPREAD_DATA_TYPE
        Else
            conventionTag = CREDIT_UPFRONT_DATA_TYPE
        End If
                            
        Dim point As MSXML2.IXMLDOMNode
        ' Set point = response.SelectSingleNode("//point[creditSpread[periodMultiplier=" & termSplit(0) & "][period='" & termSplit(1) & "']]")
        Set point = Response.selectSingleNode("//point[" & conventionTag & "[periodMultiplier=" & termSplit(0) & "][period='" & termSplit(1) & "']]")
        
        ' Check for creditSpread in RUNNING mode - upfront curves only!
        If (Not point Is Nothing And QuoteConvention = CURVE_CONVENTION_RUNNING) Then
            conventionTag = "creditSpread"
            Set point = Response.selectSingleNode("//point[" & conventionTag & "[periodMultiplier=" & termSplit(0) & "][period='" & termSplit(1) & "']]")
            
            Else
            ' Check for creditSpread in QUOTE mode
            If (point Is Nothing And QuoteConvention = CURVE_CONVENTION_QUOTED) Then
                conventionTag = "creditSpread"
                Set point = Response.selectSingleNode("//point[" & conventionTag & "[periodMultiplier=" & termSplit(0) & "][period='" & termSplit(1) & "']]")
            End If
        End If
        
        If Not point Is Nothing Then
            ' Retrieve Contractual Spread - if it exists
            ' wrong! Set pointContractualSpread = point.selectSingleNode("//creditUpfront/contractualSpread").Text
            If Not pointContractualSpread Is Nothing Then
                ' Basis Points for Contractual Spread
                ContractualSpread = pointContractualSpread.Text * 10000#
            End If
            
            ' Percents for Upfronts
            GetTermQuote = (point.LastChild.Text) * 100#
            ' Basis points for Spreads
              ' VA 2009-03-24: was a bug - CURVE_CONVENTION_SPREADS
            If (conventionTag = CREDIT_SPREAD_DATA_TYPE) Then
                GetTermQuote = GetTermQuote * 100#
            End If
        End If
        
End Function

Private Function GetIssuerRecovery(ByVal Response As MSXML2.DOMDocument) As Double
                            
        Dim point As MSXML2.IXMLDOMNode
        Set point = Response.selectSingleNode("//point[issuerRecovery]")
                                
        If Not point Is Nothing Then
            GetIssuerRecovery = (point.LastChild.Text)
        End If
        
End Function

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
                ' repl = "&nbsp;"
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

'==============================================================================================
' VA 2010-08-10: Processing Bulk GET Requests
'==============================================================================================
' DISCLAIMER :-)
' ---------------------------------------------------------------------------------------------
' The current code is messy and fragile, so we try to reuse it without any major changes.
' In particular, we'll iterate all the MarketSet nodes in original "bulk" XML response,
' and simulate a "single Curve" response to be processed by existing infrastructure.
' The result should contain a 2-dimensional array of quotes for all the requested curves.
' ---------------------------------------------------------------------------------------------

' VA 2010-08-09: New interface for "bulk" Get requests
Public Function CMS_GetBulkCurveQuoteData(ByVal CurveFourTupleAndConvention As Range, _
                                          Optional ByVal Tag = "LIVE", _
                                          Optional ByVal QuoteDate As Date, _
                                          Optional ByVal TenorStandard As Variant = TENORS_11_STANDARD) As Variant
                                          
    Dim index     As Integer
    Dim curveCount As Integer
    Dim curveIndex As Integer
    Dim itemIndex As Integer
    ' Dim marketSet As String
    ' Dim marketSets As String
    Dim currentSingleCurveXml As String
    
    ' Curve 4Tuple Dictionary
    ' Will be used for Curve-level error processing
    ' Dim curve4TupleDictionary As Dictionary
    
    ' 5Tuple & 4Tuple
    Dim Ticker As String
    Dim Ccy As String
    Dim debtClass As String
    Dim Product As String
    Dim QuoteConvention As String
    Dim fourTuple As String
    
    Dim quoteConventionTag As String
    Dim timeStamp As String

    Dim arrayHelper As New CArrayHelper
    Dim IsExtendedRequest As Boolean
    Dim quoteResults() As Variant
    Dim resultVariant As Variant
    Dim tenorColumnShift As Integer
    
    Dim rosettaUserMetaData As String
    Dim rosettaBulkXmlRequest As String
    Dim cmsBulkResponseXml As String
    Dim cmsProcessingCurveResponse As String
    
    On Error GoTo CMS_GetBulkCurveQuoteData_Err
    
    '--------------------------------------------------------
    ' XML Variables
    '--------------------------------------------------------
    ' MSXML Dom object for processing the response
    Dim marketSetNodes              As MSXML2.IXMLDOMNodeList
    Dim errorMessageContentNodes    As MSXML2.IXMLDOMNodeList
    Dim marketDataNode              As MSXML2.IXMLDOMNode
    Dim marketSetsToProcess()       As MSXML2.IXMLDOMNode
    Dim marketSetNode               As MSXML2.IXMLDOMNode
    Dim parentNode                  As MSXML2.IXMLDOMNode
    Dim marketSetNodeClone          As MSXML2.IXMLDOMNode
    Dim docProcessingCurveResponse  As MSXML2.DOMDocument
    Dim ndStatus                    As MSXML2.IXMLDOMNode
    Dim ndErrorMsg                  As MSXML2.IXMLDOMNode
    Dim loadSuccess                 As Boolean
    Dim errorMessageContents()      As String
    
    ' Curve count
    curveCount = CurveFourTupleAndConvention.Rows.Count
    
    ' Get the column shift for the current Tenor Standard
    tenorColumnShift = GetTenorColumnShift(TenorStandard)
    
    ' Redim - cleanup
    ReDim quoteResults(0 To curveCount - 1, 0 To CMS_COL_COUNT + tenorColumnShift - 1)
    
    ' Valudate the Request
    If CurveFourTupleAndConvention.Columns.Count <> 5 Then
        RaiseError ("Region parameter should contain 5 columns (Curve 4Tuple + RequestType)")
    End If
    
'''    ' Initialize Curve dictionary (for Curve-level error processing)
'''    Set curve4TupleDictionary = New Dictionary
'''    For curveIndex = 1 To curveCount
'''        fourTuple = CMS_CompileFourTuple( _
'''            CurveFourTupleAndConvention(curveIndex, BULK_REQUEST_TICKER_INDEX), _
'''            CurveFourTupleAndConvention(curveIndex, BULK_REQUEST_CCY_INDEX), _
'''            CurveFourTupleAndConvention(curveIndex, BULK_REQUEST_DEBT_CLASS_INDEX), _
'''            CurveFourTupleAndConvention(curveIndex, BULK_REQUEST_PRODUCT_INDEX))
'''
'''        ' There might be duplicates in the original Region!
'''        If (Not curve4TupleDictionary.Exists(fourTuple)) Then
'''            Call curve4TupleDictionary.Add(fourTuple, "")
'''        End If
'''    Next
    
    ' Compile current Curve
    
   
    ' User Metadata
    rosettaUserMetaData = CompileGetUserMetaData(CMS_GetUserLoginName())
    ' ' Debug
    If (DEBUG_PRINT_CMS_USER_METADATA) Then
         Debug.Print "User Metadata: "
         Debug.Print rosettaUserMetaData
        ' MsgBox rosettaUserMetaData
    End If
    
    ' Xml Request
    rosettaBulkXmlRequest = CompileBulkRosettaGetXmlRequest(CurveFourTupleAndConvention, Tag, QuoteDate)
    ' Debug
    If (DEBUG_PRINT_CMS_ROSETTA_XML) Then
        Debug.Print "BULK Rosetta XML Request: "
        Debug.Print rosettaBulkXmlRequest
        'MsgBox rosettaXmlRequest
    End If
    
    ' Get Bulk Response
    cmsBulkResponseXml = GetCmsXmlResponse(rosettaUserMetaData, rosettaBulkXmlRequest)
    cmsProcessingCurveResponse = cmsBulkResponseXml
    
    ' Debug
    If (DEBUG_PRINT_CMS_ROSETTA_XML) Then
        Debug.Print "BULK Rosetta XML Response: "
        Debug.Print cmsBulkResponseXml
        'MsgBox cmsResponse
    End If
        
    On Error GoTo CMS_GetBulkCurveQuoteData_Err
    
    ' Load docProcessingCurveResponse XML into DOMDocument
    Set docProcessingCurveResponse = New MSXML2.DOMDocument
    docProcessingCurveResponse.async = False
    docProcessingCurveResponse.validateOnParse = False
    loadSuccess = docProcessingCurveResponse.LoadXML(cmsProcessingCurveResponse)
    If (Not loadSuccess) Then
       RaiseError ("Failed to load CMS response in DOMDocument")
    End If
    
    ' Get the MarketSet nodes in original "bulk" response
    ' and move it to marketSetProcessingNodes
    
    Set marketDataNode = docProcessingCurveResponse.selectSingleNode("/Rosetta/market/marketData")
    Set marketSetNodes = docProcessingCurveResponse.selectNodes("/Rosetta/market/marketData/marketSet")
    
    ' Error Message Content nodes
    Set errorMessageContentNodes = docProcessingCurveResponse.selectNodes("/Rosetta/systemErrors/errorMessage/errorMessageContent")
    
    ' Create a placeholder for MarketSets to keep and iterate eventually
    ReDim marketSetsToProcess(0 To marketSetNodes.Length - 1)
    
    ' Create and populate an array for ErrorMessageContents
    If (errorMessageContentNodes.Length > 0) Then
    ReDim errorMessageContents(0 To errorMessageContentNodes.Length - 1)
    For index = 0 To errorMessageContentNodes.Length - 1
        errorMessageContents(index) = errorMessageContentNodes(index).Text
    Next
    Else
        ReDim errorMessageContents(0 To 0)
    End If
    
    ' http://stackoverflow.com/questions/875136/how-to-remove-an-xmlnode-from-xmlnodelist
    For index = (marketSetNodes.Length - 1) To 0 Step -1
        Set marketSetNode = marketSetNodes(index)
        Set marketSetNodeClone = marketSetNode.CloneNode(True)
        Set parentNode = marketSetNode.parentNode
        Call parentNode.removeChild(marketSetNode)
        Set marketSetsToProcess(index) = marketSetNodeClone
    Next
    
    ' Debug
    If (DEBUG_PRINT_CMS_ROSETTA_XML) Then
        Debug.Print "BULK Response with removed marketSets:"
        Debug.Print docProcessingCurveResponse.childNodes(0).XML
    
    End If
    
'    ' "Local" single Curve error processing
'    On Error GoTo CURRENT_CURVE_ERROR
    
    ' =============================================================
    ' Major cycle by Curves
    ' =============================================================
    For curveIndex = 0 To curveCount - 1
        
        ' Get current 5Tuple = 4Tuple + QuoteConvention
        Ticker = CurveFourTupleAndConvention(curveIndex + 1, BULK_REQUEST_TICKER_INDEX)
        Ccy = CurveFourTupleAndConvention(curveIndex + 1, BULK_REQUEST_CCY_INDEX)
        debtClass = CurveFourTupleAndConvention(curveIndex + 1, BULK_REQUEST_DEBT_CLASS_INDEX)
        Product = CurveFourTupleAndConvention(curveIndex + 1, BULK_REQUEST_PRODUCT_INDEX)
        QuoteConvention = CurveFourTupleAndConvention(curveIndex + 1, BULK_REQUEST_QUOTE_CONVENTION_INDEX)
        
        ' Compile current Curve 4Tuple
        fourTuple = CMS_CompileFourTuple(Ticker, Ccy, debtClass, Product)
                 
        If (fourTuple = "..." Or fourTuple = Empty) Then
            quoteResults(curveIndex, CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_FAILED
            quoteResults(curveIndex, CMS_COL_ERROR + tenorColumnShift) = "Curve FourTuple is empty"
            GoTo NEXT_CURVE_INDEX
        End If
        
        ' VA 2010-08-13: We do not allow multi-row responses for BULK requests!
        If (QuoteConvention = CURVE_CONVENTION_EXTENDED) Then
            QuoteConvention = CURVE_CONVENTION_QUOTED
        End If
        
        ' Add current marketSet to marketDataNode
        Call marketDataNode.appendChild(marketSetsToProcess(curveIndex))
        
        ' Get XML response for current single Curve
        currentSingleCurveXml = docProcessingCurveResponse.childNodes(0).XML
        
'        ' Debug - TOO VERBOSE!
'        If (DEBUG_PRINT_CMS_ROSETTA_XML) Then
'            Debug.Print "BULK Response with current marketSet # " & curveIndex + 1 & " out of " & curveCount; ":"
'            Debug.Print currentSingleCurveXml
'        End If
        
        ' Get Curve marks for the current marketSet
        resultVariant = CMS_GetCurveQuoteDataImpl(Ticker, Ccy, debtClass, _
           Product, Tag, QuoteDate, QuoteConvention, TenorStandard, currentSingleCurveXml)
        
        ' VA 2010-08-13: Now, process the potential BULK Error messages
        Call ProcessBulkErrorMessages(fourTuple, resultVariant, tenorColumnShift, errorMessageContents)
        
        ' Populate current Curve row in the output array
        ' VA 2010-08-13: Take care of Ad Hoc Tenors here!!!
        For itemIndex = 0 To UBound(resultVariant)
            quoteResults(curveIndex, itemIndex) = resultVariant(itemIndex)
        Next itemIndex
        
        ' Remove current marketSet from marketDataNode, get ready for the next iteration
        Call marketDataNode.removeChild(marketSetsToProcess(curveIndex))
        
        
        ' Go to NEXT_CURVE_INDEX
        GoTo NEXT_CURVE_INDEX
        
CURRENT_CURVE_ERROR:
        quoteResults(curveIndex, CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_FAILED
        quoteResults(curveIndex, CMS_COL_ERROR + tenorColumnShift) = Err.Description

NEXT_CURVE_INDEX:
    Next curveIndex
 
GO_TO_EXIT:
    GoTo CMS_GetBulkCurveQuoteData_Exit

CMS_GetBulkCurveQuoteData_Err:
    For curveIndex = 0 To curveCount - 1
        quoteResults(curveIndex, CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_FAILED
        quoteResults(curveIndex, CMS_COL_ERROR + tenorColumnShift) = Err.Description
    Next
    Debug.Print Err.Description
    
    ' VA 2009-06-23: !!! This excludes #VALUE!'s from the sheet
    Resume CMS_GetBulkCurveQuoteData_Exit
    
CMS_GetBulkCurveQuoteData_Exit:
    CMS_GetBulkCurveQuoteData = quoteResults
    Exit Function
End Function


' VA 2010-08-09
Private Function CompileBulkRosettaGetXmlRequest(ByVal FiveTuples As Range, _
                                ByVal Tag As String, ByVal QuoteDate As Date)

'Private Function CompileRosettaGetXmlRequest(ByVal Ticker As String, ByVal Ccy As String, ByVal DebtClass As String, ByVal Product As String, _
'                                         ByVal Tag As String, ByVal QuoteDate As Date, ByVal QuoteConvention As String)
    
    Dim curveCount As Integer
    Dim index As Integer
    Dim marketSet As String
    Dim marketSets As String
    
    ' 5Tuple & 4Tuple
    Dim Ticker As String
    Dim Ccy As String
    Dim debtClass As String
    Dim Product As String
    Dim QuoteConvention As String
    Dim fourTuple As String
    
    Dim quoteConventionTag As String
    Dim timeStamp As String
    
    marketSets = ""
    
    curveCount = FiveTuples.Rows.Count
    
    For index = 1 To curveCount
        Ticker = FiveTuples(index, BULK_REQUEST_TICKER_INDEX)
        Ccy = FiveTuples(index, BULK_REQUEST_CCY_INDEX)
        debtClass = FiveTuples(index, BULK_REQUEST_DEBT_CLASS_INDEX)
        Product = FiveTuples(index, BULK_REQUEST_PRODUCT_INDEX)
        QuoteConvention = FiveTuples(index, BULK_REQUEST_QUOTE_CONVENTION_INDEX)
        
        ' 4Tuple & upper cases
        fourTuple = CMS_CompileFourTuple(Ticker, Ccy, debtClass, Product)
        Tag = UCase(Tag)
        QuoteConvention = UCase(QuoteConvention)
        
        ' VA 2010-08-12: We do not allow multi-row responses for BULK requests!
        If ((QuoteConvention = CURVE_CONVENTION_RUNNING) Or (QuoteConvention = CURVE_CONVENTION_EXTENDED)) Then
            quoteConventionTag = CURVE_CONVENTION_UPFRONT
        Else
            quoteConventionTag = QuoteConvention
        End If
        
        ' Check the QuoteDate - it could be "default" (12/30/1899)
        If (QuoteDate = 0) Then
            QuoteDate = Now()
        End If
        timeStamp = Format(QuoteDate, "yyyyMMdd")
     
        marketSet = CompileRosettaGetMarketSetNew(fourTuple, quoteConventionTag, , Tag, Ccy)
        marketSets = marketSets + marketSet
    
    Next index
    
               
    CompileBulkRosettaGetXmlRequest = "<Rosetta version=""5.0.15""><market><marketData><action>GET</action>" _
                             & "<date>" & timeStamp & "</date>" _
                             & marketSets & "</marketData></market></Rosetta>"
                             
                             
                             
End Function


' VA 2010-08-09
Private Function CompileRosettaGetMarketSetNew(fourTuple As String, quoteConventionTag As String, Optional ByVal Location = DEFAULT_LOCATION, Optional ByVal Tag = DEFAULT_DATA_SOURCE, Optional ByVal Ccy = DEFAULT_CCY) As String
    CompileRosettaGetMarketSetNew = "<marketSet><location>" & Location & "</location>" _
                            & "<currency>" & Ccy & "</currency>" _
                            & "<dataSource>" & Tag & "</dataSource><version>CLOSE</version><type>creditSpread</type>" _
                            & "<label>" & fourTuple & "</label>" _
                            & "<genericKeys>" _
                            & "<nameValuePair name=""tag"">" & Tag & "</nameValuePair>" _
                            & "<nameValuePair name=""returnCreditCurveType"">" & quoteConventionTag & "</nameValuePair>" _
                            & "</genericKeys>" _
                            & "</marketSet>"
End Function


' VA 2010-08-13: This method maps/clears Curve Status & Error
' based on the batch-level Error collection.
' This is a temporary hack before CMS implement Curve-level Error messaging
' in XML Response.
Private Sub ProcessBulkErrorMessages(currentCurve4Tuple As String, curveQuoteResults As Variant, _
              tenorColumnShift As Integer, errorMessageContents() As String)

    Dim curveIndex As Integer
    Dim currentErrorIndex As Integer
    Dim currentCurveError As String
    Dim currentCurveRecovery As Variant
    Dim defaultErrorMessage As String
    
    defaultErrorMessage = "Unknown"
    
    ' Let's be optimistic first... :-)
    curveQuoteResults(CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_OK
    curveQuoteResults(CMS_COL_ERROR + tenorColumnShift) = ""

    ' Now, map batch-level errors to current Curve
    For currentErrorIndex = 0 To UBound(errorMessageContents)
        currentCurveError = errorMessageContents(currentErrorIndex)
        If (InStr(UCase(currentCurveError), UCase(TOO_MANY_ERRORS_CMS_MESSAGE))) Then
            defaultErrorMessage = "Too many errors; skipping current Curve."
        ElseIf (InStr(UCase(currentCurveError), UCase(currentCurve4Tuple))) Then
            ' The error refers current Curve
            curveQuoteResults(CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_FAILED
            curveQuoteResults(CMS_COL_ERROR + tenorColumnShift) = currentCurveError
            ' Debug
            Debug.Print "Error in '" & currentCurve4Tuple & "' Curve: " & currentCurveError
        End If
    Next currentErrorIndex
    
    ' Even if there is no curve-specific error yet Recovery is missing,
    ' we still treat it as an error with "Unknown" or "Too many errors" reason
'''    currentCurveRecovery = curveQuoteResults(CMS_COL_RECOVERY + tenorColumnShift)
'''    If ((curveQuoteResults(CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_OK _
'''    And (currentCurveRecovery = Empty))) _
'''    Or (curveQuoteResults(CMS_COL_STATUS + tenorColumnShift) <> REQUEST_STATUS_OK _
'''    And (curveQuoteResults(CMS_COL_ERROR + tenorColumnShift) = Empty)) Then
'''            curveQuoteResults(CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_FAILED
'''            curveQuoteResults(CMS_COL_ERROR + tenorColumnShift) = defaultErrorMessage
'''    End If
    currentCurveRecovery = curveQuoteResults(CMS_COL_RECOVERY + tenorColumnShift)
    If ((curveQuoteResults(CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_OK _
    And (currentCurveRecovery = Empty))) Then
'    Or (curveQuoteResults(CMS_COL_STATUS + tenorColumnShift) <> REQUEST_STATUS_OK _
'    And (curveQuoteResults(CMS_COL_ERROR + tenorColumnShift) = Empty)) Then
            curveQuoteResults(CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_FAILED
            curveQuoteResults(CMS_COL_ERROR + tenorColumnShift) = defaultErrorMessage
            Debug.Print "Too many errors; skipping '" & currentCurve4Tuple & "'."
    
    
    
    
    End If
    
    

End Sub


Public Function CMS_GetCurveInterpolationType(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = "DERIV", _
                                       Optional ByVal Tag = "LIVE", Optional ByVal QuoteDate As Date) As Variant

    On Error GoTo CMS_GetCurveInterpolationType_Err
    
    CMS_GetCurveInterpolationType = CMS_GetCurveInterpolationTypeImpl(Ticker, Ccy, debtClass, Product, Tag, QuoteDate)
            
    Exit Function
    
CMS_GetCurveInterpolationType_Err:
    
    CMS_GetCurveInterpolationType = "Error Occured"

End Function


Private Function CMS_GetCurveInterpolationTypeImpl(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, Optional ByVal Product = DEFAULT_PRODUCT, _
                                     Optional ByVal Tag = "LIVE", Optional ByVal QuoteDate As Date) As Variant

    Dim rosettaUserMetaData As String
    Dim rosettaXmlRequest As String
    Dim cmsResponse As String
        
    On Error GoTo CMS_GetCurveInterpolationTypeImpl_Err
    
    rosettaUserMetaData = CompileGetUserMetaData(CMS_GetUserLoginName())
    
    If (DEBUG_PRINT_CMS_USER_METADATA) Then
         Debug.Print "User Metadata: "
         Debug.Print rosettaUserMetaData
    End If
    
    rosettaXmlRequest = CompileRosettaGetXmlRequest(Ticker, Ccy, debtClass, Product, Tag, QuoteDate, CURVE_CONVENTION_QUOTED)
    
    If (DEBUG_PRINT_CMS_ROSETTA_XML) Then
        Debug.Print "Rosetta XML Request: "
        Debug.Print rosettaXmlRequest
    End If
    
    cmsResponse = GetCmsXmlResponse(rosettaUserMetaData, rosettaXmlRequest)
    
    If (DEBUG_PRINT_CMS_ROSETTA_XML) Then
        Debug.Print "Rosetta XML Response: "
        Debug.Print cmsResponse
    End If
    
    CMS_GetCurveInterpolationTypeImpl = GetInterpolationTypeFromResponse(cmsResponse)
    
    Exit Function
    
CMS_GetCurveInterpolationTypeImpl_Err:
            
    If (Err.Description <> "") Then
        RaiseError (Err.Description)
    Else
        RaiseError ("Unknown Error Occured while retrieving Interpolation Type from CMS")
    End If
    
End Function

Private Function GetInterpolationTypeFromResponse(ByVal CmsResponseXml As String)
    
    '--------------------------------------------------------
    ' XML Variables
    '--------------------------------------------------------
    ' MSXML Dom object for processing the response
    Dim docMarketDataResponse   As MSXML2.DOMDocument
    Dim ndInterpolationType     As MSXML2.IXMLDOMNode
    Dim ndStatus                As MSXML2.IXMLDOMNode
    Dim ndErrorMsg              As MSXML2.IXMLDOMNode

    Dim loadSuccess             As Boolean
        
    On Error GoTo GetInterpolationTypeFromResponse_Err
    
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
        Set ndErrorMsg = docMarketDataResponse.selectSingleNode("//systemErrors/errorMessage/errorMessageContent")
        If ndErrorMsg Is Nothing Then
            RaiseError ("Unknown Error Occured while retrieving Interpolation Type from CMS")
        Else
            RaiseError (Replace(ndErrorMsg.Text, STANDARD_JAVA_ERROR, ""))
        End If
    End If
    
    Set ndInterpolationType = docMarketDataResponse.selectSingleNode("//genericKeys/nameValuePair[@name='" & NAME_VALUE_INTERPOLATION_TYPE_TAG & "']")
     
    If (Not ndInterpolationType Is Nothing) Then
        GetInterpolationTypeFromResponse = ndInterpolationType.Text
    Else
        GetInterpolationTypeFromResponse = "Unknown"
    End If
   
   Exit Function
   
GetInterpolationTypeFromResponse_Err:

     If (Err.Description <> "") Then
        RaiseError (Err.Description)
     Else
        RaiseError ("Unknown Error Occured while retrieving Interpolation Type from CMS")
     End If
     
End Function







