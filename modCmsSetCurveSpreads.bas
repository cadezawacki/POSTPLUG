
Attribute VB_Name = "modCmsSetCurveSpreads"
'------------------------------------------------------------------------------------------------
' VA 2008-12-18: DOES ANYBODY KNOW HOW TO HIDE SOME PUBLIC MODULE FUNCTIONS IN XLA FROM THE USER?
' OH, IF YOU DON'T, ENJOY DUPLICATES HERE FROM modCmsGetCurveSpreads! :-)
'------------------------------------------------------------------------------------------------

Option Explicit

'----------------------------------------------------------------------------------
' VA 2008-11-21: XML functionality originally provided by Chaitanya Penubarthi
'----------------------------------------------------------------------------------
Private TERM_STRUCTURE() As Variant
Private FULL_TERM_STRUCTURE_NAMES() As Variant

Public Function CMS_SetCurveRecoveryRate(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, ByVal Product, _
                                         ByVal RecoveryRate As Double) As Variant
    
    Dim CurveQuoteData() As Variant
    
    ' Problems with identifications of empty array in VBA!
    Dim termQuoteArray(0) As Variant
    Dim riskQuoteData(0 To CMS_RISK_DATA_COL_COUNT - 1)
    Dim index As Integer
    Dim colDelta As Integer
    
    On Error GoTo CMS_SetCurveRecoveryRate_Err
    
    CurveQuoteData = CMS_SetCurveQuoteDataImpl(Ticker, Ccy, debtClass, Product, _
                                                       termQuoteArray(), _
                                                       RecoveryRate, _
                                                       CURVE_CONVENTION_SPREADS, _
                                                       DEFAULT_LOCATION)
                                                       
    ' Get Risk as a first column
    riskQuoteData(0) = CurveQuoteData(CMS_COL_RECOVERY)
    
    ' Now get the rest
    colDelta = CMS_COL_COUNT - CMS_RISK_DATA_COL_COUNT
    For index = (CMS_TERM_COUNT + 1) To (CMS_COL_COUNT - 1)
        riskQuoteData(index - colDelta) = CurveQuoteData(index)
    Next
    
    ' VA 2009-07-16: Fighting #VALUE!...
    Dim errorMessage As String
    errorMessage = riskQuoteData(UBound(riskQuoteData))
    If (errorMessage <> "") Then
        RaiseError (errorMessage)
    End If
    
    CMS_SetCurveRecoveryRate = riskQuoteData
    
    Exit Function
    
CMS_SetCurveRecoveryRate_Err:
    CMS_SetCurveRecoveryRate = "ERROR: " & Err.Description

End Function


Public Function CMS_SetFlatCurveQuotes(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, ByVal Product, _
                                            ByVal termQuotesValue As Double, _
                                            Optional ByVal RecoveryRate = DUMMY_DOUBLE_VALUE, _
                                            Optional ByVal QuoteConvention = CURVE_CONVENTION_SPREADS, _
                                            Optional ByVal TenorStandard As Variant = TENORS_11_STANDARD) As Variant
    
    Dim CurveQuoteData() As Variant
    Dim TermQuoteData() As Variant
    Dim termQuoteArray() As Variant
    Dim termQuoteCount As Integer
    Dim TermStructure As Variant
    Dim rngTenorStructure As Range
    
    Dim index As Integer
    Dim colDelta As Integer
    Dim tenorColumnShift As Integer
    Dim arrayHelper As New CArrayHelper
    
    On Error GoTo CMS_SetFlatCurveQuotes_Err
    
    ' VA 2009-07-14: Now we support Ad Hoc Tenors
    If (IsObject(TenorStandard) And Not IsNumeric(TenorStandard)) Then
        Set rngTenorStructure = TenorStandard
        TermStructure = GetAdHocTenorsFromRange(rngTenorStructure)
        termQuoteCount = UBound(TermStructure) + 1
    Else
        termQuoteCount = TenorStandard
    End If
    
    
    
    tenorColumnShift = GetTenorColumnShift(TenorStandard)
    
    ' VA 2009-07-14: Was it wrong all the time???
    ' ReDim TermQuoteData(0 To TenorStandard - 1)
    ReDim TermQuoteData(0 To CMS_TERM_DATA_COL_COUNT - 1)
    
    ReDim termQuoteArray(1 To termQuoteCount)

    For index = 1 To termQuoteCount
        termQuoteArray(index) = termQuotesValue
    Next

    CurveQuoteData = CMS_SetCurveQuoteDataImpl(Ticker, Ccy, debtClass, Product, _
                                                       termQuoteArray(), _
                                                       RecoveryRate, _
                                                       QuoteConvention, _
                                                       DEFAULT_LOCATION, _
                                                       TenorStandard)
                                                       
   ' VA 2009-04-20: SetFlat is an incorrect request for SNAC curves,
   ' that would return a 2D array
    If (arrayHelper.NumberOfArrayDimensions(CurveQuoteData) > 1) Then
        Dim arrayWithError() As Variant
        ReDim arrayWithError(0 To CMS_TERM_DATA_COL_COUNT - 1)
        arrayWithError(CMS_TERM_COL_STATUS) = "FAILED"
        arrayWithError(CMS_TERM_COL_ERROR) = "'Set Flat' can't by applied to Standard Contract curve"
        CMS_SetFlatCurveQuotes = arrayWithError
        Exit Function
    End If
    
    ' VA 2009-07-14: Now it might be ad hoc, so let's take 0-th.
    ' Get 5Y quote as a representative of a flat curve
    'TermQuoteData(0) = curveQuoteData(CMS_COL_5Y + tenorColumnShift)
    TermQuoteData(0) = CurveQuoteData(0)
    
'''    ' Now get the rest
    colDelta = CMS_TERM_COUNT + tenorColumnShift + 1
    ' VA 2009-07-14
    For index = 1 To UBound(TermQuoteData)
        TermQuoteData(index) = CurveQuoteData(index + termQuoteCount - 1)
    Next
    
    
    
    ' VA 2009-07-16: Fighting #VALUE!...
    Dim errorMessage As String
    errorMessage = TermQuoteData(UBound(TermQuoteData))
    If (errorMessage <> "") Then
        RaiseError (errorMessage)
    End If
    
    CMS_SetFlatCurveQuotes = TermQuoteData
    
    Exit Function
    
CMS_SetFlatCurveQuotes_Err:
    CMS_SetFlatCurveQuotes = "ERROR: " & Err.Description
End Function

Public Function CMS_SetCurveQuoteData(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, ByVal Product, _
                                            ByRef TermQuotes As Range, _
                                            Optional ByVal RecoveryRate = DUMMY_DOUBLE_VALUE, _
                                            Optional ByVal QuoteConvention = CURVE_CONVENTION_SPREADS, _
                                            Optional ByVal AdHocTenorStandard As Variant = Nothing) As Variant

    Dim functionResultArray() As Variant
    Dim TenorStandard As Variant
    Dim termQuoteArray() As Variant
    Dim termQuoteCount As Integer
    Dim tenorStandardArray() As Variant
    Dim tenorStandardCount As Integer
    Dim CurveQuoteData() As Variant
    ReDim CurveQuoteData(0 To CMS_COL_COUNT - 1)
    Dim rosettaUserMetaData As String
    Dim rosettaXmlRequest As String
    Dim cmsResponse As String
    ' Dim TenorStandard As Integer
    Dim IsSnacCurve As Boolean
    Dim ContractualSpread As Double
    Dim rowCount As Integer
    
    On Error GoTo Process_Err
    
    ' VA 2009-07-13: Now we handle dynamic Tenor list for SETs as well
    ' Check for TenorStandard - it should be either nothing,
    ' or row/column.
    If (Not AdHocTenorStandard Is Nothing) Then
        ' Transpose to a "Landscape" if necessary - and sorry, I am lazy... :-)
         If (AdHocTenorStandard.Rows.Count > AdHocTenorStandard.Columns.Count) Then
             tenorStandardArray = WorksheetFunction.Transpose(AdHocTenorStandard)
         Else
             tenorStandardArray = WorksheetFunction.Transpose(WorksheetFunction.Transpose(AdHocTenorStandard))
         End If
        tenorStandardCount = UBound(tenorStandardArray)
    End If
    
    ' VA 2009-06-19: We now handle SNAC & STEC curves both as STANDARD_CONTRACT curves.
    ' Temporarily keep using IsSnacCurve as IsStandardContract
    ' On the other hand, we now don't care about ContractualSpread value -
    ' it's just a sign for CMS that we are dealing with SNAC/STEC curves.
    ContractualSpread = CONTRACTUAL_SPREAD_DUMMY
    IsSnacCurve = IsStandardContractProduct(Product)
       
    ' -----------------------------------------------------
    ' Check the correctness of TermQuotes dimensions
    ' -----------------------------------------------------
    If (TermQuotes Is Nothing) Then
        RaiseError ("'TermQuote' Range is empty")
    End If
    
    ' Transpose to a "Landscape" if necessary - and sorry, I am lazy... :-)
    If (TermQuotes.Rows.Count > TermQuotes.Columns.Count) Then
        termQuoteArray = WorksheetFunction.Transpose(TermQuotes)
    Else
        termQuoteArray = WorksheetFunction.Transpose(WorksheetFunction.Transpose(TermQuotes))
    End If
    
    rowCount = DimensionCount(termQuoteArray)
    
    Select Case (rowCount)
        Case 1:
            termQuoteCount = UBound(termQuoteArray)
            ' VA 2009-07-13: Check for ad hoc Tenors
            If (tenorStandardCount > 0) Then
                ' Number of Tenors should match
                If (tenorStandardCount <> termQuoteCount) Then
                    RaiseError ("Number of Tenors in TenorStandard (" & tenorStandardCount & ") and in TermQuotes (" & termQuoteCount & ") do not match!")
                End If
                ' VA 2009-07-14: 'SET' issue3 again
                Set TenorStandard = AdHocTenorStandard
            
            ' STANDARD Tenors
            Else
                ' VA 2009-06-22: Adding 12 tenors
                ' VA 2009-05-11: We DON'T save 17 tenors - just yet :-)
                ' VA 2009-03-02: Check for new TENOR_STANDARD
                
                TenorStandard = UBound(termQuoteArray)
                Select Case (UBound(termQuoteArray))
                    Case TENORS_13_STANDARD:
                        ' TenorStandard = TENORS_13_STANDARD
                    Case TENORS_12_STANDARD:
                        ' TenorStandard = TENORS_12_STANDARD
                    Case TENORS_11_STANDARD:
                        ' TenorStandard = TENORS_11_STANDARD
                    Case Else
                        ' VA 2009-04-27: Now SNAC might be just one row as well...
                        ' RaiseError ("Invalid 'TermQuote' Range dimensions: should be (1x11), or (1x13) for PAR curves; (2x11), or (2x13) for SNAC curves. Transposed ranges are allowed as well.")
                        RaiseError ("'TermQuote' Range has non-standard number of Tenors (i.e., not 11,12 or 13). Please provide AdHocTenorStandard array.")
                End Select
            
            End If
        
        ' VA 2009-04-27: "Old-fashioned" SNAC curves
        ' Should be SNAC curve
        Case 2:
            ' Should be SNAC curve
            If (Not IsSnacCurve) Then
                    RaiseError ("Wrong Quote array dimensions - double set is reserved for Standard curves only.")
            End If
            
            ' VA 2009-03-02: Check for new TENOR_STANDARD
            TenorStandard = UBound(termQuoteArray, 2)
            Select Case (UBound(termQuoteArray, 2))
                Case TENORS_13_STANDARD:
                    ' TenorStandard = TENORS_13_STANDARD
                Case TENORS_12_STANDARD:
                    ' TenorStandard = TENORS_12_STANDARD
                Case TENORS_11_STANDARD:
                    ' TenorStandard = TENORS_11_STANDARD
                Case Else
                    RaiseError ("Invalid 'TermQuote' Range dimensions: should have 11,12 or 13 columns. Transposed ranges are allowed as well.")
            End Select
    End Select
    

    
    functionResultArray = CMS_SetCurveQuoteDataImpl(Ticker, Ccy, debtClass, Product, _
                                                      termQuoteArray(), _
                                                      RecoveryRate, _
                                                      QuoteConvention, _
                                                      DEFAULT_LOCATION, _
                                                      TenorStandard, _
                                                      IsSnacCurve, _
                                                      ContractualSpread)
                                                      
  
 
    ' Dim errorMessage As String
    Dim errorMessage As String
    If (UBound(functionResultArray) = 1) Then
        errorMessage = functionResultArray(0, UBound(functionResultArray, 2))
    Else
        errorMessage = functionResultArray(UBound(functionResultArray))
    End If
    If (errorMessage <> "") Then
        RaiseError (errorMessage)
    End If
    
   CMS_SetCurveQuoteData = functionResultArray
    
    Exit Function
 
Process_Err:
   CMS_SetCurveQuoteData = "ERROR: " & Err.Description
End Function


Private Function CMS_SetCurveQuoteDataImpl(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, ByVal Product, _
                                            ByRef termQuoteArray() As Variant, _
                                            Optional ByVal RecoveryRate = DUMMY_DOUBLE_VALUE, _
                                            Optional ByVal QuoteConvention = CURVE_CONVENTION_SPREADS, _
                                            Optional ByVal Location = DEFAULT_LOCATION, _
                                            Optional ByVal TenorStandard As Variant = TENORS_11_STANDARD, _
                                            Optional ByVal IsSnacCurve = False, _
                                            Optional ByVal ContractualSpread = 500)

    Dim functionResultArray() As Variant
    
    Dim CurveQuoteData() As Variant
    
    Dim rosettaUserMetaData As String
    Dim rosettaXmlRequest As String
    Dim cmsResponse As String
    Dim tenorColumnShift As Integer
    
    On Error GoTo CMS_GetCurveQuoteData_Err
    
    ' VA 2009-04-30: Now SNAC might and mostly will be 1 row...
    Dim termQuoteRowCount As Integer
    Dim arrayHelper As New CArrayHelper
    termQuoteRowCount = arrayHelper.NumberOfArrayDimensions(termQuoteArray)
    
    ' VA 2009-02-25
    tenorColumnShift = GetTenorColumnShift(TenorStandard)
    ReDim CurveQuoteData(0 To CMS_COL_COUNT + tenorColumnShift - 1)
        
    ' Compile User & MarketData parts of XML Request
    rosettaUserMetaData = CompileSetUserMetaData(ThisWorkbook.IMPERSONATED_USER_NAME)
    rosettaXmlRequest = CompileRosettaSetXmlRequest(UCase(Ticker), UCase(Ccy), UCase(debtClass), UCase(Product), termQuoteArray, UCase(QuoteConvention), RecoveryRate, UCase(Location), _
                                                    TenorStandard, IsSnacCurve, ContractualSpread)

    ' *** PRINT ***
    'Public Const DEBUG_PRINT_CMS_USER_METADATA As Boolean = False
    'Public Const DEBUG_PRINT_CMS_ROSETTA_XML As Boolean = True

     ' Debug
     
   ' ' Debug
    If (DEBUG_PRINT_CMS_USER_METADATA) Then
         Debug.Print "User Metadata: "
         Debug.Print rosettaUserMetaData
        ' MsgBox rosettaUserMetaData
    End If
    
    ' Debug
    If (DEBUG_PRINT_CMS_ROSETTA_XML) Then
        Debug.Print "Rosetta XML Request: "
        Debug.Print rosettaXmlRequest
        'MsgBox rosettaXmlRequest
    End If
                                
    cmsResponse = GetCmsXmlResponse(rosettaUserMetaData, rosettaXmlRequest)
    
    ' Debug
    If (DEBUG_PRINT_CMS_ROSETTA_XML) Then
        Debug.Print "Rosetta XML Response: "
        Debug.Print cmsResponse
        'MsgBox cmsResponse
    End If
    
    
    ' =============================================================
    ' VA 2008-03-09: Using CMarketData Class now
    ' =============================================================
    Dim currentMarketData As CMarketData
    Dim setPresenData() As Variant
    Dim index As Integer
        
    Set currentMarketData = New CMarketData
    currentMarketData.Initialize (TenorStandard)

    ' VA 2009-03-09: No time to clean it up... :-(
    Call currentMarketData.GetMarketDataFromResponse(cmsResponse, True)
    
'''    Select Case (currentMarketData.Product)
'''       Case "SNAC100", "SNAC500":
       
       ' VA 2009-06-19: Using concept of STANDARD CONTRACT Curves now
       If (IsSnacCurve) Then
            
            Dim setPresentationData() As Variant
            
            ' Old-fashioned 2-row SNAC marking
            If (termQuoteRowCount = 2) Then
                ReDim setPresentationData(0 To 1, 0 To UBound(currentMarketData.MarketDataSpreads))
                For index = 0 To UBound(currentMarketData.MarketDataSpreads)
                    setPresentationData(0, index) = currentMarketData.MarketDataSpreads(index)
                    setPresentationData(1, index) = currentMarketData.MarketDataUpfronts(index)
                Next
                CMS_SetCurveQuoteDataImpl = setPresentationData
                
                Exit Function
            
            ' New 1-row SNAC marking
            Else
                ReDim setPresentationData(0 To UBound(currentMarketData.MarketDataSpreads))
                For index = 0 To UBound(currentMarketData.MarketDataSpreads)
                    If (QuoteConvention = CURVE_CONVENTION_QUOTEDSPREADS) Then
                        setPresentationData(index) = currentMarketData.MarketDataSpreads(index)
                    Else
                    setPresentationData(index) = currentMarketData.MarketDataUpfronts(index)
                    
'''                    ' VA 2009-06-22: Hunting for #VALUE! in results...
'''                    If IsError(setPresentationData(index)) Then
'''                        MsgBox "Here!" & setPresentationData(index)
'''                    End If
                    
                    End If
                Next
                functionResultArray = setPresentationData
            End If
            
 
            
        Else
            functionResultArray = GetMarketData(cmsResponse, QuoteConvention, TenorStandard)
        End If
        
'''        ' VA 2009-07-15: It looks like Empty values are rendered as #VALUE!
'''        CMS_SetCurveQuoteDataImpl = arrayHelper.ReplaceEmptyArrayValues(CMS_SetCurveQuoteDataImpl)
        
        ' VA 2009-07-16: So far, failing to beat #VALUE! on the returned Range...
        Dim errorMessage As String
        errorMessage = functionResultArray(UBound(functionResultArray))
        If (errorMessage <> "") Then
            RaiseError (errorMessage)
        End If
        
        CMS_SetCurveQuoteDataImpl = functionResultArray
    
    Exit Function
    
    
CMS_GetCurveQuoteData_Err:
        CurveQuoteData(CMS_COL_STATUS + tenorColumnShift) = REQUEST_STATUS_FAILED
        CurveQuoteData(CMS_COL_ERROR + tenorColumnShift) = Err.Description
        CMS_SetCurveQuoteDataImpl = CurveQuoteData
    
    ' CMS_SetCurveQuoteDataImpl = "ERROR: " & Err.Description
    
    ' VA 2009-06-23: !!! This removes #VALUE!'s from the sheet
    Resume CMS_GetCurveQuoteData_Exit
CMS_GetCurveQuoteData_Exit:

End Function

' ====================================================================
' CMS WEB Service: See HY_CDX & LCDX Blotters by Victor Abramovich
' ====================================================================
Private Function GetCmsXmlResponse(rosettaUserMetaData As String, rosettaXmlRequest As String) As String

    Dim dateStart As Date
    Dim dateEnd As Date
    Dim cmsSoapClient As MSOSOAPLib30.SoapClient30
    
   On Error GoTo GetCmsXmlResponse_Err
   
'''    ' VA 2009-07-22: As suggested by Gavin Shanks & Merwyn Dsouza,
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
'''
'''   ' dateStart = Time()

    Set cmsSoapClient = GetCmsConnection()
    GetCmsXmlResponse = ThisWorkbook.CmsClient.SetMarketData(rosettaUserMetaData, rosettaXmlRequest)
   
   ' dateEnd = Time()
   
   Exit Function
   
     
'     ' DEBUG
'     Debug.Print "SET TIMEOUT: " & CMS_SERVER_SET_TIMEOUT
'     Debug.Print "SET | " & dateStart & " : " & dateEnd
'     Debug.Print "Rosetta XML Response: "
'     Debug.Print GetCmsXmlResponse
'     MsgBox GetCmsXmlResponse

GetCmsXmlResponse_Err:
    Call RaiseError("Failed to connect to CMS " & CMS_GetServerMode() & " Server: " & Err.Description)

End Function



' ==========================================================================================
' modUtils
' ==========================================================================================

Private Function CMS_ImpersonateUserAs(ImpersonatedUser As String) As String
    If (Len(ImpersonatedUser) = 0) Then
        ThisWorkbook.IMPERSONATED_USER_NAME = CMS_GetUserLoginName()
    Else
    ThisWorkbook.IMPERSONATED_USER_NAME = ImpersonatedUser
    End If
    CMS_ImpersonateUserAs = ThisWorkbook.IMPERSONATED_USER_NAME
End Function

Public Function CMS_GetUserImpersonatedName() As String
    If (Len(ThisWorkbook.IMPERSONATED_USER_NAME) = 0) Then
        ThisWorkbook.IMPERSONATED_USER_NAME = CMS_GetUserLoginName()
    End If
    CMS_GetUserImpersonatedName = ThisWorkbook.IMPERSONATED_USER_NAME
End Function

Private Function CompileSetUserMetaData(Optional ImpersonatedUserName As String) As String
    ' Get UserName
    Dim userName As String
    If (ImpersonatedUserName = "") Then
        ImpersonatedUserName = CMS_GetUserLoginName()
    End If
    
    ' Encode the ActiveWorkbook.FullName - might break XML parsing on CMS side!
    ' VA 2009-01-28: Rolled back...
    ' VA 2009-04-14: ... back in business
    ' VA 2009-07-13: Added <application-version>
    Dim workbookFullNameEncoded As String
    workbookFullNameEncoded = HTMLEncode(ActiveWorkbook.FullName)
    
    CompileSetUserMetaData = "<user-metadata><user-details>" _
                        + "<user-id>" + ImpersonatedUserName + "</user-id>" _
                        + "<user-domain>" + CMS_GetUserDomainName() + "</user-domain>" _
                        + "<application-id>" + APP_ID + "</application-id>" _
                        + "<application-version>" + ADDIN_VERSION + "</application-version>" _
                        + "<application-batch-id>" + workbookFullNameEncoded + "</application-batch-id>" _
                        + "<id>" + CMS_GetUserLoginName() + "</id>" _
                        + "<host-id>" + CMS_GetUserMachineName() + "</host-id>" _
                        + "</user-details></user-metadata>"
'    Debug.Print CompileSetUserMetaData
End Function

Private Function CompileRosettaSetXmlRequest(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, ByVal Product As String, _
                                             ByRef termQuoteArray() As Variant, _
                                             ByVal QuoteConvention As String, _
                                             ByVal RecoveryRate As Double, _
                                             ByVal Location As String, _
                                             ByVal TenorStandard As Variant, _
                                             ByVal IsSnacCurve As Boolean, _
                                             ByVal ContractualSpread As Double) As String
    
    
    Dim fourTuple As String
    Dim quoteConventionTag As String
    Dim rosettaSetMarketSet As String
    Dim dataType As String
    
    ' 4Tuple & upper cases
    fourTuple = CMS_CompileFourTuple(Ticker, Ccy, debtClass, Product)
    QuoteConvention = UCase(QuoteConvention)
    
    ' Dummy timestamp
    Dim timeStamp As String
    timeStamp = Format(Now(), "yyyyMMdd")
     
    '  VA 2009-03-10: dataType instead of QuoteConvention?
    rosettaSetMarketSet = CompileRosettaSetMarketSet(Ticker, Ccy, debtClass, Product, _
                                                     termQuoteArray, _
                                                     QuoteConvention, _
                                                     RecoveryRate, _
                                                     Location, _
                                                     TenorStandard, _
                                                     IsSnacCurve, _
                                                     ContractualSpread)
    
    CompileRosettaSetXmlRequest = "<Rosetta version=""5.0.15"">" _
                                    & "<market>" _
                                        & "<marketData>" _
                                            & "<action>SET</action>" _
                                            & "<date>" & timeStamp & "</date>" _
                                            & rosettaSetMarketSet _
                                        & "</marketData>" _
                                    & "</market>" _
                                 & "</Rosetta>"
End Function

Private Function CompileRosettaSetMarketSet(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, ByVal Product As String, _
                                            ByRef termQuoteArray() As Variant, _
                                            ByVal dataType As String, _
                                            ByVal RecoveryRate As Double, _
                                            ByVal Location As String, _
                                            ByVal TenorStandard As Variant, _
                                            ByVal IsSnacCurve As Boolean, _
                                            ByVal ContractualSpread As Double) As String
                                            
                                            
    Dim termPointQuotesXml As String
    Dim recoveryPointXml As String
    Dim tenorIndex As Integer
    Dim hasContractualSpread As Boolean
    Dim currentDataType As String
    
    ' VA 2009-04-27: Now SNAC might and mostly will be 1 row...
    Dim termQuoteRowCount As Integer
    Dim arrayHelper As New CArrayHelper
    termQuoteRowCount = arrayHelper.NumberOfArrayDimensions(termQuoteArray)
    
    
    ' VA 2009-04-28: For now, both SNAC Quoted Spreads & Upfronts should have Contractual spreads
    hasContractualSpread = IsSnacCurve
    
    ' VA 2009-07-13: We now processing Ad Hoc Tenors in SETs!
    TERM_STRUCTURE = CMS_GetTermStructure(TenorStandard)
    TERM_STRUCTURE = ConvertTermsToTermStructure(TERM_STRUCTURE)
    
'''
'''        ' VA 2009-06-22: Adding 12 Tenors
'''        Select Case (TenorStandard)
'''        Case TENORS_13_STANDARD:
'''            TERM_STRUCTURE = Array("0:M", "3:M", "6:M", "9:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
'''        Case TENORS_12_STANDARD:
'''            TERM_STRUCTURE = Array("0:M", "3:M", "6:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
'''        Case TENORS_11_STANDARD:
'''default:
'''            TERM_STRUCTURE = Array("3:M", "6:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
'''        End Select
'''
    
    termPointQuotesXml = ""
    recoveryPointXml = ""
    
    If (UBound(termQuoteArray) <> 0) Then
        
        ' Process SNAC curve
        ' VA 2009-04-27: "Old-fashioned" SNACs only
        If (IsSnacCurve And termQuoteRowCount = 2) Then
            
            ' Set the credit Upfronts first with Contractual Spreads
            ' Spreads in the first line, Upfronts in the second
            For tenorIndex = 1 To UBound(termQuoteArray, 2)
                termPointQuotesXml = termPointQuotesXml & _
                                     CompileSetPointQuoteRequest(Ticker, Ccy, debtClass, Product, _
                                                              termQuoteArray(1, tenorIndex), _
                                                              CREDIT_SPREAD_DATA_TYPE, _
                                                              TERM_STRUCTURE(tenorIndex - 1), _
                                                              True, _
                                                              ContractualSpread)
            Next
        
            ' Now set the credit Spreads (Quoted Spreads) with Contractual Spreads
            For tenorIndex = 1 To UBound(termQuoteArray, 2)
                termPointQuotesXml = termPointQuotesXml & _
                                     CompileSetPointQuoteRequest(Ticker, Ccy, debtClass, Product, _
                                                              termQuoteArray(2, tenorIndex), _
                                                              CREDIT_UPFRONT_DATA_TYPE, _
                                                              TERM_STRUCTURE(tenorIndex - 1), _
                                                              True, _
                                                              ContractualSpread)
            Next
            
            If (RecoveryRate > 0) Then
                recoveryPointXml = CompileSetPointRecoveryRequest(Ticker, Ccy, debtClass, Product, _
                                                          RecoveryRate)
            End If
        ' PAR spread
        Else
            
            ' VA 2009-03-10: GHU - A quick fix for SPREADS & UPFRONT...
            ' VA 2009-04-27: ... and new one-row SNACs
            Dim curveTypeTag As String
            Select Case (dataType)
                Case CURVE_CONVENTION_SPREADS, CURVE_CONVENTION_QUOTEDSPREADS:
                    curveTypeTag = CREDIT_SPREAD_DATA_TYPE
                Case CURVE_CONVENTION_UPFRONT:
                    curveTypeTag = CREDIT_UPFRONT_DATA_TYPE
                Case Else
                    curveTypeTag = dataType
            End Select
            
            For tenorIndex = 1 To UBound(termQuoteArray)
                termPointQuotesXml = termPointQuotesXml & _
                                     CompileSetPointQuoteRequest(Ticker, Ccy, debtClass, Product, _
                                                              termQuoteArray(tenorIndex), _
                                                              curveTypeTag, _
                                                              TERM_STRUCTURE(tenorIndex - 1), _
                                                              hasContractualSpread, _
                                                              ContractualSpread)
            Next
        End If
    End If
    
    If (RecoveryRate > 0) Then
        recoveryPointXml = CompileSetPointRecoveryRequest(Ticker, Ccy, debtClass, Product, _
                                                          RecoveryRate)
    End If
    
    
    CompileRosettaSetMarketSet = "<marketSet>" _
                                    & "<location>" & Location & "</location>" _
                                    & "<logicalTime>" & DEFAULT_LOGICAL_TIME & "</logicalTime>" _
                                    & termPointQuotesXml _
                                    & recoveryPointXml _
                                & "</marketSet>"
                            
End Function

Private Function CompileSetPointQuoteRequest(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, ByVal Product As String, _
                                          ByVal TenorQuote As Double, _
                                          ByVal dataType As String, _
                                          ByVal colonTenor As String, _
                                          ByVal hasContractualSpread As Boolean, _
                                          ByVal ContractualSpread As Double) As String
                                          
    Dim fourTuple As String
    Dim tenorPeriod As String
    Dim periodMultiplier As String
    Dim colonTenorItems() As String
    
    
    fourTuple = CMS_CompileFourTuple(Ticker, Ccy, debtClass, Product)
    colonTenorItems = Split(colonTenor, ":")
    periodMultiplier = colonTenorItems(0)
    tenorPeriod = colonTenorItems(1)
    
'    ' Get correct dataType by QuoteConvention
'    Select Case UCase(QuoteConvention)
'        Case CURVE_CONVENTION_SPREADS:
'            DataType = CREDIT_SPREAD_DATA_TYPE
'        Case CURVE_CONVENTION_UPFRONT:
'            DataType = CREDIT_UPFRONT_DATA_TYPE
'        Case default:
'            RaiseError ("QuoteConvention must be '" & CURVE_CONVENTION_SPREADS & "' or '" & CURVE_CONVENTION_UPFRONT & "'.")
'    End Select
    
    CompileSetPointQuoteRequest = "<point>" _
                                    & "<label>" & fourTuple & "</label>" _
                                    & "<" & dataType & ">" _
                                    & "<issuerTicker>" & Ticker & "</issuerTicker>" _
                                    & "<debtClass>" & debtClass & "</debtClass>" _
                                    & "<currency>" & Ccy & "</currency>" _
                                    & "<creditCurveType>" & Product & "</creditCurveType>" _
                                    & "<periodMultiplier>" & periodMultiplier & "</periodMultiplier>" _
                                    & "<period>" & tenorPeriod & "</period>"
                                        
    If (hasContractualSpread) Then
        CompileSetPointQuoteRequest = CompileSetPointQuoteRequest _
                                    & "<contractualSpread>" & ContractualSpread & "</contractualSpread>"
    End If
                                        
    CompileSetPointQuoteRequest = CompileSetPointQuoteRequest _
                                    & "</" & dataType & ">" _
                                    & "<value>" & TenorQuote & "</value>" _
                                & "</point>"
End Function

                                          
Private Function CompileSetPointRecoveryRequest(ByVal Ticker As String, ByVal Ccy As String, ByVal debtClass As String, ByVal Product As String, _
                                                ByVal RecoveryRate As Double) As String
                                          
    Dim fourTuple As String
    
    fourTuple = CMS_CompileFourTuple(Ticker, Ccy, debtClass, Product)
    
    CompileSetPointRecoveryRequest = "<point>" _
                                    & "<label>" & fourTuple & "</label>" _
                                    & "<" & ISSUER_RECOVERY_DATA_TYPE & ">" _
                                        & "<issuerTicker>" & Ticker & "</issuerTicker>" _
                                        & "<debtClass>" & debtClass & "</debtClass>" _
                                        & "<currency>" & Ccy & "</currency>" _
                                        & "<creditCurveType>" & Product & "</creditCurveType>" _
                                    & "</" & ISSUER_RECOVERY_DATA_TYPE & ">" _
                                    & "<value>" & RecoveryRate & "</value>" _
                                & "</point>"
End Function


Private Function GetMarketData(ByVal CmsResponseXml As String, ByVal QuoteConvention As String, ByVal TenorStandard As Variant) As Variant

 Dim spreadMarks() As Variant
    
 'MSXML Dom object for processing the response
 Dim getMarketDataResponse As MSXML2.DOMDocument
 Dim success As Boolean
 Dim tenorColumnShift As Integer
 
 Dim tenorColumnCount As Integer
 
 ' DEBUG
' Debug.Print "CmsResponseXml"
' Debug.Print CmsResponseXml
 
     ' VA 2009-07-13: Now we handle Ad-Hoc SETs
     TERM_STRUCTURE = CMS_GetTermStructure(TenorStandard)
     TERM_STRUCTURE = ConvertTermsToTermStructure(TERM_STRUCTURE)
     
     ' CMS_SCLR_COL_COUNT
     tenorColumnCount = UBound(TERM_STRUCTURE) + 1
     
     ' VA 2009-07-13: !?!?!?!?!?!?!?!?!?!?
     tenorColumnShift = tenorColumnCount - CMS_TERM_COUNT
     
'''     ' VA 2009-06-22: Adding 12 Tenors
'''     ' Now we handle 2 Tenor Standards (11 & 13)
'''     ' So, get the corresponding Column Shift & Term Structure
'''     Select Case (TenorStandard)
'''    Case TENORS_13_STANDARD:
'''        tenorColumnShift = CMS_TENOR_13_COL_SHIFT
'''        TERM_STRUCTURE = Array("0:M", "3:M", "6:M", "9:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
'''    Case TENORS_12_STANDARD:
'''        tenorColumnShift = CMS_TENOR_12_COL_SHIFT
'''        TERM_STRUCTURE = Array("0:M", "3:M", "6:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
'''    Case TENORS_11_STANDARD:
'''default:
'''        tenorColumnShift = 0
'''        TERM_STRUCTURE = Array("3:M", "6:M", "1:Y", "2:Y", "3:Y", "4:Y", "5:Y", "7:Y", "10:Y", "20:Y", "30:Y")
'''    End Select
 
    ReDim spreadMarks(0 To CMS_COL_COUNT + tenorColumnShift - 1)


Set getMarketDataResponse = New MSXML2.DOMDocument
    
 getMarketDataResponse.async = False
 getMarketDataResponse.validateOnParse = False

 success = getMarketDataResponse.LoadXML(CmsResponseXml)
 
 Dim status As MSXML2.IXMLDOMNode
 
 ' VA 2008-12-12
 Dim LastMarkedOn As MSXML2.IXMLDOMNode
 Dim LastMarkedBy As MSXML2.IXMLDOMNode
 Dim Owner        As MSXML2.IXMLDOMNode
 Dim IsParent     As MSXML2.IXMLDOMNode
 
 Dim ErrorMsg As MSXML2.IXMLDOMNode

 Set status = getMarketDataResponse.selectSingleNode("//systemMessages/statusMessage/statusCode")
 
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
     ' VA 2009-02-25
     GetMarketData = spreadMarks
 Else
     spreadMarks(CMS_COL_STATUS + tenorColumnShift) = "OK"
     spreadMarks(CMS_COL_ERROR + tenorColumnShift) = ""
     
     
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
    For Each Term In TERM_STRUCTURE
       spreadMarks(TermInt) = GetTermQuote(Term, getMarketDataResponse, QuoteConvention)
       TermInt = TermInt + 1
    Next Term
    
    
    GetMarketData = spreadMarks
    
  End If
  
End Function

  
  
  Private Function GetTermQuote(ByVal Term As String, ByVal Response As MSXML2.DOMDocument, ByVal QuoteConvention As String) As Double
        'Debug.Print "Extracting the spreads for term "; Term
        Dim termSplit() As String
        termSplit = Strings.Split(Term, ":")
        
        ' VA 2008-12-8: Now we retrieve CreditUpfront as well
        Dim conventionTag As String
        
        If (QuoteConvention = CURVE_CONVENTION_SPREADS) Then
            conventionTag = "creditSpread"
        Else
            conventionTag = "creditUpfront"
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
            ' Don't need it in SET!
'            GetTermQuote = (point.LastChild.Text) * 100#
'            If (conventionTag = "creditSpread") Then
'                GetTermQuote = GetTermQuote * 100#
'            End If
            GetTermQuote = point.LastChild.Text
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






