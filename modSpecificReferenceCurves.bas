
Attribute VB_Name = "modSpecificReferenceCurves"
'-----------------------------------------
' VA 2011-03-30: Specific Reference Curves
'-----------------------------------------
Option Explicit
' Option Private Module

'
Public Function CMS_Compile4TupleFrom5(ByVal Ticker As String, ByVal Ccy As String, ByVal TierOrRefId As String, ByVal DocClause As String, ByVal Product As String, _
                                       Optional ByVal Delimiter = DEFAULT_DELIMITER) As String
    
    Dim concatChar As String
    Dim debtClass As String
    
    If UCase(Product) = SPEC_REF_CURVE_PRODUCT Then
        concatChar = SPEC_REF_DEBT_CLASS_CONCAT_CHAR
    Else
        concatChar = REG_DEBT_CLASS_CONCAT_CHAR
    End If
    
    debtClass = UCase(TierOrRefId)
    If Len(DocClause) > 0 Then
        debtClass = debtClass + concatChar + UCase(DocClause)
    End If
    
    CMS_Compile4TupleFrom5 = CMS_CompileFourTuple(Ticker, Ccy, debtClass, Product, Delimiter)
End Function


Public Function CMS_CompileDebtClass(ByVal TierOrRefId As String, ByVal DocClause As String, _
                                     concatChar As String) As String
    Dim debtClass As String
    
    debtClass = TierOrRefId
    If (Len(DocClause) > 0) Then
        debtClass = debtClass + concatChar + DocClause
    End If
    CMS_CompileDebtClass = debtClass
End Function
    
Public Function CMS_CompileDebtClassByProduct(ByVal Tier As String, ByVal DocClause As String, _
                                              Product As String) As String
    Dim concatChar As String
    
    If (Product = SPEC_REF_CURVE_PRODUCT) Then
       ' Specific Reference Curve
       concatChar = SPEC_REF_DEBT_CLASS_CONCAT_CHAR
    Else
        ' "Regular" Curve
       concatChar = REG_DEBT_CLASS_CONCAT_CHAR
    End If
    CMS_CompileDebtClassByProduct = CMS_CompileDebtClass(Tier, DocClause, concatChar)
End Function


Public Function CMS_Get4TupleArrayFrom5(ByVal Ticker As String, ByVal Ccy As String, ByVal TierOrRefId As String, ByVal DocClause As String, ByVal Product As String) As String()
    Dim debtClass As String
    Dim fourTupleArray() As String
    
    debtClass = CMS_CompileDebtClassByProduct(TierOrRefId, DocClause, Product)
    ReDim fourTupleArray(0 To 3)
    fourTupleArray(0) = Ticker
    fourTupleArray(1) = Ccy
    fourTupleArray(2) = debtClass
    fourTupleArray(3) = Product
    
    CMS_Get4TupleArrayFrom5 = fourTupleArray

End Function

