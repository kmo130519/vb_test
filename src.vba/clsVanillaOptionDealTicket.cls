Option Explicit

Public prod_code As String
Public prod_type As String
Public ul_code As String
Public option_type As Long ' 0: European / 1: American
Public call_put As Long '0: call / 1: put
Public k As Double 'strike price
Public maturity_date As Date 'maturity date as date type
Public qty As Double ' number of contracts
Public conversion_ratio As Double 'multiplier
Public ul_currency As String
Public face_value As Double
Public option_on_futures As Boolean
Public futures_maturity As Date


Public Sub initialize(ByVal prod_code_in As String, _
                    ByVal prod_type_in As String, _
                    ByVal ul_code_in As String, _
                    ByVal maturity_date_in As Date, _
                    ByVal k_in As Double, _
                    ByVal qty_in As Double, _
                    ByVal conversion_ratio_in As Double, _
                    ByVal ul_currency_in As String, _
                    Optional ByVal option_type_in As Long = 0, _
                    Optional ByVal option_on_futures_in As Boolean = False, _
                    Optional ByVal futures_maturity_in As Date = -1)
    
    prod_code = prod_code_in
    prod_type = prod_type_in
    
    Select Case prod_type
    Case "C": call_put = 0
    Case "P": call_put = 1
    Case Else: call_put = -1
    End Select
    
    ul_code = ul_code_in
    maturity_date = maturity_date_in
    k = k_in
    qty = qty_in
    conversion_ratio = conversion_ratio_in
    ul_currency = ul_currency_in
    option_type = option_type_in
    option_on_futures = option_on_futures_in
    futures_maturity = futures_maturity_in
    
    'for NYSE stock options : 2024.07.25
    ul_code = Replace(ul_code, ".NS", "", , , vbTextCompare)
    
    'for S&P500 mini-options
    If Left(ul_code_in, 2) = "ES" And ul_currency_in = "USD" Then
        option_type = 1 'American
        option_on_futures = True
        ul_code = get_ua_code(ua.SPX)
    End If

End Sub