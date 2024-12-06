Option Explicit
'------------------------------------------------------------------------------------------
' Class: clsMarket
' Handles market data associated with an index
'------------------------------------------------------------------------------------------
Public sabr_surface_ As clsSABRSurface
Public rate_curve_ As clsRateCurve
Public drift_adjust_ As clsRateCurve 'drift adjustment 추가: 2023.11.21
Public s_ As Double
Public prev_s_ As Double
Public div_schedule_ As clsDividendSchedule
Public div_yield_ As Double
Public heston_parameters_ As clsHestonParameter

Public index_name As String

Public ul_currency As String
'Public ul_currency_vol As Double
Private ul_currency_vol_() As Double

Public vol_index As String
Public vol_index_value As Double

Public vix_delta_alpha As Double
Public vix_delta_beta As Double

Public option_maturity_time As Date
Public market_open_time As Date
Public market_close_time As Date

Public option_maturity_date As Integer
Public market_open_date As Integer
Public market_close_date As Integer

Public refPriceForDividend As Double '2018.7.19
'Public iv_surface_ As clsImpliedVolSurface

Public lambda_neutral As Double 'local correlation 추가 2019.3.27

Private futures_quote_() As clsQuote



Public Function copy_obj(Optional rate_curve_shift As Long = 0) As clsMarket

    Dim rtn_obj As clsMarket
    Dim inx As Integer

    Set rtn_obj = New clsMarket

    If Not sabr_surface_ Is Nothing Then
        Set rtn_obj.sabr_surface_ = sabr_surface_.copy_obj()
    End If

    If Not rate_curve_ Is Nothing Then
        Set rtn_obj.rate_curve_ = rate_curve_.copy_obj(rate_curve_shift)
    End If
    
    'drift adjustment 추가: 2023.11.21
    If Not drift_adjust_ Is Nothing Then
        Set rtn_obj.drift_adjust_ = drift_adjust_.copy_obj(rate_curve_shift)
    End If

    rtn_obj.s_ = s_
    rtn_obj.prev_s_ = prev_s_

    If Not div_schedule_ Is Nothing Then
        Set rtn_obj.div_schedule_ = div_schedule_.copy_obj
    End If

    rtn_obj.div_yield_ = div_yield_

    rtn_obj.index_name = index_name
    rtn_obj.ul_currency = ul_currency
    
    For inx = 1 To NUM_DCF
        rtn_obj.set_ul_currency_vol inx, ul_currency_vol_(inx)
    Next inx

    If Not heston_parameters_ Is Nothing Then
        Set rtn_obj.heston_parameters_ = heston_parameters_.copy_obj
    End If

    'Set rtn_obj.iv_surface_ = iv_surface_.copy_obj

    rtn_obj.vol_index = vol_index
    rtn_obj.vol_index_value = vol_index_value
    rtn_obj.vix_delta_alpha = vix_delta_alpha
    rtn_obj.vix_delta_beta = vix_delta_beta

    rtn_obj.option_maturity_time = option_maturity_time
    rtn_obj.market_open_time = market_open_time
    rtn_obj.market_close_time = market_close_time

    rtn_obj.option_maturity_date = option_maturity_date
    rtn_obj.market_open_date = market_open_date
    rtn_obj.market_close_date = market_close_date

    rtn_obj.refPriceForDividend = refPriceForDividend '2018.7.19
    rtn_obj.lambda_neutral = lambda_neutral 'local correlation 추가 2019.3.27

On Error Resume Next

    Dim no_of_futures As Integer

    no_of_futures = UBound(futures_quote_)

    If Err.number = 8 Then
        no_of_futures = 0
    End If

    For inx = 1 To no_of_futures
        rtn_obj.set_futures_quote futures_quote_(inx)
    Next inx



    Set copy_obj = rtn_obj


End Function

Public Function ul_currency_vol(in_index As Integer) As Double
   
    ul_currency_vol = ul_currency_vol_(in_index)

End Function

Public Sub set_ul_currency_vol(in_index As Integer, in_vol As Double)
   
    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next

    Dim temp_inx As Integer
    temp_inx = UBound(ul_currency_vol_)
    
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = in_index
    Else
        initial_lbound = LBound(ul_currency_vol_)
        initial_ubound = max(UBound(ul_currency_vol_), in_index)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve ul_currency_vol_(initial_lbound To initial_ubound) As Double
    
    ul_currency_vol_(in_index) = in_vol
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_ul_currency_vol"
    
End Sub


Public Function get_futures_value(asset_code As String, Optional theo_price As Boolean = True) As Double

    Dim rtn_value As Double
    Dim futures_index As Integer
    
    futures_index = find_futures_index(asset_code, futures_quote_)
    
    If futures_index > 0 Then
    
        If theo_price Then
            rtn_value = futures_quote_(futures_index).theo_price
        Else
            rtn_value = futures_quote_(futures_index).last_price
        End If
    
    End If
    
    get_futures_value = rtn_value


End Function

Public Function get_futures_quotes() As clsQuote()

    
    
On Error GoTo ErrorHandler

    get_futures_quotes = futures_quote_

    Exit Function
    
ErrorHandler:

    raise_err "clsMarket.get_futures_quotes", Err.description

End Function



Public Sub set_futures_quotes(quotes() As clsQuote)
    
    
    Dim no_of_quotes As Integer
    Dim inx As Integer
    
On Error Resume Next

    no_of_quotes = UBound(quotes)
        
    If Err.number = 9 Then
        no_of_quotes = 0
    End If
    
    For inx = 1 To no_of_quotes
        
        set_futures_quote quotes(inx)
    
    Next inx
    
    
    

End Sub
Public Sub set_futures_quote(clsQuote As clsQuote)
    
    Dim futures_index As Integer
    
    futures_index = find_futures_index(clsQuote.asset_code, futures_quote_)
    
    If futures_index > 0 Then
        Set futures_quote_(futures_index) = clsQuote.copy_obj()
        
    Else
        push_back_clsquote futures_quote_, clsQuote.copy_obj()
    End If

End Sub


Private Function find_futures_index(ByVal asset_code As String, find_from() As clsQuote) As Integer

    Dim rtn_value As Integer
    Dim inx As Integer
    
On Error GoTo ErrorHandler

    rtn_value = -1
    
    For inx = 1 To get_array_size_clsQuote(find_from)
    
        If find_from(inx).asset_code = asset_code Then
            
            rtn_value = inx
            Exit For
        
        End If
    
    Next inx
    
    find_futures_index = rtn_value

    Exit Function
    
ErrorHandler:

    raise_err "clsMarket.find_futures_index", Err.description

End Function


Public Sub set_market(rate_curve As clsRateCurve, s As Double, div_schedule As clsDividendSchedule, heston_param As clsHestonParameter, Optional div_yield As Double = 0, Optional prev_s As Double)

    On Error GoTo ErrorHandler
    
    Set rate_curve_ = rate_curve
    
    s_ = s
    prev_s_ = prev_s
    
    Set div_schedule_ = div_schedule
    
    Set heston_parameters_ = heston_param
    
    div_yield_ = div_yield
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsMarket.set_market"

End Sub


Public Function get_drift_adjust(tdate As Date) As Double

    Dim ref_date As Date
    ref_date = CDate(drift_adjust_.rate_dates()(0))
    
    get_drift_adjust = -Log(drift_adjust_.get_discount_factor(ref_date, tdate)) / (tdate - ref_date) * 365

End Function