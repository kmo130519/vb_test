'=================================
' Modified on
' 2013-10-14
' 2013-10-21
'=================================

Option Explicit

'Private index_id_() As String
Private market_() As clsMarket

Public correlation_pair_ As clsCorrelationPairs

'local correlation 추가 20190327
Public min_correlation_pair_ As clsCorrelationPairs

'Public pl_currency_rate_curve_ As clsRateCurve
Private pl_currency_rate_curve_() As clsRateCurve '2023.07.26 통화별 할인커브 관리

Private fx_rates_() As clsQuote

Public Function find_vol_index_ul_code(vol_index_in As String) As String

    Dim rtn_value As String
    Dim inx As Integer
    
On Error GoTo ErrorHandler

    For inx = 1 To UBound(market_)
        
        If market_(inx).vol_index = vol_index_in Then
        
            rtn_value = market_(inx).index_name
            Exit For
        End If
    
    Next inx
    
    find_vol_index_ul_code = rtn_value
    
    Exit Function
    
ErrorHandler:

    find_vol_index_ul_code = ""

End Function




Public Function get_spot(ul_code() As String) As Double()

    Dim rtn_array() As Double
    Dim no_of_array As Integer
    Dim inx As Integer
    
    no_of_array = get_array_size_string(ul_code)
    
    ReDim rtn_array(1 To no_of_array) As Double
    
    For inx = 1 To no_of_array
        
        rtn_array(inx) = market_by_ul(ul_code(inx)).s_
    
    Next inx
    
    get_spot = rtn_array

End Function

Public Function get_no_of_dcfs() As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next

    rtn_value = UBound(pl_currency_rate_curve_)
    
    If Err.number = 9 Then
        rtn_value = 0
    End If
    
On Error GoTo ErrorHandler
        
    
    get_no_of_dcfs = rtn_value
    
    Exit Function
    
ErrorHandler:

    raise_err "clsMarketSet.no_of_dcfs", Err.description

End Function

Public Function get_no_of_markets() As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next

    rtn_value = UBound(market_)
    
    If Err.number = 9 Then
        rtn_value = 0
    End If
    
On Error GoTo ErrorHandler
        
    
    get_no_of_markets = rtn_value
    
    Exit Function
    
ErrorHandler:

    raise_err "clsMarketSet.no_of_markets", Err.description

End Function

Public Sub replace_fx_rate(to_market_set As clsMarketSet)

    Dim no_of_markets As Integer
    Dim inx As Integer
    Dim market_index As Integer
    
On Error GoTo ErrorHandler
    
    set_fx_rates to_market_set.get_fx_rates()
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsMarketSet.replace_dividend", Err.description

End Sub
Public Sub replace_dividend(to_market_set As clsMarketSet)

    Dim no_of_markets As Integer
    Dim inx As Integer
    Dim market_index As Integer
    
On Error GoTo ErrorHandler
    
    no_of_markets = get_no_of_markets()

    For inx = 1 To no_of_markets
    
        market_index = to_market_set.find_index(market_(inx).index_name)
    
        If market_index > 0 Then
        
            If Not to_market_set.market(market_index).div_schedule_ Is Nothing Then
                Set market_(inx).div_schedule_ = to_market_set.market(market_index).div_schedule_.copy_obj()
            End If
            market_(inx).div_yield_ = to_market_set.market(market_index).div_yield_
        End If
    
    Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsMarketSet.replace_dividend", Err.description

End Sub
Public Sub replace_rate(to_market_set As clsMarketSet)

    Dim no_of_markets As Integer
    Dim inx As Integer
    Dim market_index As Integer
    
On Error GoTo ErrorHandler

    '2023-07-26
    no_of_dcfs = get_no_of_dcfs()
    
    For inx = 1 To no_of_dcfs
        Set pl_currency_rate_curve_(inx) = to_market_set.pl_currency_rate_curve_(inx).copy_obj()
    Next inx
    
    no_of_markets = get_no_of_markets()

    For inx = 1 To no_of_markets
    
        market_index = to_market_set.find_index(market_(inx).index_name)
    
        If market_index > 0 Then
            If Not to_market_set.market(market_index).rate_curve_ Is Nothing Then
                Set market_(inx).rate_curve_ = to_market_set.market(market_index).rate_curve_.copy_obj()
            End If
        End If
    
    Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsMarketSet.replace_rate", Err.description

End Sub
'
'Public Sub replace_vol_surface(to_market_set As clsMarketSet, inx As Integer) ', Optional local_vol_fwd_shift As Double = 0)
'
'    Dim market_index As Integer
'
'On Error GoTo ErrorHandler
'
'
'    market_index = to_market_set.find_index(market_(inx).index_name)
'
'    If market_index > 0 Then
'
'        If Not to_market_set.market(market_index).heston_parameters_ Is Nothing Then
'            Set market_(inx).heston_parameters_ = to_market_set.market(market_index).heston_parameters_.copy_obj()
'        End If
'
'        If Not to_market_set.market(market_index).sabr_surface_ Is Nothing Then
'            Set market_(inx).sabr_surface_ = to_market_set.market(market_index).sabr_surface_.copy_obj() '0, to_market_set.market(market_index).s_ - market_(inx).s_)
'        End If
'
'        Set market_(inx).iv_surface_ = to_market_set.market(market_index).iv_surface_.copy_obj()
'
'        market_(inx).ul_currency_vol = to_market_set.market(market_index).ul_currency_vol
'
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "clsMarketSet.replace_vol_surface", Err.description
'
'End Sub
'
'Public Sub replace_vol_surface_all(to_market_set As clsMarketSet) ', Optional local_vol_fwd_shift As Double = 0)
'
'    Dim no_of_markets As Integer
'    Dim inx As Integer
'    Dim market_index As Integer
'
'On Error GoTo ErrorHandler
'
'    no_of_markets = get_no_of_markets()
'
'    For inx = 1 To no_of_markets
'
'        replace_vol_surface to_market_set, inx
'
'    Next inx
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "clsMarketSet.replace_vol_surface", Err.description
'
'End Sub
'
'
'Public Sub recalibrate_sabr_surface_all(to_market_set As clsMarketSet)
'
'    Dim no_of_markets As Integer
'    Dim inx As Integer
'    Dim market_index As Integer
'
'On Error GoTo ErrorHandler
'
'    no_of_markets = get_no_of_markets()
'
'    For inx = 1 To no_of_markets
'
'        recalibrate_sabr_surface to_market_set, inx
'
'    Next inx
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "clsMarketSet.recalibrate_sabr_surface_all", Err.description
'
'End Sub
'
'
'Public Sub recalibrate_sabr_surface(to_market_set As clsMarketSet, inx As Integer)
'
'    Dim market_index As Integer
'
'On Error GoTo ErrorHandler
'
'    market_index = to_market_set.find_index(market_(inx).index_name)
'
'    If market_index > 0 Then
'
'        If Not to_market_set.market(market_index).sabr_surface_ Is Nothing And Not market_(inx).sabr_surface_ Is Nothing Then
'                'Set market_(inx).sabr_surface_ = market_(inx).sabr_surface_.copy_obj(0, to_market_set.market(market_index).s_ - market_(inx).s_)
'            market_(inx).sabr_surface_.shift_surface to_market_set.market(market_index).s_
'
'        End If
'
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "clsMarketSet.recalibrate_sabr_surface", Err.description
'
'End Sub

Public Sub bump_ul_price(bump_pct As Double, ul_code As String)

  
    Dim market_index As Integer
    Dim quotes() As clsQuote
    Dim inx As Integer
    
On Error GoTo ErrorHandler
    
    market_index = find_index(ul_code)
    
    If market_index > 0 Then
        market_(market_index).s_ = market_(market_index).s_ * (1 + bump_pct)
        'added(2018.8.8)
        market_(market_index).refPriceForDividend = market_(market_index).s_
        
        quotes = market(market_index).get_futures_quotes()
        
        For inx = 1 To get_array_size_clsQuote(quotes)
            
            quotes(inx).last_price = quotes(inx).last_price * (1 + bump_pct)
            quotes(inx).theo_price = quotes(inx).theo_price * (1 + bump_pct)
        
        Next inx
        
        market_(market_index).set_futures_quotes quotes

    End If
    
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsMarketSet.bump_ul_price", Err.description


End Sub
Public Sub replace_ul_price(to_market_set As clsMarketSet, inx As Integer)

    
    Dim market_index As Integer
    
On Error GoTo ErrorHandler
    
    market_index = to_market_set.find_index(market_(inx).index_name)
    
    If market_index > 0 Then
        market_(inx).s_ = to_market_set.market(market_index).s_
        market_(inx).set_futures_quotes to_market_set.market(market_index).get_futures_quotes()

    End If
    
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsMarketSet.replace_ul_price", Err.description

End Sub


Public Sub replace_ul_price_all(to_market_set As clsMarketSet)

    Dim no_of_markets As Integer
    Dim inx As Integer
    Dim market_index As Integer
    
On Error GoTo ErrorHandler
    
    no_of_markets = get_no_of_markets()

    For inx = 1 To no_of_markets
        
        replace_ul_price to_market_set, inx
    
    Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsMarketSet.replace_ul_price", Err.description

End Sub

Public Sub set_fx_rates(fx_rates() As clsQuote)

    Dim inx As Integer
    
    For inx = 1 To UBound(fx_rates)
    
        set_fx_rate fx_rates(inx).asset_code, fx_rates(inx).last_price, fx_rates(inx).prev_price
    
    Next inx

End Sub

Public Sub set_fx_rate(currency_pair As String, rate_value As Double, Optional prev_value As Double)

    Dim inx As Integer
    Dim tmp_obj As clsQuote
    
    inx = find_quote_index(currency_pair, fx_rates_)
    
    If inx < 0 Then
    
        Set tmp_obj = New clsQuote
        
        tmp_obj.asset_code = currency_pair
        tmp_obj.last_price = rate_value
        tmp_obj.prev_price = prev_value
        
        push_back_clsquote fx_rates_, tmp_obj
    Else
        
        fx_rates_(inx).last_price = rate_value
        fx_rates_(inx).prev_price = prev_value
    
    End If
        



End Sub

Public Function get_fx_rates() As clsQuote()

  
    
On Error GoTo ErrorHandler
    

    
    get_fx_rates = fx_rates_
    
    Exit Function
    
ErrorHandler:

    raise_err "clsmarketSet.get_fx_rates", Err.description

End Function
Public Function get_fx_rate(currency_pair As String) As Double

    Dim rtn_value As Double
    
    
On Error GoTo ErrorHandler
    
    If Not find_quote_value(rtn_value, currency_pair, fx_rates_) Then
        
        rtn_value = 1
    End If
    
    get_fx_rate = rtn_value
    
    Exit Function
    
ErrorHandler:

    raise_err "clsmarketSet.get_fx_rate", Err.description

End Function

Private Function find_quote_index(ByVal asset_code As String, find_from() As clsQuote) As Integer

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
    
    find_quote_index = rtn_value

    Exit Function
    
ErrorHandler:

    raise_err "clsMarketSet.find_quote_index", Err.description

End Function

Private Function find_quote_value(ByRef price As Double, ByVal asset_code As String, find_from() As clsQuote, Optional prev_price As Boolean = False) As Boolean

    Dim rtn_value As Boolean
    Dim inx As Integer
    
    
    rtn_value = False
    
    inx = find_quote_index(asset_code, find_from)
    
    If inx >= 0 Then
    
        rtn_value = True
        
        If prev_price Then
            price = find_from(inx).prev_price
        Else
            price = find_from(inx).last_price
        End If
        
    Else
    
        rtn_value = False
        
    End If
            
    
    find_quote_value = rtn_value
    

End Function
'Private Function get_array_size_clsQuote(the_array() As clsQuote) As Integer
'
'    Dim rtn_value As Integer
'
'On Error Resume Next
'
'    rtn_value = UBound(the_array)
'
'    If Err.number = 9 Then
'        rtn_value = 0
'    End If
'
'
'     get_array_size_clsQuote = rtn_value
'
'End Function

'Public Function index_id(Optional ByVal inx As Integer = 1) As String
'
'    index_id = index_id_(inx)
'
''End Function

Public Function copy_obj(Optional rate_curve_day_shift As Long) As clsMarketSet

    Dim rtn_obj As clsMarketSet
    Dim inx As Integer
    Dim no_of_dcfs As Integer
    Dim no_of_markets As Integer
    
On Error Resume Next

    no_of_dcfs = UBound(pl_currency_rate_curve_)
    no_of_markets = UBound(market_)

    If Err.number = 8 Then

        no_of_markets = 0

    End If

On Error GoTo ErrorHandler


    Set rtn_obj = New clsMarketSet

    Set rtn_obj.correlation_pair_ = correlation_pair_.copy_obj
    Set rtn_obj.min_correlation_pair_ = min_correlation_pair_.copy_obj '2019.3.27
    
    For inx = 1 To no_of_dcfs
        rtn_obj.set_pl_currency_rate_curve inx, pl_currency_rate_curve_(inx).copy_obj(rate_curve_day_shift)
    Next inx
    
    For inx = 1 To no_of_markets
        rtn_obj.set_market market_(inx).index_name, market_(inx).copy_obj(rate_curve_day_shift)
    Next inx

    'rtn_obj.set_fx_rates fx_rates_

    Set copy_obj = rtn_obj
    Set rtn_obj = Nothing

    Exit Function

ErrorHandler:

    raise_err "clsMarketSet.copy_obj", Err.descriptions


End Function


Public Function dcf_by_ccy(ByVal ccy_code As String) As clsRateCurve
    

On Error GoTo ErrorHandler
    

    Set dcf_by_ccy = pl_currency_rate_curve_(get_ccy_idx(ccy_code))
    
    Exit Function
    
ErrorHandler:

    raise_err "dcf_by_ccy", Err.description


End Function


Public Function market_by_ul(Optional ByVal ul_code As String = "KOSPI200") As clsMarket
    
    Dim market_index As Integer

On Error GoTo ErrorHandler

    market_index = find_index(ul_code)
    
    If market_index > 0 Then

        Set market_by_ul = market_(market_index)
        
    End If
    
    Exit Function
    
ErrorHandler:

    raise_err "market_by_ul", Err.description


End Function


Public Function market(Optional ByVal inx As Integer = 1) As clsMarket

    Set market = market_(inx)

End Function

Public Function s(ByVal inx As Integer) As Double
    
    s = market_(inx).s_
    
    Exit Function
    
ErrorHandler:

    raise_err "clsMarketSet.s", Err.description

End Function


Public Function s_array() As Double()

    Dim array_size As Integer
    Dim rtn_array() As Double
    Dim inx As Integer

On Error Resume Next

    array_size = UBound(market_)
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
    
        array_size = 0
        
    Else
    
        ReDim rtn_array(1 To array_size) As Double
        
        For inx = 1 To array_size
        
            rtn_array(inx) = market_(inx).s_
        
        Next inx


    End If
    
    s_array = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsMarketSet.s_array", Err.description


End Function

Public Sub set_market(in_index As String, in_market As clsMarket)

    Dim tmp_index As Integer
    
    tmp_index = find_index(in_index)
    
    If tmp_index > 0 Then
    
        'index_id_(tmp_index) = in_index
        Set market_(tmp_index) = in_market
        
    Else
    
        'push_back_string index_id_, in_index
        push_back_market market_, in_market
    
    End If

End Sub

Public Sub set_pl_currency_rate_curve(in_index As Integer, in_dcf As clsRateCurve)

    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next

    Dim temp_inx As Integer
    temp_inx = UBound(pl_currency_rate_curve_)
    
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = in_index
    Else
        initial_lbound = LBound(pl_currency_rate_curve_)
        initial_ubound = max(UBound(pl_currency_rate_curve_), in_index)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve pl_currency_rate_curve_(initial_lbound To initial_ubound) As clsRateCurve
    
    Set pl_currency_rate_curve_(in_index) = in_dcf
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_pl_currency_rate_curve"
    
End Sub

Public Function find_index(in_index As String) As Integer

    Dim inx As Integer
    Dim array_size As Integer
    Dim rtn_value As Integer
    
On Error Resume Next

    array_size = UBound(market_)
    
    If Err.number = 9 Then
    
        rtn_value = 0
        
    Else
    
On Error GoTo ErrorHandler
        
        For inx = 1 To array_size
        
            If market_(inx).index_name = in_index Then
            
                rtn_value = inx
                
                Exit For
                
            End If
        
        Next inx

    End If
    
    find_index = rtn_value
    
    Exit Function
    
ErrorHandler:

    raise_err "clsMarketSet.find_index", Err.description


End Function