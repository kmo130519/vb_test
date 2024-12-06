'---------------------------------
' Modified on
' 2013-10-16
'---------------------------------
Option Explicit

Private date_initialized__ As Boolean
Private holidays__() As Date

Public Function read_rate_curve(the_range As Range, Optional isPrevDate As Boolean = False) As clsRateCurve

    Dim rtn_obj As clsRateCurve
        
    Dim ofs As Integer
    If isPrevDate = True Then
        ofs = 5
    Else
        ofs = 0
    End If
    
    Dim no_of_rate_dates As Integer
    no_of_rate_dates = the_range.Offset(0, ofs).Cells(6, 3).value
    
    If no_of_rate_dates = 0 Then
        Set rtn_obj = Nothing
    Else
        Set rtn_obj = New clsRateCurve
        rtn_obj.initialize Range(the_range.Offset(0, ofs).Cells(8, 2), the_range.Offset(0, ofs).Cells(7 + no_of_rate_dates, 2)), Range(the_range.Offset(0, ofs).Cells(8, 3), the_range.Offset(0, ofs).Cells(7 + no_of_rate_dates, 3))
    End If
    
    Set read_rate_curve = rtn_obj

End Function

Public Function read_corr_set(Optional isPrevDate As Boolean = False) As clsCorrelationPairs

    Dim rtn_obj As clsCorrelationPairs
    Dim inx As Integer
    Dim jnx As Integer
    
    Dim the_range As Range
    Set the_range = shtMarket.Range("correlation_table")
    
    Dim ofs As Integer
    If isPrevDate = True Then
        ofs = the_range.Rows.count + 3
    Else
        ofs = 0
    End If
    
    Set the_range = the_range.Offset(ofs, 0)
    
    Set rtn_obj = New clsCorrelationPairs
    
    For inx = 1 To the_range.Rows.count
    
        For jnx = 1 To inx 'the_range.Columns.count
        
            rtn_obj.set_corr the_range.Cells(inx, 0).value, the_range.Cells(0, jnx).value, the_range.Cells(inx, jnx).value
        
        Next jnx
        
    Next inx
    
    Set read_corr_set = rtn_obj

End Function

'local correlation 추가 2019.3.27
Public Function read_min_corr_set(Optional isPrevDate As Boolean = False) As clsCorrelationPairs

    Dim rtn_obj As clsCorrelationPairs
    Dim inx As Integer
    Dim jnx As Integer
    
    Dim the_range As Range
    Set the_range = shtMarket.Range("min_correlation_table")
    
    Dim ofs As Integer
    If isPrevDate = True Then
        ofs = the_range.Rows.count + 3
    Else
        ofs = 0
    End If
    
    Set the_range = the_range.Offset(ofs, 0)
    
    Set rtn_obj = New clsCorrelationPairs
    
    For inx = 1 To the_range.Rows.count
    
        For jnx = 1 To inx 'the_range.Columns.count
        
            rtn_obj.set_corr the_range.Cells(inx, 0).value, the_range.Cells(0, jnx).value, the_range.Cells(inx, jnx).value
        
        Next jnx
        
    Next inx
    
    Set read_min_corr_set = rtn_obj
    
    Set the_range = Nothing
    Set rtn_obj = Nothing

End Function

''local correlation 추가 2019.3.27
'Public Sub read_lambda_neutral(Optional isPrevDate As Boolean = False)
'
'    Dim rtn_obj As clsCorrelationPairs
'    Dim inx As Integer
'
'    Dim the_range As Range
'    Set the_range = shtMarket.Range("correlation_table")
'
'    Dim ofs As Integer
'    If isPrevDate = True Then
'        ofs = the_range.Rows.count + 2
'    Else
'        ofs = 0
'    End If
'
'    Set the_range = the_range.Offset(ofs, 0)
'
'    For inx = 2 To the_range.Rows.count
'
'        'rtn_obj.set_min_corr the_range.Cells(inx, 1).value, the_range.Cells(1, jnx).value, the_range.Cells(inx, jnx).value
'
'    Next inx
'
'End Sub


Public Sub read_fr_mkt_tester()
    
    Dim test_market As clsMarket
    
    Set test_market = read_index_market("SPX")


End Sub

Public Function read_kospi_index_market(ByVal index_name As String _
                                      , Optional read_local_vol_flag As Boolean = False) As clsMarket

    Dim rtn_market As clsMarket
    
    Dim spot As Double
    'Dim prev_s As Double
    
    Dim rate_curve As clsRateCurve
    Dim sabr_surface As clsSABRSurface
    
    Dim eval_date As Date
    'Dim prev_date As Date
    
    Dim vol_surface_grid As clsPillarGrid
    Dim vol_dates() As Date
    Dim no_of_vol_dates As Integer
    Dim no_of_div_dates As Integer
    
'    Dim rho() As Double
'    Dim nu() As Double
'    Dim rho_coeff() As Double
'    Dim nu_coeff() As Double
    
'    Dim vol_strikes() As Double
'    Dim inx As Integer
'    Dim dummy_array() As Double
    
    Dim dividend_schedule As clsDividendSchedule
    Dim dividend_yield As Double
    
'    Dim heston_param As clsHestonParameter
    
'    Dim fwds() As Double
'    Dim atm_vols() As Double
'    Dim alpha() As Double
    
On Error GoTo ErrorHandler
    
    Set rtn_market = New clsMarket
    rtn_market.index_name = index_name
    
    Dim dates() As Date
    Dim values() As Double
        
    spot = shtMarket.Range("S").Cells(1, 1).value
    
    Set rate_curve = New clsRateCurve
    rate_curve.initialize shtMarket.Range("rate_dates"), shtMarket.Range("discount")
    
    read_div dates, values
    Set dividend_schedule = New clsDividendSchedule
    dividend_schedule.initialize_div UBound(dates), dates, values
    
    
    
    
    '----------------
    ' Currency
    '----------------
    rtn_market.ul_currency = "KRW"
    
    '----------------
    ' Dividend Schedule
    '----------------
    Set rtn_market.div_schedule_ = dividend_schedule
    rtn_market.div_yield_ = shtMarket.Range("Div_Yield").value
    
    
    '----------------
    ' Date
    '----------------
    eval_date = shtMarket.Range("market_date").value

    '------------------------
    ' Spot
    '------------------------
    rtn_market.s_ = spot

    '----------------
    ' Rate Curve
    '----------------
    Set rtn_market.rate_curve_ = rate_curve
    
    '----------------
    ' SABR Vol Surface
    '----------------
    
    Set sabr_surface = New clsSABRSurface
    
    read_local_vol sabr_surface, index_name, spot, eval_date
                              
    Set rtn_market.sabr_surface_ = sabr_surface
    
    Set read_kospi_index_market = rtn_market
    
    Exit Function
    
ErrorHandler:

    raise_err "read_kospi_index_market", Err.description

End Function

'Public Sub read_local_vol(ByRef grid As clsPillarGrid, ByRef vol_data() As Double, ByRef local_vol_data() As Double)
'
'    Set grid = New clsPillarGrid
'
'    grid.initialize
'
'End Sub

Public Sub read_fx_rates(quotes() As clsQuote)
    
    Dim no_of_data As Integer
    
    Dim inx As Integer
    Dim a_quote As clsQuote
    
    no_of_data = shtMarketForeign.Range("fx_rate_start").Cells(0, 1).value
    
    ReDim quotes(1 To no_of_data) As clsQuote
    
    For inx = 1 To no_of_data
        
        Set a_quote = New clsQuote
        
        a_quote.asset_code = shtMarketForeign.Range("fx_rate_start").Cells(inx + 1, 1).value
        a_quote.last_price = shtMarketForeign.Range("fx_rate_start").Cells(inx + 1, 4).value
        a_quote.prev_price = shtMarketForeign.Range("fx_rate_start").Cells(inx + 1, 3).value
        
        Set quotes(inx) = a_quote
        
        
    Next inx
    


End Sub

'------------------------------------------------------------------------
'Sub read_real_time_futures
' Read S&P500 futures prices from excel worksheet.
'------------------------------------------------------------------------
Public Sub read_real_time_futures(quotes() As clsQuote)

    Dim no_of_data As Integer
    Dim inx As Integer
    Dim a_quote As clsQuote
    Dim forward_factor As Double
    
    Erase quotes
    
    no_of_data = shtMarketForeign.Range("spx_futures_closing").Cells(-1, 1).value
    
    ReDim quotes(1 To no_of_data) As clsQuote
    
    For inx = 1 To no_of_data
    
        Set a_quote = New clsQuote
        
        a_quote.asset_code = shtMarketForeign.Range("spx_futures_closing").Cells(inx + 1, 1).value
        a_quote.bl_code = shtMarketForeign.Range("spx_futures_closing").Cells(inx + 1, 2).value
        a_quote.maturity_date = shtMarketForeign.Range("spx_futures_closing").Cells(inx + 1, 0).value
        
        a_quote.last_price = shtMarketForeign.Range("spx_futures_closing").Cells(inx + 1, 4).value
        a_quote.prev_price = shtMarketForeign.Range("spx_futures_closing").Cells(inx + 1, 3).value
        
        a_quote.prev_theo_price = shtMarketForeign.Range("spx_futures_closing").Cells(inx + 1, 5).value
        a_quote.theo_price = shtMarketForeign.Range("spx_futures_closing").Cells(inx + 1, 6).value
        
        Set quotes(inx) = a_quote
    
    Next inx
    

End Sub

Public Sub read_fx()

    Dim no_of_data As Integer
    Dim inx As Integer
    Dim a_quote As clsQuote
    
    Erase fx__
    
    no_of_data = shtMarketForeign.Range("fx_rate_start").Cells(0, 1).value
    
    ReDim fx__(1 To no_of_data) As clsQuote
    
    For inx = 1 To no_of_data
    
        Set a_quote = New clsQuote
        
        a_quote.asset_code = shtMarketForeign.Range("fx_rate_start").Cells(inx + 1, 1).value
        a_quote.last_price = shtMarketForeign.Range("fx_rate_start").Cells(inx + 1, 4).value
        a_quote.prev_price = shtMarketForeign.Range("fx_rate_start").Cells(inx + 1, 3).value
        
        Set fx__(inx) = a_quote
    
    Next inx
    

End Sub

'Public Function read_index_market(ByVal index_name As String _
'                                        , ByVal base_ccy As String _
'                                        , Optional read_local_vol_flag As Boolean = False _
'                                        , Optional isPrevDate As Boolean = False) As clsMarket
Public Function read_index_market(ByVal index_name As String _
                                        , Optional read_local_vol_flag As Boolean = False _
                                        , Optional isPrevDate As Boolean = False) As clsMarket

    Dim rtn_market As clsMarket
    
    Dim spot As Double
    
    Dim rate_curve As clsRateCurve
    Dim drift_adjust As clsRateCurve 'drift adjustment 추가: 2023.11.21
    Dim sabr_surface As clsSABRSurface

    Dim eval_date As Date
    
    Dim vol_surface_grid As clsPillarGrid
    Dim vol_dates() As Date
    Dim no_of_vol_dates As Integer
    Dim no_of_rate_dates As Integer
    Dim no_of_adj_dates As Integer 'drift adjustment 추가: 2023.11.21
    Dim no_of_div_dates As Integer
    
'    Dim vol_stirkes() As Double
'    Dim inx As Integer
'    Dim dummy_array() As Double
    
    Dim dividend_schedule As clsDividendSchedule
'    Dim dividend_yield As Double
    
'    Dim vol_strikes() As Double
'    Dim atm_vol() As Double
    
On Error GoTo ErrorHandler

    Dim ofs_market_rng As Integer
    Dim ofs_vol_rng As Integer
    If isPrevDate Then
        ofs_market_rng = 5
        ofs_vol_rng = 105
    Else
        ofs_market_rng = 0
        ofs_vol_rng = 0
    End If
    
    Set rtn_market = New clsMarket
    rtn_market.index_name = index_name
    
    'Set sabr_surface = Nothing

    '------------------------
    ' Range
    '------------------------
    Dim the_range As Range
    Set the_range = shtMarket.Range(index_name).Offset(0, ofs_market_rng)
    
    no_of_div_dates = the_range.Cells(36, 3).value
    If no_of_div_dates = 0 Then
        Set dividend_schedule = Nothing
    Else
        Set dividend_schedule = New clsDividendSchedule
        dividend_schedule.initialize_div no_of_div_dates, range_to_array_date(shtMarket.Range(the_range.Cells(38, 2), the_range.Cells(37 + no_of_div_dates, 2)), 1), range_to_array(shtMarket.Range(the_range.Cells(38, 3), the_range.Cells(37 + no_of_div_dates, 3)), 1)
        dividend_schedule.ratioDividend = 0
    End If
        
    '----------------
    ' Currency
    '----------------
    rtn_market.ul_currency = the_range.Cells(5, 3).value
    
    '----------------
    ' Lambda Neutral 2019.3.27
    '----------------
    rtn_market.lambda_neutral = the_range.Cells(24, 3).value
    
    If shtMarket.Range("SCENARIO_ENABLE") = True Then
        Dim corr_shock As Double
        corr_shock = 0
    
        Select Case shtMarket.Range("SCENARIO_ID").Value2
        Case "C+0.1": corr_shock = 0.1
        Case "C+0.2": corr_shock = 0.2
        Case "C+0.3": corr_shock = 0.3
        Case "SC001": corr_shock = 0.1
        Case "SC002": corr_shock = 0.2
        Case "H001_1": corr_shock = 0.25 * 0.1
        Case "H001_2": corr_shock = 0.25 * 0.2
        Case "H001_3": corr_shock = 0.25 * 0.3
        Case "H001_4": corr_shock = 0.25 * 0.4
        Case "H001_5": corr_shock = 0.25 * 0.5
        Case "H001_6": corr_shock = 0.25 * 0.6
        Case "H001_7": corr_shock = 0.25 * 0.7
        Case "H001_8": corr_shock = 0.25 * 0.8
        Case "H001_9": corr_shock = 0.25 * 0.9
        Case "H001": corr_shock = 0.25
        Case "H002_1": corr_shock = 0.1 * 1 / 7
        Case "H002_2": corr_shock = 0.1 * 2 / 7
        Case "H002_3": corr_shock = 0.1 * 3 / 7
        Case "H002_4": corr_shock = 0.1 * 4 / 7
        Case "H002_5": corr_shock = 0.1 * 5 / 7
        Case "H002_6": corr_shock = 0.1 * 6 / 7
        Case "H002": corr_shock = 0.1
        Case "H003_1": corr_shock = 0.15 * 1 / 7
        Case "H003_2": corr_shock = 0.15 * 2 / 7
        Case "H003_3": corr_shock = 0.15 * 3 / 7
        Case "H003_4": corr_shock = 0.15 * 4 / 7
        Case "H003_5": corr_shock = 0.15 * 5 / 7
        Case "H003_6": corr_shock = 0.15 * 6 / 7
        Case "H003": corr_shock = 0.15
        Case "RVS001": corr_shock = 0.55 * 0.1
        Case "RVS002": corr_shock = 0.55 * 0.2
        Case "RVS003": corr_shock = 0.55 * 0.3
        Case "RVS004": corr_shock = 0.55 * 0.4
        Case "RVS005": corr_shock = 0.55 * 0.5
        Case "RVS006": corr_shock = 0.55 * 0.6
        Case "RVS007": corr_shock = 0.55 * 0.7
        Case "RVS008": corr_shock = 0.55 * 0.8
        Case "RVS009": corr_shock = 0.55 * 0.9
        Case "RVS010": corr_shock = 0.55
        Case Else: corr_shock = 0
        End Select
        rtn_market.lambda_neutral = min(rtn_market.lambda_neutral + corr_shock, 1)
    End If
    
    '----------------
    ' FX VOL
    '----------------
'    Select Case base_ccy
'    Case "KRW": rtn_market.ul_currency_vol = the_range.Cells(25, 3).value
'    Case "USD": rtn_market.ul_currency_vol = the_range.Cells(26, 3).value
'    End Select
    rtn_market.set_ul_currency_vol DCF.KRW, the_range.Cells(25, 3).value
    rtn_market.set_ul_currency_vol DCF.USD, the_range.Cells(26, 3).value
    
    '--------------------
    ' Dividend Schedule
    '--------------------
    Set rtn_market.div_schedule_ = dividend_schedule
    rtn_market.div_yield_ = the_range.Cells(32, 3).value
    
    '----------------
    ' Date
    '----------------
    eval_date = shtMarket.Range("market_date").value
    
    '------------------------
    ' Spot
    '------------------------
    rtn_market.s_ = the_range.Cells(4, 3).value
    rtn_market.refPriceForDividend = rtn_market.s_ '2018.7.19
    
    '----------------
    ' Rate Curve
    '----------------
    no_of_rate_dates = the_range.Cells(6, 3).value
    If no_of_rate_dates = 0 Then
        Set rate_curve = Nothing
    Else
        Set rate_curve = New clsRateCurve
        rate_curve.initialize shtMarket.Range(the_range.Cells(8, 2), the_range.Cells(7 + no_of_rate_dates, 2)), shtMarket.Range(the_range.Cells(8, 3), the_range.Cells(7 + no_of_rate_dates, 3))
    End If
    
    Set rtn_market.rate_curve_ = rate_curve
    
    '----------------
    ' Drift adjustment : 2023.11.21
    '----------------
    Dim term_date() As Date
    Dim adjust() As Double
    
    no_of_adj_dates = get_drift_adjustment(term_date, adjust, eval_date, index_name)
    If no_of_adj_dates = 0 Then
        Set drift_adjust = Nothing
    Else
        Set drift_adjust = New clsRateCurve
        drift_adjust.initialize_by_array term_date, adjust
    End If
    
    Set rtn_market.drift_adjust_ = drift_adjust
    
    '----------------
    ' SABR Vol Surface
    '----------------
    
    Set sabr_surface = New clsSABRSurface
    
    'for leverage indices
    If index_name = "KRD020021147" Then
        index_name = "KOSPI200"
    End If
    read_local_vol sabr_surface, index_name, spot, eval_date, ofs_vol_rng
                              
    Set rtn_market.sabr_surface_ = sabr_surface
    
    Set read_index_market = rtn_market
    
    Set dividend_schedule = Nothing
    Set vol_surface_grid = Nothing
    Set rate_curve = Nothing
    Set sabr_surface = Nothing
    Set rtn_market = Nothing
    Set the_range = Nothing
    
    Exit Function
    
ErrorHandler:

    raise_err "read_index_market", Err.description

End Function


Public Function read_index_market_yesterday(ByVal index_name As String _
                                        , Optional read_local_vol_flag As Boolean = False) As clsMarket
                                        
    Dim ofs As Integer
    ofs = 5

    Dim rtn_market As clsMarket
    
    Dim spot As Double
    
    Dim rate_curve As clsRateCurve
    Dim sabr_surface As clsSABRSurface

    Dim eval_date As Date
    
    Dim vol_surface_grid As clsPillarGrid
    Dim vol_dates() As Date
    Dim no_of_vol_dates As Integer
    Dim no_of_rate_dates As Integer
    Dim no_of_div_dates As Integer
    
'    Dim vol_stirkes() As Double
'    Dim inx As Integer
'    Dim dummy_array() As Double
    
    Dim dividend_schedule As clsDividendSchedule
'    Dim dividend_yield As Double
    
'    Dim vol_strikes() As Double
'    Dim atm_vol() As Double
    
On Error GoTo ErrorHandler

    Set rtn_market = New clsMarket
    rtn_market.index_name = index_name
    
    'Set sabr_surface = Nothing

    '------------------------
    ' Range
    '------------------------
    Dim the_range As Range
    Set the_range = shtMarket.Range(index_name)
    
    no_of_div_dates = the_range.Cells(28, 3 + ofs).value
    If no_of_div_dates = 0 Then
        Set dividend_schedule = Nothing
    Else
        Set dividend_schedule = New clsDividendSchedule
        dividend_schedule.initialize_div no_of_div_dates, range_to_array_date(shtMarket.Range(the_range.Cells(30, 2 + ofs), the_range.Cells(29 + no_of_div_dates, 2 + ofs)), 1), range_to_array(shtMarket.Range(the_range.Cells(30, 3 + ofs), the_range.Cells(29 + no_of_div_dates, 3 + ofs)), 1)
    End If
        
    '----------------
    ' Currency
    '----------------
    rtn_market.ul_currency = the_range.Cells(5, 3 + ofs).value
    
    '----------------
    ' FX VOL
    '----------------
    rtn_market.ul_currency_vol = the_range.Cells(25, 3 + ofs).value
    
    '--------------------
    ' Dividend Schedule
    '--------------------
    Set rtn_market.div_schedule_ = dividend_schedule
    rtn_market.div_yield_ = the_range.Cells(27, 3 + ofs).value
    
    '----------------
    ' Date
    '----------------
    'eval_date = shtMarket.Range("market_date").Offset(1, 0).value
    eval_date = shtMarket.Range("market_date").value
    
    '------------------------
    ' Spot
    '------------------------
    rtn_market.s_ = the_range.Cells(4, 3 + ofs).value
   
    '----------------
    ' Rate Curve
    '----------------
    no_of_rate_dates = the_range.Cells(6, 3 + ofs).value
    If no_of_rate_dates = 0 Then
        Set rate_curve = Nothing
    Else
        Set rate_curve = New clsRateCurve
        rate_curve.initialize shtMarket.Range(the_range.Cells(8, 2 + ofs), the_range.Cells(7 + no_of_rate_dates, 2 + ofs)), shtMarket.Range(the_range.Cells(8, 3 + ofs), the_range.Cells(7 + no_of_rate_dates, 3 + ofs))
    End If
    
    Set rtn_market.rate_curve_ = rate_curve
    
    '----------------
    ' SABR Vol Surface
    '----------------
    
    Set sabr_surface = New clsSABRSurface
    
    read_local_vol_yesterday sabr_surface, index_name, spot, eval_date
                              
    Set rtn_market.sabr_surface_ = sabr_surface
    
    Set read_index_market_yesterday = rtn_market
    
    Set dividend_schedule = Nothing
    Set vol_surface_grid = Nothing
    Set rate_curve = Nothing
    Set sabr_surface = Nothing
    Set rtn_market = Nothing
    Set the_range = Nothing
    
    Exit Function
    
ErrorHandler:

    raise_err "read_index_market_yesterday", Err.description

End Function

Public Function read_local_vol(ByRef sabr_surface As clsSABRSurface, ByVal index_name As String, ByVal spot As Double, ByVal eval_date As Date, Optional ByVal ofs As Integer = 0) ', Optional max_date As Date = -1)

    Dim vol_surface_range As Range
    Dim local_vol_surface_range As Range
    
    Dim vol_surface_grid As clsPillarGrid
    Dim local_vol_surface_grid As clsPillarGrid
    
    Dim no_of_strikes As Integer
    Dim no_of_vol_dates As Integer
    Dim no_of_local_vol_dates As Integer
    Dim effective_no_of_vol_dates As Integer
    
    Dim vol_dates() As Date
    Dim local_vol_dates() As Date
    Dim vol_strikes() As Double
    Dim local_vol_strikes() As Double
    Dim vol_data() As Double
    Dim local_vol_data() As Double
    
    Dim inx As Integer
    Dim jnx As Integer
    Dim knx As Integer
    
    Dim sabr_parameters_loc As clsSABRParamArray
    
On Error GoTo ErrorHandler
    
    sabr_surface.eval_date_ = eval_date
    
    '-----------------------------------
    ' Set Worksheet Range
    '-----------------------------------
    Set local_vol_surface_range = shtLocalVol.Range(index_name & "_Local_Vol_Surface").Offset(0, ofs)
    
    '-----------------------------------
    ' Set Grid
    ' The number of local vol dates are not set, yet
    '-----------------------------------
    Set local_vol_surface_grid = New clsPillarGrid
    
    no_of_strikes = local_vol_surface_range.Cells(0, 2).value

    If no_of_strikes > 0 Then
    
    '-----------------------------------
    ' Set local_vol_dates
    '-----------------------------------
    'ReDim vol_dates(1 To no_of_vol_dates)
    For inx = 1 To local_vol_surface_range.Cells(0, 1).value
    
        'day_shift 적용으로 이 부분 삭제:2018.12.5
        'If local_vol_surface_range.Cells(inx + 1, 1).value > eval_date Then
            push_back_date local_vol_dates, local_vol_surface_range.Cells(inx + 1, 1).value
            effective_no_of_vol_dates = effective_no_of_vol_dates + 1
        'End If

    Next inx
    
  '  max_date = 0
    
'    If max_date > local_vol_dates(effective_no_of_vol_dates) Then
'        push_back_date local_vol_dates, max_date
'    End If
'
    '-----------------------------------
    ' Set vol_strikes.
    '-----------------------------------
    ReDim vol_strikes(1 To no_of_strikes)
    For inx = 1 To no_of_strikes
        vol_strikes(inx) = local_vol_surface_range.Cells(1, inx + 1).value
    Next inx
    
    
    '-----------------------------------
    ' Set local_vol_strikes.
    '-----------------------------------
    
    inx = 1
    
    Do While local_vol_surface_range.Cells(1, inx + 1).value > 0
        push_back_double local_vol_strikes, local_vol_surface_range.Cells(1, inx + 1).value
        inx = inx + 1
    Loop
        

    
    '-----------------------------------
    ' Initialize grids
    '-----------------------------------

    local_vol_surface_grid.initialize spot, local_vol_dates, eval_date
    

    local_vol_surface_grid.set_strikes local_vol_strikes
    
  '  Set sabr_surface = New clsSABRSurface
  
    '-----------------------------------
    ' Read Data
    '-----------------------------------
    ReDim local_vol_data(1 To local_vol_surface_grid.no_of_dates, 1 To local_vol_surface_grid.no_of_strikes) As Double
    
   
    For inx = 1 To local_vol_surface_grid.no_of_dates ' Loop for grid dates
            
        For knx = 1 To local_vol_surface_grid.no_of_strikes ' Loop for strikes
            
            local_vol_data(inx, knx) = local_vol_surface_range.Cells(inx + 1, knx + 1).value
            
        Next knx
        
    Next inx
    
    sabr_surface.set_local_vol_surface local_vol_surface_grid, local_vol_data

    End If
    
    Set vol_surface_range = Nothing
    Set local_vol_surface_range = Nothing
    Set vol_surface_grid = Nothing
    Set local_vol_surface_grid = Nothing
    Set sabr_parameters_loc = Nothing
    
    Exit Function
    
ErrorHandler:

    raise_err "read_local_vol", Err.description
    

End Function


Public Function read_local_vol_yesterday(ByRef sabr_surface As clsSABRSurface, ByVal index_name As String, ByVal spot As Double, ByVal eval_date As Date)  ', Optional max_date As Date = -1)

    Dim ofs As Integer
    ofs = 53
    
    sabr_surface.eval_date_ = eval_date

    Dim vol_surface_range As Range
    Dim local_vol_surface_range As Range
    
    Dim vol_surface_grid As clsPillarGrid
    Dim local_vol_surface_grid As clsPillarGrid
    
    Dim no_of_strikes As Integer
    Dim no_of_vol_dates As Integer
    Dim no_of_local_vol_dates As Integer
    Dim effective_no_of_vol_dates As Integer
    
    Dim vol_dates() As Date
    Dim local_vol_dates() As Date
    Dim vol_strikes() As Double
    Dim local_vol_strikes() As Double
    Dim vol_data() As Double
    Dim local_vol_data() As Double
    
    Dim inx As Integer
    Dim jnx As Integer
    Dim knx As Integer
    
    Dim sabr_parameters_loc As clsSABRParamArray
    
On Error GoTo ErrorHandler
    
    '-----------------------------------
    ' Set Worksheet Range
    '-----------------------------------
    Set local_vol_surface_range = shtLocalVol.Range(index_name & "_Local_Vol_Surface").Offset(0, ofs)
    
    '-----------------------------------
    ' Set Grid
    ' The number of local vol dates are not set, yet
    '-----------------------------------
    Set local_vol_surface_grid = New clsPillarGrid
    
    no_of_strikes = local_vol_surface_range.Cells(0, 2).value


    
    '-----------------------------------
    ' Set local_vol_dates
    '-----------------------------------
    'ReDim vol_dates(1 To no_of_vol_dates)
    For inx = 1 To local_vol_surface_range.Cells(0, 1).value
      
        If local_vol_surface_range.Cells(inx + 1, 1).value > eval_date Then
            push_back_date local_vol_dates, local_vol_surface_range.Cells(inx + 1, 1).value
            effective_no_of_vol_dates = effective_no_of_vol_dates + 1
        End If

    Next inx
    
  '  max_date = 0
    
'    If max_date > local_vol_dates(effective_no_of_vol_dates) Then
'        push_back_date local_vol_dates, max_date
'    End If
'
    '-----------------------------------
    ' Set vol_strikes.
    '-----------------------------------
    ReDim vol_strikes(1 To no_of_strikes)
    For inx = 1 To no_of_strikes
        vol_strikes(inx) = local_vol_surface_range.Cells(1, inx + 1).value
    Next inx
    
    
    '-----------------------------------
    ' Set local_vol_strikes.
    '-----------------------------------
    
    inx = 1
    
    Do While local_vol_surface_range.Cells(1, inx + 1).value > 0
        push_back_double local_vol_strikes, local_vol_surface_range.Cells(1, inx + 1).value
        inx = inx + 1
    Loop
        

    
    '-----------------------------------
    ' Initialize grids
    '-----------------------------------

    local_vol_surface_grid.initialize spot, local_vol_dates, eval_date
    

    local_vol_surface_grid.set_strikes local_vol_strikes
    
  '  Set sabr_surface = New clsSABRSurface
  
    '-----------------------------------
    ' Read Data
    '-----------------------------------
    ReDim local_vol_data(1 To local_vol_surface_grid.no_of_dates, 1 To local_vol_surface_grid.no_of_strikes) As Double
    
   
    For inx = 1 To local_vol_surface_grid.no_of_dates ' Loop for grid dates
            
        For knx = 1 To local_vol_surface_grid.no_of_strikes ' Loop for strikes
            
            local_vol_data(inx, knx) = local_vol_surface_range.Cells(inx + 1, knx + 1).value
            
        Next knx
        
    Next inx
    
    sabr_surface.set_local_vol_surface local_vol_surface_grid, local_vol_data

    Set vol_surface_range = Nothing
    Set local_vol_surface_range = Nothing
    Set vol_surface_grid = Nothing
    Set local_vol_surface_grid = Nothing
    Set sabr_parameters_loc = Nothing
    
    Exit Function
    
ErrorHandler:

    raise_err "read_local_vol", Err.description
    

End Function

Private Function read_sabr_parameter(ul_code As String, position As Integer, no_of_dates As Integer) As Double()
    
    Dim rtn_value() As Double
    Dim inx As Integer
    Dim tmp_last_value As Double
    
    
On Error GoTo ErrorHandler
    
    For inx = 1 To no_of_dates
    
        push_back_double rtn_value, shtLocalVol.Range(ul_code & "_Local_Vol_Surface").Cells(inx + 1, position).value
    
    Next inx
    
    
'    If push_back_last Then
'
'        tmp_last_value = rtn_value(no_of_dates)
'        push_back_double rtn_value, tmp_last_value
'    End If
        
    read_sabr_parameter = rtn_value

    Exit Function
    
ErrorHandler:

    raise_err "read_sabr_parameter", Err.description

End Function


'--------------------------------------------------------------------
' Function: read_deal_ticket
' Desc: Read deal sheet to make deal ticket object
'--------------------------------------------------------------------
Public Function read_barrier_deal_ticket() As clsBarrierDealTicket

    Dim deal_ticket As clsBarrierDealTicket
    Dim no_of_schedule As Integer
    Dim inx As Integer
    Dim call_dates() As Date
    Dim strikes() As Double
    Dim coupons() As Double
    
On Error GoTo ErrorHandler

    Set deal_ticket = New clsBarrierDealTicket
    
    deal_ticket.asset_code = shtBarrierPricer.Range("asset_code").Cells(1, 1).value
    deal_ticket.fund_code_c = shtBarrierPricer.Range("fund_code_c").Cells(1, 1).value
    deal_ticket.fund_code_m = shtBarrierPricer.Range("fund_code_m").Cells(1, 1).value
    deal_ticket.ul_code = shtBarrierPricer.Range("ul_code").Cells(1, 1).value
    
    deal_ticket.current_date = shtBarrierPricer.Range("current_date").Cells(1, 1).value
    deal_ticket.value_date = shtBarrierPricer.Range("value_date").Cells(1, 1).value
    deal_ticket.maturity_date = shtBarrierPricer.Range("maturity_date").Cells(1, 1).value
    deal_ticket.settlement_date = shtBarrierPricer.Range("settlement_date").Cells(1, 1).value
    
    deal_ticket.alive_yn = shtBarrierPricer.Range("alive_yn").Cells(1, 1).value
    deal_ticket.confirmed_yn = shtBarrierPricer.Range("confirmed_yn").Cells(1, 1).value
    
    deal_ticket.issue_cost = shtBarrierPricer.Range("issue_cost").Cells(1, 1).value
    deal_ticket.reference = shtBarrierPricer.Range("reference").Cells(1, 1).value
    
    If UCase(shtBarrierPricer.Range("buy_sell")) = "BUY" Then
        deal_ticket.quantity = shtBarrierPricer.Range("no_of_contracts").Cells(1, 1).value
    Else
        deal_ticket.quantity = shtBarrierPricer.Range("no_of_contracts").Cells(1, 1).value * -1
    End If
    
    deal_ticket.call_put = shtBarrierPricer.Range("call_put").Cells(1, 2).value
    deal_ticket.strike = shtBarrierPricer.Range("strike").Cells(1, 1).value
    
    deal_ticket.barrier_type = shtBarrierPricer.Range("barrier_type").Cells(1, 2).value
    deal_ticket.barrier = shtBarrierPricer.Range("barrier").Cells(1, 1).value
    deal_ticket.rebate = shtBarrierPricer.Range("rebate").Cells(1, 1).value
    deal_ticket.barrier_monitoring = shtBarrierPricer.Range("Barrier_monitoring_freq.").Cells(1, 1).value
    deal_ticket.rebate_only = shtBarrierPricer.Range("Rebate_Only").Cells(1, 2).value
    deal_ticket.barrier_shift = shtBarrierPricer.Range("Barrier_Shift").Cells(1, 1).value
    
    deal_ticket.participation_rate = shtBarrierPricer.Range("PR").Cells(1, 1).value
    
    deal_ticket.x_grid = shtBarrierPricer.Range("x_grid").Cells(1, 1).value
    deal_ticket.v_grid = shtBarrierPricer.Range("v_grid").Cells(1, 1).value
    deal_ticket.t_grid = shtBarrierPricer.Range("t_grid").Cells(1, 1).value
    
    deal_ticket.scheme_type = shtBarrierPricer.Range("fdm_scheme").Cells(1, 2).value
    
    deal_ticket.instrument_type = shtBarrierPricer.Range("Instrument_type").Cells(1, 2).value
    
    no_of_schedule = shtBarrierPricer.Range("no_of_schedules").value
    
    If no_of_schedule >= 1 Then
    
        ReDim call_dates(1 To no_of_schedule) As Date
        ReDim strikes(1 To no_of_schedule) As Double
        ReDim coupons(1 To no_of_schedule) As Double
        
        For inx = 1 To no_of_schedule
        
            call_dates(inx) = shtBarrierPricer.Range("schedule_start").Cells(inx, 1).value
            strikes(inx) = shtBarrierPricer.Range("schedule_start").Cells(inx, 5).value
            coupons(inx) = shtBarrierPricer.Range("schedule_start").Cells(inx, 4).value
        
        Next inx
            
    
    End If
    
    deal_ticket.funding_spread = shtBarrierPricer.Range("funding_spread").value
    
    deal_ticket.set_schedule no_of_schedule, call_dates, strikes, coupons
    
    
    no_of_schedule = shtBarrierPricer.Range("no_of_floating_leg").value
    
    If no_of_schedule >= 1 Then
    
        ReDim call_dates(1 To no_of_schedule) As Date
        ReDim coupons(1 To no_of_schedule) As Double
        
        For inx = 1 To no_of_schedule
        
            call_dates(inx) = shtBarrierPricer.Range("floating_leg_start").Cells(inx, 1).value
            coupons(inx) = shtBarrierPricer.Range("floating_leg_start").Cells(inx, 2).value
        
        Next inx
            
    
    End If
    
    deal_ticket.set_floating_schedule no_of_schedule, call_dates, coupons
    
    Set read_barrier_deal_ticket = deal_ticket
    
    Exit Function
    
ErrorHandler:

    raise_err "read_deal_ticket"
    
End Function



Public Sub clr_tester()
    Dim rtn_obj As Object
    
   Set rtn_obj = ac_deal_ticket_to_clr(read_ac_deal_ticket())

End Sub

'Modify deal_ticket for KOSPI2LG : 2018.7.9
Public Function modify_ac_deal_ticket(deal_ticket As clsACDealTicket, market_set As clsMarketSet) As clsACDealTicket
    
    Dim i As Integer
    Dim levULCode As String
    Dim baseULCode As String
    'Dim ratioLeverage As Double
    'ratioLeverage - >codecodeLeverage : 2020.6.19
    Dim codeLeverage As Integer
    
    Dim levULRef As Double
    Dim levULSpot As Double
    Dim baseULRef As Double
    Dim baseULSpot As Double
    
    Dim isLeveragedUL As Boolean
    Dim isLeveragedDeal As Boolean
    isLeveragedUL = False
    isLeveragedDeal = False

    For i = 1 To deal_ticket.no_of_ul
    
        Select Case deal_ticket.ul_code(i)
        Case "KRD020021147"
            levULCode = "KRD020021147"
            baseULCode = get_ua_code(ua.KOSPI200)
            'ratioLeverage = 2
            codeLeverage = 1
            isLeveragedUL = True
        Case Else
            codeLeverage = 0
            isLeveragedUL = False
        End Select
        isLeveragedDeal = isLeveragedDeal Or isLeveragedUL
        
        If isLeveragedUL Then
            levULSpot = market_set.market_by_ul(levULCode).s_
            baseULSpot = market_set.market_by_ul(baseULCode).s_
            levULRef = deal_ticket.reference_price(i)
            baseULRef = levULRef * baseULSpot / levULSpot
            
            deal_ticket.set_reference_price baseULRef, i
            deal_ticket.set_ul_code baseULCode, i
            'deal_ticket.set_ratioLeverage ratioLeverage, i
        Else
            'deal_ticket.set_ratioLeverage 1, i
        End If
        
        deal_ticket.set_codeLeverage codeLeverage, i
                
    Next i
    
    deal_ticket.isLeveraged = isLeveragedDeal
  
    Set modify_ac_deal_ticket = deal_ticket
    
End Function


Public Function read_ac_deal_ticket() As clsACDealTicket

    Dim deal_ticket As clsACDealTicket
    Dim schedule_list() As clsAutocallSchedule
    Dim no_of_schedule As Integer
    Dim call_dates() As Date
    Dim strikes() As Double
    Dim coupons() As Double
    Dim strike_shifts() As Double
    'Dim early_exit_touched_flags() As Double
    Dim early_exit_touched_flags() As Long 'data type 변경: dll(2018.7.17)
    Dim early_exit_performance_types() As Long 'dll(2018.8.8)
    Dim early_exit_barrier_types() As Long 'dll(2018.8.8)
    Dim inx As Integer
    Dim jnx As Integer
    
On Error GoTo ErrorHandler

    Set deal_ticket = New clsACDealTicket
    
    deal_ticket.fund_code_m = shtACPricer.Range("fund_code_m").Cells(1, 1).value
    deal_ticket.fund_code_c = shtACPricer.Range("fund_code_c").Cells(1, 1).value
    deal_ticket.asset_code = shtACPricer.Range("asset_code").Cells(1, 1).value
    
    'no_of_ul
    deal_ticket.set_ul_dim shtACPricer.Range("No_of_Underlying").Cells(1, 1).value
    If shtACPricer.Range("no_of_early_exit_schedule").Cells(1, 1).value = 1 Then
        deal_ticket.redim_early_exit_barrier shtACPricer.Range("no_of_early_exit_schedule").Cells(1, 1).value
    End If
    
    For inx = 1 To deal_ticket.no_of_ul
        deal_ticket.set_ul_code shtACPricer.Range("ul_code").Cells(1, inx).value, inx
    Next inx
    'deal_ticket.set_ul_code shtACPricer.Range("ul_code").Cells(1, 2).value, 2
    
    deal_ticket.current_date = shtACPricer.Range("current_date").Cells(1, 1).value
    deal_ticket.current_date_origin_ = deal_ticket.current_date
    deal_ticket.value_date = shtACPricer.Range("value_date").Cells(1, 1).value
    
    deal_ticket.settlement_date = shtACPricer.Range("settlement_date").Cells(1, 1).value
    
    deal_ticket.alive_yn = shtACPricer.Range("alive_yn").Cells(1, 1).value
    deal_ticket.confirmed_yn = shtACPricer.Range("confirmed_yn").Cells(1, 1).value
    
    If UCase(shtACPricer.Range("buy_sell")) = "BUY" Then
        deal_ticket.notional = shtACPricer.Range("notional").Cells(1, 1).value
    Else
        deal_ticket.notional = shtACPricer.Range("notional").Cells(1, 1).value * -1
    End If
    deal_ticket.ccy = shtACPricer.Range("CCY").value
        
    deal_ticket.call_put = shtACPricer.Range("call_put").Cells(1, 2).value
    deal_ticket.dummy_coupon = shtACPricer.Range("dummy").Cells(1, 1).value
    deal_ticket.floor_value = shtACPricer.Range("floor_value").value
        
    deal_ticket.ki_barrier_flag = shtACPricer.Range("ki_flag").Cells(1, 2).value
    deal_ticket.ki_touched_flag = shtACPricer.Range("ki_touch_flag").Cells(1, 2).value
    
    For inx = 1 To deal_ticket.no_of_ul
        deal_ticket.set_ki_barrier shtACPricer.Range("ki_barrier").Cells(1, inx).value, inx
        'deal_ticket.set_ki_barrier shtACPricer.Range("ki_barrier").Cells(1, 1).value, inx
    Next inx


    deal_ticket.put_strike = shtACPricer.Range("put_strike").Cells(1, 1).value
    deal_ticket.call_participation = shtACPricer.Range("Call_PR").Cells(1, 1).value
    deal_ticket.call_strike = shtACPricer.Range("Call_Strike").Cells(1, 1).value
    deal_ticket.put_additional_coupon = shtACPricer.Range("Put_Add_CPN").Cells(1, 1).value
    deal_ticket.ki_monitoring_freq = shtACPricer.Range("KI_Monitoring_Freq").Cells(1, 1).value
    deal_ticket.ki_adj_pct = shtACPricer.Range("ki_adj_pct").Cells(1, 1).value
    deal_ticket.put_participation = shtACPricer.Range("put_participation").Cells(1, 1).value

    
    deal_ticket.ejectable_flag = shtACPricer.Range("ejectable_flag").Cells(1, 1).value
    
    For inx = 1 To deal_ticket.no_of_ul
        deal_ticket.set_reference_price shtACPricer.Range("reference").Cells(1, inx).value, inx
        deal_ticket.set_ejected_ul_flag shtACPricer.Range("ejected_ul_flag").Cells(2, inx).value, inx
    Next inx
'    deal_ticket.reference_price_2 = shtACPricer.Range("reference").Cells(1, 3).value  '<-- 2D

    'deal_ticket.strike_shift = shtACPricer.Range("strike_shift").Cells(1, 1).value
    deal_ticket.ki_barrier_shift = shtACPricer.Range("KI_Barrier_Shift").Cells(1, 1).value
    
    deal_ticket.rate_spread = shtACPricer.Range("Rate_Spread").Cells(1, 1).value
    deal_ticket.issue_cost = shtACPricer.Range("issue_cost").Cells(1, 1).value
    deal_ticket.hedge_cost = shtACPricer.Range("Hedge_Cost").Cells(1, 1).value
    
    
    deal_ticket.issue_price = shtACPricer.Range("price").Cells(2, 1).value
    
    
    no_of_schedule = shtACPricer.Range("no_of_schedules").Cells(1, 1).value
    deal_ticket.no_of_schedule = no_of_schedule
    
    If no_of_schedule >= 1 Then
        
        Erase schedule_list 'ReDim schedule_list(1 To no_of_schedule) As clsAutocallSchedule
        
        Dim a_schedule As clsAutocallSchedule
        
       
        For inx = 1 To no_of_schedule
        
            Set a_schedule = New clsAutocallSchedule
            
            a_schedule.call_date = shtACPricer.Range("schedule_start").Cells(inx, 1).value
            
            jnx = 1
            
            Do
                a_schedule.set_percent_strike shtACPricer.Range("schedule_start").Cells(inx, 4 + jnx).value, jnx
                a_schedule.set_coupon_on_call shtACPricer.Range("schedule_start").Cells(inx, 1 + jnx).value, jnx
                
                jnx = jnx + 1
            
            Loop While shtACPricer.Range("schedule_start").Cells(inx, 4 + jnx).value > 0
            
            '<--- added
            a_schedule.performance_type = shtACPricer.Range("performance_type_on_call").Cells(inx, 1).value
            a_schedule.ejectable_order = shtACPricer.Range("ejectable_order_on_call").Cells(inx, 1).value 'dll (2121.11.12)
            
            'for the ejectable structure : dll(2021.11.12)
            If deal_ticket.ejectable_flag = True Then
                'If a_schedule.call_date = deal_ticket.current_date _
                    And shtACPricer.check_autocall(a_schedule.percent_strike, a_schedule.performance_type) = False Then
                    a_schedule.ejected_event_flag = 1
                'End If
            Else
                a_schedule.ejected_event_flag = 0
            End If
    
            If deal_ticket.no_of_ul = 3 And inx < no_of_schedule And deal_ticket.settlement_date < #6/18/2017# Then
                a_schedule.strike_shift = 0
            Else
                'a_schedule.strike_shift = shtACPricer.Range("strike_shift").value
                a_schedule.strike_shift = shtACPricer.Range("strike_smoothing_width").Cells(inx, 1).value
            End If
            '--->
        
            push_back_clsAutocallSchedule schedule_list, a_schedule
        
        Next inx
        
        deal_ticket.set_schedule_array schedule_list
        
        deal_ticket.strike_at_maturity = schedule_list(no_of_schedule).percent_strike '* deal_ticket.reference_price
        deal_ticket.coupon_at_maturity = schedule_list(no_of_schedule).coupon_on_call
        deal_ticket.maturity_date = schedule_list(no_of_schedule).call_date
    
    Else
    
        raise_err "read_ac_deal_ticket", "no schedule found"
        
    End If
    
    
    'deal_ticket.set_schedule no_of_schedule, call_dates, strikes, coupons
    
    'deal_ticket.comment = shtACPricer.Range("txtComment").value
    
    
    deal_ticket.monthly_coupon_flag = shtACPricer.Range("Monthly_Coupon_Flag").Cells(1, 2).value
        
    '----------------------------------------------------------------------
    ' Read Monthly Coupon Schedule
    '----------------------------------------------------------------------
    no_of_schedule = shtACPricer.Range("no_of_coupon_schedule").Cells(1, 1).value
    
    If no_of_schedule >= 1 Then
    
        ReDim call_dates(1 To no_of_schedule) As Date
        ReDim strikes(1 To no_of_schedule) As Double
        ReDim coupons(1 To no_of_schedule) As Double
         
        For inx = 1 To no_of_schedule
        
            call_dates(inx) = shtACPricer.Range("cpn_schedule_start").Cells(inx, 1).value
            strikes(inx) = shtACPricer.Range("cpn_schedule_start").Cells(inx, 2).value
            coupons(inx) = shtACPricer.Range("cpn_schedule_start").Cells(inx, 3).value
        
        Next inx
                
    End If
        
    deal_ticket.set_coupon_schedule no_of_schedule, call_dates, strikes, coupons
    '----------------------------------------------------------------------

    deal_ticket.monthly_coupon_amount = shtACPricer.Range("Monthly_Cpn").Cells(1, 1).value
    
    

        
    '----------------------------------------------------------------------
    ' Read EE Schedule
    '----------------------------------------------------------------------
    deal_ticket.early_exit_flag = shtACPricer.Range("Early_Exit_Flag").Cells(1, 2).value
    
'    For inx = 1 To deal_ticket.no_of_ul
'        deal_ticket.set_early_exit_barrier shtACPricer.Range("EE_Barrier").Cells(1, inx).value, inx
'        'deal_ticket.set_ki_barrier shtACPricer.Range("ki_barrier").Cells(1, 1).value, inx
'    Next inx
    
    no_of_schedule = shtACPricer.Range("no_of_early_exit_schedule").Cells(1, 1).value
    
    If no_of_schedule >= 1 Then
    
        deal_ticket.redim_early_exit_barrier no_of_schedule
    
        ReDim call_dates(1 To no_of_schedule) As Date
        ReDim coupons(1 To no_of_schedule) As Double
        ReDim strike_shifts(1 To no_of_schedule) As Double
        'ReDim early_exit_touched_flags(1 To no_of_schedule) As Double
        ReDim early_exit_touched_flags(1 To no_of_schedule) As Long  'data type 변경: dll(2018.7.17)
        ReDim early_exit_performace_types(1 To no_of_schedule) As Long  'dll(2018.8.8)
        ReDim early_exit_barrier_types(1 To no_of_schedule) As Long  'dll(2018.8.8)
        
        For inx = 1 To no_of_schedule
        
            call_dates(inx) = shtACPricer.Range("early_exit_schedule_start").Cells(inx, 1).value
            coupons(inx) = shtACPricer.Range("early_exit_schedule_start").Cells(inx, 2).value
            strike_shifts(inx) = 0
            
            For jnx = 1 To deal_ticket.no_of_ul
                deal_ticket.set_early_exit_barrier shtACPricer.Range("early_exit_schedule_start").Cells(inx, 5).value, deal_ticket.no_of_ul * (inx - 1) + jnx
            Next jnx
            
            If shtACPricer.Range("early_exit_schedule_start").Cells(inx, 6).value = "Y" Then
                early_exit_touched_flags(inx) = 1#
            ElseIf shtACPricer.Range("early_exit_schedule_start").Cells(inx, 6).value = "N" Then
                early_exit_touched_flags(inx) = 0#
            End If
            
            early_exit_performace_types(inx) = shtACPricer.Range("early_exit_schedule_start").Cells(inx, 7).value  'dll(2018.8.8)
            early_exit_barrier_types(inx) = shtACPricer.Range("early_exit_schedule_start").Cells(inx, 8).value  'dll(2018.8.8)
            
        Next inx
                
    End If
        
    'deal_ticket.set_early_exit_schedule no_of_schedule, call_dates, coupons
    'deal_ticket.set_early_exit_schedule no_of_schedule, call_dates, coupons, strike_shifts, early_exit_touched_flags
    deal_ticket.set_early_exit_schedule no_of_schedule, call_dates, coupons, strike_shifts, early_exit_touched_flags, early_exit_performace_types, early_exit_barrier_types
    '----------------------------------------------------------------------

    
    
    no_of_schedule = shtACPricer.Range("no_of_floating_leg").value
    
    If no_of_schedule >= 1 Then
    
        ReDim call_dates(1 To no_of_schedule) As Date
        ReDim coupons(1 To no_of_schedule) As Double
        
        For inx = 1 To no_of_schedule
        
            call_dates(inx) = shtACPricer.Range("floating_leg_start").Cells(inx, 1).value
            coupons(inx) = shtACPricer.Range("floating_leg_start").Cells(inx, 2).value
        
        Next inx
            
    
    End If
    
    deal_ticket.set_floating_schedule no_of_schedule, call_dates, coupons
    
    
    deal_ticket.x_grid = shtACPricer.Range("x_grid").Cells(1, 1).value
    deal_ticket.v_grid = shtACPricer.Range("v_grid").Cells(1, 1).value
    deal_ticket.t_grid = shtACPricer.Range("t_grid").Cells(1, 1).value
    deal_ticket.days_per_step = shtACPricer.Range("days_per_step").Cells(1, 1).value
    deal_ticket.scheme_type = shtACPricer.Range("fdm_scheme").Cells(1, 2).value
    deal_ticket.mid_day_greek = (shtACPricer.Range("mid_day_greek").Cells(1, 1).value = "Y")
    deal_ticket.vol_scheme_type = shtACPricer.Range("vol_scheme").Cells(1, 2).value
    deal_ticket.no_of_trials = shtACPricer.Range("no_of_trials").Cells(1, 1).value
    
    deal_ticket.instrument_type = shtACPricer.Range("Instrument_type").Cells(1, 2).value
    'deal_ticket.performance_type = shtACPricer.Range("performance_type").Cells(1, 2).value
    deal_ticket.ki_performance_type = shtACPricer.Range("ki_performance_type").Cells(1, 2).value
    
    deal_ticket.ra_flag = shtACPricer.Range("Range_Accrual_Flag").Cells(1, 2).value
    deal_ticket.ra_cpn = shtACPricer.Range("ra_cpn").value
    deal_ticket.ra_tenor = shtACPricer.Range("ra_tenor").Cells(1, 2).value
    deal_ticket.ra_min_percent = shtACPricer.Range("ra_min").value
    deal_ticket.ra_max_percent = shtACPricer.Range("ra_max").value
    
    '-------------------------------
    '2015-10-05
    '-------------------------------
    Dim no_of_term As Integer
    Dim term_array() As Date
    
    no_of_term = shtACPricer.Range("no_of_term").value
    
    For inx = 1 To no_of_term
        
        push_back_date term_array, shtACPricer.Range("Term_Vega_Start").Cells(inx, 1).value
    
    Next inx
    
    deal_ticket.set_term_vega_tenor term_array
    
    '----------------------------------------------------------------------
    
    Set read_ac_deal_ticket = deal_ticket
    
    Set deal_ticket = Nothing
    For inx = 1 To no_of_schedule
        Set schedule_list(inx) = Nothing
    Next inx
    
    Exit Function

ErrorHandler:

    raise_err "read_ac_deal_ticket", Err.description


End Function




'Public Function read_ac_deal_ticket(deal_code As String) As clsACDealTicket
'
'    Dim deal_ticket As clsACDealTicket
'    Dim no_of_schedule As Integer
'    Dim call_dates() As Date
'    Dim strikes() As Double
'    Dim coupons() As Double
'    Dim strike_shifts() As Double
'    'Dim early_exit_touched_flags() As Double
'    Dim early_exit_touched_flags() As Long 'data type 변경: dll(2018.7.17)
'    Dim early_exit_performance_types() As Long 'dll(2018.8.8)
'    Dim early_exit_barrier_types() As Long 'dll(2018.8.8)
'    Dim inx As Integer
'    Dim jnx As Integer
'
'    Dim schedule_list() As clsAutocallSchedule
'
'On Error GoTo ErrorHandler
'
'    Set deal_ticket = New clsACDealTicket
'
'    Dim SQL As String
'
'    Dim oCmd As New ADODB.Command
'    Dim oRS As New ADODB.Recordset
'
'    Dim today As Date
'    today = shtACPricer.Range("current_date").value
'
'    Dim tDayStr As String
'    tDayStr = date2str(today)
'
'    Dim oDB As New ADODB.Connection
'    oDB.Open connStr
'
'    deal_ticket.current_date = today
'    deal_ticket.current_date_origin_ = deal_ticket.current_date
'    deal_ticket.asset_code = deal_code
'
'    deal_ticket.alive_yn = "Y"
'    deal_ticket.confirmed_yn = "Y"
'    deal_ticket.call_put = "CALL"
'    deal_ticket.put_participation = 1
'    deal_ticket.put_strike = 1
'    deal_ticket.call_participation = 0
'    deal_ticket.call_strike = 1
'    deal_ticket.put_additional_coupon = 0
'    deal_ticket.ki_adj_pct = 1
'
'    deal_ticket.x_grid = 200
'    deal_ticket.v_grid = 100
'    deal_ticket.t_grid = 254
'    deal_ticket.mid_day_greek = False
'    deal_ticket.scheme_type = 1
'    deal_ticket.rate_spread = 0
'    deal_ticket.hedge_cost = 0
'    deal_ticket.issue_price = 0
'
'    deal_ticket.vol_scheme_type = 1
'    deal_ticket.no_of_trials = 16383
'    deal_ticket.instrument_type = 0
'    deal_ticket.ra_flag = 0
'    deal_ticket.ra_cpn = 0
'    deal_ticket.ra_tenor = 0
'    deal_ticket.ra_min_percent = 0
'    deal_ticket.ra_max_percent = 9999
'    deal_ticket.floor_value = 0
'
'    '---------- From Front DB
'    SQL = " select * from sps.ac_deal where asset_code = '" + productCode + "' "
'
'    With oCmd
'        .ActiveConnection = oDB
'        .CommandType = adCmdText
'        .CommandText = SQL
'
'        oRS.Open .Execute
'    End With
'
'    i = 1
'    Dim value_date As Date
'    Dim maturity_date As Date
'    Dim settlement_date As Date
'    Dim dummy_coupon As Double
'    Dim ki_flag As Long
'    Dim ki_touch_flag As Long
'    Dim ki_monitoring_freq As Integer
'    Dim notional As Double
'    Dim issue_cost As Double
'    Dim strike_smoothing As Double
'    Dim ki_barrier_shift As Double
'
'    Do Until oRS.EOF
'
'        'value date
'        value_date = str2date(oRS("VALUE_DATE"))
'        deal_ticket.value_date = value_date
'        'maturity date
'        maturity_date = str2date(oRS("EXPIRY_DATE"))
'        shtACPricer.Range("maturity_date").Cells(1, i) = maturity_date
'        'issue date
'        settlement_date = str2date(oRS("SETTLEMENT_DATE"))
'        deal_ticket.settlement_date = settlement_date
'
'        'dummy coupon
'        dummy_coupon = oRS("DUMMY_COUPON")
'        deal_ticket.dummy_coupon = dummy_coupon
'
'        'KI barrier flag
'        If (oRS("KI_BARRIER_YN") = "Y") Then
'            ki_flag = 1
'        Else
'            ki_flag = 0
'        End If
'        deal_ticket.ki_barrier_flag = ki_flag
'
'        'KI touched flag
'        If (oRS("KI_TOUCHED_YN") = "Y") Then
'            ki_touch_flag = 1
'        Else
'            ki_touch_flag = 0
'        End If
'        deal_ticket.ki_touched_flag = ki_touch_flag
'
'        'KI monitoring freq.
'        ki_monitoring_freq = oRS("KI_MONITORING_FREQ")
'        deal_ticket.ki_monitoring_freq = ki_monitoring_freq
'
'        'notional
'        If oRS("NOTIONAL") = 0 Then
'            notional = 1
'        Else
'            notional = oRS("NOTIONAL")
'        End If
'        deal_ticket.notional = -1 * notional 'SELL 가정
'
'        'issue cost
'        issue_cost = oRS("ISSUE_COST")
'        deal_ticket.issue_cost = issue_cost
'
'        'KI barrier shift
'        ki_barrier_shift = oRS("KIBARRIER_SHIFT_SIZE")
'        deal_ticket.ki_barrier_shift = ki_barrier_shift
'
'        i = i + 1
'        oRS.MoveNext
'    Loop
'
'    oRS.Close
'
'    '---------- From BizOne
'    '기초자산 코드, 최초기준가, KI수준
'    SQL = " SELECT indv_iscd, decode(unas_iscd,'NIKKEI225','NKY',unas_iscd), unas_intl_prc, clrd_sdrt/100, barr_val/100 from BSYS.TBSIMO202D00@GDW where indv_iscd = '" + productCode + "' order by 2"
'
'    With oCmd
'        .ActiveConnection = oDB
'        .CommandType = adCmdText
'        .CommandText = SQL
'
'        oRS.Open .Execute
'    End With
'
'    i = 1
'    Dim redeem_shift As Boolean
'    redeem_shift = False
'
'    Dim ua_code As String
'    Dim ua_ref_spot() As Double
'    Dim ua_ki_barr() As Double
'    Dim ua_close_spot() As Double
'    ReDim ua_ref_spot(0) As Double
'    ReDim ua_ki_barr(0) As Double
'    ReDim ua_close_spot(0) As Double
'
'    Do Until oRS.EOF
'
'        '단축코드 -> KR코드 변환 2019.3.5
'        If isQuote = True Then
'            Select Case oRS(1)
'            Case "005380": ua_code = "KR7005380001" '현대차
'            Case "005930": ua_code = "KR7005930003" '삼성전자
'            Case "000030": ua_code = "KR7000030007" '우리은행
'            Case "028260": ua_code = "KR7028260008" '삼성물산
'            Case "105560": ua_code = "KR7105560007" 'KB금융
'            Case "035420": ua_code = "KR7035420009" 'NAVER
'            Case "018260": ua_code = "KR7018260000" '삼성SDS
'            Case "005490": ua_code = "KR7005490008" 'POSCO
'            Case "034220": ua_code = "KR7034220004" 'LG디스플레이
'            Case "D02002": ua_code = "KRD020021147" 'KOSPI200 레버리지
'            Case Else: ua_code = oRS(1)
'            End Select
'        Else
'            ua_code = oRS(1)
'        End If
'
'        'underlying code
'        deal_ticket.set_ul_code ua_code(i), i
'
'        ReDim Preserve ua_ref_spot(UBound(ua_ref_spot) + 1) As Double
'        ReDim Preserve ua_ki_barr(UBound(ua_ki_barr) + 1) As Double
'        ReDim Preserve ua_close_spot(UBound(ua_close_spot) + 1) As Double
'
'        ua_close_spot(i) = WorksheetFunction.VLookup(ua_code, Range("ua_close_spot"), 2, False)
'
'        'ref. spot: 신규발행종목 최초기준가 100으로 설정
'        '스트레스테스트 일 경우는 제외 2018.12.20
'        If (scenario_test = False) And (today = value_date) Then
'            ua_ref_spot(i) = ua_close_spot(i)
'        Else
'            ua_ref_spot(i) = oRS(2)
'            '최초기준가 검증(발행일에만)
'            If today = settlement_date And check_spot(oRS(1), date2str(value_date), oRS(2), oDB) = False Then
'                MsgBox deal_code & ": " & oRS(1) & " 기초자산최초기준가 오류"
'            End If
'        End If
'
'        deal_ticket.set_reference_price ua_ref_spot(i), i
'
'        'KI barrier
'        ua_ki_barr(i) = oRS(4)
'        deal_ticket.set_ki_barrier ua_ki_barr(i), i
'
'        If oRS(1) = "SX5E" Then redeem_shift = True
'        If oRS(1) = "SPX" Then redeem_shift = True
'
'        i = i + 1
'        oRS.MoveNext
'    Loop
'
'    'no of underlying
'    deal_ticket.no_of_ul = i - 1
'
'    If redeem_shift Then
'        maturity_date = maturity_date + 1
'    End If
'
'    oRS.Close
'
'    Dim early_exit_flag As Long
'    '----------------------------------------
'    'performance type, fund_code, ccy
'    SQL = " SELECT UNAS_CHOC_MTHD_CODE, CLRD_TYPE_CODE, substr(PROD_FNCD,1,2), PROD_FNCD, STLM_CRCD " _
'        & " FROM    bsys.TBSIMO201M00@gdw " _
'        & " WHERE   INDV_ISCD = '" + productCode + "' "
'
'    With oCmd
'        .ActiveConnection = oDB
'        .CommandType = adCmdText
'        .CommandText = SQL
'
'        oRS.Open .Execute
'    End With
'
'    Dim isFiveWins As Boolean
'    isFiveWins = False
'
'    Do Until oRS.EOF
'
'        Select Case oRS(0)
'        Case "1"
'            deal_ticket.ki_performance_type = -1
'        Case "2"
'            deal_ticket.ki_performance_type = 1
'        Case "4"
'            deal_ticket.ki_performance_type = 0
'        Case "5"
'            deal_ticket.ki_performance_type = -1
'        End Select
'
'        If oRS(1) = "29" Then
'            early_exit_flag = 1
'        Else
'            early_exit_flag = 0
'        End If
'        deal_ticket.early_exit_flag = early_exit_flag
'
'        If oRS(1) = "38" Then
'            isFiveWins = True
'        Else
'            isFiveWins = False
'        End If
'
'        deal_ticket.fund_code_m = oRS(2)
'        deal_ticket.fund_code_c = oRS(3)
'        deal_ticket.ccy = oRS(4)
'
'        oRS.MoveNext
'    Loop
'
'
'    '----------------------------------------
'    'Call schedule
'    Dim no_of_ac_schedule As Integer
'    Dim call_date() As Date
'    Dim coupon_on_call() As Double
'    Dim barr_strike() As Double
'    Dim barr_perform_type() As Integer
'    ReDim call_date(0) As Date
'    ReDim coupon_on_call(0) As Double
'    ReDim barr_strike(0) As Double
'    ReDim barr_perform_type(0) As Integer
'
'    If isFiveWins Then
'        SQL = " SELECT CLRD_DTRM_DATE, null, "
'        SQL = SQL & "                 NVL(CLRD_INRT,0)/100 CLRD_ERT,"
'        SQL = SQL & "                 NVL(CLRD_BARR_VAL,0)/100 UNAS_SDRT,"
'        SQL = SQL & "                'N' AVRG_APLY_YN"
'        SQL = SQL & "         FROM   BSYS.TBSIMO227L00@GDW"
'        SQL = SQL & "         WHERE  INDV_ISCD = '" + productCode + "'"
'        SQL = SQL & " ORDER BY CLRD_DTRM_DATE"
'    Else
'        SQL = " SELECT  TRTH_CLRD_DTRM_DATE   DT  " _
'            & "        ,RMS.GET_WORKDATE(TRTH_CLRD_DTRM_DATE,1)   BFRPY_BASE_DT " _
'            & "    ,NVL(CLRD_ERT,0) / 100       BFRPY_BASE_RT " _
'            & "    ,NVL(UNAS_SDRT1,0)/100     BASERT1 " _
'            & "    ,NVL(AVRG_APLY_YN,'N')     AVRG_APLY_YN " _
'            & "FROM    BSYS.TBSIMO203D00@gdw " _
'            & "WHERE   INDV_ISCD = '" + productCode + "' " _
'            & "ORDER BY CLRD_DTRM_DATE "
'    End If
'
'    With oCmd
'        .ActiveConnection = oDB
'        .CommandType = adCmdText
'        .CommandText = SQL
'
'        oRS.Open .Execute
'    End With
'
'    i = 1
'    Do Until oRS.EOF
'
'        ReDim Preserve call_date(UBound(call_date) + 1) As Date
'        ReDim Preserve coupon_on_call(UBound(coupon_on_call) + 1) As Double
'        ReDim Preserve barr_strike(UBound(barr_strike) + 1) As Double
'        ReDim Preserve barr_perform_type(UBound(barr_perform_type) + 1) As Integer
'
'        If redeem_shift Then
'            '2019.4.18
'            '오늘이 실제 조기상환평가일 후 첫 영업일이면 call date를 오늘로 설정: 상환여부 판단 때문
'            If today = WorksheetFunction.WorkDay(str2date(oRS(0)), 1, shtHolidays.Range("Holidays")) Then
'                call_date(i) = today
'            Else
'                '기초자산 지역에 따라 상환스케줄 1차 조정: 실제 평가일 + 1달력일
'                call_date(i) = str2date(oRS(0)) + 1
'            End If
'        Else
'            call_date(i) = str2date(oRS(0))
'        End If
'        shtACPricer.Range("schedule_start").Cells(i, 1) = call_date(i)
'        coupon_on_call(i) = oRS(2)
'        barr_strike(i) = oRS(3)
'        shtACPricer.Range("cpn_on_call").Cells(i, 1) = coupon_on_call(i)
'        shtACPricer.Range("strike_rate").Cells(i, 1) = barr_strike(i)
'
'        'performance_type_on_call
'        If oRS(4) = "Y" Then
'            barr_perform_type(i) = 0
'        Else
'            barr_perform_type(i) = -1
'        End If
'        shtACPricer.Range("performance_type_on_call").Cells(i, 1) = barr_perform_type(i)
'
'        ''strike smoothing
'        'shtACPricer.Range("strike_smoothing_width").Cells(i, 1) = oRS(5)
'
'        i = i + 1
'        oRS.MoveNext
'    Loop
'
'    oRS.Close
'
'    '----------------------------------------
'    'Strike smoothing factor: sps DB의 call date를 shift할 경우 비즈원 call date와 join되지 않는 문제가 있어 부득이 독립적으로 입수함
'    SQL = "SELECT call_date, strike_smoothing_width from sps.ac_schedule where ASSET_CODE = '" + productCode + "' ORDER BY 1 "
'
'    With oCmd
'        .ActiveConnection = oDB
'        .CommandType = adCmdText
'        .CommandText = SQL
'
'        oRS.Open .Execute
'    End With
'
'    i = 1
'    Do Until oRS.EOF
'
'        'strike smoothing
'        shtACPricer.Range("strike_smoothing_width").Cells(i, 1) = oRS(1)
'
'        i = i + 1
'        oRS.MoveNext
'    Loop
'
'    oRS.Close
'
'    no_of_ac_schedule = i - 1
'    deal_ticket.no_of_schedule = no_of_ac_schedule
'
'
'
'    If no_of_ac_schedule >= 1 Then
'
'        Erase schedule_list 'ReDim schedule_list(1 To no_of_schedule) As clsAutocallSchedule
'
'        Dim a_schedule As clsAutocallSchedule
'
'
'        For inx = 1 To no_of_ac_schedule
'
'            Set a_schedule = New clsAutocallSchedule
'
'            a_schedule.call_date = shtACPricer.Range("schedule_start").Cells(inx, 1).value
'
'            jnx = 1
'
'            Do
'                a_schedule.set_percent_strike shtACPricer.Range("schedule_start").Cells(inx, 4 + jnx).value, jnx
'                a_schedule.set_coupon_on_call shtACPricer.Range("schedule_start").Cells(inx, 1 + jnx).value, jnx
'
'                jnx = jnx + 1
'
'            Loop While shtACPricer.Range("schedule_start").Cells(inx, 4 + jnx).value > 0
'
'            '<--- added
'            a_schedule.performance_type = shtACPricer.Range("performance_type_on_call").Cells(inx, 1).value
'            If deal_ticket.no_of_ul = 3 And inx < no_of_ac_schedule And deal_ticket.settlement_date < #6/18/2017# Then
'                a_schedule.strike_shift = 0
'            Else
'                'a_schedule.strike_shift = shtACPricer.Range("strike_shift").value
'                a_schedule.strike_shift = shtACPricer.Range("strike_smoothing_width").Cells(inx, 1).value
'            End If
'            '--->
'
'            push_back_clsAutocallSchedule schedule_list, a_schedule
'
'        Next inx
'
'        deal_ticket.set_schedule_array schedule_list
'
'        deal_ticket.strike_at_maturity = schedule_list(no_of_ac_schedule).percent_strike '* deal_ticket.reference_price
'        deal_ticket.coupon_at_maturity = schedule_list(no_of_ac_schedule).coupon_on_call
'        deal_ticket.maturity_date = schedule_list(no_of_ac_schedule).call_date
'
'    Else
'
'        raise_err "read_ac_deal_ticket", "no schedule found"
'
'    End If
'
'
'
'
'
'
'
'    If shtACPricer.Range("no_of_early_exit_schedule").Cells(1, 1).value = 0 Then
'        deal_ticket.set_ul_dim shtACPricer.Range("No_of_Underlying").Cells(1, 1).value
'    Else
'        deal_ticket.set_ul_dim shtACPricer.Range("No_of_Underlying").Cells(1, 1).value, shtACPricer.Range("no_of_early_exit_schedule").Cells(1, 1).value
'    End If
'
'
'
'
'
'
'
''----------------------------------------
'    'Monthly coupon schedule
'    Dim monthly_coupon_flag As Boolean
'    monthly_coupon_flag = False
'
'    Dim coupon_date() As Date
'    Dim montly_coupon() As Double
'    Dim coupon_barr() As Double
'    ReDim coupon_date(0) As Date
'    ReDim montly_coupon(0) As Double
'    ReDim coupon_barr(0) As Double
'
'    Dim no_of_monthly_coupon_schedule As Integer
'
'
'    '쿠폰구간관리
'    SQL = " SELECT NVL(BONS_CUPN_STA_SDRT/100, 0) CouponBarrier, " _
'        & "         NVL(BONS_CUPN_FIN_SDRT/100, 100) CouponUpperBarrier, " _
'        & "         NVL(BONS_CUPN_INRT/100, 0) CouponRate, SCTN_STA_CTNU, SCTN_FIN_CTNU " _
'        & " FROM    bsys.TBSIMO210L00@gdw " _
'        & " WHERE   INDV_ISCD = '" + productCode + "' ORDER BY SCTN_STA_CTNU ASC"
'
'    With oCmd
'        .ActiveConnection = oDB
'        .CommandType = adCmdText
'        .CommandText = SQL
'
'        oRS.Open .Execute
'    End With
'
'    i = 1
'    Do Until oRS.EOF
'
'        ReDim Preserve coupon_barr(UBound(coupon_barr) + 1) As Double
'        ReDim Preserve montly_coupon(UBound(montly_coupon) + 1) As Double
'
'        For i = oRS(3) To oRS(4)
'            'coupon barrier
'            coupon_barr(i) = oRS(0)
'            'coupon on call
'            montly_coupon(i) = oRS(2)
'        Next i
'
'        oRS.MoveNext
'
'    Loop
'
'    oRS.Close
'
'    '쿠폰지급일관리
'    SQL = " SELECT TRTH_CUPN_VLTN_DATE COUPON_DATE " _
'        & " FROM    BSYS.TBSIMO213L00@GDW   " _
'        & " WHERE   INDV_ISCD = '" + productCode + "' ORDER BY 1 ASC"
'
'    With oCmd
'        .ActiveConnection = oDB
'        .CommandType = adCmdText
'        .CommandText = SQL
'
'        oRS.Open .Execute
'    End With
'
'    i = 1
'    'coupon date
'    Do Until oRS.EOF
'        monthly_coupon_flag = True
'        ReDim coupon_date(UBound(coupon_date) + 1) As Date
'
'        If redeem_shift Then
'            '2019.4.18
'            '오늘이 실제 조기상환평가일 후 첫 영업일이면 call date를 오늘로 설정: 상환여부 판단 때문
'            If today = WorksheetFunction.WorkDay(str2date(oRS(0)), 1, shtHolidays.Range("Holidays")) Then
'                coupon_date(i) = today
'            Else
'                '기초자산 지역에 따라 상환스케줄 1차 조정: 실제 평가일 + 1달력일
'                coupon_date(i) = str2date(oRS(0)) + 1
'            End If
'        Else
'            coupon_date(i) = str2date(oRS(0))
'        End If
'
'        i = i + 1
'        oRS.MoveNext
'
'    Loop
'
'    oRS.Close
'
'
'
'    no_of_monthly_coupon_schedule = i - 1
'
'
'
'    If monthly_coupon_flag Then
'        deal_ticket.monthly_coupon_flag = 1
'    Else
'        deal_ticket.monthly_coupon_flag = 0
'    End If
'
'    deal_ticket.set_coupon_schedule no_of_monthly_coupon_schedule, coupon_date, coupon_barr, montly_coupon
'    deal_ticket.monthly_coupon_amount = montly_coupon(1)
'
'
'
'
'    '----------------------------------------------------------------------
'    ' Read EE Schedule
'    '----------------------------------------------------------------------
'
'
''    For inx = 1 To deal_ticket.no_of_ul
''        deal_ticket.set_early_exit_barrier shtACPricer.Range("EE_Barrier").Cells(1, inx).value, inx
''        'deal_ticket.set_ki_barrier shtACPricer.Range("ki_barrier").Cells(1, 1).value, inx
''    Next inx
'
'    no_of_schedule = shtACPricer.Range("no_of_early_exit_schedule").Cells(1, 1).value
'
'    If no_of_schedule >= 1 Then
'
'        ReDim call_dates(1 To no_of_schedule) As Date
'        ReDim coupons(1 To no_of_schedule) As Double
'        ReDim strike_shifts(1 To no_of_schedule) As Double
'        'ReDim early_exit_touched_flags(1 To no_of_schedule) As Double
'        ReDim early_exit_touched_flags(1 To no_of_schedule) As Long  'data type 변경: dll(2018.7.17)
'        ReDim early_exit_performace_types(1 To no_of_schedule) As Long  'dll(2018.8.8)
'        ReDim early_exit_barrier_types(1 To no_of_schedule) As Long  'dll(2018.8.8)
'
'        For inx = 1 To no_of_schedule
'
'            call_dates(inx) = shtACPricer.Range("early_exit_schedule_start").Cells(inx, 1).value
'            coupons(inx) = shtACPricer.Range("early_exit_schedule_start").Cells(inx, 2).value
'            strike_shifts(inx) = 0
'
'            For jnx = 1 To deal_ticket.no_of_ul
'                deal_ticket.set_early_exit_barrier shtACPricer.Range("early_exit_schedule_start").Cells(inx, 5).value, deal_ticket.no_of_ul * (inx - 1) + jnx
'            Next jnx
'
'            If shtACPricer.Range("early_exit_schedule_start").Cells(inx, 6).value = "Y" Then
'                early_exit_touched_flags(inx) = 1#
'            ElseIf shtACPricer.Range("early_exit_schedule_start").Cells(inx, 6).value = "N" Then
'                early_exit_touched_flags(inx) = 0#
'            End If
'
'            early_exit_performace_types(inx) = shtACPricer.Range("early_exit_schedule_start").Cells(inx, 7).value  'dll(2018.8.8)
'            early_exit_barrier_types(inx) = shtACPricer.Range("early_exit_schedule_start").Cells(inx, 8).value  'dll(2018.8.8)
'
'        Next inx
'
'    End If
'
'    'deal_ticket.set_early_exit_schedule no_of_schedule, call_dates, coupons
'    'deal_ticket.set_early_exit_schedule no_of_schedule, call_dates, coupons, strike_shifts, early_exit_touched_flags
'    deal_ticket.set_early_exit_schedule no_of_schedule, call_dates, coupons, strike_shifts, early_exit_touched_flags, early_exit_performace_types, early_exit_barrier_types
'    '----------------------------------------------------------------------
'
'
'
'    no_of_schedule = shtACPricer.Range("no_of_floating_leg").value
'
'    If no_of_schedule >= 1 Then
'
'        ReDim call_dates(1 To no_of_schedule) As Date
'        ReDim coupons(1 To no_of_schedule) As Double
'
'        For inx = 1 To no_of_schedule
'
'            call_dates(inx) = shtACPricer.Range("floating_leg_start").Cells(inx, 1).value
'            coupons(inx) = shtACPricer.Range("floating_leg_start").Cells(inx, 2).value
'
'        Next inx
'
'
'    End If
'
'    deal_ticket.set_floating_schedule no_of_schedule, call_dates, coupons
'
'
'
'
'    Set read_ac_deal_ticket = deal_ticket
'
'    Set deal_ticket = Nothing
'    For inx = 1 To no_of_schedule
'        Set schedule_list(inx) = Nothing
'    Next inx
'
'    Exit Function
'
'ErrorHandler:
'
'    raise_err "read_ac_deal_ticket", Err.description
'
'
'End Function

'--------------------------------------------------------------------
' Function: read_deal_ticket
' Desc: Read deal sheet to make deal ticket object
'--------------------------------------------------------------------
Public Function read_deal_ticket(Optional time_step As Integer = -1, Optional trials As Long = -1, Optional required_tolerance = -1 _
                               , Optional homogeneous As Boolean = False, Optional local_floor As Double, Optional local_cap As Double)

    Dim schedule_dates() As Date
    Dim floors() As Double
    Dim caps() As Double
    Dim fixing_values() As Double
    
    Dim deal_ticket As clsCliquetDealTicket
    
    Dim no_of_schedule As Integer
    Dim inx As Integer
    
    
    On Error GoTo ErrorHandler

    Set deal_ticket = New clsCliquetDealTicket
    
    no_of_schedule = shtCliquetPricer.Range("no_of_schedule").Cells(1, 1).value
    
    ReDim schedule_dates(0 To no_of_schedule - 1) As Date
    ReDim floors(0 To no_of_schedule - 1) As Double
    ReDim caps(0 To no_of_schedule - 1) As Double
    ReDim fixing_values(0 To no_of_schedule - 1) As Double
    
    
    For inx = 0 To no_of_schedule - 1
    
        schedule_dates(inx) = shtCliquetPricer.Range("schedule_dates").Cells(inx + 1, 1).value
        
        If homogeneous Then
            floors(inx) = local_floor
            caps(inx) = local_cap
            fixing_values(inx) = -1
        Else
            floors(inx) = shtCliquetPricer.Range("floors").Cells(inx + 1, 1).value
            caps(inx) = shtCliquetPricer.Range("caps").Cells(inx + 1, 1).value
            fixing_values(inx) = shtCliquetPricer.Range("fixing_values").Cells(inx + 1, 1).value
        End If
    
    Next inx
    
    deal_ticket.current_date = shtCliquetPricer.Range("current_date").Cells(1, 1).value
    deal_ticket.value_date = shtCliquetPricer.Range("value_date").Cells(1, 1).value
'    deal_ticket.maturity = shtCliquetPricer.Range("maturity").Cells(1, 1).value
    
    deal_ticket.set_fixing_schedule no_of_schedule, schedule_dates, floors, caps, fixing_values
    
    deal_ticket.asset_code = shtCliquetPricer.Range("asset_code").Cells(1, 1).value
    deal_ticket.fund_code_c = shtCliquetPricer.Range("fund_code_c").Cells(1, 1).value
    deal_ticket.fund_code_m = shtCliquetPricer.Range("fund_code_m").Cells(1, 1).value
    deal_ticket.ul_code = shtCliquetPricer.Range("ul_code").Cells(1, 1).value
    
    deal_ticket.maturity_date = shtCliquetPricer.Range("maturity_date").Cells(1, 1).value
    deal_ticket.settlement_date = shtCliquetPricer.Range("settlement_date").Cells(1, 1).value
    deal_ticket.global_cap = shtCliquetPricer.Range("global_cap").Cells(1, 1).value
    deal_ticket.global_floor = shtCliquetPricer.Range("global_floor").Cells(1, 1).value
    
    deal_ticket.alive_yn = shtCliquetPricer.Range("alive_yn").Cells(1, 1).value
    deal_ticket.confirmed_yn = shtCliquetPricer.Range("confirmed_yn").Cells(1, 1).value
    deal_ticket.cliquet_type = shtCliquetPricer.Range("cliquet_type").Cells(1, 1).value
    deal_ticket.r_cliquet_cap = shtCliquetPricer.Range("r_cliquet_cap").Cells(1, 1).value
    deal_ticket.replace_period_no = shtCliquetPricer.Range("replace_period_no").Cells(1, 1).value
    deal_ticket.spread = shtCliquetPricer.Range("spread").Cells(1, 1).value
    deal_ticket.issue_cost = shtCliquetPricer.Range("issue_cost").Cells(1, 1).value
    
    If UCase(shtCliquetPricer.Range("buy_sell")) = "BUY" Then
        deal_ticket.notional = shtCliquetPricer.Range("notional").Cells(1, 1).value
    Else
        deal_ticket.notional = shtCliquetPricer.Range("notional").Cells(1, 1).value * -1
    End If
    
    'deal_ticket.comment = shtCliquetPricer.Range("comment").Cells(1, 1).value
    deal_ticket.comment = shtCliquetPricer.Range("txtComment").value
    
    deal_ticket.bump_vega = shtCliquetPricer.Range("bump_greek").Cells(1, 1).value
    deal_ticket.bump_skew = shtCliquetPricer.Range("bump_greek").Cells(1, 2).value
    deal_ticket.bump_theta = shtCliquetPricer.Range("bump_greek").Cells(1, 3).value
    
    
    If time_step = -1 Then
        deal_ticket.time_step = shtCliquetPricer.Range("time_step").Cells(1, 1).value
    Else
        deal_ticket.time_step = time_step
    End If
    
    If trials = -1 Then
        deal_ticket.no_of_trials = shtCliquetPricer.Range("trial").Cells(1, 1).value
    Else
        deal_ticket.no_of_trials = trials
    End If
    
    If required_tolerance = -1 Then
        deal_ticket.required_tolerance = shtCliquetPricer.Range("required_tolerance").Cells(1, 1).value
    Else
        deal_ticket.required_tolerance = required_tolerance
        
    End If
    
    deal_ticket.day_fraction_ = shtCliquetPricer.Range("day_fraction").value
    
    Set read_deal_ticket = deal_ticket
    
    Exit Function
    
ErrorHandler:

    raise_err "read_deal_ticket"
    
End Function



'--------------------------------------------------------------------
' Function: read_vanilla_ticket
' Desc: Read deal sheet to make deal ticket object
'--------------------------------------------------------------------
Public Function read_vanilla_ticket(Optional time_step As Integer = -1, Optional trials As Long = -1, Optional homogeneous As Boolean = False, Optional local_floor As Double, Optional local_cap As Double)

    Dim schedule_dates() As Date
    Dim floors() As Double
    Dim caps() As Double
    Dim fixing_values() As Double
    
    Dim deal_ticket As clsCliquetDealTicket
    
    Dim no_of_schedule As Integer
    Dim inx As Integer
    
    
    Dim vanilla_sheet As Worksheet
    
    Set vanilla_sheet = Sheets("vanilla")
    

    Set deal_ticket = New clsCliquetDealTicket
    
    no_of_schedule = vanilla_sheet.Range("no_of_schedule").Cells(1, 1).value
    
    ReDim schedule_dates(0 To no_of_schedule - 1) As Date
    ReDim floors(0 To no_of_schedule - 1) As Double
    ReDim caps(0 To no_of_schedule - 1) As Double
    ReDim fixing_values(0 To no_of_schedule - 1) As Double
    
    
    For inx = 0 To no_of_schedule - 1
    
        schedule_dates(inx) = vanilla_sheet.Range("schedule_dates").Cells(inx + 1, 1).value
        
        If homogeneous Then
            floors(inx) = local_floor
            caps(inx) = local_cap
            fixing_values(inx) = -1
        Else
            floors(inx) = vanilla_sheet.Range("floors").Cells(inx + 1, 1).value
            caps(inx) = vanilla_sheet.Range("caps").Cells(inx + 1, 1).value
            fixing_values(inx) = vanilla_sheet.Range("fixing_values").Cells(inx + 1, 1).value
        End If
    
    Next inx
    
    deal_ticket.set_fixing_schedule no_of_schedule, schedule_dates, floors, caps, fixing_values
    
    deal_ticket.asset_code = vanilla_sheet.Range("asset_code").Cells(1, 1).value
    deal_ticket.ul_code = vanilla_sheet.Range("ul_code").Cells(1, 1).value
    deal_ticket.current_date = vanilla_sheet.Range("current_date").Cells(1, 1).value
    deal_ticket.value_date = vanilla_sheet.Range("value_date").Cells(1, 1).value
    deal_ticket.Maturity = vanilla_sheet.Range("maturity").Cells(1, 1).value
    deal_ticket.maturity_date = vanilla_sheet.Range("maturity_date").Cells(1, 1).value
    deal_ticket.settlement_date = vanilla_sheet.Range("settlement_date").Cells(1, 1).value
    deal_ticket.global_cap = vanilla_sheet.Range("global_cap").Cells(1, 1).value
    deal_ticket.global_floor = vanilla_sheet.Range("global_floor").Cells(1, 1).value
    
    If UCase(Application.Range("buy_sell")) = "BUY" Then
        deal_ticket.notional = vanilla_sheet.Range("notional").Cells(1, 1).value
    Else
        deal_ticket.notional = vanilla_sheet.Range("notional").Cells(1, 1).value * -1
    End If
    
    deal_ticket.comment = vanilla_sheet.Range("comment").Cells(1, 1).value
'    deal_ticket.bump_greek = vanilla_sheet.Range("bump_greek").Cells(1, 1).value
    
    
    If time_step = -1 Then
        deal_ticket.time_step = vanilla_sheet.Range("time_step").Cells(1, 1).value
    Else
        deal_ticket.time_step = time_step
    End If
    
    If trials = -1 Then
        deal_ticket.no_of_trials = vanilla_sheet.Range("trial").Cells(1, 1).value
    Else
        deal_ticket.no_of_trials = trials
    End If
        
    
    Set read_vanilla_ticket = deal_ticket
    
    
End Function



Public Sub read_config(config As clsConfig)

On Error GoTo ErrorHandler
                                      
    Set config = New clsConfig
    
    config.ip_address_ = get_ip_address
    
    
    
    config.current_date_ = shtConfig.Range("date_config").Cells(1, 1).value
    
    config.time_step_closing_ = shtConfig.Range("mc_config").Cells(1, 1).value
    config.no_of_trials_closing_ = 2 ^ shtConfig.Range("mc_config").Cells(3, 1).value - 1
    
    config.time_step_ = shtConfig.Range("mc_config").Cells(4, 1).value
    config.required_tolerance_ = shtConfig.Range("mc_config").Cells(2, 1).value
    config.no_of_trials_ = 2 ^ shtConfig.Range("mc_config").Cells(5, 1).value - 1
    
    config.grid_interval_ = shtConfig.Range("closing_config").Cells(1, 1).value
    config.min_s_ = shtConfig.Range("closing_config").Cells(2, 1).value
    config.max_s_ = shtConfig.Range("closing_config").Cells(3, 1).value
    config.file_path_ = shtConfig.Range("closing_config").Cells(4, 1).value
    
    config.sparse_grid_level = shtConfig.Range("closing_config").Cells(5, 1).value
    config.sparse_grid_min = shtConfig.Range("closing_config").Cells(6, 1).value
    config.sparse_grid_max = shtConfig.Range("closing_config").Cells(7, 1).value
    
    config.max_retrial_count_ = shtConfig.Range("closing_config").Cells(8, 1).value
    config.snapshot_file_extension = shtConfig.Range("closing_config").Cells(9, 1).value
    config.position_file_extension = shtConfig.Range("closing_config").Cells(10, 1).value
    config.position_summary_file_extension = shtConfig.Range("closing_config").Cells(11, 1).value
    config.realtime_file_name = shtConfig.Range("closing_config").Cells(12, 1).value
    config.batch_size = shtConfig.Range("closing_config").Cells(14, 1).value
    
    
    config.market_refresh_interval_ = shtConfig.Range("calculation_config").Cells(1, 1).value
    
    config.x_grid_ = shtConfig.Range("fdm_config").Cells(1, 1).value
    config.v_grid_ = shtConfig.Range("fdm_config").Cells(2, 1).value
    config.time_step_per_day = shtConfig.Range("fdm_config").Cells(3, 1).value
    config.fdm_scheme_ = shtConfig.Range("fdm_config").Cells(4, 2).value
    
    config.empirical_vega_weighting = shtConfig.Range("misc_config").Cells(2, 1).value
    config.vega_reference_maturity = shtConfig.Range("misc_config").Cells(1, 1).value
    
    config.no_of_strike_grid = shtConfig.Range("misc_config").Cells(3, 1).value
    config.width_of_strike = shtConfig.Range("misc_config").Cells(4, 1).value
    
    
    If UCase(shtConfig.Range("calculation_config").Cells(2, 1).value) = "ON" Then
        config.auto_calculation_ = True
    Else
        config.auto_calculation_ = False
    End If
    
    config.intra_day_greek_ = shtConfig.Range("calculation_config").Cells(3, 1).value
    config.neglect_barrier_smoothing_ = shtConfig.Range("calculation_config").Cells(4, 1).value
    
    config.adjust_strike_shift_percent = shtConfig.Range("closing_config").Cells(13, 1).value
    
    config.set_term_vega_tenor read_term_vega_tenor
    
    Exit Sub
    
ErrorHandler:

    raise_err "config"

End Sub

Public Function read_implied_spot(ByRef spot As Double, ByRef prev_s As Double, ByRef first_futures_date As Date)

        
    spot = shtMarket.Range("S").Cells(4, 1).value
    prev_s = shtMarket.Range("prev_s").Cells(4, 1).value
    first_futures_date = shtMarket.Range("S").Cells(4, 0).value
    
End Function


Public Function read_market(ByRef spot As Double, ByRef rate_curve As clsRateCurve, ByRef div_schedule As clsDividendSchedule, ByRef heston_param As clsHestonParameter, ByRef prev_s As Double)

    
    Dim dates() As Date
    Dim values() As Double
    Dim inx As Integer
        
    spot = shtMarket.Range("S").Cells(1, 1).value
'    prev_s = shtMarket.Range("prev_s").Cells(1, 1).value
    
    Set rate_curve = New clsRateCurve
    
    rate_curve.initialize shtMarket.Range("rate_dates"), shtMarket.Range("discount")
    
    read_div dates, values
    Set div_schedule = New clsDividendSchedule
    
    div_schedule.initialize_div UBound(dates), dates, values
    

End Function

Public Sub read_sabr_parameters(ByRef maturities() As Date, ByRef rho() As Double, ByRef nu() As Double, ByRef rho_coeff() As Double, ByRef nu_coeff() As Double)

    Dim no_of_maturities As Integer
    Dim no_of_rho_coeff As Integer
    Dim no_of_nu_coeff As Integer
    Dim inx As Integer
    
    
    no_of_maturities = shtMarket.Range("no_sabr_maturities").Cells(1, 1).value
    no_of_rho_coeff = shtMarket.Range("no_of_rho_coeff").Cells(1, 1).value
    no_of_nu_coeff = shtMarket.Range("no_of_nu_coeff").Cells(1, 1).value
    
    ReDim maturities(1 To no_of_maturities) As Date
    ReDim rho(1 To no_of_maturities) As Double
    ReDim nu(1 To no_of_maturities) As Double
    ReDim rho_coeff(1 To no_of_rho_coeff) As Double
    ReDim nu_coeff(1 To no_of_nu_coeff) As Double
    
    For inx = 1 To no_of_maturities
        maturities(inx) = shtMarket.Range("sabr_parameters").Cells(inx, 1).value
        rho(inx) = shtMarket.Range("sabr_parameters").Cells(inx, 3).value
        nu(inx) = shtMarket.Range("sabr_parameters").Cells(inx, 4).value
    Next inx
    
    For inx = 1 To no_of_rho_coeff
        rho_coeff(inx) = shtMarket.Range("sabr_coefficient").Cells(inx, 1).value
    Next inx
    
    For inx = 1 To no_of_nu_coeff
        nu_coeff(inx) = shtMarket.Range("sabr_coefficient").Cells(inx, 2).value
    Next inx
        

End Sub

Public Sub read_div(ByRef dates() As Date, ByRef values() As Double)

    Dim inx As Integer
    Dim count As Integer
    
    count = shtMarket.Range("dividend_dates").count
    
    ReDim dates(1 To count) As Date
    ReDim values(1 To count) As Double
    
    For inx = 1 To count
    
        If shtMarket.Range("dividend_dates").Cells(inx, 1).value <> "" Then
    
            dates(inx) = shtMarket.Range("dividend_dates").Cells(inx, 1).value
            values(inx) = shtMarket.Range("dividend_values").Cells(inx, 1).value
        
        Else
            
            Exit For
        
        End If
    
    Next inx
    
    
    ReDim Preserve dates(1 To inx - 1) As Date
    ReDim Preserve values(1 To inx - 1) As Double
    

End Sub

Public Function get_barrier_code_name(range_name As String, code_value As Integer) As String


    Dim inx As Integer
    Dim rtn_value As String
    
On Error GoTo ErrorHandler

    For inx = 1 To shtBarrierPricer.Range(range_name).Rows.count
    
        If shtBarrierPricer.Range(range_name).Cells(inx, 2).value = code_value Then
            rtn_value = shtBarrierPricer.Range(range_name).Cells(inx, 1)
            Exit For
        End If
    
    Next inx

    get_barrier_code_name = rtn_value
    
    Exit Function
    
ErrorHandler:

    raise_err "get_barrier_code_name"



End Function

Public Function get_ac_code_name(range_name As String, code_value As Integer) As String


    Dim inx As Integer
    Dim rtn_value As String
    
On Error GoTo ErrorHandler

    For inx = 1 To shtACPricer.Range(range_name).Rows.count
    
        If shtACPricer.Range(range_name).Cells(inx, 2).value = code_value Then
            rtn_value = shtACPricer.Range(range_name).Cells(inx, 1)
            Exit For
        End If
    
    Next inx

    get_ac_code_name = rtn_value
    
    Exit Function
    
ErrorHandler:

    raise_err "get_ac_code_name"



End Function