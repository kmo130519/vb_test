Public Declare PtrSafe Function local_vol_vanilla_pricer Lib "C:\cpp_dll\indexPricer_v1.90.dll" _
(ByRef value As Double, ByRef delta As Double, ByRef gamma As Double, ByRef theta As Double _
, ByVal s As Double, ByVal r As Double, ByVal q As Double, ByVal t As Double, ByVal k As Double _
, ByVal call_put As Long, ByVal evaluationDate As Long, ByVal optionType As Long _
, volDates_in() As Long, volStrikes_in() As Double, volMatrix_in() As Double _
, divDates_in() As Long, divs_() As Double _
, ByVal ratio_dividend As Double, ByVal refprice_for_dividend As Double) As Long



'Name           In/Out  VBA Type    Desc
'-----------------------------------------------------
'value           Out     double     평가값. Local Currency. KRW 아님.
'delta           Out     double     delta. dValue / dS.
'gamma           Out     double     gamma. dDelta / dS.
'theta           Out     double     theta. dValue / dt. t 단위는 년.
's               In      double     지수.
'r               In      double     금리.
'q               In      double     배당율.
't               In      double     time to maturity. 년도 단위.
'k               In      double     행사가.
'call_put        In      Long       call / put 구분. 0: Call, 1: Put
'evaluation_date In      Long       평가일.
'option_type     In      Long       European / American 구분. 0: European, 1: American.
'vol_dates       In      Long()     로컬볼 서피스 테너 리스트. 1 부터 시작. e.g. vol_dates(1) = 42601, vol_dates(2) = 42603, .... vol_dates(15) = 43084.
'vol_strikes     In      double()   로컬볼 서피스 행사가 리스트. vol_strikes(1) = 1264.2, vol_strike(2) = 1292.13,..... vol_strike(51) = 3712.95
'vol_matrix      In      double()   로컬볼 데이터. 테너 X 행사가의 2 차원 배열. e.g. vol_matrix(1,1 ) = 0.999,.... vol_matrix(1,51)= 0.9999,... vol_matrix(15,1) = ..., vol_matrixk( 15, 51 ) =....
'div_dates       In      Long()     배당락일 리스트.
'divs            In      double()   배당금액.


Public Sub run_futopt_pricing(eval_date As Date _
                            , ByRef the_greeks As clsGreeks _
                            , deal_ticket As clsVanillaOptionDealTicket _
                            , market As clsMarket _
                            , Optional calc_greeks As Boolean = False _
                            , Optional is_ul_vol_index As Boolean = False)

'<------Adjustment for options on futures: 기초자산 가격을 선물이론가로 넣고 보정하는 것 = 기초자산 가격을 SPX로 넣고 보정을 안하는 것 일거 같은 생각이 드는데 검토 부탁드립니다.
'        If ul_code = "SPX" Then
'
'            pv_divs = 0
'            For m = 0 To UBound(divs)
'                If (CDate(div_dates(m)) - market_date) / 365 < t Then
'                    pv_divs = pv_divs + divs(m) * market.rate_curve_.get_discount_factor(CDate(div_dates(m)), market_date)
'                Else
'                    Exit For
'                End If
'            Next m
'
'            '선도이론가
'            s = (market.s_ - pv_divs) * Exp((r - q) * t)
'
'            '기존 SPX기준 스트라이크 포인트에 대응되는 선도가격을 만기별로 계산 후, 선도가격이 기존 SPX기준 스트라이크 포인트일 때 로컬볼 계산(만기별 로컬볼 보간)
'            Dim tj As Double
'            Dim rj As Double
'            Dim fj() As Double
'            ReDim fj(UBound(vol_strikes))
'
'            For j = 1 To UBound(vol_dates)
'
'                tj = (maturity_date - CDate(vol_dates(j))) / 365
'                rj = market.rate_curve_.get_fwd_rate(maturity_date, CDate(vol_dates(j)))
'
'                pv_divs = 0
'                For m = 0 To UBound(divs)
'                    If div_dates(m) >= vol_dates(j) And (div_dates(m) - vol_dates(j)) / 365 < tj Then
'                        pv_divs = pv_divs + divs(m) * market.rate_curve_.get_discount_factor(CDate(div_dates(m)), CDate(vol_dates(j)))
'                    End If
'                Next m
'
'                For m = 1 To UBound(vol_strikes)
'                    fj(m) = (vol_strikes(m) - pv_divs) * Exp((rj - q) * tj)
'                Next m
'
'                Dim vj() As Double
'                ReDim vj(UBound(vol_strikes))
'                Dim vj_new() As Double
'                ReDim vj_new(UBound(vol_strikes))
'
'                For m = 1 To UBound(vol_strikes)
'                    vj(m) = vol_matrix(j, m)
'                Next m
'
'                For m = 1 To UBound(vol_strikes)
'                    vj_new(m) = quadratic_interpolation(vol_strikes(m), fj, vj)
'                Next m
'
'                For m = 1 To UBound(vol_strikes)
'                    vol_matrix(j, m) = vj_new(m)
'                Next m
'
'            Next j
'
'
'        End If
'------>

    If is_ul_vol_index = True Then
        market.s_ = get_spot_price(deal_ticket.ul_code, date2str(eval_date))
        market.div_yield_ = 0
        Set market.div_schedule_ = Nothing
        Set market.drift_adjust_ = Nothing
    End If
            
    Dim i As Integer
    
    Dim k As Double
    Dim s As Double
    Dim t As Double
    Dim r As Double
    Dim q As Double

    Dim div_dates() As Long
    Dim divs() As Double
            
    Dim price As Double
    Dim delta As Double
    Dim gamma As Double
    Dim theta As Double

    Dim stickymoneyness_delta_greek_up As clsGreeks
    Dim stickymoneyness_delta_greek_down As clsGreeks
    Dim vol_bump_greek As clsGreeks

    Dim shifted_strikes() As Double
    Dim pv_divs As Double
            
    k = deal_ticket.k
    t = (deal_ticket.maturity_date - eval_date) / 365
    r = market.rate_curve_.get_fwd_rate(eval_date, deal_ticket.maturity_date)
    s = market.s_
    
    If market.drift_adjust_ Is Nothing Then
        q = 0
    Else
        q = market.div_yield_ - market.get_drift_adjust(deal_ticket.maturity_date) 'drift adjustment 추가: 2023.11.21
    End If
    
    If market.div_schedule_ Is Nothing Then
        pv_divs = 0
    Else
        div_dates = market.div_schedule_.get_div_dates
        divs = market.div_schedule_.get_divs
    
        For i = 0 To UBound(divs)
            If (div_dates(i) - CLng(eval_date)) / 365 < t Then
                pv_divs = pv_divs + divs(i) * market.rate_curve_.get_discount_factor(CDate(div_dates(i)), eval_date)
            End If
        Next i
    End If
    
'        vol_dates = market.sabr_surface_.local_vol_surface.grid_.get_all_dates_as_long(1, 0, eval_date)
'        vol_strikes = market.sabr_surface_.local_vol_surface.grid_.get_all_strikes(1)
'        vol_matrix = market.sabr_surface_.local_vol_surface.vol_surface

    Select Case deal_ticket.prod_type
    Case "F"
       
        price = (s - pv_divs) * Exp((r - q) * t)
        
    Case "C", "P"

        Call local_vol_vanilla_pricer(price, delta, gamma, theta, _
                                        s, _
                                        r, _
                                        q, _
                                        t, _
                                        k, _
                                        deal_ticket.call_put, CLng(eval_date), deal_ticket.option_type, _
                                        market.sabr_surface_.local_vol_surface.grid_.get_all_dates_as_long(1, 0, eval_date), _
                                        market.sabr_surface_.local_vol_surface.grid_.get_all_strikes(1), _
                                        market.sabr_surface_.local_vol_surface.vol_surface, _
                                        div_dates, divs, _
                                        0, market.refPriceForDividend)
    End Select
    
    'from unit metrics to dollar values
    the_greeks.ul_price = s
    the_greeks.value = price * deal_ticket.conversion_ratio * deal_ticket.qty
            
    If calc_greeks = True Then
    
        Select Case deal_ticket.prod_type
        Case "F"
            
            If is_ul_vol_index = True Then
                delta = 0
                the_greeks.vega = Exp((r - q) * t) * deal_ticket.conversion_ratio * deal_ticket.qty
            Else
                delta = Exp((r - q) * t)
                the_greeks.vega = 0
            End If
            
            gamma = 0
            theta = -1 * (s - pv_divs) * Exp((r - q) * t) * (r - q)

            the_greeks.sticky_moneyness_delta = delta * s * deal_ticket.conversion_ratio * deal_ticket.qty
            the_greeks.sticky_moneyness_gamma = 0
                
        Case "C", "P"
        
            Dim backup_market As clsMarket
            Set backup_market = market.copy_obj
            
            'Set market = backup_market.copy_obj
            
            'sticky monenyness delta (1% up)
            Set stickymoneyness_delta_greek_up = New clsGreeks

            'market.s_ = s * 1.01
            backup_market.s_ = s * 1.01
            
            'shifted_strikes = market.sabr_surface_.local_vol_surface.grid_.get_all_strikes
            shifted_strikes = backup_market.sabr_surface_.local_vol_surface.grid_.get_all_strikes
            'For i = 1 To market.sabr_surface_.local_vol_surface.grid_.no_of_strikes
            For i = 1 To backup_market.sabr_surface_.local_vol_surface.grid_.no_of_strikes
                shifted_strikes(i) = shifted_strikes(i) * 1.01
            Next i
            'market.sabr_surface_.local_vol_surface.grid_.set_strikes shifted_strikes
            backup_market.sabr_surface_.local_vol_surface.grid_.set_strikes shifted_strikes
            
            'Call run_futopt_pricing(eval_date, stickymoneyness_delta_greek_up, deal_ticket, market)
            Call run_futopt_pricing(eval_date, stickymoneyness_delta_greek_up, deal_ticket, backup_market)
            
            'Set market = backup_market.copy_obj
            Set backup_market = market.copy_obj
            
            'sticky monenyness delta (1% down)
            Set stickymoneyness_delta_greek_down = New clsGreeks

            'market.s_ = s * 0.99
            backup_market.s_ = s * 0.99
            
            'shifted_strikes = market.sabr_surface_.local_vol_surface.grid_.get_all_strikes
            shifted_strikes = backup_market.sabr_surface_.local_vol_surface.grid_.get_all_strikes
            'For i = 1 To market.sabr_surface_.local_vol_surface.grid_.no_of_strikes
            For i = 1 To backup_market.sabr_surface_.local_vol_surface.grid_.no_of_strikes
                shifted_strikes(i) = shifted_strikes(i) * 0.99
            Next i
            'market.sabr_surface_.local_vol_surface.grid_.set_strikes shifted_strikes
            backup_market.sabr_surface_.local_vol_surface.grid_.set_strikes shifted_strikes
            
            'Call run_futopt_pricing(eval_date, stickymoneyness_delta_greek_down, deal_ticket, market)
            Call run_futopt_pricing(eval_date, stickymoneyness_delta_greek_down, deal_ticket, backup_market)
            
            'Set market = backup_market.copy_obj
            Set backup_market = market.copy_obj
            
            '1%p vega
            Set vol_bump_greek = New clsGreeks
            
            'market.sabr_surface_.bump_vol_surface 0.01
            backup_market.sabr_surface_.bump_vol_surface 0.01
            
            'Call run_futopt_pricing(eval_date, vol_bump_greek, deal_ticket, market)
            Call run_futopt_pricing(eval_date, vol_bump_greek, deal_ticket, backup_market)
            
            'market.sabr_surface_.rewind_vol_bump
            backup_market.sabr_surface_.rewind_vol_bump
            
            'Set market = backup_market.copy_obj
            Set backup_market = market.copy_obj
        
            'pertubation: from dollar values to dollar values
            the_greeks.vega = vol_bump_greek.value - the_greeks.value
            the_greeks.sticky_moneyness_delta = (stickymoneyness_delta_greek_up.value - stickymoneyness_delta_greek_down.value) / (2 * 0.01 * the_greeks.ul_price) * s
            the_greeks.sticky_moneyness_gamma = (stickymoneyness_delta_greek_up.value + stickymoneyness_delta_greek_down.value - 2 * the_greeks.value) / (the_greeks.ul_price * 0.01) ^ 2
            
        End Select
    
        'from unit metrics to dollar values
        the_greeks.delta = delta * s * deal_ticket.conversion_ratio * deal_ticket.qty
        the_greeks.gamma = 0.5 * gamma * (s * 0.01) ^ 2 * deal_ticket.conversion_ratio * deal_ticket.qty
        the_greeks.theta = theta * deal_ticket.conversion_ratio * deal_ticket.qty
        
    End If

    Set stickymoneyness_delta_greek_up = Nothing
    Set stickymoneyness_delta_greek_down = Nothing
    Set vol_bump_greek = Nothing
    Set backup_market = Nothing
    
End Sub