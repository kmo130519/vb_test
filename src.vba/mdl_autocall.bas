Option Explicit

'Const NUM_UA As Integer = 16

'Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private pricer_deal_index__ As Integer

Public Sub cmd_calc_ac_price(ByVal deal_ticket As clsACDealTicket, ByVal market_set As clsMarketSet, Optional ByRef greeks As clsGreeks = Null)

    'Dim greeks As New clsGreeks
    'Dim tStart As Long
    
    'Dim deal_ticket As clsACDealTicket
    
    'Dim market As clsMarket
    'Dim market_set As clsMarketSet
    
    'Dim bump_greek As Boolean
    'Dim bump_greek_set As clsGreekSet

On Error GoTo ErrorHandler
    
    'Set deal_ticket = read_ac_deal_ticket()
    
    'bump_greek = (shtACPricer.Range("bump_greek").value = "Y")
    
   
    '-------------------------
    ' 1 Index : Pricing w/ Stoc.Vol
    '-------------------------
'    If deal_ticket.no_of_ul = 1 And deal_ticket.vol_scheme_type = 0 Then
'
'        read_market s, rate_curve, div_schedule, heston_param, prev_s
'
'        Set market = New clsMarket
'        market.set_market rate_curve, s, div_schedule, heston_param
'
'        bump_greek = (shtACPricer.Range("bump_greek").value = "Y")
'        'mid_day_greek = (shtACPricer.range("mid_day_greek").value = "Y")
'
'        run_ac_pricing greeks, deal_ticket, market, last_node, bump_greek ', , , mid_day_greek
'    'theta_adjustment greeks, deal_ticket.current_date, deal_ticket.current_date + 1
'        display_greeks greeks
'
'    Else
    '-------------------------
    ' 2 Indices : Pricing w/ local Vol
    '-------------------------
        
        'Set market_set = read_market_set(deal_ticket.ccy)
        Dim pricing_mode As String
        
'        If deal_ticket.isLeveraged Or deal_ticket.no_of_ul >= 3 Then
'            '기초자산이 레버리지 지수를 포함하거나 기초자산이 3개 이상인 경우
'            pricing_mode = "MC"
'        Else
'            '스트레스테스트 일 경우는 MC
'            If shtMarket.Range("SCENARIO_ENABLE") = False Then
'                pricing_mode = "FDM"
'            Else
'                pricing_mode = "MC"
'            End If
'        End If
        'local correaltion 적용하면서 무조건 MC 적용 2019.3.27
        pricing_mode = "MC"
            
        Select Case pricing_mode
        Case "FDM":
            run_ac_pricing_fdm greeks, deal_ticket, market_set _
                                , shtACPricer.Range("chkDeltaGamma").value _
                                , shtACPricer.Range("chkStickyMoneynessDeltaGamma").value _
                                , shtACPricer.Range("chkStickyStrikeDelta").value _
                                , shtACPricer.Range("chkCrossGamma").value _
                                , shtACPricer.Range("chkVega").value _
                                , shtACPricer.Range("chkTermVega").value _
                                , shtACPricer.Range("chkSkew").value _
                                , shtACPricer.Range("chkCorr").value _
                                , shtACPricer.Range("chkRho").value _
                                , shtACPricer.Range("chkTheta").value
                                
        Case "MC"
            run_ac_pricing_mc greeks, deal_ticket, market_set _
                                , shtACPricer.Range("chkDeltaGamma").value _
                                , shtACPricer.Range("chkStickyMoneynessDeltaGamma").value _
                                , shtACPricer.Range("chkStickyStrikeDelta").value _
                                , shtACPricer.Range("chkCrossGamma").value _
                                , shtACPricer.Range("chkVega").value _
                                , shtACPricer.Range("chkTermVega").value _
                                , shtACPricer.Range("chkSkew").value _
                                , shtACPricer.Range("chkCorr").value _
                                , shtACPricer.Range("chkRho").value _
                                , shtACPricer.Range("chkTheta").value _
                                , shtACPricer.Range("chkLocalCorrelation").value
        End Select
        
'        Select Case deal_ticket.no_of_ul
'        Case 1:
'            run_ac_pricing_1d greeks, deal_ticket, market_set, bump_greek_set _
'                            , shtACPricer.Range("chkVega").value _
'                            , shtACPricer.Range("chkSkew").value _
'                            , shtACPricer.Range("chkCorr").value _
'                            , shtACPricer.Range("chkRho").value _
'                            , _
'                            , _
'                            , shtACPricer.Range("chkBump").value
'        Case 2:
'            run_ac_pricing_2d greeks, deal_ticket, market_set, bump_greek_set _
'                            , shtACPricer.Range("chkVega").value _
'                            , shtACPricer.Range("chkSkew").value _
'                            , shtACPricer.Range("chkCorr").value _
'                            , shtACPricer.Range("chkRho").value _
'                            , _
'                            , _
'                            , shtACPricer.Range("chkSnapshot").value _
'                            , shtACPricer.Range("chkBump").value _
'                            , shtACPricer.Range("chkDelta").value _
'                            , shtACPricer.Range("chkTermVega").value
'
'        Case 3:
'            run_ac_pricing_3d greeks, deal_ticket, market_set, bump_greek_set _
'                            , shtACPricer.Range("chkVega").value _
'                            , shtACPricer.Range("chkSkew").value _
'                            , shtACPricer.Range("chkCorr").value _
'                            , shtACPricer.Range("chkRho").value _
'                            , shtACPricer.Range("chkTheta").value _
'                            , _
'                            , _
'                            , shtACPricer.Range("chkSnapshot").value _
'                            , shtACPricer.Range("chkBump").value _
'                            , shtACPricer.Range("chkDelta").value _
'                            , shtACPricer.Range("chkTermVega").value
'        End Select
        
        clear_ac_greeks
        
        display_greeks_nd greeks, deal_ticket.no_of_ul
            
'    End If

    'Set greeks = Nothing

    Exit Sub
    
ErrorHandler:

    raise_err "cmd_calc_ac_price", Err.description

End Sub

'Public Function read_market_set(base_ccy As String, Optional isPrevDate As Boolean = False) As clsMarketSet
Public Function read_market_set(Optional riskfree_dcf_enable As Boolean = False, Optional isPrevDate As Boolean = False) As clsMarketSet '.... 2024.03.29 riskfree_dcf_enable 추가

    Dim market_set As clsMarketSet
    
On Error GoTo ErrorHandler

    Set market_set = New clsMarketSet
    
    '기초자산별 시장정보
    Dim i As Integer
    For i = 1 To NUM_UA
        If is_active_ua(get_ua_code(i)) = True Then
            market_set.set_market get_ua_code(i), read_index_market(get_ua_code(i), True, isPrevDate)
        End If
    Next i

    '할인금리커브
    Dim rate_curve As clsRateCurve
    Dim no_of_rate_dates As Integer
    
    If riskfree_dcf_enable = True Then '.... 2024.03.29 riskfree_dcf_enable 추가
        Call market_set.set_pl_currency_rate_curve(DCF.KRW, read_rate_curve(shtMarket.Range("KOSPI200"), isPrevDate))
        Call market_set.set_pl_currency_rate_curve(DCF.USD, read_rate_curve(shtMarket.Range("SPX"), isPrevDate))
    Else
        For i = 1 To NUM_DCF
            Call market_set.set_pl_currency_rate_curve(i, read_rate_curve(shtMarket.Range("DCF_" & get_dcf_ccy(i)), isPrevDate))
        Next i
    End If
    
    Set market_set.correlation_pair_ = read_corr_set(isPrevDate)
    
    '<----- local correlation 추가 2019.3.27
    Set market_set.min_correlation_pair_ = read_min_corr_set(isPrevDate)
    '----->/
    
    Set read_market_set = market_set
    
    Set market_set = Nothing

    Exit Function
    
ErrorHandler:

    raise_err "read_market_set", Err.description


End Function

Public Function update_market_set(ByVal market_set_yesterday As clsMarketSet, ByVal ul_code As String) As clsMarketSet
    
    '특정 지수의 변동성만 평가일로 update한다.
   
On Error GoTo ErrorHandler
    
    Dim market_set_tmp As clsMarketSet
    Set market_set_tmp = market_set_yesterday.copy_obj
    
    Dim index_name As String
    index_name = market_set_tmp.market_by_ul(ul_code).index_name
    
    Dim spot As Double

    Dim eval_date As Date
    eval_date = shtMarket.Range("market_date").value

    Dim sabr_surface As clsSABRSurface
    Set sabr_surface = New clsSABRSurface
    
    read_local_vol sabr_surface, index_name, spot, eval_date
    
    Set market_set_tmp.market(market_set_tmp.find_index(ul_code)).sabr_surface_ = sabr_surface
    
    Set update_market_set = market_set_tmp
    
    Exit Function
    
ErrorHandler:

    raise_err "update_market_set", Err.description


End Function

Public Sub cmd_load_ac_deal()

    Dim deal_ticket As clsACDealTicket

On Error GoTo ErrorHandler
    
   ' clear_deal

    Set deal_ticket = New clsACDealTicket
    deal_ticket.asset_code = shtACPricer.Range("asset_code").Cells(1, 1).value
    
    '--------------------------------------
    ' Load deal information
    '--------------------------------------
    If retrieve_ac_deal(deal_ticket) Then
    
        
        '---------------------------------------
        ' Display deal information
        '---------------------------------------
        display_ac_deal deal_ticket
        
    Else
    
        MsgBox "Invalid asset code"
        
    End If
    
    Exit Sub
    
ErrorHandler:

'    show_error
    raise_err "cmd_load_ac_deal"
    
    
End Sub

Public Sub cmd_new_ac_deal()

    Dim deal_ticket As clsACDealTicket
    
On Error GoTo ErrorHandler
    
    'Read deal information from the deal sheet
    Set deal_ticket = read_ac_deal_ticket()
    
    'insert into the database
    insert_ac_deal deal_ticket
    
    MsgBox "Insertion Successful!"
    
  '  cmd_load_ac_deal
    
    Exit Sub
    
ErrorHandler:
        
'    cmd_load_deal
    raise_err "cmd_new_ac_deal"
    
End Sub



Public Sub display_greeks_nd(greeks As clsGreeks, no_of_ul As Integer)

    Dim inx As Integer
    Dim jnx As Integer
    
On Error Resume Next
    
    shtACPricer.Range("price").Cells(2, 1).value = greeks.value
    
    Dim adoCon As New adoDB.Connection
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
    For inx = 1 To no_of_ul
        shtACPricer.Range("delta").Cells(2 * inx, 1).value = greeks.deltas(inx)
        
        shtACPricer.Range("gamma").Cells(2 * inx, inx).value = greeks.gammas(inx)
        
        shtACPricer.Range("vega").Cells(2 * inx, 1).value = greeks.vegas(inx)
        If Err.number = 9 Then
            shtACPricer.Range("vega").Cells(2 * inx, 1).value = 0
        ElseIf Err.number <> 0 Then
            raise_err "display_greeks_nd", Err.description
        End If
        
        shtACPricer.Range("sticky_moneyness_delta").Cells(2 * inx, 1).value = greeks.sticky_moneyness_deltas(inx)
        If Err.number = 9 Then
            shtACPricer.Range("sticky_moneyness_delta").Cells(2 * inx, 1).value = 0
        ElseIf Err.number <> 0 Then
            raise_err "display_greeks_nd", Err.description
        End If
        
        shtACPricer.Range("skew_sensitivity").Cells(2 * inx, 1).value = greeks.skew_s_s(inx)
        If Err.number = 9 Then
            shtACPricer.Range("skew_sensitivity").Cells(2 * inx, 1).value = 0
        ElseIf Err.number <> 0 Then
            raise_err "display_greeks_nd", Err.description
        End If
        
        shtACPricer.Range("vanna").Cells(2 * inx, 1).value = greeks.vannas(inx)
        If Err.number = 9 Then
            shtACPricer.Range("vanna").Cells(2 * inx, 1).value = 0
        ElseIf Err.number <> 0 Then
            raise_err "display_greeks_nd", Err.description
        End If
        
        shtACPricer.Range("rho_ul").Cells(2 * inx, 1).value = greeks.rho_ul(inx)
        If Err.number = 9 Then
            shtACPricer.Range("rho_ul").Cells(2 * inx, 1).value = 0
        ElseIf Err.number <> 0 Then
            raise_err "display_greeks_nd", Err.description
        End If
        
        If (Abs(greeks.rho_ul(inx)) < 1) Or (Abs(greeks.deltas(inx)) < 1) Then
            shtACPricer.Range("Duration").Cells(2 * inx, 1).value = 0
        Else
            shtACPricer.Range("Duration").Cells(2 * inx, 1).value = greeks.rho_ul(inx) / (greeks.deltas(inx) * greeks.ul_prices(inx)) * 10000
        End If
        If Err.number = 9 Then
            shtACPricer.Range("Duration").Cells(2 * inx, 1).value = 0
        ElseIf Err.number <> 0 Then
            raise_err "display_greeks_nd", Err.description
        End If
        
        'Skew
        Dim today As Date
        today = shtACPricer.Range("current_date")
        
        Dim tDayStr As String
        tDayStr = Left(CStr(today), 4) & Mid(CStr(today), 6, 2) & Right(CStr(today), 2)
        
        Dim indexid As String
        indexid = shtACPricer.Range("ul_code").Cells(1, inx).value
        
        Dim td_spot As Double
        td_spot = shtACPricer.Range("ul_spot").Cells(2 * inx, 1).value
        
        Dim tau As Double
        tau = shtACPricer.Range("Duration").Cells(2 * inx, 1).value
        
        Dim iv_110 As Double
        Dim iv_90 As Double
        iv_110 = get_vol_on_surface(tDayStr, indexid, td_spot, tau, 1.1, "Implied", "FRONT", adoCon)
        iv_90 = get_vol_on_surface(tDayStr, indexid, td_spot, tau, 0.9, "Implied", "FRONT", adoCon)
        
        Dim skew_tau As Double
        skew_tau = (iv_110 - iv_90) / 0.2
        
        shtACPricer.Range("skew").Cells(2 * inx, 1).value = skew_tau
        If Err.number = 9 Then
            shtACPricer.Range("skew").Cells(2 * inx, 1).value = 0
        ElseIf Err.number <> 0 Then
            raise_err "display_greeks_nd", Err.description
        End If

        shtACPricer.Range("delta_adj").Cells(2 * inx, 1).value = -0.5 * greeks.vegas(inx) * skew_tau / td_spot * 100
        If Err.number = 9 Then
            shtACPricer.Range("delta_adj").Cells(2 * inx, 1).value = 0
        ElseIf Err.number <> 0 Then
            raise_err "display_greeks_nd", Err.description
        End If
        
    Next inx
    
    If no_of_ul = 2 Then
        shtACPricer.Range("gamma").Cells(4, 1) = greeks.cross_gamma12
    ElseIf no_of_ul = 3 Then
        shtACPricer.Range("gamma").Cells(4, 1) = greeks.cross_gamma12
        shtACPricer.Range("gamma").Cells(6, 1) = greeks.cross_gamma13
        shtACPricer.Range("gamma").Cells(6, 2) = greeks.cross_gamma23
    End If
    
    For inx = 1 To no_of_ul
        For jnx = inx + 1 To no_of_ul
            shtACPricer.Range("Corr_Sens.").Cells(jnx * 2, inx) = greeks.corr_sens(inx, jnx)
        Next jnx
        
    Next inx
    
    '-------------------------------
    '2015-10-05
'    For inx = 1 To no_of_ul
'        For jnx = 1 To greeks.no_of_tenors
'            shtACPricer.Range("Term_Vega_Start").Cells(jnx, inx + 1).value = greeks.term_vega(inx, jnx)
'            shtACPricer.Range("Term_Vega_Start").Cells(jnx, inx + 4).value = greeks.term_skew(inx, jnx)
'            shtACPricer.Range("Term_Vega_Start").Cells(jnx, inx + 7).value = greeks.term_conv(inx, jnx)
'        Next jnx
'    Next inx
        
    '-----------------------------------
    ' 2016-03-15
'    For inx = 1 To no_of_ul
'        For jnx = 1 To greeks.no_of_tenors
'            shtACPricer.Range("Term_Vega_Start").Cells(jnx, inx + 1).value = greeks.partial_vega(inx, jnx) / 100
'            shtACPricer.Range("Term_Vega_Start").Cells(jnx, inx + 4).value = greeks.partial_skew(inx, jnx)  '/ 100
'            shtACPricer.Range("Term_Vega_Start").Cells(jnx, inx + 7).value = greeks.partial_conv(inx, jnx) '/ 100
'        Next jnx
'    Next inx
    
    shtACPricer.Range("theta").Cells(2, 1).value = greeks.theta / 365
    shtACPricer.Range("rho").Cells(2, 1).value = greeks.rho

    Call disconnectDB(adoCon)
    
End Sub


Public Sub clear_ac_greeks()
    Dim inx As Integer
    Dim jnx As Integer
    Const max_ul As Integer = 3
    
    shtACPricer.Range("price").Cells(2, 1).ClearContents
    shtACPricer.Range("theta").Cells(2, 1).ClearContents
    shtACPricer.Range("Rho").Cells(2, 1).ClearContents
    
    For inx = 1 To max_ul
        shtACPricer.Range("delta").Cells(2 * inx, 1).ClearContents
        shtACPricer.Range("sticky_moneyness_delta").Cells(2 * inx, 1).ClearContents
        shtACPricer.Range("vega").Cells(2 * inx, 1).ClearContents
        shtACPricer.Range("skew_sensitivity").Cells(2 * inx, 1).ClearContents
        shtACPricer.Range("vanna").Cells(2 * inx, 1).ClearContents
        shtACPricer.Range("rho_ul").Cells(2 * inx, 1).ClearContents
        shtACPricer.Range("duration").Cells(2 * inx, 1).ClearContents
        shtACPricer.Range("skew").Cells(2 * inx, 1).ClearContents
        shtACPricer.Range("delta_adj").Cells(2 * inx, 1).ClearContents
        shtACPricer.Range("PriceByVolChg").Cells(2 * inx, 1).ClearContents
    Next inx
    
    For inx = 1 To max_ul
        For jnx = 1 To max_ul
            shtACPricer.Range("gamma").Cells(2 * inx, jnx).ClearContents
            shtACPricer.Range("Corr_Sens.").Cells(2 * inx, jnx).ClearContents
        Next jnx
    Next inx
    
    
    For inx = 1 To max_ul
        For jnx = 1 To shtACPricer.Range("no_of_term").value
            shtACPricer.Range("Term_Vega_Start").Cells(jnx, inx + 1).ClearContents
            shtACPricer.Range("Term_Vega_Start").Cells(jnx, inx + 4).ClearContents
            shtACPricer.Range("Term_Vega_Start").Cells(jnx, inx + 7).ClearContents
        Next jnx
    Next inx

End Sub

Public Sub clear_ac_deal()

    shtACPricer.Range("fund_code_m").ClearContents
    shtACPricer.Range("fund_code_c").ClearContents
    shtACPricer.Range("ul_code").ClearContents
    shtACPricer.Range("current_date").ClearContents
    shtACPricer.Range("value_date").ClearContents

    shtACPricer.Range("settlement_date").ClearContents
    shtACPricer.Range("notional").ClearContents
    shtACPricer.Range("issue_cost").ClearContents
'    shtACPricer.Range("current_notional").ClearContents

    shtACPricer.Range("call_put").ClearContents
    shtACPricer.Range("dummy").ClearContents
    shtACPricer.Range("ki_flag").ClearContents
    shtACPricer.Range("ki_touch_flag").ClearContents
    shtACPricer.Range("ki_barrier").ClearContents
    shtACPricer.Range("ki_barrier").Cells(1, 3).ClearContents
    shtACPricer.Range("ki_barrier").Cells(1, 4).ClearContents

    shtACPricer.Range("put_strike").ClearContents
    shtACPricer.Range("put_participation").ClearContents
    shtACPricer.Range("KI_Monitoring_Freq").ClearContents
    
    shtACPricer.Range("reference").ClearContents
    
    shtACPricer.Range("buy_sell").ClearContents
    
    
    shtACPricer.Range("alive_yn").ClearContents
    
    shtACPricer.Range("confirmed_yn").ClearContents
    
    shtACPricer.Range("Rate_Spread").ClearContents
    shtACPricer.Range("strike_shift").ClearContents
    shtACPricer.Range("KI_Barrier_Shift").ClearContents
    shtACPricer.Range("Hedge_Cost").ClearContents
    shtACPricer.Range("instrument_type").ClearContents
    shtACPricer.Range("performance_type").ClearContents
    
    If shtACPricer.Range("no_of_schedules").value > 0 Then
        shtACPricer.Range("schedule_start").Range("A1:g" & shtACPricer.Range("no_of_schedules").value).ClearContents
    End If
    
    If shtACPricer.Range("no_of_coupon_schedule").value > 0 Then
        shtACPricer.Range("cpn_schedule_start").Range("A1:D" & shtACPricer.Range("no_of_coupon_schedule").value).ClearContents
    End If
    
     If shtACPricer.Range("no_of_floating_leg").value > 0 Then
        shtACPricer.Range("floating_leg_start").Range("A1:D" & shtACPricer.Range("no_of_floating_leg").value).ClearContents
    End If

    

End Sub

Public Sub display_ac_deal(deal_ticket As clsACDealTicket)

    Dim inx As Integer
    Dim jnx As Integer
    Const max_no_of_ul As Integer = 3
    
On Error GoTo ErrorHandler

    
    shtACPricer.Range("fund_code_m").value = deal_ticket.fund_code_m
    shtACPricer.Range("fund_code_c").value = deal_ticket.fund_code_c
    shtACPricer.Range("asset_code").value = deal_ticket.asset_code
    shtACPricer.Range("No_of_Underlying").value = deal_ticket.no_of_ul
    
    shtACPricer.Range("current_date").value = deal_ticket.current_date
    shtACPricer.Range("value_date").value = deal_ticket.value_date
    
    shtACPricer.Range("settlement_date").value = deal_ticket.settlement_date
    
    shtACPricer.Range("call_put").value = get_ac_code_name("call_put_code", deal_ticket.call_put)
    shtACPricer.Range("dummy").value = deal_ticket.dummy_coupon
    
    shtACPricer.Range("ki_flag").value = get_ac_code_name("bool_code", deal_ticket.ki_barrier_flag)
    shtACPricer.Range("ki_touch_flag").value = get_ac_code_name("bool_code", deal_ticket.ki_touched_flag)
    
    For inx = 1 To deal_ticket.no_of_ul
        shtACPricer.Range("ki_barrier").Cells(1, inx).value = deal_ticket.ki_barrier(inx)
        shtACPricer.Range("ul_code").Cells(1, inx).value = deal_ticket.ul_code(inx)
    Next inx
    For inx = 1 To max_no_of_ul
        shtACPricer.Range("ki_barrier").Cells(2, inx).value = "= " & shtACPricer.Range("ki_barrier").Cells(1, inx).Address & "*" & shtACPricer.Range("reference").Cells(1, inx).Address
    Next inx
    
    shtACPricer.Range("KI_Monitoring_Freq").value = deal_ticket.ki_monitoring_freq
    
    shtACPricer.Range("floor_value").value = deal_ticket.floor_value
    shtACPricer.Range("ki_adj_pct").value = deal_ticket.ki_adj_pct
    

'-----
' KO Feature
'-----

    shtACPricer.Range("put_strike").value = deal_ticket.put_strike
    shtACPricer.Range("put_participation").value = deal_ticket.put_participation
    shtACPricer.Range("Put_Add_CPN").value = deal_ticket.put_additional_coupon
    
    For inx = 1 To deal_ticket.no_of_ul
        shtACPricer.Range("reference").Cells(1, inx).value = deal_ticket.reference_price(inx)
    Next inx
    shtACPricer.Range("notional").value = Abs(deal_ticket.notional)
    shtACPricer.Range("issue_cost").value = deal_ticket.issue_cost
    
    shtACPricer.Range("Rate_Spread").value = deal_ticket.rate_spread
    shtACPricer.Range("strike_shift").value = deal_ticket.strike_shift
    shtACPricer.Range("KI_Barrier_Shift").value = deal_ticket.ki_barrier_shift
    shtACPricer.Range("Hedge_Cost").value = deal_ticket.hedge_cost
    
'    If deal_ticket.notional > 0 Then
'        shtACPricer.Range("buy_sell").value = "BUY"
'    Else
'        shtACPricer.Range("buy_sell").value = "SELL"
'    End If
    
'    shtacpricer.Range("comment").value = deal_ticket.comment
    
    shtACPricer.Range("txtComment").value = deal_ticket.comment
    
    For inx = 1 To deal_ticket.no_of_schedule
        
        shtACPricer.Range("schedule_start").Cells(inx, 1).value = deal_ticket.autocall_schedules(inx).call_date
        
        For jnx = 1 To deal_ticket.autocall_schedules(inx).no_of_jumps
        
            shtACPricer.Range("schedule_start").Cells(inx, 4 + jnx).value = deal_ticket.autocall_schedules(inx).percent_strike(jnx)
            shtACPricer.Range("schedule_start").Cells(inx, 1 + jnx).value = deal_ticket.autocall_schedules(inx).coupon_on_call(jnx)
        
        Next jnx
'        For jnx = 1 To deal_ticket.no_of_ul
'            shtACPricer.Range("schedule_start").Cells(inx, jnx + 3).value = "=" & shtACPricer.Range("schedule_start").Cells(inx, 3).Address & " * " & shtACPricer.Range("reference").Cells(1, jnx).Address
'        Next jnx
    
    Next inx
    
    If deal_ticket.alive_yn = "" Then
        shtACPricer.Range("alive_yn").value = "N"
    Else
        shtACPricer.Range("alive_yn").value = deal_ticket.alive_yn
    End If
    
    If deal_ticket.confirmed_yn = "" Then
        shtACPricer.Range("confirmed_yn").value = "N"
    Else
        shtACPricer.Range("confirmed_yn").value = deal_ticket.confirmed_yn
    End If
    
    shtACPricer.Range("Monthly_Coupon_Flag").value = get_ac_code_name("bool_code", deal_ticket.monthly_coupon_flag)
    
    shtACPricer.Range("instrument_type").value = get_ac_code_name("Instrument_type_code", deal_ticket.instrument_type)
    shtACPricer.Range("performance_type").value = get_ac_code_name("performance_type_code", deal_ticket.performance_type)
    
    
    If deal_ticket.monthly_coupon_flag = 1 Then
    
        For inx = 1 To deal_ticket.no_of_coupon_schedule
            
            shtACPricer.Range("cpn_schedule_start").Cells(inx, 1).value = deal_ticket.monthly_coupon_schedules(inx).call_date
            shtACPricer.Range("cpn_schedule_start").Cells(inx, 2).value = deal_ticket.monthly_coupon_schedules(inx).percent_strike
            shtACPricer.Range("cpn_schedule_start").Cells(inx, 3).value = deal_ticket.monthly_coupon_schedules(inx).coupon_on_call
            
            For jnx = 1 To deal_ticket.no_of_ul
                shtACPricer.Range("cpn_schedule_start").Cells(inx, 3 + jnx).value = "=" & Chr(65 + shtACPricer.Range("cpn_schedule_start").Column) & (shtACPricer.Range("cpn_schedule_start").Row + inx - 1) & " * " & shtACPricer.Range("reference").Cells(1, jnx).Address
            Next jnx
        
        Next inx
        
        shtACPricer.Range("Monthly_Cpn").value = deal_ticket.monthly_coupon_amount
        
    End If
    
    If deal_ticket.no_of_floating_coupon_schedule > 0 Then
    
        For inx = 1 To deal_ticket.no_of_floating_coupon_schedule
        
            shtACPricer.Range("floating_leg_start").Cells(inx, 1).value = deal_ticket.floating_coupon_dates()(inx)
            shtACPricer.Range("floating_leg_start").Cells(inx, 2).value = deal_ticket.floating_fixing_values()(inx)
        
        Next inx
    
    End If
    
    
    shtACPricer.Range("Range_Accrual_Flag").value = get_ac_code_name("bool_code", deal_ticket.ra_flag)
    
     shtACPricer.Range("ra_cpn").value = deal_ticket.ra_cpn
     shtACPricer.Range("ra_tenor").value = get_ac_code_name("tenor_code", deal_ticket.ra_tenor)
     shtACPricer.Range("ra_min").value = deal_ticket.ra_min_percent
     shtACPricer.Range("ra_max").value = deal_ticket.ra_max_percent
     

    Exit Sub
    
ErrorHandler:

    raise_err "display_ac_deal"
        

'    Err.Raise vbObjectError + 1000, "display_a_deal:" & Chr(13) & Chr(13) & Err.source, Err.description

End Sub

'--------------------------------
' Modified on
' 2013-10-16
' 2013-10-21
' 2013-10-23
'--------------------------------


'Private Function get_range_cnt(from_date As Date, to_date As Date, ticker As String, range_min As Double, range_max As Double) As Integer
'
'    Dim no_of_dates As Integer
'    Dim base_dates() As Date
'    Dim prices() As Double
'    Dim inx As Integer
'    Dim rtn_value As Integer
'
'On Error GoTo ErrorHandler
'
'    DBConnector
'
'    rtn_value = 0
'
'    no_of_dates = retrieve_bl_history(base_dates, prices, ticker, from_date, to_date)
'
'    For inx = 1 To no_of_dates
'
'        If prices(inx) >= range_min And prices(inx) <= range_max Then
'
'            rtn_value = rtn_value + 1
'
'        End If
'
'    Next inx
'
'    DBDisConnector
'
'    get_range_cnt = rtn_value
'
'    Exit Function
'
'ErrorHandler:
'
'    DBDisConnector
'
'    raise_err Err.description, "get_range_cnt"
'
'
'End Function
'
'Private Function get_accrued_cpn(deal_ticket As clsACDealTicket, current_date As Date, rate_curve As clsRateCurve) As Double
'
'    Dim total_days As Integer
'    Dim next_schedule As clsAutocallSchedule
'    Dim prev_schedule As clsAutocallSchedule
'
'    Dim from_date As Date
'    Dim to_date As Date
'
'    Dim rtn_value As Double
'
'    Dim ticker As String
'
'On Error GoTo ErrorHandler
'
'    If deal_ticket.get_next_schedule(next_schedule, current_date + 1) Then
'        to_date = next_schedule.call_date
'    Else
'        raise_err "get_accrued_cpn", "No Schedule"
'    End If
'
'    If deal_ticket.get_prev_schedule(prev_schedule, current_date + 1) Then
'        from_date = prev_schedule.call_date
'    Else
'        from_date = deal_ticket.value_date
'    End If
'
'    If deal_ticket.ul_code() = "KOSPI200" Then
'        ticker = "KOSPI2"
'    Else
'        ticker = deal_ticket.ul_code()
'    End If
'
'    rtn_value = get_range_cnt(from_date + 1, current_date, ticker, deal_ticket.ra_min_percent * deal_ticket.reference_price(), deal_ticket.ra_max_percent * deal_ticket.reference_price()) _
'              / business_days_between(from_date + 1, to_date) _
'              * deal_ticket.ra_cpn _
'              * rate_curve.get_discount_factor(current_date, to_date)
'
'    get_accrued_cpn = rtn_value
'
'    Exit Function
'
'ErrorHandler:
'
'    raise_err "get_accrued_cpn", Err.description
'
'End Function






Public Sub ac_deal_to_midday(ac_deals() As clsACDealTicket, ByVal no_of_deals As Integer)

    Dim inx As Integer
    
    For inx = 1 To no_of_deals
    
        ac_deals(inx).mid_day_greek = True
        ac_deals(inx).current_date = ac_deals(inx).current_date - 1
        ac_deals(inx).current_date_origin_ = ac_deals(inx).current_date_origin_ - 1
    
    Next inx


End Sub
Public Sub run_ac_pricing_1d(ByRef the_greeks As clsGreeks _
                        , deal_ticket As clsACDealTicket _
                        , market As clsMarketSet _
                        , ByRef bump_greek_set As clsGreekSet _
                        , Optional calc_vega As Boolean = False _
                        , Optional calc_skew_s As Boolean = False _
                        , Optional calc_corr As Boolean = False _
                        , Optional calc_rho As Boolean = False _
                        , Optional snapshot_time As Double = 0.001 _
                        , Optional ignore_smoothing As Boolean = False _
                        , Optional bump_delta As Boolean = False)

    Dim holiday_list__(0 To 0) As Long
    holiday_list__(0) = 42000
    
    Dim theNote As Object
    Dim theEngine As Object
    Dim rTS As Object
    Dim qTs As Object
    Dim fTs As Object
    Dim volTs As Object
    Dim quantoHelper As Object
    Dim process As Object
    Dim rate_spread As Double
    
    Dim day_shift As Long
    Dim ul_prices(1) As Double
    
    Dim vegas(1 To 1) As Double
    Dim vannas(1 To 1) As Double
    Dim skew_s(1 To 1) As Double
    Dim deltas(1 To 1) As Double
    
    Dim rho_ul(1 To 1) As Double
    
    Dim partial_vegas() As Double
    
    Dim rho As Double
    
    Dim origin_pl_currency_curve As clsRateCurve
    
On Error GoTo ErrorHandler

    Set theNote = ac_deal_ticket_to_clr(deal_ticket)
    
    If deal_ticket.instrument_type = 0 Then
        rate_spread = deal_ticket.rate_spread
    End If
    
    day_shift = deal_ticket.current_date - market.pl_currency_rate_curve_.rate_dates()(0)
    
'    Dim tmp As Double
'    tmp = theNote.testAutocallStrike
'
    'Set rTs = New ql_handle_YieldTermStructure
    'Set fTs = New ql_handle_YieldTermStructure
    'Set qTs = New ql_handle_YieldTermStructure
    'Set volTs = New ql_handle_BlackVarianceSurface
    
    Set rTS = get_rTs(market, rate_spread, day_shift)
'    If rTs.initializeTs(market.pl_currency_rate_curve_.rate_dates(), market.pl_currency_rate_curve_.dcfs(), rate_spread) = 0 Then
'        raise_err "run_ac_pricing_1d", "Failed to generate rate term Structure"
'    End If
    
    Set qTs = get_qTs(market, deal_ticket.ul_code(), rate_spread, day_shift)
    
'    If qTs.initializeFlat(market.market_by_ul(deal_ticket.ul_code()).div_yield_ + rate_spread) = 0 Then
'        raise_err "run_ac_pricing_1d", "Failed to generate div term Structure"
'    End If
    
    Set fTs = get_fTs(market, deal_ticket.ul_code(), rate_spread, day_shift)
    
'    If fTs.initializeTs(market.market_by_ul(deal_ticket.ul_code()).rate_curve_.rate_dates(), market.market_by_ul(deal_ticket.ul_code()).rate_curve_.dcfs(), rate_spread) = 0 Then
'        raise_err "run_ac_pricing_1d", "Failed to generate rate term Structure"
'    End If
    
    Set quantoHelper = New ql_shared_ptr_FdmQuantoHelper
    quantoHelper.initializeQH rTS _
                            , fTs _
                            , market.market_by_ul(deal_ticket.ul_code()).ul_currency_vol _
                            , market.correlation_pair_.get_corr(deal_ticket.ul_code(), market.market_by_ul(deal_ticket.ul_code()).ul_currency & "KRW")
                            
    Set volTs = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(), day_shift)
    
'    If volTs.initializeVS(deal_ticket.current_date _
'                      , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.grid_.get_all_dates_as_long(0, day_shift) _
'                      , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.grid_.get_all_strikes(0) _
'                      , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.get_vol_surface(0) _
'                      ) = 0 Then
'        raise_err "run_ac_pricing_1d", "Failed to generate vol term Structure"
'    End If

    Set process = New ql_shared_ptr_blackScholesMertonProcess
    process.initializeProcess market.market_by_ul(deal_ticket.ul_code()).s_, qTs, rTS, volTs
    
    Set theEngine = New sy_shared_ptr_FdAutocallableEngine1D
    
'    theEngine.initializeEngine process _
'                             , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
'                             , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , deal_ticket.strike_shift _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
'   <----delete deal_ticket.strike_shift---->
'    theEngine.initializeEngine process _
'                             , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
'                             , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
    Dim ratioDividendIn(1) As Double
    ratioDividendIn(0) = 0
    theEngine.initializeEngine process _
                             , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
                             , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
                             , deal_ticket.t_grid, deal_ticket.x_grid _
                             , deal_ticket.scheme_type _
                             , True _
                             , quantoHelper _
                             , snapshot_time _
                             , ratioDividendIn
                             
    theNote.setPricingEngine theEngine
    
    the_greeks.value = theNote.NPV() * deal_ticket.notional
    the_greeks.delta = theNote.delta()(0) * deal_ticket.notional
    the_greeks.gamma = theNote.gamma()(0) * deal_ticket.notional
    the_greeks.theta = theNote.theta() * deal_ticket.notional
    the_greeks.ul_price = market.market_by_ul(deal_ticket.ul_code(1)).s_
    Dim tmp_delta() As Double
    Dim tmp_gamma() As Double
    tmp_delta = theNote.delta()
    tmp_gamma = theNote.gamma()
    the_greeks.set_all_deltas tmp_delta, 0, deal_ticket.notional
    the_greeks.set_all_gammas tmp_gamma, 0, deal_ticket.notional
    
    '------------------------
    ' 2015-01-05
    '------------------------
    ul_prices(1) = market.market_by_ul(deal_ticket.ul_code(1)).s_
    the_greeks.set_all_ul_prices ul_prices
    
    If snapshot_time >= 0 Then
        Dim tmp_xAxis() As Double
        Dim tmp_snapshotValues() As Double
    
        tmp_xAxis = theNote.xAxis()
        tmp_snapshotValues = theNote.snapShotValues()
        
        the_greeks.set_xAxis tmp_xAxis, True, deal_ticket.reference_price(1)
        the_greeks.set_snapshot_value tmp_snapshotValues, Abs(deal_ticket.notional)
    End If
    
    '======================================================
    ' VEGA
    '======================================================

    If calc_vega Then
    
        market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.bump_vol_surface 0.01
        
        Set volTs = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(), day_shift)
'        If volTs.initializeVS(deal_ticket.current_date _
'                          , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.grid_.get_all_dates_as_long(0) _
'                          , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.grid_.get_all_strikes(0) _
'                          , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.get_vol_surface(0) _
'                          ) = 0 Then
'            raise_err "run_ac_pricing_1d", "Failed to generate vol term Structure"
'        End If

        process.initializeProcess market.market_by_ul(deal_ticket.ul_code()).s_, qTs, rTS, volTs
        
'        theEngine.initializeEngine process _
'                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date) _
'                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date) _
'                                 , deal_ticket.t_grid, deal_ticket.x_grid, deal_ticket.scheme_type, deal_ticket.strike_shift, True, quantoHelper, snapshot_time
'   <----delete deal_ticket.strike_shift---->
'        theEngine.initializeEngine process _
'                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date) _
'                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date) _
'                                 , deal_ticket.t_grid, deal_ticket.x_grid, deal_ticket.scheme_type, True, quantoHelper, snapshot_time
        ratioDividendIn(0) = 0
        theEngine.initializeEngine process _
                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
                                 , deal_ticket.t_grid, deal_ticket.x_grid _
                                 , deal_ticket.scheme_type _
                                 , True _
                                 , quantoHelper _
                                 , snapshot_time _
                                 , ratioDividendIn
                             
        theNote.setPricingEngine theEngine
        
        the_greeks.vega = theNote.NPV() * deal_ticket.notional - the_greeks.value
        the_greeks.vanna = theNote.delta()(0) * deal_ticket.notional - the_greeks.delta
        
        the_greeks.vega = the_greeks.vega
        the_greeks.vanna = the_greeks.vanna
        
        market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.rewind_vol_bump
        
        vegas(1) = the_greeks.vega
        vannas(1) = the_greeks.vanna
        
        the_greeks.set_all_vegas vegas
        the_greeks.set_all_vannas vannas
        
    End If
    
'    If calc_term_vega Then
'    End If
    
    If calc_skew_s Then
    
        market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.bump_skew -0.01
        
        Set volTs = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(), day_shift)
'        If volTs.initializeVS(deal_ticket.current_date _
'                          , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.grid_.get_all_dates_as_long(0) _
'                          , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.grid_.get_all_strikes(0) _
'                          , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.get_vol_surface(0) _
'                          ) = 0 Then
'            raise_err "run_ac_pricing_1d", "Failed to generate vol term Structure"
'        End If
    
        process.initializeProcess market.market_by_ul(deal_ticket.ul_code()).s_, qTs, rTS, volTs
        
'        theEngine.initializeEngine process _
'                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date) _
'                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date) _
'                                 , deal_ticket.t_grid, deal_ticket.x_grid, deal_ticket.scheme_type, deal_ticket.strike_shift, True, quantoHelper, snapshot_time
'   <----delete deal_ticket.strike_shift---->
'        theEngine.initializeEngine process _
'                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date) _
'                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date) _
'                                 , deal_ticket.t_grid, deal_ticket.x_grid, deal_ticket.scheme_type, True, quantoHelper, snapshot_time
        ratioDividendIn(0) = 0
        theEngine.initializeEngine process _
                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
                                 , deal_ticket.t_grid, deal_ticket.x_grid _
                                 , deal_ticket.scheme_type _
                                 , True _
                                 , quantoHelper _
                                 , snapshot_time _
                                 , ratioDividendIn
        
        theNote.setPricingEngine theEngine
        
        the_greeks.skew_s = (theNote.NPV() * deal_ticket.notional - the_greeks.value) * -1
        
        market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.rewind_vol_bump
        
        the_greeks.set_all_skews skew_s
        
    End If

    If calc_rho Then
    
        Dim inx As Integer
        Dim tmp_greek As clsGreeks
        Dim tmp_rate_curves() As clsRateCurve
        ReDim tmp_rate_curves(1 To deal_ticket.no_of_ul) As clsRateCurve
     '   Dim bump_greek_set As clsGreekSet

        Set tmp_greek = New clsGreeks
        Set origin_pl_currency_curve = market.pl_currency_rate_curve_.copy_obj()
        'Bumping DCF: +10bp
        Set market.pl_currency_rate_curve_ = market.pl_currency_rate_curve_.copy_obj(0, 0.001)
        
        '---
        run_ac_pricing_1d tmp_greek, deal_ticket, market, bump_greek_set, False, False, False, False, snapshot_time, False, False
        'Rho(DCF)
        rho = (tmp_greek.value - the_greeks.value) * 0.1
        
        Set market.pl_currency_rate_curve_ = origin_pl_currency_curve
        
        'Rho(UL)
        'Bumping riskfree rate curve by adjusting dividend yield: +10bp
        For inx = 1 To deal_ticket.no_of_ul
            
            Set tmp_greek = New clsGreeks
            Dim tmp_div_yield As Double
            tmp_div_yield = market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_
            
            market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_ = tmp_div_yield - 0.001
            
            '---
            run_ac_pricing_1d tmp_greek, deal_ticket, market, bump_greek_set, False, False, False, False, snapshot_time, False, False
            
            rho_ul(inx) = (tmp_greek.value - the_greeks.value) * 0.1
            
            market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_ = tmp_div_yield
            
        Next inx
        
        the_greeks.rho = rho
        the_greeks.set_all_rho_ul rho_ul
        
    End If
    
'    If bump_delta Then
'
'        market.bump_ul_price 0.001, deal_ticket.ul_code(1)
'        market.market_by_ul(deal_ticket.ul_code(1)).sabr_surface_.shift_surface market.market_by_ul(deal_ticket.ul_code(1)).s_
'
'        Set volTs = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(), day_shift)
''        If volTs.initializeVS(deal_ticket.current_date _
''                             , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.grid_.get_all_dates_as_long(0) _
''                             , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.grid_.get_all_strikes(0) _
''                             , market.market_by_ul(deal_ticket.ul_code()).sabr_surface_.local_vol_surface.get_vol_surface(0) _
''                             ) = 0 Then
''               raise_err "run_ac_pricing_1d", "Failed to generate vol term Structure"
''        End If
'
'         process.initializeProcess market.market_by_ul(deal_ticket.ul_code()).s_, qTs, rTs, volTs
'
'         theEngine.initializeEngine process _
'                                  , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date) _
'                                  , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date) _
'                                  , deal_ticket.t_grid, deal_ticket.x_grid, deal_ticket.scheme_type, deal_ticket.strike_shift, True, quantoHelper, snapshot_time
'
'         theNote.setPricingEngine theEngine
'
'         the_greeks.sticky_strike_delta = (theNote.NPV() * deal_ticket.notional - the_greeks.value) / (0.001 * market.market_by_ul(deal_ticket.ul_code(1)).s_)
'
'    End If

'    If deal_ticket.ra_flag Then
'        the_greeks.value = the_greeks.value + get_accrued_cpn(deal_ticket, deal_ticket.current_date, market.market().rate_curve_) * deal_ticket.notional
'    End If
'
'    If snapshot_time > 0 Then 'snapshot_time > 0 Then
'        write_closing_file_1d "C:\log\", deal_ticket.asset_code & "." & config__.snapshot_file_extension, the_greeks.get_xAxis, the_greeks.get_snapshot_value
'    End If
'
'    theNote.dispose_com
    
    Exit Sub
    
ErrorHandler:

    raise_err "run_ac_pricing_1d", Err.description

End Sub


'Public Function get_rTs(market As clsMarketSet, spread As Double, Optional day_shift As Long = 0) As Object
'
'    Dim rTs As Object
'
'On Error GoTo ErrorHandler
'
'    Set rTs = New ql_handle_YieldTermStructure
'
'    If rTs.initializeTs(market.pl_currency_rate_curve_.rate_dates_shifted(day_shift), market.pl_currency_rate_curve_.dcfs(), spread) = 0 Then
'            raise_err "get_rTs", "Failed to generate rate term Structure"
'    End If
'
'    Set get_rTs = rTs
'
'    Exit Function
'
'ErrorHandler:
'
'    raise_err "get_rTs", Err.description
'
'End Function

Public Function get_rTs(pl_currency_rate_curve As clsRateCurve, spread As Double, Optional day_shift As Long = 0) As Object

    Dim rTS As Object
    
On Error GoTo ErrorHandler
    
    Set rTS = New ql_handle_YieldTermStructure
    
    If rTS.initializeTs(pl_currency_rate_curve.rate_dates_shifted(day_shift), pl_currency_rate_curve.dcfs(), spread) = 0 Then
            raise_err "get_rTs", "Failed to generate rate term Structure"
    End If
    
    Set get_rTs = rTS
    
    Exit Function
    
ErrorHandler:

    raise_err "get_rTs", Err.description

End Function

Public Function get_fTs(market As clsMarketSet, ul_code As String, ByVal spread As Double, Optional day_shift As Long = 0) As Object

    Dim fTs As Object
    Dim theMarket As clsMarket
    
On Error GoTo ErrorHandler
    
    Set theMarket = market.market_by_ul(ul_code)
    Set fTs = New ql_handle_YieldTermStructure
    
    '<<<<<<<=======================================
    If spread < 0 Then
        spread = 0
    End If
    
    If fTs.initializeTs(theMarket.rate_curve_.rate_dates_shifted(day_shift), theMarket.rate_curve_.dcfs(), spread) = 0 Then
            raise_err "get_fTs", "Failed to generate rate term Structure"
    End If
    
    Set get_fTs = fTs
    
    Exit Function
    
ErrorHandler:

    raise_err "get_fTs", Err.description

End Function

Public Function get_aTs(market As clsMarketSet, ul_code As String, ByVal spread As Double, Optional day_shift As Long = 0) As Object

    Dim aTs As Object
    Dim theMarket As clsMarket
    
On Error GoTo ErrorHandler
    
    Set theMarket = market.market_by_ul(ul_code)
    Set aTs = New ql_handle_YieldTermStructure
    
    '<<<<<<<=======================================
    If spread < 0 Then
        spread = 0
    End If
    
    If aTs.initializeTs(theMarket.drift_adjust_.rate_dates_shifted(day_shift), theMarket.drift_adjust_.dcfs(), spread) = 0 Then
        raise_err "get_aTs", "Failed to generate rate term Structure"
    End If
    
    Set get_aTs = aTs
    
    Exit Function
    
ErrorHandler:

    raise_err "get_aTs", Err.description

End Function

Public Function get_qTs(market As clsMarketSet, ul_code As String, spread As Double, Optional day_shift As Long = 0) As Object

    Dim qTs As Object
    Dim theMarket As clsMarket
    
On Error GoTo ErrorHandler
    
    Set theMarket = market.market_by_ul(ul_code)
    Set qTs = New ql_handle_YieldTermStructure
    
    If qTs.initializeFlat(theMarket.div_yield_ + spread) = 0 Then
            raise_err "get_qTs", "Failed to generate rate term Structure"
    End If
    
    Set get_qTs = qTs
    
    Exit Function
    
ErrorHandler:

    raise_err "get_qTs", Err.description

End Function

'Public Function get_quantoHelper(rTs As Object, market As clsMarketSet, ul_code As String, spread As Double) As Object
Public Function get_quantoHelper(rTS As Object, market As clsMarketSet, ul_code As String, deal_ccy As String, spread As Double, day_shift As Long) As Object '2018.7.10
    Dim quantoHelper As Object
    Dim theMarket As clsMarket
    Dim fTs As Object
    
On Error GoTo ErrorHandler
    
    Set theMarket = market.market_by_ul(ul_code)
    Set quantoHelper = New ql_shared_ptr_FdmQuantoHelper
    'Set fTs = get_fTs(market, ul_code, spread)
    Set fTs = get_fTs(market, ul_code, spread, day_shift) '2018.7.10
    
    'If quantoHelper.initializeQH(rTs, fTs, theMarket.ul_currency_vol, market.correlation_pair_.get_corr(ul_code, theMarket.ul_currency & "KRW")) = 0 Then
    If quantoHelper.initializeQH(rTS, fTs, theMarket.ul_currency_vol, market.correlation_pair_.get_corr(ul_code, theMarket.ul_currency & deal_ccy)) = 0 Then
                            
        raise_err "get_quantoHelper", "Failed to generate rate term Structure"
                            
    End If
        
    Set get_quantoHelper = quantoHelper
    
    Exit Function
    
ErrorHandler:

    raise_err "get_quantoHelper", Err.description

End Function

Public Function get_vol_surface_(current_date As Date, market As clsMarketSet, ul_code As String, Optional ByVal day_shift As Long = 0) As Object
    
    Dim volTs As Object
    Dim theMarket As clsMarket
    
On Error GoTo ErrorHandler

    Set theMarket = market.market_by_ul(ul_code)
    Set volTs = New ql_handle_BlackVarianceSurface
          
    'day_shift = day_shift + market.pl_currency_rate_curve_.rate_dates()(0) - theMarket.sabr_surface_.eval_date
      
    If volTs.initializeVS(current_date _
                      , theMarket.sabr_surface_.local_vol_surface.grid_.get_all_dates_as_long(0, day_shift, current_date) _
                      , theMarket.sabr_surface_.local_vol_surface.grid_.get_all_strikes(0) _
                      , theMarket.sabr_surface_.local_vol_surface.get_vol_surface(0) _
                      ) = 0 Then
        raise_err "get_vol_surface_", "Failed to generate vol term Structure"
        
    End If
    
    Set get_vol_surface_ = volTs
    
    Exit Function
    
ErrorHandler:

    raise_err "get_vol_surface_", Err.description

End Function


Public Function get_process(rTS As Object, qTs As Object, volTs As Object, market As clsMarketSet, ul_code As String) As Object
    
    Dim process As Object
    Dim theMarket As clsMarket
    
On Error GoTo ErrorHandler

    Set theMarket = market.market_by_ul(ul_code)
    Set process = New ql_shared_ptr_blackScholesMertonProcess
      
    If process.initializeProcess(theMarket.s_, qTs, rTS, volTs) = 0 Then
        raise_err "get_process", "Failed to generate vol term Structure"
        
    End If
    
    Set get_process = process
    
    Exit Function
    
ErrorHandler:

    raise_err "get_process", Err.description

End Function

'Public Function get_div_process(rTs As Object, qTs As Object, fTs As Object, volTs As Object, market As clsMarketSet, ul_code As String, ccy As String, ratioLeverage As Double) As Object
'day_shift 추가: 2018.12.05
'Public Function get_div_process(rTs As Object, qTs As Object, fTs As Object, volTs As Object, market As clsMarketSet, ul_code As String, ccy As String, ratioLeverage As Double, Optional day_shift As Long = 0) As Object
'double ratio_leverage (1.0: 일반 상품, 2.0: 2배 레버리지, default=1.0) -> int code_leverage (0: 일반 상품, 1: 2배 레버리지, 2: 1.5배 레버리지, default=0) 로 변경 : 2020.6.19
'Public Function get_div_process(rTs As Object, qTs As Object, fTs As Object, volTs As Object, market As clsMarketSet, ul_code As String, CCY As String, Optional codeLeverage As Integer = 0, Optional day_shift As Long = 0) As Object
'drift adjustment aTs 추가: 2023.11.21
Public Function get_div_process(rTS As Object, qTs As Object, fTs As Object, aTs As Object, volTs As Object, market As clsMarketSet, ul_code As String, ccy As String, Optional codeLeverage As Integer = 0, Optional day_shift As Long = 0) As Object

    Dim process As Object
    Dim theMarket As clsMarket
    Dim div_dates() As Long
    Dim divs() As Double
    Dim div_size() As Long
    Dim ul_code_arr(1 To 1) As String
    
    ul_code_arr(1) = ul_code
    
On Error GoTo ErrorHandler

    Set theMarket = market.market_by_ul(ul_code)
    Set process = New sy_shared_ptr_dividendBSMProcess
    
    'get_dividend_array div_dates, divs, div_size, ul_code_arr, market
    'day_shift 추가: 2018.12.05
    get_dividend_array div_dates, divs, div_size, ul_code_arr, market, day_shift
    
        'tmp_fx_vol(inx) = market.market(index_seq(inx)).ul_currency_vol
        'tmp_fx_corr(inx) = market.correlation_pair_.get_corr(deal_ticket.ul_code(inx), market.market(index_seq(inx)).ul_currency & "KRW")
      
'    If process.initializeProcess(theMarket.s_ _
'                               , qTs _
'                               , rTs _
'                               , volTs _
'                               , div_dates _
'                               , divs _
'                               , theMarket.ul_currency_vol _
'                               , market.correlation_pair_.get_corr(ul_code, theMarket.ul_currency & "KRW") _
'                               , fTs _
'                               ) = 0
'    Dim ratioDividend As Double
'    ratioDividend = 0
'    Dim refPriceForDividend As Double
'    refPriceForDividend = market.market_by_ul(ul_code).s_
'    Dim ratioLeverage As Double
'    ratioLeverage = 1
    If process.initializeProcess(theMarket.s_ _
                               , qTs _
                               , rTS _
                               , fTs _
                               , aTs _
                               , volTs _
                               , div_dates _
                               , divs _
                               , theMarket.ul_currency_vol(get_dcf_idx(ccy)) _
                               , market.correlation_pair_.get_corr(ul_code, theMarket.ul_currency & ccy) _
                               , theMarket.div_schedule_.ratioDividend _
                               , theMarket.refPriceForDividend _
                               , codeLeverage _
                               ) = 0 _
    Then
                               
        raise_err "get_process", "Failed to generate div process"
        
    End If
    
    Set get_div_process = process
    
    Exit Function
    
ErrorHandler:

    raise_err "get_process", Err.description

End Function
'
'
''---For MC
'Public Function get_div_process_array(rTs As Object, qTs As Object, volTs As Object, market As clsMarketSet, ul_code As String) As Object
'
'    Dim process As Object
'    Dim theMarket As clsMarket
'    Dim div_dates() As Long
'    Dim divs() As Double
'    Dim div_size() As Long
'    Dim ul_code_arr(1 To 1) As String
'
'    ul_code_arr(1) = ul_code
'
'On Error GoTo ErrorHandler
'
'    Set theMarket = market.market_by_ul(ul_code)
'    Set process = New sy_shared_ptr_dividendBSMProcess
'
'    get_dividend_array div_dates, divs, div_size, ul_code_arr, market
'
'        'tmp_fx_vol(inx) = market.market(index_seq(inx)).ul_currency_vol
'        'tmp_fx_corr(inx) = market.correlation_pair_.get_corr(deal_ticket.ul_code(inx), market.market(index_seq(inx)).ul_currency & "KRW")
'
'    If process.initializeProcess(theMarket.s_, qTs, rTs, volTs, div_dates, divs, theMarket.ul_currency_vol, market.correlation_pair_.get_corr(ul_code, theMarket.ul_currency & "KRW")) = 0 Then
'        raise_err "get_process", "Failed to generate vol term Structure"
'
'    End If
'
'    Set get_div_process = process
'
'    Exit Function
'
'ErrorHandler:
'
'    raise_err "get_process", Err.description
'
'End Function


'Public Sub get_dividend_array(div_dates() As Long, divs() As Double, div_size() As Long, ul_code() As String, market As clsMarketSet)
'day_shift 추가: 2018.12.05
Public Sub get_dividend_array(div_dates() As Long, divs() As Double, div_size() As Long, ul_code() As String, market As clsMarketSet, Optional day_shift As Long = 0)

    Dim no_of_ul As Integer
    Dim inx As Integer
    Dim jnx As Integer
    Dim tmp_divs() As Double
    Dim tmp_div_dates() As Long
    Dim div_size_counter As Long
    
On Error GoTo ErrorHandler

    
    no_of_ul = get_array_size_string(ul_code)
    
    For inx = 1 To no_of_ul
        
'        If Not market.market_by_ul(ul_code(inx)).div_schedule_ Is Nothing Then
        
            tmp_divs = market.market_by_ul(ul_code(inx)).div_schedule_.get_divs(, , ul_code(inx))
            'tmp_div_dates = market.market_by_ul(ul_code(inx)).div_schedule_.get_div_dates(, , ul_code(inx))
            'day_shift 추가: 2018.12.05
            tmp_div_dates = market.market_by_ul(ul_code(inx)).div_schedule_.get_div_dates(, , ul_code(inx), day_shift)
            'push_back_long div_size, get_array_size_double(tmp_divs), 0
            div_size_counter = 0
             
            For jnx = 0 To get_array_size_double(tmp_divs) - 1
               
                If tmp_div_dates(jnx) > market.market_by_ul(ul_code(inx)).rate_curve_.rate_dates()(0) Then
                    div_size_counter = div_size_counter + 1
                    push_back_long div_dates, tmp_div_dates(jnx), 0
                    push_back_double divs, tmp_divs(jnx), 0
                    
                End If
            
            Next jnx
            
            push_back_long div_size, div_size_counter, 0
        
            Erase tmp_divs
            Erase tmp_div_dates
            
'        Else
'
'            push_back_long div_dates, 73415, 0
'            push_back_double divs, 0, 0
'            push_back_long div_size, 1, 0
'
'        End If
        
    Next inx
    

    Exit Sub
    
ErrorHandler:

    raise_err "get_dividend_array", Err.description
    
End Sub

Public Sub run_ac_pricing_2d(ByRef the_greeks As clsGreeks _
                        , deal_ticket As clsACDealTicket _
                        , ByVal market As clsMarketSet _
                        , ByRef bump_greek_set As clsGreekSet _
                        , Optional calc_vega As Boolean = False _
                        , Optional calc_skew_s As Boolean = False _
                        , Optional calc_corr As Boolean = False _
                        , Optional calc_rho As Boolean = False _
                        , Optional snapshot_time As Double = 0.001 _
                        , Optional ignore_smoothing As Boolean = False _
                        , Optional log_file As Boolean = False _
                        , Optional bump_delta As Boolean = False _
                        , Optional calc_delta As Boolean = False _
                        , Optional calc_term_vega As Boolean = False)

    Dim holiday_list__(0 To 0) As Long
    holiday_list__(0) = 42000
    
    Dim theNote As Object
    Dim theEngine As Object
    Dim rTS As Object
    Dim qTs(0 To 1) As Object
    Dim fTs(0 To 1) As Object
    Dim volTs(0 To 1) As Object
    Dim quantoHelper(0 To 1) As Variant
    Dim process(0 To 1) As Variant
    Dim rate_spread As Double
    
    Dim day_shift As Long
    Dim ul_prices(1 To 2) As Double '---- 2015-01-05
    
    Dim vegas(1 To 2) As Double
    Dim vannas(1 To 2) As Double
    Dim skew_s(1 To 2) As Double
    Dim deltas(1 To 2) As Double
    Dim corr_sens(1 To 2, 1 To 2) As Double
    
    Dim rho_ul(1 To 2) As Double
    
    Dim partial_vegas() As Double ' ---- 2015-10-05
    
    Dim inx As Integer
    Dim jnx As Integer
    
    Dim divs() As Double
    Dim div_dates() As Long
    Dim div_array_size() As Long
    
    Dim rho As Double
    
    Dim backup_market As clsMarketSet
    
    Const no_of_ul As Integer = 2
   
    Dim origin_pl_currency_curve As clsRateCurve
    
On Error GoTo ErrorHandler
    
    Set theNote = ac_deal_ticket_to_clr_2d(deal_ticket)
    
    If deal_ticket.instrument_type = 0 Then
        rate_spread = deal_ticket.rate_spread
    End If
    
    day_shift = deal_ticket.current_date - market.pl_currency_rate_curve_.rate_dates()(0)
    
    Set rTS = get_rTs(market, rate_spread, day_shift) ' DC Curve
        
    For inx = 0 To no_of_ul - 1
        Set qTs(inx) = get_qTs(market, deal_ticket.ul_code(inx + 1), rate_spread)
        Set quantoHelper(inx) = get_quantoHelper(rTS, market, deal_ticket.ul_code(inx + 1), rate_spread)
        Set volTs(inx) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx + 1), day_shift)
        
        Set process(inx) = get_process(rTS, qTs(inx), volTs(inx), market, deal_ticket.ul_code(inx + 1))
    Next inx
            
    get_dividend_array div_dates, divs, div_array_size, deal_ticket.get_ul_codes(), market
    
    Set theEngine = New sy_shared_ptr_FdAutocallableEngine2D
    
'    theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , deal_ticket.strike_shift _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
'   <----delete deal_ticket.strike_shift---->
'    theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
' dll: 2017.11.28
    Dim ratioDividendIn(1) As Double
    ratioDividendIn(0) = 0
    ratioDividendIn(1) = 0
'    theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time _
'                             , ratioDividendIn
' dll: 2018.5.3
    Dim div_refPriceForDividend(1) As Double
    div_refPriceForDividend(0) = market.market_by_ul(deal_ticket.ul_code(1)).s_
    div_refPriceForDividend(1) = market.market_by_ul(deal_ticket.ul_code(2)).s_
    theEngine.initializeEngine process _
                             , div_dates _
                             , divs _
                             , div_array_size _
                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
                             , deal_ticket.t_grid, deal_ticket.x_grid _
                             , deal_ticket.scheme_type _
                             , True _
                             , quantoHelper _
                             , snapshot_time _
                             , ratioDividendIn _
                             , div_refPriceForDividend
                             
    theNote.setPricingEngine theEngine
    
    the_greeks.value = theNote.NPV() * deal_ticket.notional
    Dim tmp_delta() As Double
    Dim tmp_gamma() As Double
    tmp_delta = theNote.delta()
    tmp_gamma = theNote.gamma()
    the_greeks.set_all_deltas tmp_delta, 0, deal_ticket.notional
    the_greeks.set_all_gammas tmp_gamma, 0, deal_ticket.notional
    the_greeks.theta = theNote.theta() * deal_ticket.notional
    
    '------------------------
    ' CrossGamma: +0.1%/0.1%
    '------------------------
    If calc_delta Then
        Set backup_market = market.copy_obj()
    
        market.bump_ul_price 0.001, deal_ticket.ul_code(1)
    
        Set volTs(0) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(1), day_shift)
        Set process(0) = get_process(rTS, qTs(0), volTs(0), market, deal_ticket.ul_code(1))
    
'        theEngine.initializeEngine process _
'                     , div_dates _
'                     , divs _
'                     , div_array_size _
'                     , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                     , deal_ticket.t_grid, deal_ticket.x_grid _
'                     , deal_ticket.scheme_type _
'                     , deal_ticket.strike_shift _
'                     , True _
'                     , quantoHelper _
'                     , snapshot_time
'   <----delete deal_ticket.strike_shift---->
'        theEngine.initializeEngine process _
'                     , div_dates _
'                     , divs _
'                     , div_array_size _
'                     , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                     , deal_ticket.t_grid, deal_ticket.x_grid _
'                     , deal_ticket.scheme_type _
'                     , True _
'                     , quantoHelper _
'                     , snapshot_time

' dll: 2017.11.28
        ratioDividendIn(0) = 0
        ratioDividendIn(1) = 0
'        theEngine.initializeEngine process _
'                                 , div_dates _
'                                 , divs _
'                                 , div_array_size _
'                                 , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                                 , deal_ticket.t_grid, deal_ticket.x_grid _
'                                 , deal_ticket.scheme_type _
'                                 , True _
'                                 , quantoHelper _
'                                 , snapshot_time _
'                                 , ratioDividendIn
' dll: 2018.5.3
    div_refPriceForDividend(0) = market.market_by_ul(deal_ticket.ul_code(1)).s_
    div_refPriceForDividend(1) = market.market_by_ul(deal_ticket.ul_code(2)).s_
    theEngine.initializeEngine process _
                             , div_dates _
                             , divs _
                             , div_array_size _
                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
                             , deal_ticket.t_grid, deal_ticket.x_grid _
                             , deal_ticket.scheme_type _
                             , True _
                             , quantoHelper _
                             , snapshot_time _
                             , ratioDividendIn _
                             , div_refPriceForDividend
                             
        theNote.setPricingEngine theEngine
    
        theNote.NPV
    
        Dim tmp_delta_up_ua1() As Double
        tmp_delta_up_ua1 = theNote.delta
    
        the_greeks.cross_gamma12 = (tmp_delta_up_ua1(1) - tmp_delta(1)) / (0.001 * market.market_by_ul(deal_ticket.ul_code(1)).s_) * deal_ticket.notional
    
        Set market = backup_market.copy_obj
        Set volTs(0) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(1), day_shift)
        Set process(0) = get_process(rTS, qTs(0), volTs(0), market, deal_ticket.ul_code(1))
    End If
    
    '------------------------
    ' 2015-01-05
    '------------------------
    For inx = 1 To no_of_ul
        ul_prices(inx) = market.market_by_ul(deal_ticket.ul_code(inx)).s_
    Next inx
    the_greeks.set_all_ul_prices ul_prices
   
    If snapshot_time >= 0 Then
        Dim tmp_xAxis() As Double
        Dim tmp_yAxis() As Double
        Dim tmp_snapshotValues() As Double
    
        tmp_xAxis = theNote.xAxis()
        tmp_yAxis = theNote.yAxis()
        tmp_snapshotValues = theNote.snapShotValues()
        
        the_greeks.set_xAxis tmp_xAxis, True, deal_ticket.reference_price(1)
        the_greeks.set_yAxis tmp_yAxis, True, deal_ticket.reference_price(2)
        the_greeks.set_snapshot_value tmp_snapshotValues, Abs(deal_ticket.notional)
    End If
    
    '======================================================
    ' VEGA
    '======================================================

    If calc_vega Then '+1%p
    
        Set backup_market = market.copy_obj()
    
        For inx = 1 To no_of_ul
            
            market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.bump_vol_surface 0.01
            Set volTs(inx - 1) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx), day_shift)
            Set process(inx - 1) = get_process(rTS, qTs(inx - 1), volTs(inx - 1), market, deal_ticket.ul_code(inx))
            
'            theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , deal_ticket.strike_shift _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
'   <----delete deal_ticket.strike_shift---->
'            theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
' dll: 2017.11.28
            ratioDividendIn(0) = 0
            ratioDividendIn(1) = 0
'            theEngine.initializeEngine process _
'                                     , div_dates _
'                                     , divs _
'                                     , div_array_size _
'                                     , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                                     , deal_ticket.t_grid, deal_ticket.x_grid _
'                                     , deal_ticket.scheme_type _
'                                     , True _
'                                     , quantoHelper _
'                                     , snapshot_time _
'                                     , ratioDividendIn
' dll: 2018.5.3
    div_refPriceForDividend(0) = market.market_by_ul(deal_ticket.ul_code(1)).s_
    div_refPriceForDividend(1) = market.market_by_ul(deal_ticket.ul_code(2)).s_
    theEngine.initializeEngine process _
                             , div_dates _
                             , divs _
                             , div_array_size _
                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
                             , deal_ticket.t_grid, deal_ticket.x_grid _
                             , deal_ticket.scheme_type _
                             , True _
                             , quantoHelper _
                             , snapshot_time _
                             , ratioDividendIn _
                             , div_refPriceForDividend
                             
            theNote.setPricingEngine theEngine
    
            vegas(inx) = theNote.NPV() * deal_ticket.notional - the_greeks.value
            vannas(inx) = theNote.delta()(inx - 1) * deal_ticket.notional - the_greeks.deltas(inx)
            
            vegas(inx) = vegas(inx)
            vannas(inx) = vannas(inx)
            
            Set market = backup_market.copy_obj
            'market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.rewind_vol_bump
            Set volTs(inx - 1) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx), day_shift)
            Set process(inx - 1) = get_process(rTS, qTs(inx - 1), volTs(inx - 1), market, deal_ticket.ul_code(inx))
            
        Next inx
                
        the_greeks.set_all_vegas vegas
        the_greeks.set_all_vannas vannas
    
    End If
    
    If calc_term_vega Then '+1%p
    
        Set backup_market = market.copy_obj()
        ReDim partial_vegas(1 To no_of_ul, 1 To get_array_size_date(deal_ticket.term_vega_tenor_array())) As Double
        
        For inx = 1 To no_of_ul
            
            For jnx = 1 To get_array_size_date(deal_ticket.term_vega_tenor_array())
                
                market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.bump_local_vol_upto 0.01, deal_ticket.term_vega_tenor(jnx)
            
                Set volTs(inx - 1) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx), day_shift)
                Set process(inx - 1) = get_process(rTS, qTs(inx - 1), volTs(inx - 1), market, deal_ticket.ul_code(inx))
                
'                theEngine.initializeEngine process _
'                                 , div_dates _
'                                 , divs _
'                                 , div_array_size _
'                                 , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                                 , deal_ticket.t_grid, deal_ticket.x_grid _
'                                 , deal_ticket.scheme_type _
'                                 , deal_ticket.strike_shift _
'                                 , True _
'                                 , quantoHelper _
'                                 , snapshot_time
'   <----delete deal_ticket.strike_shift---->
'                theEngine.initializeEngine process _
'                                 , div_dates _
'                                 , divs _
'                                 , div_array_size _
'                                 , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                                 , deal_ticket.t_grid, deal_ticket.x_grid _
'                                 , deal_ticket.scheme_type _
'                                 , True _
'                                 , quantoHelper _
'                                 , snapshot_time
' dll: 2017.11.28
                ratioDividendIn(0) = 0
                ratioDividendIn(1) = 0
'                theEngine.initializeEngine process _
'                                         , div_dates _
'                                         , divs _
'                                         , div_array_size _
'                                         , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                                         , deal_ticket.t_grid, deal_ticket.x_grid _
'                                         , deal_ticket.scheme_type _
'                                         , True _
'                                         , quantoHelper _
'                                         , snapshot_time _
'                                         , ratioDividendIn
' dll: 2017.11.28
    div_refPriceForDividend(0) = market.market_by_ul(deal_ticket.ul_code(1)).s_
    div_refPriceForDividend(1) = market.market_by_ul(deal_ticket.ul_code(2)).s_
    theEngine.initializeEngine process _
                             , div_dates _
                             , divs _
                             , div_array_size _
                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
                             , deal_ticket.t_grid, deal_ticket.x_grid _
                             , deal_ticket.scheme_type _
                             , True _
                             , quantoHelper _
                             , snapshot_time _
                             , ratioDividendIn _
                             , div_refPriceForDividend
                             
                theNote.setPricingEngine theEngine
                
                partial_vegas(inx, jnx) = (theNote.NPV() * deal_ticket.notional - the_greeks.value)
                
                Set market = backup_market.copy_obj
                'market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.rewind_bump_vol_upto
                Set volTs(inx - 1) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx), day_shift)
                Set process(inx - 1) = get_process(rTS, qTs(inx - 1), volTs(inx - 1), market, deal_ticket.ul_code(inx))
            
            Next jnx
                        
        Next inx
        
        the_greeks.no_of_tenors = get_array_size_date(deal_ticket.term_vega_tenor_array())
        the_greeks.redim_arrays 2, the_greeks.no_of_tenors
        the_greeks.set_term_dates_per_ul 1, deal_ticket.term_vega_tenor_array
        the_greeks.set_term_dates_per_ul 2, deal_ticket.term_vega_tenor_array
        
        the_greeks.set_all_partial_vega partial_vegas
        
        the_greeks.to_term_vega deal_ticket.current_date
    
    End If
    
    If calc_skew_s Then '-1%p
    
        Set backup_market = market.copy_obj()
        For inx = 1 To no_of_ul
    
            market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.s_ = market.market_by_ul(deal_ticket.ul_code(inx)).s_
            market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.bump_skew -0.01
            Set volTs(inx - 1) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx), day_shift)
            Set process(inx - 1) = get_process(rTS, qTs(inx - 1), volTs(inx - 1), market, deal_ticket.ul_code(inx))
            
'            theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , deal_ticket.strike_shift _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
'   <----delete deal_ticket.strike_shift---->
'            theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
            ratioDividendIn(0) = 0
            ratioDividendIn(1) = 0
'            theEngine.initializeEngine process _
'                                     , div_dates _
'                                     , divs _
'                                     , div_array_size _
'                                     , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                                     , deal_ticket.t_grid, deal_ticket.x_grid _
'                                     , deal_ticket.scheme_type _
'                                     , True _
'                                     , quantoHelper _
'                                     , snapshot_time _
'                                     , ratioDividendIn
' dll: 2017.11.28
    div_refPriceForDividend(0) = market.market_by_ul(deal_ticket.ul_code(1)).s_
    div_refPriceForDividend(1) = market.market_by_ul(deal_ticket.ul_code(2)).s_
    theEngine.initializeEngine process _
                             , div_dates _
                             , divs _
                             , div_array_size _
                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
                             , deal_ticket.t_grid, deal_ticket.x_grid _
                             , deal_ticket.scheme_type _
                             , True _
                             , quantoHelper _
                             , snapshot_time _
                             , ratioDividendIn _
                             , div_refPriceForDividend
                             
            theNote.setPricingEngine theEngine
    
            skew_s(inx) = (theNote.NPV() * deal_ticket.notional - the_greeks.value) * -1
            
            Set market = backup_market.copy_obj
            'market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.rewind_vol_bump
            Set volTs(inx - 1) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx), day_shift)
            Set process(inx - 1) = get_process(rTS, qTs(inx - 1), volTs(inx - 1), market, deal_ticket.ul_code(inx))
            
        Next inx
        
        the_greeks.set_all_skews skew_s
    
    End If
 
    If calc_corr Then '+0.05/5
        
        Set backup_market = market.copy_obj()
    
'        theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) + 0.05 _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , deal_ticket.strike_shift _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
'   <----delete deal_ticket.strike_shift---->
'        theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) + 0.05 _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
        ratioDividendIn(0) = 0
        ratioDividendIn(1) = 0
'        theEngine.initializeEngine process _
'                                 , div_dates _
'                                 , divs _
'                                 , div_array_size _
'                                 , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2) + 0.05) _
'                                 , deal_ticket.t_grid, deal_ticket.x_grid _
'                                 , deal_ticket.scheme_type _
'                                 , True _
'                                 , quantoHelper _
'                                 , snapshot_time _
'                                 , ratioDividendIn
' dll: 2017.11.28
    div_refPriceForDividend(0) = market.market_by_ul(deal_ticket.ul_code(1)).s_
    div_refPriceForDividend(1) = market.market_by_ul(deal_ticket.ul_code(2)).s_
    theEngine.initializeEngine process _
                             , div_dates _
                             , divs _
                             , div_array_size _
                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
                             , deal_ticket.t_grid, deal_ticket.x_grid _
                             , deal_ticket.scheme_type _
                             , True _
                             , quantoHelper _
                             , snapshot_time _
                             , ratioDividendIn _
                             , div_refPriceForDividend
                             
        theNote.setPricingEngine theEngine

        corr_sens(1, 2) = (theNote.NPV() * deal_ticket.notional - the_greeks.value) / 5
        
        Set market = backup_market.copy_obj
        'market.correlation_pair_.rewind deal_ticket.ul_code(1), deal_ticket.ul_code(2)

        the_greeks.set_all_corr_sens corr_sens
    
    End If

    If calc_rho Then '+10bp/10
    
        Set backup_market = market.copy_obj()
    
        Dim tmp_greek As clsGreeks
        Dim tmp_rate_curves() As clsRateCurve
        ReDim tmp_rate_curves(1 To deal_ticket.no_of_ul) As clsRateCurve
     '   Dim bump_greek_set As clsGreekSet

        Set tmp_greek = New clsGreeks
        Set origin_pl_currency_curve = market.pl_currency_rate_curve_.copy_obj()
        'Bumping DCF: +10bp
        Set market.pl_currency_rate_curve_ = market.pl_currency_rate_curve_.copy_obj(0, 0.001)
        
        run_ac_pricing_2d tmp_greek, deal_ticket, market, bump_greek_set, False, False, False, False, snapshot_time, False, False
        'Rho(DCF)
        rho = (tmp_greek.value - the_greeks.value) * 0.1
        
        Set market.pl_currency_rate_curve_ = origin_pl_currency_curve
        
        'Rho(UL)
        'Bumping riskfree rate curve by adjusting dividend yield: +10bp
        For inx = 1 To deal_ticket.no_of_ul
            
            Set tmp_greek = New clsGreeks
            Dim tmp_div_yield As Double
            tmp_div_yield = market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_
            
            market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_ = tmp_div_yield - 0.001
            
            run_ac_pricing_2d tmp_greek, deal_ticket, market, bump_greek_set, False, False, False, False, snapshot_time, False, False
            
            rho_ul(inx) = (tmp_greek.value - the_greeks.value) * 0.1
            
            Set market = backup_market.copy_obj
            'market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_ = tmp_div_yield
            
        Next inx
        
        the_greeks.rho = rho
        the_greeks.set_all_rho_ul rho_ul
        
    End If

    If bump_delta Then '+0.1%/+0.1%
    
        Set backup_market = market.copy_obj()

        For inx = 1 To no_of_ul

            market.bump_ul_price 0.001, deal_ticket.ul_code(inx)
            'market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.shift_surface market.market_by_ul(deal_ticket.ul_code(inx)).s_, 0.5

            Set volTs(inx - 1) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx), day_shift)
            Set process(inx - 1) = get_process(rTS, qTs(inx - 1), volTs(inx - 1), market, deal_ticket.ul_code(inx))

'            theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , deal_ticket.strike_shift _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
'   <----delete deal_ticket.strike_shift---->
'            theEngine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
            ratioDividendIn(0) = 0
            ratioDividendIn(1) = 0
'            theEngine.initializeEngine process _
'                                     , div_dates _
'                                     , divs _
'                                     , div_array_size _
'                                     , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                                     , deal_ticket.t_grid, deal_ticket.x_grid _
'                                     , deal_ticket.scheme_type _
'                                     , True _
'                                     , quantoHelper _
'                                     , snapshot_time _
'                                     , ratioDividendIn
' dll: 2017.11.28
    div_refPriceForDividend(0) = market.market_by_ul(deal_ticket.ul_code(1)).s_
    div_refPriceForDividend(1) = market.market_by_ul(deal_ticket.ul_code(2)).s_
    theEngine.initializeEngine process _
                             , div_dates _
                             , divs _
                             , div_array_size _
                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
                             , deal_ticket.t_grid, deal_ticket.x_grid _
                             , deal_ticket.scheme_type _
                             , True _
                             , quantoHelper _
                             , snapshot_time _
                             , ratioDividendIn _
                             , div_refPriceForDividend
                             
            theNote.setPricingEngine theEngine

            deltas(inx) = (theNote.NPV() * deal_ticket.notional - the_greeks.value) / (0.001 * market.market_by_ul(deal_ticket.ul_code(inx)).s_)

            Set market = backup_market.copy_obj
            Set volTs(inx - 1) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx), day_shift)
            Set process(inx - 1) = get_process(rTS, qTs(inx - 1), volTs(inx - 1), market, deal_ticket.ul_code(inx))

        Next inx

        the_greeks.set_all_sticky_strike_deltas deltas

    End If


'    If log_file Then 'snapshot_time > 0 Then
'        write_closing_file_2d "C:\log\", deal_ticket.asset_code & "." & config__.snapshot_file_extension, the_greeks.get_xAxis, the_greeks.get_yAxis, the_greeks.get_snapshot_value
'    End If

    theNote.dispose_com
   
    Exit Sub
    
ErrorHandler:

    raise_err "run_ac_pricing_2d", Err.description

End Sub

Public Sub run_ac_pricing_fdm(ByRef the_greeks As clsGreeks _
                        , ByVal deal_ticket As clsACDealTicket _
                        , ByVal market As clsMarketSet _
                        , Optional calc_delta_gamma As Boolean = False _
                        , Optional calc_stickymoneyness_delta As Boolean = False _
                        , Optional calc_stickystrike_delta As Boolean = False _
                        , Optional calc_cross_gamma As Boolean = False _
                        , Optional calc_vega As Boolean = False _
                        , Optional calc_term_vega As Boolean = False _
                        , Optional calc_skew_s As Boolean = False _
                        , Optional calc_corr As Boolean = False _
                        , Optional calc_rho As Boolean = False _
                        , Optional calc_theta As Boolean = False _
                        , Optional snapshot_time As Double = 0.001)

    Dim no_of_ul As Integer
    no_of_ul = deal_ticket.no_of_ul
    
    Dim the_note As Object
    Dim the_engine As Object
    Dim rTS As Object
    
    'base=0
    'for no_of_ul=1, not array
    Dim qTs() As Object
    Dim fTs() As Object
    Dim volTs() As Object
    Dim quantoHelper() As Variant
    Dim process() As Variant
    ReDim qTs(0 To no_of_ul - 1) As Object
    ReDim fTs(0 To no_of_ul - 1) As Object
    ReDim volTs(0 To no_of_ul - 1) As Object
    ReDim quantoHelper(0 To no_of_ul - 1) As Variant
    ReDim process(0 To no_of_ul - 1) As Variant
    
    Dim ratioDividendIn() As Double 'dll: 2017.11.28
    Dim div_refPriceForDividend() As Double 'dll: 2018.5.3
    ReDim ratioDividendIn(0 To no_of_ul - 1) As Double
    ReDim div_refPriceForDividend(0 To no_of_ul - 1) As Double
    
    'base=1
    Dim ul_prices() As Double '---- 2015-01-05
    Dim vegas() As Double
    Dim vannas() As Double
    Dim skew_s() As Double
    Dim deltas() As Double
    Dim gammas() As Double
    Dim rho_ul() As Double
    ReDim ul_prices(1 To no_of_ul) As Double '---- 2015-01-05
    ReDim vegas(1 To no_of_ul) As Double
    ReDim vannas(1 To no_of_ul) As Double
    ReDim skew_s(1 To no_of_ul) As Double
    ReDim deltas(1 To no_of_ul) As Double
    ReDim gammas(1 To no_of_ul) As Double
    ReDim rho_ul(1 To no_of_ul) As Double
    
    If no_of_ul > 1 Then
        Dim corr_sens() As Double
        ReDim corr_sens(1 To no_of_ul, 1 To no_of_ul) As Double
    End If
    
    Dim partial_vegas() As Double ' ---- 2015-10-05
   
    Dim rate_spread As Double
    Dim day_shift As Long
    Dim rho As Double
    
    Dim backup_market As clsMarketSet
    
    Dim inx As Integer
    Dim jnx As Integer
    
On Error GoTo ErrorHandler
    
    If deal_ticket.instrument_type = 0 Then
        rate_spread = deal_ticket.rate_spread
    End If
    
    day_shift = deal_ticket.current_date - market.pl_currency_rate_curve_.rate_dates()(0)
    
    Set rTS = get_rTs(market, rate_spread, day_shift) 'discount curve
    
    For inx = 0 To no_of_ul - 1
        ' dll: 2017.11.28
        ratioDividendIn(inx) = 0
        ' dll: 2018.5.3
        div_refPriceForDividend(inx) = market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_
        
        Set qTs(inx) = get_qTs(market, deal_ticket.ul_code(inx + 1), rate_spread)
        'Set quantoHelper(inx) = get_quantoHelper(rTs, market, deal_ticket.ul_code(inx + 1), rate_spread)
        Set quantoHelper(inx) = get_quantoHelper(rTS, market, deal_ticket.ul_code(inx + 1), deal_ticket.ccy, rate_spread, day_shift) '2018.7.10
        Set volTs(inx) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx + 1), day_shift)
        Set process(inx) = get_process(rTS, qTs(inx), volTs(inx), market, deal_ticket.ul_code(inx + 1))
    Next inx
            
    ' dll: 2018.5.3
    If no_of_ul = 1 Then
    
        Set the_note = ac_deal_ticket_to_clr(deal_ticket)
        Set the_engine = New sy_shared_ptr_FdAutocallableEngine1D
        
'        the_engine.initializeEngine process(0) _
'                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
'                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
'                                 , deal_ticket.t_grid, deal_ticket.x_grid _
'                                 , deal_ticket.scheme_type _
'                                 , True _
'                                 , quantoHelper(0) _
'                                 , snapshot_time _
'                                 , ratioDividendIn
        'day_shift 추가: 2018.12.05
        the_engine.initializeEngine process(0) _
                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_div_dates(0, deal_ticket.maturity_date, deal_ticket.ul_code(), day_shift) _
                                 , market.market_by_ul(deal_ticket.ul_code()).div_schedule_.get_divs(0, deal_ticket.maturity_date, deal_ticket.ul_code()) _
                                 , deal_ticket.t_grid, deal_ticket.x_grid _
                                 , deal_ticket.scheme_type _
                                 , True _
                                 , quantoHelper(0) _
                                 , snapshot_time _
                                 , ratioDividendIn
    ElseIf no_of_ul = 2 Then
    
        Set the_note = ac_deal_ticket_to_clr_2d(deal_ticket)
        Set the_engine = New sy_shared_ptr_FdAutocallableEngine2D
        
        Dim divs() As Double
        Dim div_dates() As Long
        Dim div_array_size() As Long
    
        'get_dividend_array div_dates, divs, div_array_size, deal_ticket.get_ul_codes(), market
        'day_shift 추가: 2018.12.05
        get_dividend_array div_dates, divs, div_array_size, deal_ticket.get_ul_codes(), market, day_shift
        
        the_engine.initializeEngine process _
                                 , div_dates _
                                 , divs _
                                 , div_array_size _
                                 , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
                                 , deal_ticket.t_grid, deal_ticket.x_grid _
                                 , deal_ticket.scheme_type _
                                 , True _
                                 , quantoHelper _
                                 , snapshot_time _
                                 , ratioDividendIn _
                                 , div_refPriceForDividend
                                 
    End If
    
'    the_engine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , deal_ticket.strike_shift _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time
'   <----delete deal_ticket.strike_shift---->
'    the_engine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time

'    the_engine.initializeEngine process _
'                             , div_dates _
'                             , divs _
'                             , div_array_size _
'                             , market.correlation_pair_.get_corr(deal_ticket.ul_code(1), deal_ticket.ul_code(2)) _
'                             , deal_ticket.t_grid, deal_ticket.x_grid _
'                             , deal_ticket.scheme_type _
'                             , True _
'                             , quantoHelper _
'                             , snapshot_time _
'                             , ratioDividendIn
                             
    the_note.setPricingEngine the_engine
    
    the_greeks.value = the_note.NPV() * deal_ticket.notional
    
    For inx = 1 To no_of_ul
        ul_prices(inx) = market.market_by_ul(deal_ticket.ul_code(inx)).s_
    Next inx

    the_greeks.set_all_ul_prices ul_prices
    
    Set backup_market = market.copy_obj()
    
    Dim tmp_delta() As Double
    Dim tmp_gamma() As Double
    tmp_delta = the_note.delta()
    tmp_gamma = the_note.gamma()
    
    If calc_delta_gamma Then
        the_greeks.set_all_deltas tmp_delta, 0, deal_ticket.notional
        the_greeks.set_all_gammas tmp_gamma, 0, deal_ticket.notional
    End If
    
    'cross gamma: +0.1%/0.1%
    If calc_cross_gamma And no_of_ul = 2 Then
    
        Dim xgamma_greeks As New clsGreeks
        
        market.bump_ul_price 0.001, deal_ticket.ul_code(1)

        run_ac_pricing_fdm xgamma_greeks, deal_ticket, market, True
    
        the_greeks.cross_gamma12 = (xgamma_greeks.deltas(2) - the_greeks.deltas(2)) / (0.001 * market.market_by_ul(deal_ticket.ul_code(1)).s_)
    
        Set market = backup_market.copy_obj
        
    End If

    If calc_stickymoneyness_delta Then
    
        Set market = backup_market.copy_obj
        
        Dim stickymoneyness_delta_greek_up As clsGreeks
        Dim stickymoneyness_delta_greek_down As clsGreeks

        For inx = 1 To no_of_ul

            Set stickymoneyness_delta_greek_up = New clsGreeks
            Set stickymoneyness_delta_greek_down = New clsGreeks

            market.bump_ul_price 0.01, deal_ticket.ul_code(inx)
            
            '<----- for the sticky moneyness model 2018.09.07
            Dim shifted_strikes() As Double
            shifted_strikes = market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.get_all_strikes
            Dim i_strike As Integer
            For i_strike = 1 To market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.no_of_strikes
                shifted_strikes(i_strike) = shifted_strikes(i_strike) * 1.01
            Next i_strike
            market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.set_strikes shifted_strikes
            '------>
            
            run_ac_pricing_fdm stickymoneyness_delta_greek_up, deal_ticket, market

            Set market = backup_market.copy_obj

            market.bump_ul_price -0.01, deal_ticket.ul_code(inx)
            
            '<----- for the sticky moneyness model 2018.09.07
            shifted_strikes = market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.get_all_strikes
            For i_strike = 1 To market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.no_of_strikes
                shifted_strikes(i_strike) = shifted_strikes(i_strike) * 0.99
            Next i_strike
            market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.set_strikes shifted_strikes
            '------>
            
            run_ac_pricing_fdm stickymoneyness_delta_greek_down, deal_ticket, market

            deltas(inx) = (stickymoneyness_delta_greek_up.value - stickymoneyness_delta_greek_down.value) / (the_greeks.ul_prices(inx) * 0.02)
            gammas(inx) = (stickymoneyness_delta_greek_up.value + stickymoneyness_delta_greek_down.value - 2 * the_greeks.value) / (the_greeks.ul_prices(inx) * 0.01) ^ 2

            Set market = backup_market.copy_obj
            
        Next inx
        
        the_greeks.set_all_sticky_moneyness_deltas deltas
        the_greeks.set_all_sticky_moneyness_gammas gammas
    
    End If
    
    'sticky_strike_delta: +0.1%/+0.1%
    If calc_stickystrike_delta Then

        For inx = 1 To no_of_ul

            market.bump_ul_price 0.001, deal_ticket.ul_code(inx)
            'shift vol surface: under working
            'market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.shift_surface market.market_by_ul(deal_ticket.ul_code(inx)).s_, 0.5
            
            Dim bump_ul_greeks As New clsGreeks

            deltas(inx) = (bump_ul_greeks.value - the_greeks.value) / (0.001 * market.market_by_ul(deal_ticket.ul_code(inx)).s_)

            Set market = backup_market.copy_obj

        Next inx

        the_greeks.set_all_sticky_strike_deltas deltas

    End If
    
    If calc_vega Then '+1%p
    
        For inx = 1 To no_of_ul
        
            Dim vol_bump_greeks As New clsGreeks
            
            market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.bump_vol_surface 0.01
            
            run_ac_pricing_fdm vol_bump_greeks, deal_ticket, market
    
            vegas(inx) = vol_bump_greeks.value - the_greeks.value
            vannas(inx) = vol_bump_greeks.deltas(inx) - the_greeks.deltas(inx)
            
            Set market = backup_market.copy_obj
            
        Next inx
                
        the_greeks.set_all_vegas vegas
        the_greeks.set_all_vannas vannas
    
    End If
    
    If calc_term_vega Then '+1%p
    
        ReDim partial_vegas(1 To no_of_ul, 1 To get_array_size_date(deal_ticket.term_vega_tenor_array())) As Double
        
        For inx = 1 To no_of_ul
            
            For jnx = 1 To get_array_size_date(deal_ticket.term_vega_tenor_array())
                
                Dim t_volup_greeks As New clsGreeks
                
                market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.bump_local_vol_upto 0.01, deal_ticket.term_vega_tenor(jnx)
                
                run_ac_pricing_fdm t_volup_greeks, deal_ticket, market
            
                partial_vegas(inx, jnx) = t_volup_greeks.value - the_greeks.value
                
                Set market = backup_market.copy_obj
            
            Next jnx
            
        Next inx
        
        the_greeks.no_of_tenors = get_array_size_date(deal_ticket.term_vega_tenor_array())
        the_greeks.redim_arrays 2, the_greeks.no_of_tenors
        the_greeks.set_term_dates_per_ul 1, deal_ticket.term_vega_tenor_array
        the_greeks.set_term_dates_per_ul 2, deal_ticket.term_vega_tenor_array
        
        the_greeks.set_all_partial_vega partial_vegas
        the_greeks.to_term_vega deal_ticket.current_date
    
    End If
    
    If calc_skew_s Then '-1%p
    
        For inx = 1 To no_of_ul
    
            market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.s_ = market.market_by_ul(deal_ticket.ul_code(inx)).s_
            market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.bump_skew -0.01
            
            Dim skew_greeks As New clsGreeks
            
            run_ac_pricing_fdm skew_greeks, deal_ticket, market
    
            skew_s(inx) = (skew_greeks.value - the_greeks.value) * -1
            
            Set market = backup_market.copy_obj
            
        Next inx
        
        the_greeks.set_all_skews skew_s
    
    End If
 
    If calc_corr Then '+0.05/5
        
        market.correlation_pair_.bump_corr deal_ticket.ul_code(1), deal_ticket.ul_code(2), 0.05
        
        Dim corr_greeks As New clsGreeks
        
        run_ac_pricing_fdm corr_greeks, deal_ticket, market

        corr_sens(1, 2) = (corr_greeks.value - the_greeks.value) / 5
        
        Set market = backup_market.copy_obj

        the_greeks.set_all_corr_sens corr_sens
    
    End If

    If calc_rho Then '+10bp/10
    
        Dim rho_ccy_greek As New clsGreeks
        
        'Bumping DCF: +10bp
        Set market.pl_currency_rate_curve_ = market.pl_currency_rate_curve_.copy_obj(0, 0.001)
        
        run_ac_pricing_fdm rho_ccy_greek, deal_ticket, market
        
        Set market = backup_market.copy_obj
        
        'Rho(DCF)
        rho = (rho_ccy_greek.value - the_greeks.value) * 0.1
        the_greeks.rho = rho
        
        'Rho(UL)
        'Bumping riskfree rate curve by adjusting dividend yield: +10bp
        For inx = 1 To deal_ticket.no_of_ul
            
            Dim rho_ul_greek As New clsGreeks
            
            market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_ = market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_ - 0.001
            
            run_ac_pricing_fdm rho_ul_greek, deal_ticket, market
            
            rho_ul(inx) = (rho_ul_greek.value - the_greeks.value) * 0.1
            
            Set market = backup_market.copy_obj
            
        Next inx

        the_greeks.set_all_rho_ul rho_ul
        
    End If
    
    If calc_theta Then
        the_greeks.theta = the_note.theta() * deal_ticket.notional
    End If

'    If log_file Then 'snapshot_time > 0 Then
'        write_closing_file_2d "C:\log\", deal_ticket.asset_code & "." & config__.snapshot_file_extension, the_greeks.get_xAxis, the_greeks.get_yAxis, the_greeks.get_snapshot_value
'    End If

    'base=0
    'for no_of_ul=1, not array
    
    rTS.dispose_com
    
    For inx = 0 To no_of_ul - 1
        qTs(inx).dispose_com
        fTs(inx).dispose_com
        volTs(inx).dispose_com
        process(inx).dispose_com
        quantoHelper(inx).dispose_com
        Set qTs(inx) = Nothing
        Set fTs(inx) = Nothing
        Set volTs(inx) = Nothing
        Set process(inx) = Nothing
        Set quantoHelper(inx) = Nothing
    Next inx
    
    the_note.dispose_com
    the_engine.dispose_com
    
    Set the_note = Nothing
    Set the_engine = Nothing
   
    Set backup_market = Nothing
    Set rTS = Nothing
    
    Exit Sub
    
ErrorHandler:

    raise_err "run_ac_pricing_fdm", Err.description

End Sub

Private Sub draw_test_surface(market As clsMarket, row_position As Integer, column_position As Integer)
    
    Dim inx As Integer
    Dim jnx As Integer
    
    For jnx = 1 To get_array_size_double(market.sabr_surface_.local_vol_surface.grid_.get_all_strikes())
        Sheets("tester").Cells(row_position + 1, column_position + jnx + 7) = market.sabr_surface_.local_vol_surface.grid_.get_all_strikes()(jnx)
    Next jnx
    
    For inx = 1 To get_array_size_date(market.sabr_surface_.local_vol_surface.grid_.get_all_dates())
    
        Sheets("tester").Cells(row_position + inx + 1, column_position + 1) = market.sabr_surface_.sabr_parameters_loc_.sabr_param(inx).forward
        Sheets("tester").Cells(row_position + inx + 1, column_position + 2) = market.sabr_surface_.sabr_parameters_loc_.sabr_param(inx).alpha
        Sheets("tester").Cells(row_position + inx + 1, column_position + 3) = market.sabr_surface_.sabr_parameters_loc_.sabr_param(inx).beta
        Sheets("tester").Cells(row_position + inx + 1, column_position + 4) = market.sabr_surface_.sabr_parameters_loc_.sabr_param(inx).nu
        Sheets("tester").Cells(row_position + inx + 1, column_position + 5) = market.sabr_surface_.sabr_parameters_loc_.sabr_param(inx).rho
        Sheets("tester").Cells(row_position + inx + 1, column_position + 6) = market.sabr_surface_.sabr_parameters_loc_.sabr_param(inx).vol_atm
    
    
        For jnx = 1 To get_array_size_double(market.sabr_surface_.local_vol_surface.grid_.get_all_strikes())
    
            Sheets("tester").Cells(row_position + inx + 1, column_position + jnx + 7) = market.sabr_surface_.local_vol_surface.vol_surface()(inx, jnx)
    
        Next jnx
        
    Next inx


End Sub


Public Sub run_ac_closing_3d(value() As Double, deal_ticket As clsACDealTicket, market As clsMarketSet _
                                , sparse_grid() As Double _
                                , Optional min_percentage As Double = 0.8, Optional max_percentage As Double = 1.2)

    Dim the_greeks As clsGreeks
    Dim reference_prices As Double
    Dim adj_market As clsMarketSet
    
    Const Dimension As Integer = 3
    
    Dim inx As Integer
    Dim jnx As Integer
    
On Error GoTo ErrorHandler
    
    Set the_greeks = New clsGreeks
    
    If deal_ticket.current_date <> market.pl_currency_rate_curve_.rate_dates()(0) Then
        Set adj_market = market.copy_obj(deal_ticket.current_date - market.pl_currency_rate_curve_.rate_dates()(0))
    Else
        Set adj_market = market.copy_obj()
    End If
    
    For inx = 1 To get_array_size_double(sparse_grid)
        
        For jnx = 1 To Dimension
            
            
            adj_market.market_by_ul(deal_ticket.ul_code(jnx)).s_ = ((max_percentage - min_percentage) * sparse_grid(inx, jnx) + min_percentage) _
                                                                 * market.market_by_ul(deal_ticket.ul_code(jnx)).s_
            
        Next jnx
                
        Dim dummy_set As clsGreekSet
        run_ac_pricing_3d the_greeks, deal_ticket, adj_market, dummy_set, False, False, False, False, False, , , , , False
        
        push_back_double value, the_greeks.value ' Sgn(deal_ticket.notional)

    Next inx
    
    
    Exit Sub
    
ErrorHandler:
    
    raise_err "run_ac_closing_3d", Err.description

End Sub

Public Function get_corr_array(no_of_ul As Integer, market As clsMarketSet, deal_ticket As clsACDealTicket) As Double()
    Dim inx As Integer
    Dim jnx As Integer
    Dim rtn_array() As Double
    
    For inx = 1 To no_of_ul
        For jnx = 1 To no_of_ul
            push_back_double rtn_array, market.correlation_pair_.get_corr(deal_ticket.ul_code(inx), deal_ticket.ul_code(jnx)), 0
        Next jnx
    Next inx
    
    get_corr_array = rtn_array
    
End Function

'local correlation 추가 2019. 3. 27
Public Function get_min_corr_array(no_of_ul As Integer, market As clsMarketSet, deal_ticket As clsACDealTicket) As Double()
    Dim inx As Integer
    Dim jnx As Integer
    Dim rtn_array() As Double
    
    For inx = 1 To no_of_ul
        For jnx = 1 To no_of_ul
            push_back_double rtn_array, market.min_correlation_pair_.get_corr(deal_ticket.ul_code(inx), deal_ticket.ul_code(jnx)), 0
        Next jnx
    Next inx
    
    get_min_corr_array = rtn_array
    
End Function


Public Sub run_ac_pricing_3d(ByRef the_greeks As clsGreeks _
                        , deal_ticket As clsACDealTicket _
                        , market As clsMarketSet _
                        , ByRef bump_greek_set As clsGreekSet _
                        , Optional calc_vega As Boolean = False _
                        , Optional calc_skew_s As Boolean = False _
                        , Optional calc_corr As Boolean = False _
                        , Optional calc_rho As Boolean = False _
                        , Optional calc_theta As Boolean = False _
                        , Optional snapshot_time As Double = 0.001 _
                        , Optional ignore_smoothing As Boolean = False _
                        , Optional calc_snapshot As Boolean = False _
                        , Optional bump_delta As Boolean = False _
                        , Optional calc_delta As Boolean = False _
                        , Optional calc_term_vega As Boolean = False _
                        , Optional max_recursive As Integer = 1)
                        
    Dim the_note As Object
    Dim process(0 To 2) As Variant
    Dim the_process_array As Object
    Dim the_engine As Object
    
    Dim rTS As Object
    Dim qTs(0 To 2) As Object
    Dim fTs(0 To 2) As Object
    Dim volTs(0 To 2) As Object
        
    Dim rate_spread As Double
    
    Dim theta As Double
    
    Dim deltas(1 To 3) As Double
    Dim sticky_strike_deltas(1 To 3) As Double
    Dim gammas(1 To 3) As Double
    Dim vegas(1 To 3) As Double
    Dim skew_s(1 To 3) As Double
    Dim vannas(1 To 3) As Double
    Dim ul_prices(1 To 3) As Double
    Dim corr_sens(1 To 3, 1 To 3) As Double
    Dim rho_ul(1 To 3) As Double
    
    Dim partial_vegas() As Double
    
    Dim tmp_corr() As Double
    
    Dim origin_pl_currency_curve As clsRateCurve

    Dim origin_curves(1 To 3) As clsRateCurve
    
    Dim inx As Integer
    Dim jnx As Integer
    
    Dim rho As Double
    
    Dim day_shift As Long
    
    Const no_of_ul As Integer = 3
    
    day_shift = deal_ticket.current_date - market.pl_currency_rate_curve_.rate_dates()(0)
    
    Set the_note = ac_deal_ticket_to_clr_nd(deal_ticket)
    
    If deal_ticket.instrument_type = 0 Then
        rate_spread = deal_ticket.rate_spread
    End If
    
    
    Set rTS = get_rTs(market, rate_spread, day_shift) ' DC Curve
        
    For inx = 0 To no_of_ul - 1
        Set qTs(inx) = get_qTs(market, deal_ticket.ul_code(inx + 1), rate_spread, day_shift)
        Set fTs(inx) = get_fTs(market, deal_ticket.ul_code(inx + 1), rate_spread, day_shift)
       ' Set quantoHelper(inx) = get_quantoHelper(rTs, market, deal_ticket.ul_code(inx + 1), rate_spread)
        Set volTs(inx) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx + 1), day_shift)
        
        Set process(inx) = get_div_process(rTS, qTs(inx), fTs(inx), volTs(inx), market, deal_ticket.ul_code(inx + 1))
    Next inx
        
    Set the_process_array = New sy_shared_ptr_StochasticProcessArray
    
    
    tmp_corr = get_corr_array(no_of_ul, market, deal_ticket)
    
    'the_process_array.initializeProcess process, tmp_corr
    Dim inLocCorr_UseLocalCorr As Boolean
    Dim inLocCorr_MinCorr(0) As Double
    Dim lambdaSlope_down(0) As Double
    Dim lambdaSlope_up(0) As Double
    Dim lambdaNeutral(0) As Double
    Dim lambdaPriceChangeInterval(0) As Double
    inLocCorr_UseLocalCorr = False
    inLocCorr_MinCorr(0) = 0
    lambdaSlope_down(0) = 0
    lambdaSlope_up(0) = 0
    lambdaNeutral(0) = 0
    lambdaPriceChangeInterval(0) = 0
    the_process_array.initializeProcess process, tmp_corr, inLocCorr_UseLocalCorr, inLocCorr_MinCorr, lambdaSlope_down, lambdaSlope_up, lambdaNeutral, lambdaPriceChangeInterval
    
    Set the_engine = New sy_shared_ptr_McAutocallableEngineNd
    
    'the_engine.initializeEngine the_process_array, deal_ticket.t_grid, deal_ticket.t_grid, deal_ticket.no_of_trials, 0#, deal_ticket.no_of_trials, calc_theta
    the_engine.initializeEngine the_process_array, deal_ticket.t_grid, deal_ticket.t_grid, deal_ticket.no_of_trials, 0#, deal_ticket.no_of_trials, calc_theta, inLocCorr_UseLocalCorr
    
    the_note.setPricingEngine the_engine
    
    the_greeks.value = the_note.NPV() * deal_ticket.notional '<----------------- Value
    
    '------------------------
    ' 2015-01-05
    '------------------------
    For inx = 1 To no_of_ul
        ul_prices(inx) = market.market_by_ul(deal_ticket.ul_code(inx)).s_
    Next inx
    
    the_greeks.set_all_ul_prices ul_prices
    
    If calc_theta Then
        the_greeks.theta = the_note.theta() * deal_ticket.notional
    End If
    
    '--------------------------Delta Gamma
    If calc_delta Then
    
        If max_recursive = 0 Then
            calc_delta = False
        Else
            max_recursive = max_recursive - 1
        End If
        
        Dim delta_greek_up1_deltas2 As Double
        Dim delta_greek_down1_deltas2 As Double
        Dim delta_greek_up1_deltas3 As Double
        Dim delta_greek_down1_deltas3 As Double
        Dim delta_greek_up2_deltas3 As Double
        Dim delta_greek_down2_deltas3 As Double
        
        For inx = 0 To no_of_ul - 1
       
            If inx = no_of_ul - 1 Then
                calc_delta = False
            End If

            Dim s_origin As Double
            
            Dim delta_greek_up As clsGreeks
            Dim delta_greek_down As clsGreeks
            
            Set delta_greek_up = New clsGreeks
            Set delta_greek_down = New clsGreeks
            
            
'            Dim value_up As Double
'            Dim value_down As Double


            s_origin = market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_

            '------------------------- Calc value up
            market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_ = s_origin * 1.01
            run_ac_pricing_3d delta_greek_up, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, ignore_smoothing, False, False, calc_delta, False, max_recursive

            market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_ = s_origin * 0.99
            run_ac_pricing_3d delta_greek_down, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, ignore_smoothing, False, False, calc_delta, False, max_recursive

            deltas(inx + 1) = (delta_greek_up.value - delta_greek_down.value) / (s_origin * 0.02)
            gammas(inx + 1) = (delta_greek_up.value + delta_greek_down.value - 2 * the_greeks.value) / (s_origin * 0.01) ^ 2

            'REWIND
            market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_ = s_origin
            
            If inx = 0 And calc_delta Then
                delta_greek_up1_deltas2 = delta_greek_up.deltas(2)
                delta_greek_up1_deltas3 = delta_greek_up.deltas(3)
                delta_greek_down1_deltas2 = delta_greek_down.deltas(2)
                delta_greek_down1_deltas3 = delta_greek_down.deltas(3)
            End If

            If inx = 1 And calc_delta Then
                delta_greek_up2_deltas3 = delta_greek_up.deltas(3)
                delta_greek_down2_deltas3 = delta_greek_down.deltas(3)
            End If

        Next inx

        the_greeks.set_all_deltas deltas
        the_greeks.set_all_gammas gammas
        
        the_greeks.cross_gamma12 = (delta_greek_up1_deltas2 - delta_greek_down1_deltas2) / (market.market_by_ul(deal_ticket.ul_code(1)).s_ * 0.02)
        the_greeks.cross_gamma13 = (delta_greek_up1_deltas3 - delta_greek_down1_deltas3) / (market.market_by_ul(deal_ticket.ul_code(1)).s_ * 0.02)
        the_greeks.cross_gamma23 = (delta_greek_up2_deltas3 - delta_greek_down2_deltas3) / (market.market_by_ul(deal_ticket.ul_code(2)).s_ * 0.02)

    End If
    '<--------------------------Delta Gamma

    '--------------------------Vega : +1%p
    If calc_vega Then
        For inx = 0 To no_of_ul - 1
        
            Dim vol_bump_greek As clsGreeks
            Dim vanna_greek_up As clsGreeks
            Dim vanna_greek_down As clsGreeks
            
            Set vol_bump_greek = New clsGreeks
            Set vanna_greek_up = New clsGreeks
            Set vanna_greek_down = New clsGreeks
            'Set origin_pl_currency_curve = market.pl_currency_rate_curve_
            market.market_by_ul(deal_ticket.ul_code(inx + 1)).sabr_surface_.bump_vol_surface 0.01
            'Set market.pl_currency_rate_curve_ = market.pl_currency_rate_curve_.copy_obj(0, 0.001)
    
            run_ac_pricing_3d vol_bump_greek, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, ignore_smoothing, False, False, False
            vegas(inx + 1) = vol_bump_greek.value - the_greeks.value
            vegas(inx + 1) = vegas(inx + 1)
            
            s_origin = market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_
            '------------------------- Calc value up
            market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_ = s_origin * 1.01
            run_ac_pricing_3d vanna_greek_up, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, ignore_smoothing, False, False, False
            
            market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_ = s_origin * 0.99
            run_ac_pricing_3d vanna_greek_down, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, ignore_smoothing, False, False, False
            
            market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_ = s_origin
            
            vannas(inx + 1) = (vanna_greek_up.value - vanna_greek_down.value) / (s_origin * 0.02) - the_greeks.deltas(inx + 1)
            vannas(inx + 1) = vannas(inx + 1)
    
            market.market_by_ul(deal_ticket.ul_code(inx + 1)).sabr_surface_.rewind_vol_bump

        Next inx
    
    End If
    
    the_greeks.set_all_vegas vegas
    the_greeks.set_all_vannas vannas

    '<--------------------------Vega
    
    
    
    If calc_term_vega Then '+1%p
    
        ReDim partial_vegas(1 To no_of_ul, 1 To get_array_size_date(deal_ticket.term_vega_tenor_array())) As Double
        
        For inx = 1 To no_of_ul
            
            For jnx = 1 To get_array_size_date(deal_ticket.term_vega_tenor_array())
                
                market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.bump_local_vol_upto 0.01, deal_ticket.term_vega_tenor(jnx)
            
                Set vol_bump_greek = New clsGreeks
                run_ac_pricing_3d vol_bump_greek, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, ignore_smoothing, False, False, False
                
                partial_vegas(inx, jnx) = (vol_bump_greek.value - the_greeks.value)
                               
                market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.rewind_bump_vol_upto
            
            Next jnx
    
                        
        Next inx
        
        the_greeks.no_of_tenors = get_array_size_date(deal_ticket.term_vega_tenor_array())
        the_greeks.set_all_partial_vega partial_vegas
        
        the_greeks.to_term_vega deal_ticket.current_date
    
    End If
    
    


    '--------------------------Skew S.

    If calc_skew_s Then
        For inx = 0 To no_of_ul - 1
        
        
            Dim skew_bump_greek As clsGreeks
            Set skew_bump_greek = New clsGreeks
            
            market.market_by_ul(deal_ticket.ul_code(inx + 1)).sabr_surface_.s_ = market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_
            market.market_by_ul(deal_ticket.ul_code(inx + 1)).sabr_surface_.bump_skew -0.01
            
            run_ac_pricing_3d skew_bump_greek, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, ignore_smoothing, False, False, False
            skew_s(inx + 1) = (skew_bump_greek.value - the_greeks.value) * -1

            '--- Rewind
            market.market_by_ul(deal_ticket.ul_code(inx + 1)).sabr_surface_.rewind_vol_bump
            

        Next inx
    End If
    the_greeks.set_all_skews skew_s

    '<--------------------------Skew S.


    '-------------------------- Corr S.
    If calc_corr Then

        For inx = 0 To no_of_ul - 1
            For jnx = inx + 1 To no_of_ul - 1
            
                Dim corr_bump_greek As clsGreeks
                Set corr_bump_greek = New clsGreeks

                market.correlation_pair_.bump_corr deal_ticket.ul_code(inx + 1), deal_ticket.ul_code(jnx + 1), 0.05
                
                run_ac_pricing_3d corr_bump_greek, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, ignore_smoothing, False, False, False

'                tmp_corr = get_corr_array(no_of_ul, market, deal_ticket)
'
'                the_process_array.initializeProcess process, tmp_corr
'                the_engine.initializeEngine the_process_array, deal_ticket.t_grid, deal_ticket.t_grid, deal_ticket.no_of_trials, 0#, deal_ticket.no_of_trials, False
'                the_note.setPricingEngine the_engine
'                value_up = the_note.NPV() * deal_ticket.notional

                corr_sens(inx + 1, jnx + 1) = (corr_bump_greek.value - the_greeks.value) / 5
                
                market.correlation_pair_.rewind deal_ticket.ul_code(inx + 1), deal_ticket.ul_code(jnx + 1)

            Next jnx
        Next inx

    End If

    the_greeks.set_all_corr_sens corr_sens
    
    
'<-------------------------- Corr S.

   
    If calc_rho Then '+10bp/10
        Dim tmp_greek As clsGreeks
        Dim tmp_rate_curves() As clsRateCurve
        ReDim tmp_rate_curves(1 To deal_ticket.no_of_ul) As clsRateCurve
     '   Dim bump_greek_set As clsGreekSet

        Set tmp_greek = New clsGreeks
        Set origin_pl_currency_curve = market.pl_currency_rate_curve_.copy_obj()
        Set market.pl_currency_rate_curve_ = market.pl_currency_rate_curve_.copy_obj(0, 0.001)
                
        
        run_ac_pricing_3d tmp_greek, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, False, False, False, False
       
        rho = (tmp_greek.value - the_greeks.value) * 0.1
        
        Set market.pl_currency_rate_curve_ = origin_pl_currency_curve
        
        For inx = 1 To deal_ticket.no_of_ul
            
            Set tmp_greek = New clsGreeks

            Dim tmp_div_yield As Double
            tmp_div_yield = market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_
            
            market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_ = tmp_div_yield - 0.001
                        
            run_ac_pricing_3d tmp_greek, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, False, False, False, False
            
            rho_ul(inx) = (tmp_greek.value - the_greeks.value) * 0.1
            
            market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_ = tmp_div_yield
            
        Next inx

    End If

    the_greeks.rho = rho
    the_greeks.set_all_rho_ul rho_ul

    '###
    ' Lee, 2017-11-03
    rTS.dispose_com
    
    For inx = 0 To no_of_ul - 1
        qTs(inx).dispose_com
        fTs(inx).dispose_com
        volTs(inx).dispose_com
        process(inx).dispose_com
    Next inx
    
    the_process_array.dispose_com
    the_engine.dispose_com
    the_note.dispose_com

    Exit Sub
    
ErrorHandler:
    '###
    ' Lee, 2017-11-03
    rTS.dispose_com
    
    For inx = 0 To no_of_ul - 1
        qTs(inx).dispose_com
        fTs(inx).dispose_com
        volTs(inx).dispose_com
        process(inx).dispose_com
    Next inx
    
    the_process_array.dispose_com
    the_engine.dispose_com

    the_note.dispose_com
    
    raise_err "run_ac_pricing_3d", Err.description

End Sub


Public Sub run_ac_pricing_mc(ByRef the_greeks As clsGreeks _
                        , deal_ticket As clsACDealTicket _
                        , market As clsMarketSet _
                        , Optional calc_delta_gamma As Boolean = False _
                        , Optional calc_stickymoneyness_delta As Boolean = False _
                        , Optional calc_stickystrike_delta As Boolean = False _
                        , Optional calc_cross_gamma As Boolean = False _
                        , Optional calc_vega As Boolean = False _
                        , Optional calc_term_vega As Boolean = False _
                        , Optional calc_skew_s As Boolean = False _
                        , Optional calc_corr As Boolean = False _
                        , Optional calc_rho As Boolean = False _
                        , Optional calc_theta As Boolean = False _
                        , Optional inLocCorr_UseLocalCorr As Boolean = True _
                        , Optional snapshot_time As Double = 0.001)
                        ', Optional max_recursive As Integer = 1)
                        
    Dim no_of_ul As Integer
    no_of_ul = deal_ticket.no_of_ul
    
    Dim the_note As Object
    Dim the_engine As Object
    Dim rTS As Object
    Dim qTs() As Object
    Dim fTs() As Object
    Dim aTs() As Object 'drift adjustment 추가: 2023.11.21
    Dim volTs() As Object
    Dim process() As Variant
    ReDim qTs(0 To no_of_ul - 1) As Object
    ReDim fTs(0 To no_of_ul - 1) As Object
    ReDim aTs(0 To no_of_ul - 1) As Object 'drift adjustment 추가: 2023.11.21
    ReDim volTs(0 To no_of_ul - 1) As Object
    ReDim process(0 To no_of_ul - 1) As Variant
    Dim the_process_array As Object
    
    Dim ul_prices() As Double
    Dim vegas() As Double
    Dim vannas() As Double
    Dim skew_s() As Double
    Dim deltas() As Double
    Dim sticky_strike_deltas() As Double
    Dim gammas() As Double
    Dim theta As Double
    Dim rho_ul() As Double
    ReDim ul_prices(1 To no_of_ul) As Double
    ReDim vegas(1 To no_of_ul) As Double
    ReDim vannas(1 To no_of_ul) As Double
    ReDim skew_s(1 To no_of_ul) As Double
    ReDim deltas(1 To no_of_ul) As Double
    ReDim sticky_strike_deltas(1 To no_of_ul) As Double
    ReDim gammas(1 To no_of_ul) As Double
    ReDim rho_ul(1 To no_of_ul) As Double
    
    If no_of_ul > 1 Then
        Dim corr_sens() As Double
        ReDim corr_sens(1 To no_of_ul, 1 To no_of_ul) As Double
    End If
    
    Dim partial_vegas() As Double

    Dim rate_spread As Double
    Dim day_shift As Long
    Dim rho As Double
    
    Dim backup_market As clsMarketSet
    Dim origin_pl_currency_curve As clsRateCurve
    
    Dim inx As Integer
    Dim jnx As Integer
    
    '<------ local correlation 추가 2019.3.27
    'Dim inLocCorr_UseLocalCorr As Boolean
    'inLocCorr_UseLocalCorr = True
    
    Dim lambdaNeutral() As Double
    Dim lambdaSlope_down() As Double
    Dim lambdaSlope_up() As Double
    Dim lambdaPriceChangeInterval() As Double
    ReDim lambdaNeutral(0 To no_of_ul - 1) As Double
    ReDim lambdaSlope_down(0 To no_of_ul - 1) As Double
    ReDim lambdaSlope_up(0 To no_of_ul - 1) As Double
    ReDim lambdaPriceChangeInterval(0 To no_of_ul - 1) As Double
    '------>/
    
On Error GoTo ErrorHandler
    
    '[추가 필요] 상환여부 체크 -> 상환확정시 상환가 리턴 + 그릭 계산 X
    
    If deal_ticket.instrument_type = 0 Then
        rate_spread = deal_ticket.rate_spread
    End If
    
    day_shift = deal_ticket.current_date - market.dcf_by_ccy(deal_ticket.ccy).rate_dates()(0)
    
    'Set rTs = get_rTs(market, rate_spread, day_shift) ' discount curve
    Set rTS = get_rTs(market.dcf_by_ccy(deal_ticket.ccy), rate_spread, day_shift) ' discount curve
    
    For inx = 0 To no_of_ul - 1
        Set qTs(inx) = get_qTs(market, deal_ticket.ul_code(inx + 1), rate_spread, day_shift)
        Set fTs(inx) = get_fTs(market, deal_ticket.ul_code(inx + 1), rate_spread, day_shift)
        Set aTs(inx) = get_aTs(market, deal_ticket.ul_code(inx + 1), rate_spread, day_shift) 'drift adjustment 추가: 2023.11.21
        
        'Set quantoHelper(inx) = get_quantoHelper(rTs, market, deal_ticket.ul_code(inx + 1), rate_spread)
        Set volTs(inx) = get_vol_surface_(deal_ticket.current_date, market, deal_ticket.ul_code(inx + 1), day_shift)
        'Set process(inx) = get_div_process(rTs, qTs(inx), fTs(inx), volTs(inx), market, deal_ticket.ul_code(inx + 1), deal_ticket.ccy, deal_ticket.ratioLeverage(inx + 1))
        'div_dates에 day_shift 추가: 2018.12.05
        'Set process(inx) = get_div_process(rTs, qTs(inx), fTs(inx), volTs(inx), market, deal_ticket.ul_code(inx + 1), deal_ticket.ccy, deal_ticket.ratioLeverage(inx + 1), day_shift)
        'ratioLeverage -> codeLeverage : 2020.6.19
        'Set process(inx) = get_div_process(rTS, qTs(inx), fTs(inx), volTs(inx), market, deal_ticket.ul_code(inx + 1), deal_ticket.ccy, deal_ticket.codeLeverage(inx + 1), day_shift)
        'drift adjustment 추가: 2023.11.21
        Set process(inx) = get_div_process(rTS, qTs(inx), fTs(inx), aTs(inx), volTs(inx), market, deal_ticket.ul_code(inx + 1), deal_ticket.ccy, deal_ticket.codeLeverage(inx + 1), day_shift)
        
        '<------ local correlation 추가 2019.3.27
        lambdaNeutral(inx) = market.market_by_ul(deal_ticket.ul_code(inx + 1)).lambda_neutral
        lambdaSlope_down(inx) = -0.5
        lambdaSlope_up(inx) = -0.5
        lambdaPriceChangeInterval(inx) = 0.5 '6개월
        '------>/
    Next inx
    
    Set the_process_array = New sy_shared_ptr_StochasticProcessArray

    the_process_array.initializeProcess process _
                                        , get_corr_array(no_of_ul, market, deal_ticket) _
                                        , inLocCorr_UseLocalCorr _
                                        , get_min_corr_array(no_of_ul, market, deal_ticket) _
                                        , lambdaSlope_down _
                                        , lambdaSlope_up _
                                        , lambdaNeutral _
                                        , lambdaPriceChangeInterval
    
    '<------ local correlation 추가 2019.3.27
    Dim buffer_days As Integer
    buffer_days = 5
    
    For inx = 0 To no_of_ul - 1
    
        '과거 주가 입수
        ' 과거 날짜: datesAsLong()
        ' 과거 주가: historical_prices()
        Dim start_date As String
        Dim end_date As String
        start_date = date2str(deal_ticket.current_date - Round(lambdaPriceChangeInterval(inx) * 365 + buffer_days, 0))
        end_date = date2str(deal_ticket.current_date - 1)
        
        Dim historical_date() As Long
        Dim historical_price() As Double
        
        Call get_closing_s_series(deal_ticket.ul_code(inx + 1), start_date, end_date, historical_date(), historical_price())
        
        the_process_array.import_past_path inx, historical_date, historical_price
        
        Erase historical_date
        Erase historical_price
        
    Next inx
    '------>/
    
    Set the_note = ac_deal_ticket_to_clr_nd(deal_ticket)
    Set the_engine = New sy_shared_ptr_McAutocallableEngineNd
    
'    the_engine.initializeEngine the_process_array _
'                                , deal_ticket.t_grid, deal_ticket.t_grid _
'                                , deal_ticket.no_of_trials _
'                                , 0# _
'                                , deal_ticket.no_of_trials _
'                                , calc_theta _
'                                , inLocCorr_UseLocalCorr '<--추가

    'time step 조정 반영 (2023.06.15)
    the_engine.initializeEngine the_process_array _
                                , rTS _
                                , deal_ticket.no_of_trials _
                                , deal_ticket.days_per_step _
                                , inLocCorr_UseLocalCorr
    
    the_note.setPricingEngine the_engine
    
    the_greeks.value = the_note.NPV() * deal_ticket.notional
    
    For inx = 1 To no_of_ul
        ul_prices(inx) = market.market_by_ul(deal_ticket.ul_code(inx)).s_
    Next inx
    
    If GREEKS_ENABLE = True Then
    
        the_greeks.set_all_ul_prices ul_prices
                
        Set backup_market = market.copy_obj()
            
        If calc_delta_gamma Then
        
            Set market = backup_market.copy_obj
            
            Dim delta_greek_up As clsGreeks
            Dim delta_greek_down As clsGreeks
    
            For inx = 1 To no_of_ul
    
                Set delta_greek_up = New clsGreeks
                Set delta_greek_down = New clsGreeks
    
                market.bump_ul_price 0.01, deal_ticket.ul_code(inx)
                
                run_ac_pricing_mc delta_greek_up, deal_ticket, market
    
                Set market = backup_market.copy_obj
    
                market.bump_ul_price -0.01, deal_ticket.ul_code(inx)
                
                run_ac_pricing_mc delta_greek_down, deal_ticket, market
    
                deltas(inx) = (delta_greek_up.value - delta_greek_down.value) / (the_greeks.ul_prices(inx) * 0.02)
                gammas(inx) = (delta_greek_up.value + delta_greek_down.value - 2 * the_greeks.value) / (the_greeks.ul_prices(inx) * 0.01) ^ 2
    
                Set market = backup_market.copy_obj
                
            Next inx
            
            Set delta_greek_up = Nothing
            Set delta_greek_down = Nothing
                
            the_greeks.set_all_deltas deltas
            the_greeks.set_all_gammas gammas
    
        End If
        
        If calc_stickymoneyness_delta Then
        
            Set market = backup_market.copy_obj
            
            Dim stickymoneyness_delta_greek_up As clsGreeks
            Dim stickymoneyness_delta_greek_down As clsGreeks
    
            For inx = 1 To no_of_ul
    
                Set stickymoneyness_delta_greek_up = New clsGreeks
                Set stickymoneyness_delta_greek_down = New clsGreeks
    
                market.bump_ul_price 0.01, deal_ticket.ul_code(inx)
                
                '<----- for the sticky moneyness model 2018.09.07
                Dim shifted_strikes() As Double
                shifted_strikes = market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.get_all_strikes
                Dim i_strike As Integer
                For i_strike = 1 To market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.no_of_strikes
                    shifted_strikes(i_strike) = shifted_strikes(i_strike) * 1.01
                Next i_strike
                market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.set_strikes shifted_strikes
                '------>
                
                run_ac_pricing_mc stickymoneyness_delta_greek_up, deal_ticket, market
    
                Set market = backup_market.copy_obj
    
                market.bump_ul_price -0.01, deal_ticket.ul_code(inx)
                
                '<----- for the sticky moneyness model 2018.09.07
                shifted_strikes = market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.get_all_strikes
                For i_strike = 1 To market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.no_of_strikes
                    shifted_strikes(i_strike) = shifted_strikes(i_strike) * 0.99
                Next i_strike
                market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.grid_.set_strikes shifted_strikes
                '------>
                
                run_ac_pricing_mc stickymoneyness_delta_greek_down, deal_ticket, market
    
                deltas(inx) = (stickymoneyness_delta_greek_up.value - stickymoneyness_delta_greek_down.value) / (the_greeks.ul_prices(inx) * 0.02)
                gammas(inx) = (stickymoneyness_delta_greek_up.value + stickymoneyness_delta_greek_down.value - 2 * the_greeks.value) / (the_greeks.ul_prices(inx) * 0.01) ^ 2
    
                Set market = backup_market.copy_obj
                
            Next inx
            
            Set stickymoneyness_delta_greek_up = Nothing
            Set stickymoneyness_delta_greek_down = Nothing
                
            the_greeks.set_all_sticky_moneyness_deltas deltas
            the_greeks.set_all_sticky_moneyness_gammas gammas
        
        End If
        
        If calc_cross_gamma Then
        
            Set market = backup_market.copy_obj
                
            Dim delta_greek_up1_deltas2 As Double
            Dim delta_greek_down1_deltas2 As Double
            Dim delta_greek_up1_deltas3 As Double
            Dim delta_greek_down1_deltas3 As Double
            Dim delta_greek_up2_deltas3 As Double
            Dim delta_greek_down2_deltas3 As Double
    
            For inx = 1 To no_of_ul - 1
    
                Set delta_greek_up = New clsGreeks
                Set delta_greek_down = New clsGreeks
                
                market.bump_ul_price 0.01, deal_ticket.ul_code(inx)
                
                run_ac_pricing_mc delta_greek_up, deal_ticket, market, True
                
                Set market = backup_market.copy_obj
                
                market.bump_ul_price -0.01, deal_ticket.ul_code(inx)
                
                run_ac_pricing_mc delta_greek_down, deal_ticket, market, True
    
                If inx = 1 Then
                    delta_greek_up1_deltas2 = delta_greek_up.deltas(2)
                    delta_greek_up1_deltas3 = delta_greek_up.deltas(3)
                    delta_greek_down1_deltas2 = delta_greek_down.deltas(2)
                    delta_greek_down1_deltas3 = delta_greek_down.deltas(3)
                End If
    
                If inx = 2 Then
                    delta_greek_up2_deltas3 = delta_greek_up.deltas(3)
                    delta_greek_down2_deltas3 = delta_greek_down.deltas(3)
                End If
    
                Set market = backup_market.copy_obj
    
            Next inx
            
            Set delta_greek_up = Nothing
            Set delta_greek_down = Nothing
    
            the_greeks.cross_gamma12 = (delta_greek_up1_deltas2 - delta_greek_down1_deltas2) / (ul_prices(1) * 0.02)
            the_greeks.cross_gamma13 = (delta_greek_up1_deltas3 - delta_greek_down1_deltas3) / (ul_prices(1) * 0.02)
            the_greeks.cross_gamma23 = (delta_greek_up2_deltas3 - delta_greek_down2_deltas3) / (ul_prices(2) * 0.02)
    
        End If
    
        'vega : +1%p
        If calc_vega Then
        
            Set market = backup_market.copy_obj
    
            Dim vol_bump_greek As clsGreeks
            
            For inx = 1 To no_of_ul
    
                Set vol_bump_greek = New clsGreeks
    
                market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.bump_vol_surface 0.01
    
                run_ac_pricing_mc vol_bump_greek, deal_ticket, market
    
                vegas(inx) = vol_bump_greek.value - the_greeks.value
                'vannas(inx) = vol_bump_greek.deltas(inx) - the_greeks.deltas(inx)
    
                Set market = backup_market.copy_obj
                
            Next inx
            
            Set vol_bump_greek = Nothing
                
            the_greeks.set_all_vegas vegas
            'the_greeks.set_all_vannas vannas
    
        End If
    
    '    If calc_term_vega Then '+1%p
    '
    '        ReDim partial_vegas(1 To no_of_ul, 1 To get_array_size_date(deal_ticket.term_vega_tenor_array())) As Double
    '
    '        For inx = 1 To no_of_ul
    '
    '            For jnx = 1 To get_array_size_date(deal_ticket.term_vega_tenor_array())
    '
    '                market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.bump_local_vol_upto 0.01, deal_ticket.term_vega_tenor(jnx)
    '
    '                Set vol_bump_greek = New clsGreeks
    '                run_ac_pricing_3d vol_bump_greek, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, ignore_smoothing, False, False, False
    '
    '                partial_vegas(inx, jnx) = (vol_bump_greek.value - the_greeks.value)
    '
    '                market.market_by_ul(deal_ticket.ul_code(inx)).sabr_surface_.local_vol_surface.rewind_bump_vol_upto
    '
    '            Next jnx
    '
    '
    '        Next inx
    '
    '        the_greeks.no_of_tenors = get_array_size_date(deal_ticket.term_vega_tenor_array())
    '        the_greeks.set_all_partial_vega partial_vegas
    '
    '        the_greeks.to_term_vega deal_ticket.current_date
    '
    '    End If
    '
    '    If calc_skew_s Then
    '        For inx = 0 To no_of_ul - 1
    '
    '
    '            Dim skew_bump_greek As clsGreeks
    '            Set skew_bump_greek = New clsGreeks
    '
    '            market.market_by_ul(deal_ticket.ul_code(inx + 1)).sabr_surface_.s_ = market.market_by_ul(deal_ticket.ul_code(inx + 1)).s_
    '            market.market_by_ul(deal_ticket.ul_code(inx + 1)).sabr_surface_.bump_skew -0.01
    '
    '            run_ac_pricing_3d skew_bump_greek, deal_ticket, market, bump_greek_set, False, False, False, False, False, snapshot_time, ignore_smoothing, False, False, False
    '            skew_s(inx + 1) = (skew_bump_greek.value - the_greeks.value) * -1
    '
    '            '--- Rewind
    '            market.market_by_ul(deal_ticket.ul_code(inx + 1)).sabr_surface_.rewind_vol_bump
    '
    '
    '        Next inx
    '
    '        the_greeks.set_all_skews skew_s
    '
    '    End If
    
    
    '    If calc_corr Then
    '
    '        Set market = backup_market.copy_obj
    '
    '        For inx = 0 To no_of_ul - 1
    '
    '            Dim lambda_bump_greek As clsGreeks
    '            Set lambda_bump_greek = New clsGreeks
    '
    '            market.market_by_ul(deal_ticket.ul_code(inx + 1)).lambda_neutral = market.market_by_ul(deal_ticket.ul_code(inx + 1)).lambda_neutral + 0.01
    '
    '            run_ac_pricing_mc lambda_bump_greek, deal_ticket, market
    '
    '            deltas(inx) = (lambda_bump_greek.value - the_greeks.ul_prices(inx)) / 0.01
    '
    '            Set market = backup_market.copy_obj
    '
    '        Next inx
    '
    '        the_greeks.set_all_corr_sens corr_sens
    '
    '    End If
    
    
        If calc_rho Then '+10bp/10
        
            Set market = backup_market.copy_obj
    
            Dim tmp_greek As New clsGreeks
            
            'notional
            market.set_pl_currency_rate_curve get_ccy_idx(deal_ticket.ccy), market.dcf_by_ccy(deal_ticket.ccy).copy_obj(0, 0.001)
            
            run_ac_pricing_mc tmp_greek, deal_ticket, market
    
            rho = (tmp_greek.value - the_greeks.value) * 0.1
    
            the_greeks.rho = rho
    
            Set market = backup_market.copy_obj
    
            'underlying asset
            For inx = 1 To deal_ticket.no_of_ul
                
                Set tmp_greek = New clsGreeks
    
                market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_ = market.market_by_ul(deal_ticket.ul_code(inx)).div_yield_ - 0.001
    
                run_ac_pricing_mc tmp_greek, deal_ticket, market
    
                rho_ul(inx) = (tmp_greek.value - the_greeks.value) * 0.1
                
                Set market = backup_market.copy_obj
                
            Next inx
            
            Set tmp_greek = Nothing
                
            the_greeks.set_all_rho_ul rho_ul
            
        End If
        
        If calc_theta Then
            
            'dll(2023.05.30)
            Dim theta_greek_1d As New clsGreeks
            
            deal_ticket.current_date = deal_ticket.current_date + 1
            
            run_ac_pricing_mc theta_greek_1d, deal_ticket, market
            
            'the_greeks.theta = the_note.theta() * deal_ticket.notional
            the_greeks.theta = (theta_greek_1d.value - the_greeks.value) * 365
            
            deal_ticket.current_date = deal_ticket.current_date - 1
    
            Set theta_greek_1d = Nothing
    
        End If
    
    End If


    '###
    ' Lee, 2017-11-03
    rTS.dispose_com
    Set rTS = Nothing
    
    For inx = 0 To no_of_ul - 1
        qTs(inx).dispose_com
        fTs(inx).dispose_com
        volTs(inx).dispose_com
        process(inx).dispose_com
        Set qTs(inx) = Nothing
        Set fTs(inx) = Nothing
        Set volTs(inx) = Nothing
        Set process(inx) = Nothing
    Next inx
    
    the_process_array.dispose_com
    Set the_process_array = Nothing
    
    the_engine.dispose_com
    Set the_engine = Nothing
    
    the_note.dispose_com
    Set the_note = Nothing
       
    Set backup_market = Nothing
    Set origin_pl_currency_curve = Nothing
    
    Exit Sub
    
ErrorHandler:
    '###
    ' Lee, 2017-11-03
    rTS.dispose_com
    Set rTS = Nothing
    
    For inx = 0 To no_of_ul - 1
        qTs(inx).dispose_com
        fTs(inx).dispose_com
        volTs(inx).dispose_com
        process(inx).dispose_com
        Set qTs(inx) = Nothing
        Set fTs(inx) = Nothing
        Set volTs(inx) = Nothing
        Set process(inx) = Nothing
    Next inx
    
    the_process_array.dispose_com
    Set the_process_array = Nothing
    
    the_engine.dispose_com
    Set the_engine = Nothing
    
    the_note.dispose_com
    Set the_note = Nothing
       
    Set backup_market = Nothing
    Set origin_pl_currency_curve = Nothing
    
    raise_err "run_ac_pricing_mc", Err.description

End Sub


Public Function get_ac_duration_2d(deal_ticket As clsACDealTicket, market As clsMarketSet) As Double
    
    Dim tmp_deal_ticket As clsACDealTicket
    Dim last_node() As Double
    Dim tmp_greeks As New clsGreeks
    
On Error GoTo ErrorHandler

    Set tmp_deal_ticket = deal_ticket.copy_obj
    
    tmp_deal_ticket.set_to_duration_mode 0.035
    
    'run_ac_calculation_2d tmp_greeks, tmp_deal_ticket, market
    Dim dummy_set As clsGreekSet
    run_ac_pricing_2d tmp_greeks, tmp_deal_ticket, market, dummy_set
        
    'End If
    
        
    get_ac_duration_2d = tmp_greeks.value / tmp_deal_ticket.notional
    
    Exit Function
    
ErrorHandler:


    raise_err "get_ac_duration", Err.description
    
End Function

Public Function get_ac_duration_3d(deal_ticket As clsACDealTicket, market As clsMarketSet) As Double
    
    Dim tmp_deal_ticket As clsACDealTicket
    Dim last_node() As Double
    Dim tmp_greeks As New clsGreeks
    
On Error GoTo ErrorHandler

    Set tmp_deal_ticket = deal_ticket.copy_obj
    
    tmp_deal_ticket.set_to_duration_mode 0.035
    
    Dim dummy_set As clsGreekSet
    'run_ac_calculation_3d tmp_greeks, tmp_deal_ticket, market
    run_ac_pricing_3d tmp_greeks, tmp_deal_ticket, market, dummy_set, , , , , , , , , , False
        
    'End If
    
        
    get_ac_duration_3d = tmp_greeks.value / tmp_deal_ticket.notional
    
    Exit Function
    
ErrorHandler:


    raise_err "get_ac_duration_3d"
    
End Function
Public Function get_ac_duration(deal_ticket As clsACDealTicket, market As clsMarketSet) As Double
    
    Dim tmp_deal_ticket As clsACDealTicket
    Dim last_node() As Double
    Dim tmp_greeks As New clsGreeks
    Dim bump_greek_set As New clsGreekSet
    
On Error GoTo ErrorHandler

    Set tmp_deal_ticket = deal_ticket.copy_obj
    
    tmp_deal_ticket.set_to_duration_mode 0.025
    
    'If tmp_deal_ticket.no_of_ul = 1 Then

    run_ac_pricing_1d tmp_greeks, tmp_deal_ticket, market, bump_greek_set, False, False, False, False, 0.001, False, False
        
    'End If
    
        
    get_ac_duration = tmp_greeks.value / tmp_deal_ticket.notional
    
    Exit Function
    
ErrorHandler:


    raise_err "get_ac_duration", Err.description
    
End Function
Public Function get_ac_call_prob_for_all_seq_2d(deal_ticket As clsACDealTicket, market As clsMarketSet) As Double()
    
    Dim tmp_deal_ticket As clsACDealTicket
    Dim last_node() As Double
    Dim tmp_greeks As New clsGreeks
    
    Dim inx As Integer
    Dim call_prob() As Double
    
On Error GoTo ErrorHandler

    ReDim call_prob(1 To deal_ticket.no_of_schedule) As Double

    For inx = 1 To deal_ticket.no_of_schedule
    
        call_prob(inx) = get_ac_call_prob_by_seq_2d(deal_ticket, market, inx)
        
    Next inx
    
        
    get_ac_call_prob_for_all_seq_2d = call_prob
    
    Exit Function
    
ErrorHandler:


    raise_err "get_ac_call_prob_for_all_seq_2d", Err.description
    
End Function
Public Function get_ac_call_prob_by_seq_2d(deal_ticket As clsACDealTicket, market As clsMarketSet, seq As Integer) As Double
    
    Dim tmp_deal_ticket As clsACDealTicket
    Dim last_node() As Double
    Dim tmp_greeks As New clsGreeks
        
    Dim call_prob As Double
    
On Error GoTo ErrorHandler

    call_prob = 0

    
    
    If seq <= deal_ticket.no_of_schedule Then
    
        Set tmp_deal_ticket = deal_ticket.copy_obj
    
        tmp_deal_ticket.set_to_call_prob_mode 0.035, seq
        
        'run_ac_calculation_2d tmp_greeks, tmp_deal_ticket, market
        Dim dummy_set As clsGreekSet
        run_ac_pricing_2d tmp_greeks, tmp_deal_ticket, market, dummy_set
        
        call_prob = call_prob + tmp_greeks.value / tmp_deal_ticket.notional
                
    End If
        
    get_ac_call_prob_by_seq_2d = call_prob
    
    Exit Function
    
ErrorHandler:


    raise_err "get_ac_call_prob_by_seq_2d", Err.description
    
End Function
Public Function get_ac_call_prob_2d(deal_ticket As clsACDealTicket, market As clsMarketSet, call_until As Date) As Double
    
    'Dim tmp_deal_ticket As clsACDealTicket
    Dim last_node() As Double
    Dim tmp_greeks As New clsGreeks
    
    Dim inx As Integer
    Dim call_prob As Double
    
On Error GoTo ErrorHandler

    call_prob = 0

  '  Set tmp_deal_ticket = deal_ticket.copy_obj
    
    For inx = 1 To deal_ticket.no_of_schedule
    
        If deal_ticket.call_dates()(inx) > deal_ticket.current_date _
        And deal_ticket.call_dates()(inx) <= call_until Then
            
            call_prob = call_prob + get_ac_call_prob_by_seq_2d(deal_ticket, market, inx)  'tmp_greeks.value / tmp_deal_ticket.notional
                    
        End If
            
    Next inx
        
    get_ac_call_prob_2d = call_prob
    
    Exit Function
    
ErrorHandler:


    raise_err "get_ac_duration", Err.description
    
End Function

Public Function ac_deal_ticket_to_clr_nd(ac_deal_ticket As clsACDealTicket) As SYPricerInterop.shared_ptr_AutocallableNoteND
    
    Dim rtn_obj As Object
    Dim thePayoff As Object
    Dim theExercise As Object
    Dim percentStrikes() As Double
    Dim percentStrikesAtMaturity() As Double
    Dim couponOnCalls() As Double
    Dim no_of_schedule As Integer
    Dim inx As Integer
    
    Set rtn_obj = New shared_ptr_AutocallableNoteND
    Set thePayoff = New shared_ptr_AutocallablePayoffND
    Set theExercise = New sy_shared_ptr_EuropeanExercise
     
    thePayoff.initializePayoff (1 - 2 * ac_deal_ticket.call_put) _
                             , ac_deal_ticket.percent_strikes_at_maturity(0) _
                             , ac_deal_ticket.coupon_at_maturity _
                             , ac_deal_ticket.abs_ki_barriers(0) _
                             , ac_deal_ticket.dummy_coupon _
                             , ac_deal_ticket.reference_prices(0) _
                             , ac_deal_ticket.put_strike _
                             , ac_deal_ticket.put_participation _
                             , ac_deal_ticket.put_additional_coupon _
                             , ac_deal_ticket.call_strike _
                             , ac_deal_ticket.call_participation _
                             , (ac_deal_ticket.ki_barrier_flag = 1) _
                             , (ac_deal_ticket.ki_touched_flag = 1) _
                             , ac_deal_ticket.strike_shift_at_maturity _
                             , ac_deal_ticket.no_of_ul _
                             , 0 _
                             , ac_deal_ticket.performance_type_at_maturity _
                             , ac_deal_ticket.floor_value _
                             , ac_deal_ticket.ki_adj_pct _
                             , ac_deal_ticket.ki_performance_type
                             
    theExercise.initializeExerciseInt ac_deal_ticket.maturity_date
    
'    rtn_obj.initializeNote thePayoff, theExercise _
'                         , ac_deal_ticket.call_dates(0) _
'                         , ac_deal_ticket.coupon_on_call_array(0) _
'                         , ac_deal_ticket.abs_strikes(0) _
'                         , ac_deal_ticket.coupon_dates(0), ac_deal_ticket.abs_coupon_barriers(0) _
'                         , (ac_deal_ticket.monthly_coupon_flag = 1), ac_deal_ticket.monthly_coupon_amt(0) _
'                         , ac_deal_ticket.ki_barrier_shift, ac_deal_ticket.ki_monitoring_freq
'    rtn_obj.initializeNote thePayoff, theExercise _
'                         , ac_deal_ticket.call_dates(0) _
'                         , ac_deal_ticket.coupon_on_call_array(0) _
'                         , ac_deal_ticket.abs_strikes(0) _
'                         , ac_deal_ticket.strike_shift(0) _
'                         , ac_deal_ticket.coupon_dates(0), ac_deal_ticket.abs_coupon_barriers(0) _
'                         , (ac_deal_ticket.monthly_coupon_flag = 1), ac_deal_ticket.monthly_coupon_amt(0) _
'                         , ac_deal_ticket.early_exit_flag, ac_deal_ticket.early_exit_touched_flag, ac_deal_ticket.abs_early_exit_barriers(0), ac_deal_ticket.early_exit_dates(0), ac_deal_ticket.early_exit_coupon_amt(0), ac_deal_ticket.early_exit_strike_shift(0) _
'                         , ac_deal_ticket.ki_barrier_shift, ac_deal_ticket.ki_monitoring_freq
'    rtn_obj.initializeNote thePayoff, theExercise _
'                         , ac_deal_ticket.call_dates(0) _
'                         , ac_deal_ticket.coupon_on_call_array(0) _
'                         , ac_deal_ticket.abs_strikes(0) _
'                         , ac_deal_ticket.strike_shift(0) _
'                         , ac_deal_ticket.performance_type(0) _
'                         , ac_deal_ticket.coupon_dates(0), ac_deal_ticket.abs_coupon_barriers(0) _
'                         , (ac_deal_ticket.monthly_coupon_flag = 1), ac_deal_ticket.monthly_coupon_amt(0) _
'                         , ac_deal_ticket.early_exit_flag, ac_deal_ticket.early_exit_touched_flag, ac_deal_ticket.abs_early_exit_barriers(0), ac_deal_ticket.early_exit_dates(0), ac_deal_ticket.early_exit_coupon_amt(0), ac_deal_ticket.early_exit_strike_shift(0) _
'                         , ac_deal_ticket.ki_barrier_shift, ac_deal_ticket.ki_monitoring_freq
'    rtn_obj.initializeNote thePayoff, theExercise _
'                         , ac_deal_ticket.call_dates(0) _
'                         , ac_deal_ticket.coupon_on_call_array(0) _
'                         , ac_deal_ticket.abs_strikes(0) _
'                         , ac_deal_ticket.strike_shift(0) _
'                         , ac_deal_ticket.performance_type(0) _
'                         , ac_deal_ticket.coupon_dates(0), ac_deal_ticket.abs_coupon_barriers(0) _
'                         , (ac_deal_ticket.monthly_coupon_flag = 1), ac_deal_ticket.monthly_coupon_amt(0) _
'                         , ac_deal_ticket.early_exit_flag, ac_deal_ticket.early_exit_touched_flags(0), ac_deal_ticket.abs_early_exit_barriers(0), ac_deal_ticket.early_exit_dates(0), ac_deal_ticket.early_exit_coupon_amt(0), ac_deal_ticket.early_exit_strike_shift(0) _
'                         , ac_deal_ticket.ki_barrier_shift, ac_deal_ticket.ki_monitoring_freq
'dll: 2018.8.8
'    rtn_obj.initializeNote thePayoff, theExercise _
'                         , ac_deal_ticket.value_date _
'                         , ac_deal_ticket.call_dates(0) _
'                         , ac_deal_ticket.coupon_on_call_array(0) _
'                         , ac_deal_ticket.abs_strikes(0) _
'                         , ac_deal_ticket.strike_shift(0) _
'                         , ac_deal_ticket.performance_type(0) _
'                         , ac_deal_ticket.coupon_dates(0), ac_deal_ticket.abs_coupon_barriers(0) _
'                         , (ac_deal_ticket.monthly_coupon_flag = 1), ac_deal_ticket.monthly_coupon_amt(0) _
'                         , ac_deal_ticket.early_exit_flag, ac_deal_ticket.early_exit_touched_flags(0), ac_deal_ticket.abs_early_exit_barriers(0), ac_deal_ticket.early_exit_dates(0), ac_deal_ticket.early_exit_coupon_amt(0), ac_deal_ticket.early_exit_strike_shift(0) _
'                         , ac_deal_ticket.early_exit_performance_types(0), ac_deal_ticket.early_exit_barrier_types(0) _
'                         , ac_deal_ticket.ki_barrier_shift, ac_deal_ticket.ki_monitoring_freq
'
'dll: 2021.11.22 ... ejectable 구조
'dll: 2023.05.31 ... time step 변경 (ac_deal_ticket.value_date 추가)
'dll: 2024.06.10 ... memory 구조 반영 (monthly_coupon_memory_flag, unpaid_coupon 추가)
    rtn_obj.initializeNote thePayoff, theExercise _
                         , CLng(ac_deal_ticket.value_date) _
                         , ac_deal_ticket.call_dates(0) _
                         , ac_deal_ticket.coupon_on_call_array(0) _
                         , ac_deal_ticket.abs_strikes(0) _
                         , ac_deal_ticket.strike_shift(0) _
                         , ac_deal_ticket.performance_type(0) _
                         , ac_deal_ticket.coupon_dates(0), ac_deal_ticket.abs_coupon_barriers(0) _
                         , (ac_deal_ticket.monthly_coupon_flag = 1), ac_deal_ticket.monthly_coupon_amt(0) _
                         , (ac_deal_ticket.monthly_coupon_memory_flag = 1), (ac_deal_ticket.unpaid_coupon = 0) _
                         , ac_deal_ticket.early_exit_flag, ac_deal_ticket.early_exit_touched_flags(0), ac_deal_ticket.abs_early_exit_barriers(0), ac_deal_ticket.early_exit_dates(0), ac_deal_ticket.early_exit_coupon_amt(0), ac_deal_ticket.early_exit_strike_shift(0) _
                         , ac_deal_ticket.early_exit_performance_types(0), ac_deal_ticket.early_exit_barrier_types(0) _
                         , ac_deal_ticket.ejectable_flag, ac_deal_ticket.ejected_ul_flag(0), ac_deal_ticket.ejected_event_flag(0), ac_deal_ticket.ejectable_order(0) _
                         , ac_deal_ticket.ki_barrier_shift, ac_deal_ticket.ki_monitoring_freq
                        
    Set ac_deal_ticket_to_clr_nd = rtn_obj

End Function


Public Function ac_deal_ticket_to_clr_2d(ac_deal_ticket As clsACDealTicket) As SYPricerInterop.shared_ptr_AutocallableNote2D
    Dim holiday_list__(0 To 0) As Long

    holiday_list__(0) = 42000
    
    Dim rtn_obj As Object
    Dim thePayoff As Object
    Dim theExercise As Object
    Dim percentStrikes() As Double
    Dim percentStrikesAtMaturity() As Double
    Dim couponOnCalls() As Double
    Dim no_of_schedule As Integer
    Dim inx As Integer
    
    Set rtn_obj = New shared_ptr_AutocallableNote2D
    Set thePayoff = New shared_ptr_AutocallablePayoffND
    Set theExercise = New sy_shared_ptr_EuropeanExercise
     
    thePayoff.initializePayoff (1 - 2 * ac_deal_ticket.call_put) _
                             , ac_deal_ticket.percent_strikes_at_maturity(0) _
                             , ac_deal_ticket.coupon_at_maturity _
                             , ac_deal_ticket.abs_ki_barriers(0) _
                             , ac_deal_ticket.dummy_coupon _
                             , ac_deal_ticket.reference_prices(0) _
                             , ac_deal_ticket.put_strike _
                             , ac_deal_ticket.put_participation _
                             , ac_deal_ticket.put_additional_coupon _
                             , ac_deal_ticket.call_strike _
                             , ac_deal_ticket.call_participation _
                             , (ac_deal_ticket.ki_barrier_flag = 1) _
                             , (ac_deal_ticket.ki_touched_flag = 1) _
                             , ac_deal_ticket.strike_shift_at_maturity _
                             , ac_deal_ticket.no_of_ul _
                             , 0 _
                             , ac_deal_ticket.performance_type_at_maturity _
                             , ac_deal_ticket.floor_value _
                             , ac_deal_ticket.ki_adj_pct _
                             , ac_deal_ticket.ki_performance_type
                             
                             
                             
    theExercise.initializeExerciseInt ac_deal_ticket.maturity_date
                             
                              
    
'    rtn_obj.initializeNote thePayoff, theExercise, ac_deal_ticket.value_date _
'                         , ac_deal_ticket.call_dates(0) _
'                         , ac_deal_ticket.coupon_on_call_array(0) _
'                         , ac_deal_ticket.abs_strikes(0) _
'                         , ac_deal_ticket.coupon_dates(0), ac_deal_ticket.abs_coupon_barriers(0) _
'                         , ac_deal_ticket.floating_coupon_dates(0), ac_deal_ticket.floating_fixing_values(0), (ac_deal_ticket.monthly_coupon_flag = 1), ac_deal_ticket.monthly_coupon_amt(0) _
'                         , (ac_deal_ticket.early_exit_flag = 1), ac_deal_ticket.abs_early_exit_barriers(0), ac_deal_ticket.early_exit_dates(0), ac_deal_ticket.early_exit_coupon_amt(0) _
'                         , ac_deal_ticket.ra_tenor, ac_deal_ticket.ra_cpn, holiday_list__, ac_deal_ticket.abs_ra_min(0), ac_deal_ticket.abs_ra_max(0), (ac_deal_ticket.ra_flag = 1) _
'                         , ac_deal_ticket.ki_barrier_shift, ac_deal_ticket.ki_monitoring_freq _
'                         , ac_deal_ticket.rate_spread, ac_deal_ticket.instrument_type, 0
    rtn_obj.initializeNote thePayoff, theExercise, ac_deal_ticket.value_date _
                         , ac_deal_ticket.call_dates(0) _
                         , ac_deal_ticket.coupon_on_call_array(0) _
                         , ac_deal_ticket.abs_strikes(0) _
                         , ac_deal_ticket.strike_shift(0) _
                         , ac_deal_ticket.performance_type(0) _
                         , ac_deal_ticket.coupon_dates(0), ac_deal_ticket.abs_coupon_barriers(0) _
                         , ac_deal_ticket.floating_coupon_dates(0), ac_deal_ticket.floating_fixing_values(0), (ac_deal_ticket.monthly_coupon_flag = 1), ac_deal_ticket.monthly_coupon_amt(0) _
                         , (ac_deal_ticket.early_exit_flag = 1), ac_deal_ticket.early_exit_touched_flag, ac_deal_ticket.abs_early_exit_barriers(0), ac_deal_ticket.early_exit_dates(0), ac_deal_ticket.early_exit_coupon_amt(0), ac_deal_ticket.early_exit_strike_shift(0) _
                         , ac_deal_ticket.ra_tenor, ac_deal_ticket.ra_cpn, holiday_list__, ac_deal_ticket.abs_ra_min(0), ac_deal_ticket.abs_ra_max(0), (ac_deal_ticket.ra_flag = 1) _
                         , ac_deal_ticket.ki_barrier_shift, ac_deal_ticket.ki_monitoring_freq _
                         , ac_deal_ticket.rate_spread, ac_deal_ticket.instrument_type, 0
                         
    
    Set ac_deal_ticket_to_clr_2d = rtn_obj
    
    

End Function

Public Function ac_deal_ticket_to_clr(ac_deal_ticket As clsACDealTicket) As SYPricerInterop.shared_ptr_AutocallableNote
    
    Dim holiday_list__(0 To 0) As Long

    holiday_list__(0) = 42000
    
    Dim rtn_obj As Object
    Dim thePayoff As Object
    Dim theExercise As Object
    Dim percentStrikes() As Double
    Dim percentStrikesAtMaturity() As Double
    Dim couponOnCalls() As Double
    Dim no_of_schedule As Integer
    Dim inx As Integer
    
    Set rtn_obj = New shared_ptr_AutocallableNote
    Set thePayoff = New shared_ptr_AutocallablePayoffND
    Set theExercise = New sy_shared_ptr_EuropeanExercise
     
    thePayoff.initializePayoff (1 - 2 * ac_deal_ticket.call_put) _
                             , ac_deal_ticket.percent_strikes_at_maturity(0) _
                             , ac_deal_ticket.coupon_at_maturity _
                             , ac_deal_ticket.abs_ki_barriers(0) _
                             , ac_deal_ticket.dummy_coupon _
                             , ac_deal_ticket.reference_prices(0) _
                             , ac_deal_ticket.put_strike _
                             , ac_deal_ticket.put_participation _
                             , ac_deal_ticket.put_additional_coupon _
                             , ac_deal_ticket.call_strike _
                             , ac_deal_ticket.call_participation _
                             , (ac_deal_ticket.ki_barrier_flag = 1) _
                             , (ac_deal_ticket.ki_touched_flag = 1) _
                             , ac_deal_ticket.strike_shift_at_maturity _
                             , ac_deal_ticket.no_of_ul _
                             , 0 _
                             , ac_deal_ticket.performance_type_at_maturity _
                             , ac_deal_ticket.floor_value _
                             , ac_deal_ticket.ki_adj_pct _
                             , ac_deal_ticket.ki_performance_type

                             
                             
    theExercise.initializeExerciseInt ac_deal_ticket.maturity_date
                             
                              
    
'    rtn_obj.initializeNote thePayoff, theExercise, ac_deal_ticket.value_date _
'                         , ac_deal_ticket.call_dates(0) _
'                         , ac_deal_ticket.coupon_on_calls(0) _
'                         , ac_deal_ticket.abs_strikes(0) _
'                         , ac_deal_ticket.coupon_dates(0), ac_deal_ticket.monthly_coupon_amt(0), ac_deal_ticket.percent_coupon_barriers(0), (ac_deal_ticket.monthly_coupon_flag = 1) _
'                         , ac_deal_ticket.ra_tenor, ac_deal_ticket.ra_cpn, holiday_list__, ac_deal_ticket.ra_min_percent * ac_deal_ticket.reference_price, ac_deal_ticket.ra_max_percent * ac_deal_ticket.reference_price, (ac_deal_ticket.ra_flag = 1) _
'                         , ac_deal_ticket.ki_barrier_shift, ac_deal_ticket.ki_monitoring_freq _
'                         , ac_deal_ticket.rate_spread, 0
    rtn_obj.initializeNote thePayoff, theExercise, ac_deal_ticket.value_date _
                         , ac_deal_ticket.call_dates(0) _
                         , ac_deal_ticket.coupon_on_call_array(0) _
                         , ac_deal_ticket.abs_strikes(0) _
                         , ac_deal_ticket.strike_shift(0) _
                         , ac_deal_ticket.coupon_dates(0), ac_deal_ticket.monthly_coupon_amt(0), ac_deal_ticket.abs_coupon_barriers(0), (ac_deal_ticket.monthly_coupon_flag = 1) _
                         , (ac_deal_ticket.early_exit_flag = 1), ac_deal_ticket.early_exit_barrier, ac_deal_ticket.early_exit_touched_flag, ac_deal_ticket.early_exit_dates(0), ac_deal_ticket.early_exit_coupon_amt(0), ac_deal_ticket.early_exit_strike_shift(0) _
                         , ac_deal_ticket.ra_tenor, ac_deal_ticket.ra_cpn, holiday_list__, ac_deal_ticket.ra_min_percent * ac_deal_ticket.reference_price, ac_deal_ticket.ra_max_percent * ac_deal_ticket.reference_price, (ac_deal_ticket.ra_flag = 1) _
                         , ac_deal_ticket.ki_barrier_shift, ac_deal_ticket.ki_monitoring_freq _
                         , ac_deal_ticket.rate_spread
                         
    
    Set ac_deal_ticket_to_clr = rtn_obj
    
    

End Function
'
'Private Sub run_ac_calculation(ByRef the_greeks As clsGreeks, deal_ticket As clsACDealTicket, market As clsMarket, last_node() As Double _
'                            , Optional flush_last_node As Long = 0, Optional ignore_smoothing As Boolean = False) ', Optional mid_day_greek As Boolean = False)
'
'    Dim success_fail As Long
'
'    Dim value As Double
'    Dim delta As Double
'    Dim gamma As Double
'
'    Dim vega As Double
'    Dim theta As Double
'    Dim tmp_strike_shift As Double
'    Dim tmp_ki_barrier As Double
'
'    Dim tmp_ki_touched As Integer
'
'    Dim tmp_mid_day_greek As Integer
'
'On Error GoTo ErrorHandler
'
'    If ignore_smoothing Then
'
'        tmp_strike_shift = 0
'
'    Else
'
'        tmp_strike_shift = deal_ticket.strike_shift
'
'    End If
'
'    If deal_ticket.mid_day_greek Then
'
'        tmp_mid_day_greek = 1
'
'    Else
'
'        tmp_mid_day_greek = 0
'
'    End If
'
'
'    If (deal_ticket.ki_barrier * deal_ticket.reference_price >= market.s_ And deal_ticket.call_put = 0) _
'    Or (deal_ticket.ki_barrier * deal_ticket.reference_price <= market.s_ And deal_ticket.call_put = 1) Then
'
'        tmp_ki_touched = 1
'
'    Else
'
'        tmp_ki_touched = deal_ticket.ki_touched_flag
'
'    End If
'
'    '--------------------------------------------------
'    ' KI Barrier. Refer to Kou.
'    '--------------------------------------------------
'
'
'    tmp_ki_barrier = deal_ticket.ki_barrier * Exp(-0.5826 * Sqr(market.heston_parameters_.v_long) * Sqr(deal_ticket.ki_monitoring_freq / 250))
'
''Dim tmp_array_1(1 To 1) As Long
''Dim tmp_array_2(1 To 1) As Double
''Dim tmp_var_1 As Double
''
''tmp_array_1(1) = 10000
''tmp_array_2(1) = 0.5
'
'
'   ' success_fail = heston_ac_greeks(value, delta, gamma, vega, theta _
'                 , market.S_ _
'                 , market.rate_curve_.rate_dates, market.rate_curve_.spread_dcf(deal_ticket.rate_spread) _
'                 , market.div_schedule_.get_div_dates, market.div_schedule_.get_divs, deal_ticket.rate_spread _
'                 , market.heston_parameters_.v_initial, market.heston_parameters_.lamda, market.heston_parameters_.v_long, market.heston_parameters_.eta, market.heston_parameters_.rho _
'                 , deal_ticket.t_grid, deal_ticket.x_grid, deal_ticket.v_grid _
'                 , deal_ticket.call_put, deal_ticket.strike_at_maturity, deal_ticket.coupon_at_maturity, deal_ticket.dummy_coupon, CLng(deal_ticket.maturity_date) _
'                 , tmp_ki_barrier, deal_ticket.reference_price, deal_ticket.ki_barrier_flag, deal_ticket.ki_touched_flag, deal_ticket.put_strike, deal_ticket.put_participation _
'                 , deal_ticket.call_dates, deal_ticket.percent_strikes, deal_ticket.coupon_on_calls, tmp_strike_shift _
'                 , CLng(deal_ticket.current_date), deal_ticket.scheme_type, last_node, flush_last_node _
'                 )
'
'    success_fail = heston_ac_greeks(value, delta, gamma, vega, theta _
'                 , market.s_ _
'                 , market.rate_curve_.rate_dates, market.rate_curve_.spread_dcf(deal_ticket.rate_spread) _
'                 , market.div_schedule_.get_div_dates, market.div_schedule_.get_divs, deal_ticket.rate_spread _
'                 , market.heston_parameters_.v_initial, market.heston_parameters_.lamda, market.heston_parameters_.v_long, market.heston_parameters_.eta, market.heston_parameters_.rho _
'                 , deal_ticket.t_grid, deal_ticket.x_grid, deal_ticket.v_grid _
'                 , deal_ticket.call_put, deal_ticket.strike_at_maturity, deal_ticket.coupon_at_maturity, deal_ticket.dummy_coupon, CLng(deal_ticket.maturity_date) _
'                 , tmp_ki_barrier, deal_ticket.reference_price, deal_ticket.ki_barrier_flag, tmp_ki_touched, deal_ticket.put_strike, deal_ticket.put_participation _
'                 , deal_ticket.call_dates, deal_ticket.percent_strikes, deal_ticket.coupon_on_calls, tmp_strike_shift _
'                 , CLng(deal_ticket.current_date), deal_ticket.scheme_type, last_node, flush_last_node _
'                 , deal_ticket.monthly_coupon_flag, deal_ticket.coupon_dates, deal_ticket.percent_coupon_barriers, deal_ticket.monthly_coupon_amount, tmp_mid_day_greek)
'
'    If success_fail = 1 Then
'
'        the_greeks.value = value * deal_ticket.notional
'        the_greeks.delta = delta * deal_ticket.notional
'        the_greeks.gamma = gamma * deal_ticket.notional
'        the_greeks.vega = vega * deal_ticket.notional
'        the_greeks.theta = theta * deal_ticket.notional
'
'
'    Else
'
'        raise_err "run_ac_calculation"
'
'    End If
'
'    market.heston_parameters_.rewind
'
'
'    Exit Sub
'
'ErrorHandler:
'
'    market.heston_parameters_.rewind
'
'    raise_err "run_ac_calculation", Err.description
'
'End Sub

'----------------------------------------------
' Sub: retrieve_a_deal
' DESC: RETRIEVE a deal from DB.
'-----------------------------------------------
Public Function retrieve_ac_deal(ByRef deal_ticket As clsACDealTicket) As Boolean

    Dim no_of_schedule As Integer
    Dim schedule() As Date
    Dim coupon_on_call() As Double
    Dim strike() As Double
    Dim data_found As Boolean
    
    Dim no_of_coupon_schedule As Integer
    Dim coupon_schedule() As Date
    Dim monthly_coupon_amount() As Double
    Dim coupon_barrier() As Double
    
    Dim no_of_floating_schedule As Integer
    Dim floating_schedule() As Date
    Dim fixing_value() As Double
    
    
    Dim no_of_dim As Integer
    
On Error GoTo ErrorHandler

    DBConnector
    
    no_of_dim = retrieve_ac_ul_cnt(deal_ticket.asset_code)
    
    If no_of_dim >= 1 Then
    
        deal_ticket.set_ul_dim no_of_dim
        retrieve_ac_deal_header_sql deal_ticket, deal_ticket.asset_code
        retrieve_ac_deal_ul_sql deal_ticket, deal_ticket.asset_code
        
        deal_ticket.current_date = config__.current_date_
    
        data_found = True
                    
        retrieve_ac_schedule_sql no_of_schedule, schedule, coupon_on_call, strike, deal_ticket.asset_code
       
        deal_ticket.set_schedule no_of_schedule, schedule, strike, coupon_on_call
       
    '   retrieve_fixing_schedule deal_ticket, deal_ticket.asset_code
    '-----
    ' In case monthly coupon
    '-----
        If retrieve_ac_coupon_schedule_sql(no_of_coupon_schedule, coupon_schedule, monthly_coupon_amount, coupon_barrier, deal_ticket.asset_code) > 0 Then
        
            deal_ticket.set_coupon_schedule no_of_coupon_schedule, coupon_schedule, coupon_barrier, monthly_coupon_amount
            'deal_ticket.monthly_coupon_amount = monthly_coupon_amount(1)
            deal_ticket.monthly_coupon_flag = 1
        
        Else
            ReDim coupon_schedule(1 To 1) As Date
            ReDim coupon_barrier(1 To 1) As Double
             ReDim monthly_coupon_amount(1 To 1) As Double
        
            deal_ticket.set_coupon_schedule 1, coupon_schedule, coupon_barrier, monthly_coupon_amount
        
        End If
        
    '-----
    ' In case swap
    '-----
        If retrieve_ac_floating_leg_schedule_sql(no_of_floating_schedule, floating_schedule, fixing_value, deal_ticket.asset_code) > 0 Then
        
            deal_ticket.set_floating_schedule no_of_floating_schedule, floating_schedule, fixing_value
             
        End If
        
        
    '-------------------------------
    ' 2015-10-19
    ' Term VEGA
    '-------------------------------
    Dim no_of_term As Integer
    Dim term_array() As Date
    
    no_of_term = shtACPricer.Range("no_of_term").value
    
'    For inx = 1 To no_of_term
'
'        push_back_date term_array, shtACPricer.Range("Term_Vega_Start").Cells(inx, 1).value
'
'    Next inx
    
    deal_ticket.set_term_vega_tenor config__.term_vega_tenor_array ' term_array '-----
    
        
        ''------- TEMP
        Dim ee_schedule(1 To 1) As Date
        Dim ee_value(1 To 1) As Double
        deal_ticket.set_early_exit_schedule 1, ee_schedule, ee_value
        
        '---------------------------
    
    Else
    
        data_found = False
    
    End If
    
    DBDisConnector
    
    retrieve_ac_deal = data_found
    
    Exit Function
    
ErrorHandler:

    DBDisConnector
    raise_err "retrieve_ac_deal", Err.description


End Function

'=======================
' 2014-06-26, Lee
'=======================
Private Function make_virtual_bond_schedule(ByRef swap_schedules() As clsSwapSchedule, a_deal As clsACDealTicket, Optional period As Integer = 3) As Integer
    
    Dim rtn_value As Integer
    Dim no_of_schedule As Integer
    Dim an_obj As clsSwapSchedule
    Dim following_payment As Date
    
On Error GoTo ErrorHandler
    
    no_of_schedule = 0
    
    following_payment = a_deal.settlement_date
    
    Do
        no_of_schedule = no_of_schedule + 1
        
        Set an_obj = New clsSwapSchedule
        
        an_obj.fixing_date = yyyymmdd_to_date(get_minusbusiness_day(Format(following_payment, "YYYYMMDD")))
        an_obj.start_date = following_payment
        an_obj.pay_date = yyyymmdd_to_date(get_elapsedmonths_day(Format(a_deal.settlement_date, "YYYYMMDD"), no_of_schedule * period))
        an_obj.end_date = an_obj.pay_date
        an_obj.fixing_rate = get_cd_rate(Format(an_obj.fixing_date, "YYYYMMDD"))
        an_obj.add_margin = a_deal.rate_spread
        
        push_back_clsSwapSchedule swap_schedules, an_obj
        
        following_payment = an_obj.pay_date
        
    Loop While an_obj.pay_date < a_deal.maturity_date
        
    
    make_virtual_bond_schedule = no_of_schedule
    
    Exit Function
    
ErrorHandler:

    raise_err "make_vitual_bond_schedule", Err.description

End Function


'<======================PATTERN
Public Sub kill_all_market_deal()


On Error GoTo ErrorHandler

    'deal_ticket_check deal_ticket
    
    DBConnector
    
    GConn_SPS_SPT.BeginTrans
        
    update_ac_status_dml "N", , "Y", "M"

    GConn_SPS_SPT.CommitTrans
    

    DBDisConnector
    
    Exit Sub
    
ErrorHandler:

    If GConn_SPS_SPT.State <> 0 Then
        GConn_SPS_SPT.RollbackTrans
    End If
    
    DBDisConnector
    
   raise_err "kill_all_market_deal", Err.description

End Sub

Public Sub insert_ac_deal(deal_ticket As clsACDealTicket)
    

On Error GoTo ErrorHandler

    'deal_ticket_check deal_ticket
    
    Dim swap_schedules() As clsSwapSchedule
    Dim no_of_swap_schedule As Integer
        
    DBConnector
    
    no_of_swap_schedule = make_virtual_bond_schedule(swap_schedules, deal_ticket)
    
    
    GConn_SPS_SPT.BeginTrans
    
'    dbconn
    GConn_SPT_ALTI.BeginTrans
    GConn_SPT_RMS01.BeginTrans
    
    insert_ac_deal_dml deal_ticket ', deal_sheet.Range("comment").Cells(1, 1).value
    insert_ac_underlying_dml deal_ticket
    insert_ac_schedule_dml deal_ticket
    
    
    insert_virtual_bond_schedule deal_ticket.asset_code, swap_schedules, no_of_swap_schedule
    
    If deal_ticket.monthly_coupon_flag Then
        insert_ac_coupon_schedule_dml deal_ticket
    End If
    
    If deal_ticket.no_of_floating_coupon_schedule > 0 Then
        insert_ac_floating_schedule_dml deal_ticket
    End If
    
    GConn_SPT_ALTI.CommitTrans
    GConn_SPS_SPT.CommitTrans
    GConn_SPT_RMS01.CommitTrans

    DBDisConnector
    
    Exit Sub
    
ErrorHandler:

    If GConn_SPS_SPT.State <> 0 Then
        GConn_SPS_SPT.RollbackTrans
        GConn_SPT_ALTI.RollbackTrans
        GConn_SPT_RMS01.RollbackTrans

    End If
    
    DBDisConnector
    
   raise_err "insert_ac_deal", Err.description

End Sub



Public Sub retrieve_ac_deals(ByRef deals() As clsACDealTicket, asset_codes() As String, deal_count As Long, Optional mid_day_greek As Boolean = False)

    Dim inx As Integer
    Dim calldate(1 To 1) As Date
    Dim strike_percent(1 To 1) As Double
    Dim dummy_cpn_on_call(1 To 1) As Double
    
On Error GoTo ErrorHandler

    If deal_count > 0 Then
    
        ReDim deals(LBound(asset_codes) To UBound(asset_codes)) As clsACDealTicket
        
        DBConnector
        
        ' Loop for the asset code list
        For inx = LBound(asset_codes) To UBound(asset_codes)
            
            Set deals(inx) = New clsACDealTicket
            
            deals(inx).asset_code = asset_codes(inx)
            
            'Load a deal
            retrieve_ac_deal deals(inx)
            
            deals(inx).current_date = config__.current_date_
            deals(inx).current_date_origin_ = config__.current_date_
            
            deals(inx).mid_day_greek = mid_day_greek
                        
            If deals(inx).monthly_coupon_flag = 0 Then
                deals(inx).set_coupon_schedule 1, calldate, strike_percent, dummy_cpn_on_call
            End If
                        
    
        Next inx
        
        DBDisConnector
        
    End If
    
    Exit Sub
    
ErrorHandler:

    DBDisConnector

    raise_err "retrieve_ac_deals", Err.description
        

End Sub


Private Sub fill_ac_deal_info(ByRef the_deal As clsACDealTicket)

    Dim no_of_schedule As Integer
    Dim no_of_coupon_schedule As Integer
    Dim schedule() As Date
    Dim strike() As Double
    Dim coupon_on_call() As Double
    Dim coupon_schedule() As Date
    Dim monthly_coupon_amount() As Double
    Dim coupon_barrier() As Double
    
    
On Error GoTo ErrorHandler

    retrieve_ac_deal_ul_sql the_deal, the_deal.asset_code
    
    the_deal.current_date = config__.current_date_
    the_deal.current_date_origin_ = config__.current_date_
    
    retrieve_ac_schedule_sql no_of_schedule, schedule, coupon_on_call, strike, the_deal.asset_code
    
'-----
' 2015-11-18 TEMP
' Strike Shift for 3 index
'-----
'    Dim inx As Integer
'    If the_deal.no_of_ul >= 3 Then
'
'        For inx = 1 To get_array_size_double(strike)
'
'            strike(inx) = strike(inx) - 0.001
'
'        Next inx
'
'    End If


    the_deal.set_schedule no_of_schedule, schedule, strike, coupon_on_call
    
    If retrieve_ac_coupon_schedule_sql(no_of_coupon_schedule, coupon_schedule, monthly_coupon_amount, coupon_barrier, the_deal.asset_code) > 0 Then
    
        the_deal.set_coupon_schedule no_of_coupon_schedule, coupon_schedule, coupon_barrier, monthly_coupon_amount
        the_deal.monthly_coupon_amount = monthly_coupon_amount(1)
        the_deal.monthly_coupon_flag = 1
    
    Else
        ReDim coupon_schedule(1 To 1) As Date
        ReDim coupon_barrier(1 To 1) As Double
        ReDim monthly_coupon_amount(1 To 1) As Double
    
        the_deal.set_coupon_schedule 1, coupon_schedule, coupon_barrier, monthly_coupon_amount
    
    End If
    
    Exit Sub
    
ErrorHandler:

    raise_err "fill_ac_deal_info", Err.description


End Sub

Private Sub reorder_ac_deal(sorted_deals() As clsACDealTicket, deals() As clsACDealTicket)
    
    
    Dim index_max As Integer
    Dim index_min As Integer
    
    Dim temp As Date
    
    Dim inx As Integer
    Dim jnx As Integer
    
    
On Error GoTo ErrorHandler


    index_max = UBound(deals)
    index_min = LBound(deals)
    
    ReDim seq(index_min To index_max) As Integer
    ReDim reorder_seq(index_min To index_max) As Integer
    ReDim SwapArray(index_min To index_max) As Date
    
    ReDim sorted_deals(index_min To index_max) As clsACDealTicket
    
    For inx = index_min To index_max
    
        SwapArray(inx) = deals(inx).get_next_call_date(config__.current_date_)
        seq(inx) = inx
        reorder_seq(inx) = inx
    
    Next inx
    
    For inx = index_min + 1 To index_max
    
        temp = deals(inx).get_next_call_date(config__.current_date_)
        
        For jnx = inx - 1 To index_min Step -1
    
            If SwapArray(jnx) > temp Then
            
                SwapArray(jnx + 1) = SwapArray(jnx)
                SwapArray(jnx) = temp
                
                reorder_seq(jnx + 1) = reorder_seq(jnx)
                reorder_seq(jnx) = inx
                
            End If
    
        Next jnx
        
    Next inx
    
'    If LCase(DataOdering) = LCase("ASC") Then
        
        For inx = index_min To index_max
        
            Set sorted_deals(inx) = deals(reorder_seq(inx))
        
        Next inx
'
'    Else
'
'        For inx = index_min To index_max
'
'            sorted_deals(inx) = deals(reorder_seq(index_max - inx + index_min))
'
'        Next inx
'
'    End If
        


    Exit Sub
    
ErrorHandler:


    raise_err "recorder_ac_deal", Err.description


End Sub


Public Sub retrieve_ac_deal_list(deals() As clsACDealTicket, greeks() As clsGreeks, LIVE_YN As String, CONFIRM_YN As String, Optional ul_cnt As Integer = 0, Optional exclude_intraday As Boolean = False)

    Dim inx As Integer
    Dim jnx As Integer
    Dim knx As Integer
    Dim msg_str As String
    
    Dim no_of_schedule As Integer
    Dim no_of_coupon_schedule As Integer
    Dim schedule() As Date
    Dim strike() As Double
    
    Dim coupon_schedule() As Date
    Dim coupon_on_call() As Double
    Dim monthly_coupon_amount() As Double
    Dim coupon_barrier() As Double
    
    Dim deals_pre_sorting() As clsACDealTicket
    
On Error GoTo ErrorHandler

    If Not initialized__ Then
    
        Err.Raise vbObjectError + 10000, , "[PRO] Not initialized!!"
        
    End If
    
    
    '---------------------------------
    ' DB Access
    DBConnector
        
    '----------------------------
    ' Retrieves Deal list from the DB according to search condition.
    '----------------------------
    
    If (retrieve_ac_deal_header_list_sql(deals_pre_sorting, LIVE_YN, CONFIRM_YN, ul_cnt, config__.current_date_, exclude_intraday, config__.adjust_strike_shift_percent)) > 0 Then   ' Retrieve Header List

        
        ReDim ac_greek_cache__(LBound(deals_pre_sorting) To UBound(deals_pre_sorting)) As clsGreekCache
        ReDim greeks(LBound(deals_pre_sorting) To UBound(deals_pre_sorting)) As clsGreeks
        
        '--------------------------------------
        ' Fill Deal information
        '--------------------------------------
        For inx = LBound(deals_pre_sorting) To UBound(deals_pre_sorting)
            fill_ac_deal_info deals_pre_sorting(inx)
        Next inx
        
        '--------------------------------------
        ' Reorder according to call date
        '--------------------------------------
        reorder_ac_deal deals, deals_pre_sorting
        
        '--------------------------------------
        ' Fill previous date Greek information
        '--------------------------------------
        For inx = LBound(deals) To UBound(deals)
            
            Set greeks(inx) = New clsGreeks
            
            
            If deals(inx).no_of_ul = 1 Then
                greeks(inx).value = retrieve_greek(deals(inx).asset_code, config__.last_date_, "VALUE")
                greeks(inx).delta = retrieve_greek(deals(inx).asset_code, config__.last_date_, "DELTA")
                greeks(inx).gamma = retrieve_greek(deals(inx).asset_code, config__.last_date_, "GAMMA")
                greeks(inx).vega = retrieve_greek(deals(inx).asset_code, config__.last_date_, "VEGA") ' / 100  <---- 2014-08-18
                greeks(inx).theta = retrieve_greek(deals(inx).asset_code, config__.last_date_, "THETA")
                greeks(inx).skew_s = retrieve_greek(deals(inx).asset_code, config__.last_date_, "SKEW")
                greeks(inx).rho = retrieve_greek(deals(inx).asset_code, config__.last_date_, "RHO")
                greeks(inx).duration = retrieve_greek(deals(inx).asset_code, config__.last_date_, "DURATION")
                greeks(inx).vanna = retrieve_greek(deals(inx).asset_code, config__.last_date_, "VANNA") '* 100
                greeks(inx).conv_s = retrieve_greek(deals(inx).asset_code, config__.last_date_, "CONV") '* 100
                
                
                greeks(inx).implied_tree_delta = greeks(inx).delta
                greeks(inx).implied_tree_gamma = greeks(inx).gamma
                
            Else
                
                greeks(inx).no_of_tenors = retrieve_no_of_tenors(deals(inx).asset_code, config__.last_date_)
                                
                greeks(inx).redim_arrays deals(inx).no_of_ul, greeks(inx).no_of_tenors
                
                
                greeks(inx).value = retrieve_greek(deals(inx).asset_code, config__.last_date_, "VALUE")
                greeks(inx).theta = retrieve_greek(deals(inx).asset_code, config__.last_date_, "THETA")
                greeks(inx).rho = retrieve_greek(deals(inx).asset_code, config__.last_date_, "RHO")
                greeks(inx).duration = retrieve_greek(deals(inx).asset_code, config__.last_date_, "DURATION")
                
                '----
                ' Appended on 2014-08-18
                '----
                
                
                For jnx = 1 To deals(inx).no_of_ul
                     
                    greeks(inx).set_deltas jnx, retrieve_greek(deals(inx).asset_code, config__.last_date_, "DELTA", , deals(inx).ul_code(jnx))
                    greeks(inx).set_gammas jnx, retrieve_greek(deals(inx).asset_code, config__.last_date_, "GAMMA", , deals(inx).ul_code(jnx))
                    greeks(inx).set_vegas jnx, retrieve_greek(deals(inx).asset_code, config__.last_date_, "VEGA", , deals(inx).ul_code(jnx)) ' / 100 <---- 2014-08-18
                    greeks(inx).set_skews jnx, retrieve_greek(deals(inx).asset_code, config__.last_date_, "SKEW", , deals(inx).ul_code(jnx))
                    greeks(inx).set_vannas jnx, retrieve_greek(deals(inx).asset_code, config__.last_date_, "VANNA", , deals(inx).ul_code(jnx)) '* 100
                    greeks(inx).set_rho_ul jnx, retrieve_greek(deals(inx).asset_code, config__.last_date_, "RHO_UL", , deals(inx).ul_code(jnx))  '* 100
                    greeks(inx).set_convs jnx, retrieve_greek(deals(inx).asset_code, config__.last_date_, "CONV", , deals(inx).ul_code(jnx))
                    
                    greeks(inx).set_all_implied_tree_deltas greeks(inx).get_all_deltas
                    greeks(inx).set_all_implied_tree_gammas greeks(inx).get_all_gammas
                
                    For knx = jnx + 1 To deals(inx).no_of_ul
                        greeks(inx).set_corr_sens jnx, knx, retrieve_cross_greek(deals(inx).asset_code, config__.last_date_, "CORR", , deals(inx).ul_code(jnx), deals(inx).ul_code(knx))
                    Next knx
                    
                    
                    greeks(inx).set_term_dates_per_ul jnx, retrieve_term_dates(deals(inx).asset_code, config__.last_date_)
                    
                    greeks(inx).set_term_skews_per_ul jnx, retrieve_term_greek(deals(inx).asset_code, config__.last_date_, "TERM_SKEW", deals(inx).ul_code(jnx))
                    greeks(inx).set_term_vegas_per_ul jnx, retrieve_term_greek(deals(inx).asset_code, config__.last_date_, "TERM_VEGA", deals(inx).ul_code(jnx))
                    greeks(inx).set_term_convs_per_ul jnx, retrieve_term_greek(deals(inx).asset_code, config__.last_date_, "TERM_CONV", deals(inx).ul_code(jnx))
                    
                    
                Next jnx

            End If
            
            '----
            ' Appended on 2014-08-18
            '----
            calc_sticky_strike_delta greeks(inx), greeks(inx).duration, deals(inx).no_of_ul, deals(inx).get_ul_codes(), config__.last_date_, prev_market_set__
            calc_sticky_strike_gamma greeks(inx), greeks(inx).duration, deals(inx).no_of_ul, deals(inx).get_ul_codes(), config__.last_date_, prev_market_set__
            
            deals(inx).duration = greeks(inx).duration
            
            theta_adjustment greeks(inx), config__.current_date_, config__.next_date_
            
            Set ac_greek_cache__(inx) = New clsGreekCache
            ac_greek_cache__(inx).asset_code = deals(inx).asset_code
                        
        
        Next inx
    
    End If


    DBDisConnector
    
    If msg_str <> "" Then
    
        MsgBox msg_str
        
    End If
    
    
    Exit Sub
    
ErrorHandler:

    If UCase(Left(Err.description, 5)) = "[BIZ]" Then
    
        msg_str = msg_str & deals(inx).asset_code & ":" & Err.description & Chr(13)
        Resume Next
        
    End If

    DBDisConnector
    raise_err "retrieve_ac_deal_list", Err.description
    

End Sub

Public Sub retrieve_3d_deal(deals_3d() As clsACDealTicket, greeks() As clsGreeks)
    
    
    Dim asset_codes() As String
    Dim ac_deal_count As Long
    Dim inx As Integer
    
On Error GoTo ErrorHandler
        
        
    ac_deal_count = retrieve_ac_asset_code_list(asset_codes, "Y", "Y", 3)
        
    retrieve_ac_deals deals_3d, asset_codes, ac_deal_count

    retrieve_duration deals_3d, ac_deal_count
    
    
    ReDim greeks(1 To ac_deal_count) As clsGreeks
    
    retrieve_prev_day_greek deals_3d, greeks, ac_deal_count
    
    Exit Sub
    
ErrorHandler:
    
    raise_err "retrieve_3d_deal", Err.description
    

End Sub


Public Sub retrieve_duration(deals_3d() As clsACDealTicket, ac_deal_count As Long)

    Dim inx As Integer
    
On Error GoTo ErrorHandler

    DBConnector
        
    
    For inx = 1 To ac_deal_count
        
        deals_3d(inx).duration = retrieve_greek(deals_3d(inx).asset_code, config__.last_date_, "DURATION")
    
    Next inx
    
    
    DBDisConnector
    
    Exit Sub
    
ErrorHandler:

    DBDisConnector

    raise_err "retrieve_duration", Err.description


End Sub




Public Sub retrieve_prev_day_greek(deals_3d() As clsACDealTicket, greeks() As clsGreeks, ac_deal_count As Long)

    Dim inx As Integer
    Dim jnx As Integer
    Const no_of_ul As Integer = 3
    
On Error GoTo ErrorHandler

    DBConnector
        
    
    For inx = 1 To ac_deal_count
    
        Set greeks(inx) = New clsGreeks
        greeks(inx).redim_arrays 3
        
        For jnx = 1 To no_of_ul
    
            greeks(inx).set_vegas jnx, retrieve_greek(deals_3d(inx).asset_code, config__.last_date_, "VEGA", , deals_3d(inx).ul_code(jnx)) ' / 100 <---- 2014-08-18
            greeks(inx).set_skews jnx, retrieve_greek(deals_3d(inx).asset_code, config__.last_date_, "SKEW", , deals_3d(inx).ul_code(jnx))
            greeks(inx).set_vannas jnx, retrieve_greek(deals_3d(inx).asset_code, config__.last_date_, "VANNA", , deals_3d(inx).ul_code(jnx)) '* 100
        
        Next jnx
        
    Next inx
    
    
    DBDisConnector
    
    Exit Sub
    
ErrorHandler:

    DBDisConnector
    raise_err "retrieve_duration", Err.description


End Sub
'
'
'Public Sub file_to_greek_realtime(ByRef greeks() As clsGreeks, deals() As clsACDealTicket, Optional eval_date_in As Date = -1)
'
'    Dim inx As Integer
'    Dim deltas() As Double
'    Dim gammas() As Double
'    Dim implied_tree_deltas() As Double
'    Dim sticky_strike_deltas() As Double
'    Dim implied_tree_gammas() As Double
'    Dim sticky_strike_gammas() As Double
'
'    Dim rtn_value As Boolean
'
'On Error GoTo ErrorHandler
'
'    Dim file_system As Variant
'    Dim txt_file As Variant
'    Dim line_str As String
'    Dim line_str_array() As String
'    Dim aGreek As clsGreeks
'    Dim counter As Integer
'
'    rtn_value = False
'
'    ReDim implied_tree_deltas(1 To no_of_ul) As Double
'    ReDim sticky_strike_deltas(1 To no_of_ul) As Double
'    ReDim implied_tree_gammas(1 To no_of_ul) As Double
'    ReDim sticky_strike_gammas(1 To no_of_ul) As Double
'
'    Set file_system = CreateObject("Scripting.FileSystemObject")
'    Set txt_file = file_system.OpenTextFile(path_name & file_name, 1)
'
'    Do While txt_file.AtEndOfStream = False
'
'        line_str = txt_file.ReadLine
'        line_str_array = Split(CStr(line_str), "*")
'
'        counter = 0
'
'        If line_str_array(0) = asset_code Then
'
'            greeks.redim_arrays no_of_ul
'
'            counter = counter + 1
'            greeks.value = CDbl(Trim(line_str_array(counter)))
'
'            For inx = 1 To no_of_ul
'                counter = counter + 1
'                implied_tree_deltas(inx) = CDbl(Trim(line_str_array(counter)))
'                counter = counter + 1
'                sticky_strike_deltas(inx) = CDbl(Trim(line_str_array(counter)))
'                counter = counter + 1
'                implied_tree_gammas(inx) = CDbl(Trim(line_str_array(counter)))
'                counter = counter + 1
'                sticky_strike_gammas(inx) = CDbl(Trim(line_str_array(counter)))
'            Next inx
'
'            greeks.set_all_deltas implied_tree_deltas
'            greeks.set_all_gammas implied_tree_gammas
'
'            greeks.set_all_implied_tree_deltas implied_tree_deltas
'            greeks.set_all_implied_tree_gammas implied_tree_gammas
'
'            greeks.set_all_sticky_strike_deltas sticky_strike_deltas
'            greeks.set_all_sticky_strike_gammas sticky_strike_gammas
'
'            Exit Do
'
'            rtn_value = True
'
'        End If
'
'    Loop
'
'
'    txt_file.Close
'
'    file_to_greek_3d = rtn_value
'
'    Exit Function
'
'ErrorHandler:
'
'    file_to_greek_3d = False
'
'    If txt_file <> Null Then
'        txt_file.Close
'    End If
'
'    Exit Function
'
'
'End Sub

Private Sub read_position_file_to_array(ByRef asset_codes() As String, ByRef data_array() As clsDoubleArray, PATH_NAME As String, file_name As String)
    
    Dim inx As Integer
    
On Error GoTo ErrorHandler

    Dim file_system As Variant
    Dim txt_file As Variant
    
    Dim line_str As String
    Dim line_str_array() As String
    
    Dim tmp_array() As Double
    Dim aDataArray As clsDoubleArray
    
    Erase asset_codes
    Erase data_array
    
    Set file_system = CreateObject("Scripting.FileSystemObject")
    Set txt_file = file_system.OpenTextFile(PATH_NAME & file_name, 1)
    
    Dim counter As Integer
    
    line_str = txt_file.ReadLine '<-- Dummy. The first line contains the prices of underliers.
    
    Do While txt_file.AtEndOfStream = False
    
    
        Erase line_str_array
        Erase tmp_array
    
        line_str = txt_file.ReadLine
        line_str_array = Split(CStr(line_str), "*")
            
        push_back_string asset_codes, line_str_array(0)
        
        
        For inx = 1 To UBound(line_str_array)
            
            push_back_double tmp_array, CDbl(Trim(line_str_array(inx)))
        
        Next inx
        
        Set aDataArray = New clsDoubleArray
        aDataArray.set_array tmp_array
        
        push_back_clsDoubleArray data_array, aDataArray
        
    Loop
    

    Exit Sub
    
ErrorHandler:

    If Err.number = 76 Then
        MsgBox "no file found"
        Exit Sub
    Else
        raise_err "read_position_file_to_array", Err.description
    End If


End Sub

Private Function find_asset_code_sequence(asset_codes() As String, asset_code As String) As Integer
    
    Dim inx As Integer
    Dim rtn_value As Integer
    
    
On Error GoTo ErrorHandler

    rtn_value = 0
    
    
    For inx = LBound(asset_codes) To UBound(asset_codes)
    
        If asset_codes(inx) = asset_code Then
        
            rtn_value = inx
        
            Exit For
            
        End If
        
    Next inx
        
    find_asset_code_sequence = rtn_value
    
    Exit Function
    
ErrorHandler:

    find_asset_code_sequence = 0

End Function

Private Sub parse_greeks_data(ByRef aGreek As clsGreeks, data_array() As Double, no_of_ul As Integer)


    Dim deltas() As Double
    Dim gammas() As Double
    Dim implied_tree_deltas() As Double
    Dim sticky_strike_deltas() As Double
    Dim implied_tree_gammas() As Double
    Dim sticky_strike_gammas() As Double
    Dim vegas() As Double
    Dim skews() As Double
    Dim vannas() As Double
    
    Dim inx As Integer
    Dim counter As Integer
    
    Dim data_size As Integer
        
On Error GoTo ErrorHandler

    counter = 1
    
    data_size = get_array_size_double(data_array)
    
    ReDim implied_tree_deltas(1 To no_of_ul) As Double
    ReDim sticky_strike_deltas(1 To no_of_ul) As Double
    ReDim implied_tree_gammas(1 To no_of_ul) As Double
    ReDim sticky_strike_gammas(1 To no_of_ul) As Double
    ReDim vegas(1 To no_of_ul) As Double
    ReDim skews(1 To no_of_ul) As Double
    ReDim vannas(1 To no_of_ul) As Double
    
    ReDim rho_ul(1 To no_of_ul) As Double  ' 2016-01-29
    ReDim eff_dur(1 To no_of_ul) As Double  ' 2016-01-29
    
    aGreek.value = data_array(counter)

    For inx = 1 To no_of_ul
        
        counter = counter + 1
        
        implied_tree_deltas(inx) = data_array(counter)
        counter = counter + 1
        sticky_strike_deltas(inx) = data_array(counter)
        counter = counter + 1
        implied_tree_gammas(inx) = data_array(counter)
        counter = counter + 1
        sticky_strike_gammas(inx) = data_array(counter)
        counter = counter + 1
        vegas(inx) = data_array(counter)
        counter = counter + 1
        skews(inx) = data_array(counter)
        counter = counter + 1
        vannas(inx) = data_array(counter)
        
        If data_size > counter Then
            counter = counter + 1
            rho_ul(inx) = data_array(counter)
        End If
        
        If data_size > counter Then
            counter = counter + 1
            eff_dur(inx) = data_array(counter)
        End If
    
    Next inx
    
    counter = counter + 1
    aGreek.rho = data_array(counter)

    aGreek.set_all_deltas implied_tree_deltas
    aGreek.set_all_gammas implied_tree_gammas
    
    aGreek.set_all_implied_tree_deltas implied_tree_deltas
    aGreek.set_all_implied_tree_gammas implied_tree_gammas
    
    aGreek.set_all_sticky_strike_deltas sticky_strike_deltas
    aGreek.set_all_sticky_strike_gammas sticky_strike_gammas
    
    aGreek.set_all_vegas vegas
    aGreek.set_all_skews skews
    aGreek.set_all_vannas vannas
    
    aGreek.set_all_rho_ul rho_ul
    
    aGreek.set_all_eff_durations eff_dur
    
    'agreek.rho =

    Exit Sub
    
ErrorHandler:

    raise_err "parse_greeks_data", Err.description
    
End Sub

Public Sub file_to_greeks(ByRef greeks() As clsGreeks, deals() As clsACDealTicket, PATH_NAME As String, file_name As String)

    Dim inx As Integer
    
    Dim asset_codes() As String
    Dim data_array() As clsDoubleArray
    
    Dim asset_seq As Integer
        
On Error GoTo ErrorHandler
    
    Dim counter As Integer
  
    read_position_file_to_array asset_codes, data_array, PATH_NAME, file_name
    
    For inx = 1 To get_array_size_clsAcDealTicket(deals)
    
        Set greeks(inx) = New clsGreeks
        greeks(inx).redim_arrays deals(inx).no_of_ul
        
        asset_seq = find_asset_code_sequence(asset_codes, deals(inx).asset_code)
        
        If asset_seq > 0 Then
            
            parse_greeks_data greeks(inx), data_array(asset_seq).get_array(), deals(inx).no_of_ul
        
        End If
            
    
    Next inx
        
    
    Exit Sub
    
ErrorHandler:

    raise_err "file_to_greeks", Err.description

    Exit Sub
    

End Sub

Public Sub greek_to_file(greeks() As clsGreeks, deals() As clsACDealTicket, PATH_NAME As String, file_name As String)

On Error GoTo ErrorHandler
    
    Dim file_system As Variant
    Dim txt_file As Variant
    Dim line_str As String
    Dim inx As Integer
    Dim jnx As Integer
    
    
    Set file_system = CreateObject("Scripting.FileSystemObject")
    Set txt_file = file_system.CreateTextFile(PATH_NAME & file_name, True)
    
    line_str = market_set__.market_by_ul("KOSPI200").s_
    line_str = line_str & "*" & market_set__.market_by_ul("SPX").s_
    line_str = line_str & "*" & market_set__.market_by_ul("SX5E").s_
    line_str = line_str & "*" & market_set__.market_by_ul("NKY").s_
    line_str = line_str & "*" & market_set__.market_by_ul("HSCEI").s_
    
    line_str = line_str & "*" & market_set__.get_fx_rate("USDKRW")
    line_str = line_str & "*" & market_set__.get_fx_rate("EURKRW")
    line_str = line_str & "*" & market_set__.get_fx_rate("JPYKRW")
    line_str = line_str & "*" & market_set__.get_fx_rate("HKDKRW")
    
    txt_file.writeline line_str
    
    For inx = 1 To get_array_size_clsgreeks(greeks)
        
        line_str = ""
        line_str = line_str & deals(inx).asset_code
        
        line_str = line_str & "*" & greeks(inx).value
        
        For jnx = 1 To deals(inx).no_of_ul
    
            line_str = line_str & "*" & greeks(inx).implied_tree_deltas(jnx)
            line_str = line_str & "*" & greeks(inx).sticky_strike_deltas(jnx)
            line_str = line_str & "*" & greeks(inx).implied_tree_gammas(jnx)
            line_str = line_str & "*" & greeks(inx).sticky_strike_gammas(jnx)
            line_str = line_str & "*" & greeks(inx).vegas(jnx)
            line_str = line_str & "*" & greeks(inx).skew_s_s(jnx)
            line_str = line_str & "*" & greeks(inx).vannas(jnx)
            line_str = line_str & "*" & greeks(inx).rho_ul(jnx)
            line_str = line_str & "*" & greeks(inx).eff_duration(jnx)
            
        Next jnx
        
        line_str = line_str & "*" & greeks(inx).rho
        
        txt_file.writeline line_str
        
    Next inx
    
    txt_file.Close
    

    Exit Sub
    
ErrorHandler:

    raise_err "greek_to_file_3d", Err.description

End Sub


'----------------------------------------------------
' refresh_current_greek_ac
' retrieve_current_greek :Read cache greeks. If it is not initialized, trigger initialization
' initialize_cache_greek
' retreive_cache_greek : read db.
'
Public Sub refresh_current_greek_AC(ByRef greeks() As clsGreeks, deals() As clsACDealTicket, Optional end_of_day_greek As Boolean = False, _
                                    Optional intraday_greek As String = "", Optional error_recovery As Boolean = False, Optional vol_bump As Boolean = False, Optional bumping_this As String = "")


On Error GoTo ErrorHandler
    
    Dim intraday_greek_type As String
    
    If intraday_greek = "" Then
        intraday_greek_type = config__.intra_day_greek_
    Else
        intraday_greek_type = intraday_greek
    End If
    
    If Not initialized__ Then
    
        Err.Raise vbObjectError + 10000, , "[PRO] Not initialized!!"
        
    End If
    
    If end_of_day_greek Then
        
        retrieve_end_of_day_greek_ac greeks, deals
        
    Else
       
       ReDim greeks(1 To get_array_size_clsAcDealTicket(deals)) As clsGreeks
        
        If intraday_greek_type = "DB" Then
        
            file_to_greeks greeks, deals, config__.file_path_, "position_" & Format(config__.current_date_, "YYYYMMDD") & "." & config__.position_file_extension
            
        ElseIf intraday_greek_type = "Distributor" Then
        
            '---------------------------------------------------
            ' Firstly, read text file for 2D
            '---------------------------------------------------
            read_greek_file greeks, deals, config__.current_date_, vol_bump, bumping_this
            
            If Not error_recovery Then
                Set GResultHandler = New clsACResultHandlerMulti
                GResultHandler.initialize deals
                GResultHandler.set_greeks greeks
            End If
            
                        
            make_greek_to_distribute_dist GResultHandler, "DISTRIBUTOR", error_recovery
        
        ElseIf intraday_greek_type = "Calculate" Then
        
            If Not error_recovery Then
                 Set GResultHandler = New clsACResultHandlerMulti
                ' calculate_greeks_ac greeks, deals
                
                 GResultHandler.initialize deals
            End If
            
            calculate_greeks_ac_dist GResultHandler, "FULLCALC", , True, False, , True, , , , , , True, , , , error_recovery

        End If
            
    End If

    Exit Sub
    
ErrorHandler:

    raise_err "refresh_current_greek_AC", Err.description

End Sub

Public Sub calculate_greeks_ac(ByRef greeks() As clsGreeks, deals() As clsACDealTicket)

    Dim dummy_intraday_greeks() As clsGreeks
    Dim dummy_bump_greeks_sets() As clsGreekSet

On Error GoTo ErrorHandler

    '------------------------
    ' Changing config
    '----------------------------
    Dim origin_x_grid As Integer
    Dim origin_t_steps_per_day As Double
    Dim origin_no_of_trials As Integer
    Dim inx As Integer
    Dim jnx As Integer
    
    origin_x_grid = config__.x_grid_
    'config__.x_grid_ = config__.x_grid_ / 2
    
    origin_t_steps_per_day = config__.time_step_per_day
    '  config__.time_step_per_day = config__.time_step_per_day / 4
    
    origin_no_of_trials = config__.no_of_trials_closing_
    '  config__.no_of_trials_closing_ = (config__.no_of_trials_closing_ + 1) / 4 - 1
    
    '------------------------
    ' Calculation
    '----------------------------
    calculate_ac_greeks greeks, dummy_intraday_greeks, deals, get_array_size_clsAcDealTicket(deals), market_set__, , "N", , False, False, False, False, False
    calculate_ac_greeks_2D greeks, deals, get_array_size_clsAcDealTicket(deals), market_set__, dummy_bump_greeks_sets, False, False, False, False, , , False, , False
    calculate_ac_greeks_3D greeks, deals, get_array_size_clsAcDealTicket(deals), market_set__, dummy_bump_greeks_sets, False, False, False, False, False, , , False, , False
    
    '------------------------
    ' Rewinding config
    '----------------------------
    
    config__.x_grid_ = origin_x_grid
    config__.time_step_per_day = origin_t_steps_per_day
    config__.no_of_trials_closing_ = origin_no_of_trials
    
    
    For inx = 1 To get_array_size_clsgreeks(greeks)
        greeks(inx).set_all_implied_tree_deltas greeks(inx).get_all_deltas
        greeks(inx).implied_tree_delta = greeks(inx).delta
        greeks(inx).implied_tree_gamma = greeks(inx).gamma
        
        For jnx = 1 To deals(inx).no_of_ul
            calc_sticky_strike_delta greeks(inx), deals(inx).duration, deals(inx).no_of_ul, deals(inx).get_ul_codes(), config__.current_date_, market_set__
            calc_sticky_strike_gamma greeks(inx), deals(inx).duration, deals(inx).no_of_ul, deals(inx).get_ul_codes(), config__.current_date_, market_set__
        Next jnx
        
    Next inx


    Exit Sub
    
ErrorHandler:

    raise_err "calculate_greeks_ac", Err.description

End Sub
'Private Sub copy_ac_deal_list(to_deal_list() As clsACDealTicket, from_deal_list() As clsACDealTicket, dimension As Integer)
'
'    Dim inx As Integer
'    Dim no_of_deals As Integer
'
'On Error GoTo ErrorHandler
'
'    no_of_deals = get_array_size_clsAcDealTicket(from_deal_list)
'
'    For inx = 1 To no_of_deals
'
'        If from_deal_list(inx).no_of_ul = dimension Then
'
'            push_back_clsAcDealTicket deal_list, from_deal_list(inx)
'
'        End If
'
'    Next inx
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "copy_ac_deal_list", Err.description
'
'End Sub


Private Sub retrieve_end_of_day_greek_2d(ByRef greeks() As clsGreeks, deals() As clsACDealTicket, Optional eval_date_in As Date = -1)


On Error GoTo ErrorHandler

    Dim inx As Integer
    Dim eval_date As Date
    Dim data_found As Boolean
    Dim tmp_greek As clsGreeks
    
    If eval_date_in > 0 Then
        eval_date = eval_date_in
    Else
        eval_date = config__.current_date_
    End If
    
    
    DBConnector


    For inx = LBound(deals) To UBound(deals)
    
        If deals(inx).no_of_ul > 1 Then
                        
            Set tmp_greek = New clsGreeks
            tmp_greek.redim_arrays deals(inx).no_of_ul
            If retrieve_ul_greek(tmp_greek, deals(inx).asset_code, deals(inx).get_ul_codes(), eval_date) Then
                Set greeks(inx) = tmp_greek
            End If
            
        End If
    Next inx
    
    DBDisConnector

    Exit Sub
    
ErrorHandler:

    DBDisConnector

    Err.Raise Err.number, "retrieve_end_of_day_greek_2d : " & Chr(13) & Err.source, Err.description

End Sub

Public Sub make_greek_to_distribute(ByRef greeks() As clsGreeks, deals() As clsACDealTicket, Optional eval_date_in As Date = -1)

    Dim inx As Integer
    Dim jnx As Integer
    
    Dim dummy_bump_greeks_sets() As clsGreekSet

On Error GoTo ErrorHandler
    
    '---------------------------------------------------
    ' Firstly, read text file for 2D
    '---------------------------------------------------
    read_greek_file greeks, deals, config__.current_date_

    '---------------------------------------------------
    ' And then, calculate 3D autocallable.
    '---------------------------------------------------
    calculate_ac_greeks_3D greeks, deals, get_array_size_clsAcDealTicket(deals), market_set__, dummy_bump_greeks_sets, False, False, False, False, False, , , False, , False
        
    For inx = 1 To get_array_size_clsAcDealTicket(deals)
    
        If deals(inx).no_of_ul = 1 Then
            greeks(inx).vega = AC_closing_greeks__(inx).vega
            greeks(inx).vanna = AC_closing_greeks__(inx).vanna
            greeks(inx).skew_s = AC_closing_greeks__(inx).skew_s
        Else
            greeks(inx).set_all_vegas AC_closing_greeks__(inx).get_all_vegas()
            greeks(inx).set_all_skews AC_closing_greeks__(inx).get_all_skews()
            greeks(inx).set_all_vannas AC_closing_greeks__(inx).get_all_vannas()
        End If
    
        greeks(inx).set_all_implied_tree_deltas greeks(inx).get_all_deltas
        greeks(inx).set_all_implied_tree_gammas greeks(inx).get_all_gammas
        greeks(inx).implied_tree_delta = greeks(inx).delta
        greeks(inx).implied_tree_gamma = greeks(inx).gamma
        
        
                    
        'For jnx = 1 To deals(inx).no_of_ul

            calc_sticky_strike_delta greeks(inx), deals(inx).duration, deals(inx).no_of_ul, deals(inx).get_ul_codes(), config__.current_date_, market_set__
            calc_sticky_strike_gamma greeks(inx), deals(inx).duration, deals(inx).no_of_ul, deals(inx).get_ul_codes(), config__.current_date_, market_set__
            
            modify_vega_by_vanna greeks(inx), deals(inx).no_of_ul, deals(inx).get_ul_codes(), market_set__
            
       ' Next jnx
    
    Next inx
    
    greek_to_file greeks, AC_deals__, config__.file_path_, "position_" & Format(config__.current_date_, "YYYYMMDD") & "." & config__.position_file_extension

    Exit Sub
    
ErrorHandler:

    raise_err "make_greek_to_distribute", Err.description

End Sub




'---------------------------------------------------------------
' To read text file for 2d ac.
'---------------------------------------------------------------
Private Sub read_greek_file(ByRef greeks() As clsGreeks, deals() As clsACDealTicket, Optional eval_date_in As Date = -1, Optional vol_bump As Boolean = False, Optional bumping_this As String = "")

  
    Dim inx As Integer
    Dim jnx As Integer
    
    Dim file_str As String
    Dim no_of_deals As Integer
    Dim eval_date As Date
    Dim nodevalue_tmp() As Double
    Dim px_tmp() As Double
    Dim dx_tmp() As Double
    Dim nmin_tmp() As Long
    Dim nmax_tmp() As Long
    'Dim file_cls As New FileIO
    Dim file_found As Boolean
    
    Dim current_prices() As Double
        
    
    'Dim greek_table As clsACGreekTable
    Dim sparse_greek_table As clsACSparseGridTable
   ' Dim sp_legacy_object As Object
    
    Dim ul_Index As Integer

On Error Resume Next

    no_of_deals = UBound(deals)
    
    If Err.number = 9 Then
        no_of_deals = 0
    End If
    

On Error GoTo ErrorHandler

    If eval_date_in > 0 Then
        eval_date = eval_date_in
    Else
        eval_date = config__.current_date_
    End If
    
   ' Set sp_legacy_object = find_com_addIn()

    For inx = 1 To no_of_deals
    
        file_found = False
    
        ReDim nmin_tmp(1 To deals(inx).no_of_ul) As Long
        ReDim nmax_tmp(1 To deals(inx).no_of_ul) As Long
        ReDim current_prices(1 To deals(inx).no_of_ul) As Double
        

        
        For jnx = 1 To deals(inx).no_of_ul
            nmin_tmp(jnx) = 0
            
            If deals(inx).no_of_ul <= 2 Then
                nmax_tmp(jnx) = 200 '<-------
            Else
                nmax_tmp(jnx) = 5 '<-------
            End If
            
            current_prices(jnx) = market_set__.market(market_set__.find_index(deals(inx).ul_code(jnx))).s_
            
        Next jnx
        
        '-------------------------------------------------------------------------------------------
        '
        '-------------------------------------------------------------------------------------------
        If greek_file_cache__ Is Nothing Then

            Set greek_file_cache__ = New PositionFileRepository
            greek_file_cache__.initialize "W:\", "." & config__.snapshot_file_extension

        End If

        
        '-----------------------------------------------
        ' I'm not so proud of this piece. Really.
        '-----------------------------------------------
        If deals(inx).no_of_ul <= 2 Then
        
            Dim tmpValue As Double
            Dim tmp_delta() As Double
            Dim tmp_gamma() As Double
            
            tmpValue = 0
            
            ReDim tmp_delta_array(1 To deals(inx).no_of_ul) As Double
            ReDim tmp_gamma_array(1 To deals(inx).no_of_ul) As Double
            
            Set greeks(inx) = New clsGreeks
            
            greeks(inx).redim_arrays deals(inx).no_of_ul
            
            Dim tmp_file_name As String
            
            tmp_file_name = deals(inx).asset_code
            
            If vol_bump Then
                For jnx = 1 To deals(inx).no_of_ul
                    If deals(inx).ul_code(jnx) = bumping_this Then
                        tmp_file_name = deals(inx).asset_code & "_vol_bump_" & jnx & "_"
                        Exit For
                    End If
                Next jnx
            End If
            
            tmpValue = greek_file_cache__.getValue(tmp_file_name, array_base_zero(current_prices), array_base_zero(deals(inx).get_reference_price()), Format(eval_date, "YYYYMMDD"))
            
            greeks(inx).value = tmpValue * Sgn(deals(inx).notional)
            
            '------------------
            greeks(inx).value = tmpValue * Sgn(deals(inx).notional)
            
            
            tmp_delta = greek_file_cache__.getAllDeltas(tmp_file_name, array_base_zero(current_prices), array_base_zero(deals(inx).get_reference_price()), Format(eval_date, "YYYYMMDD"))
                                
            For jnx = 1 To deals(inx).no_of_ul
                tmp_delta_array(jnx) = tmp_delta(jnx + LBound(tmp_delta) - 1) * Sgn(deals(inx).notional)
            Next jnx
            
            greeks(inx).set_all_deltas tmp_delta_array ' * Sgn(deals(inx).notional)
            greeks(inx).set_all_implied_tree_deltas greeks(inx).get_all_deltas
            greeks(inx).implied_tree_delta = greeks(inx).delta
            
            
            '------------------------
            
            
            tmp_gamma = greek_file_cache__.getAllGammas(tmp_file_name, array_base_zero(current_prices), array_base_zero(deals(inx).get_reference_price()), Format(eval_date, "YYYYMMDD"))
                      
            For jnx = 1 To deals(inx).no_of_ul
                tmp_gamma_array(jnx) = tmp_gamma(jnx + LBound(tmp_gamma) - 1) * Sgn(deals(inx).notional)
            Next jnx
            
            greeks(inx).set_all_gammas tmp_gamma_array ' * Sgn(deals(inx).notional)
            greeks(inx).set_all_implied_tree_gammas greeks(inx).get_all_gammas
            greeks(inx).implied_tree_gamma = greeks(inx).gamma

        
        End If

    Next inx

    Exit Sub
    
ErrorHandler:

    raise_err "read_greek_file. Index: " & inx & ", asset_code: " & deals(inx).asset_code, Err.description


End Sub

Private Function FileExists(ByVal sFullPath As String) As Boolean
    Dim oFile As New Scripting.FileSystemObject
    FileExists = oFile.FileExists(sFullPath)
End Function

Private Sub retrieve_current_greek(ByRef greeks() As clsGreeks, ByRef deals() As clsACDealTicket, Optional eval_date_in As Date = -1) ', Optional eval_date_in As Date = -1)

    Dim inx As Integer
    Dim eval_date As Date
    
On Error GoTo ErrorHandler

    If eval_date_in < 0 Then
    
        eval_date = config__.current_date_
    
    Else
    
        eval_date = eval_date_in
        
    End If

    If Not IsNull(ac_greek_cache__) And UBound(ac_greek_cache__) >= LBound(ac_greek_cache__) Then

        ReDim greeks(LBound(ac_greek_cache__) To UBound(ac_greek_cache__)) As clsGreeks

        For inx = LBound(ac_greek_cache__) To UBound(ac_greek_cache__)
        
            If ac_greek_cache__(inx).initialized = False Then
            
                initialize_cache_greek ac_greek_cache__(inx), eval_date_in
                
            End If

            Set greeks(inx) = ac_greek_cache__(inx).get_greeks(market_set__.market_by_ul().s_)
            
            If deals(inx).no_of_ul = 1 Then
                greeks(inx).vega = AC_closing_greeks__(inx).vega
                greeks(inx).vanna = AC_closing_greeks__(inx).vanna
                greeks(inx).skew_s = AC_closing_greeks__(inx).skew_s
            Else
                greeks(inx).set_all_vegas AC_closing_greeks__(inx).get_all_vegas()
                greeks(inx).set_all_skews AC_closing_greeks__(inx).get_all_skews()
                greeks(inx).set_all_vannas AC_closing_greeks__(inx).get_all_vannas()
            End If
            
            'calc_implied_tree_delta greeks(inx), deals(inx).duration, deals(inx).no_of_ul, deals(inx).get_ul_codes(), eval_date, market_set__
            'greeks(inx).set_implied_tree_deltas deals(inx).duration, deals(inx).no_of_ul, deals(inx).get_ul_codes()

        Next inx

'        deals = deals__

    End If



    Exit Sub

ErrorHandler:

    raise_err "retrieve_current_greek", Err.description

End Sub

'-------------
'
'-------------
Private Sub calc_implied_tree_delta(the_greeks As clsGreeks, duration As Double, no_of_ul As Integer, ul_codes() As String, eval_date As Date, themarketset As clsMarketSet)
    
    Dim inx As Integer
    Dim skew As Double
    Dim market_index As Integer
    Dim s As Double
    
On Error GoTo ErrorHandler
    
    If no_of_ul > 1 Then
        
        If no_of_ul = get_array_size_double(the_greeks.get_all_deltas) Then
        
            For inx = 1 To no_of_ul
                
                the_greeks.set_implied_tree_deltas inx, the_greeks.deltas(inx) '+ the_greeks.vegas(inx) * get_skew(eval_date, eval_date + 365 * duration, ul_codes(inx), S) * 100 / S
                                      
            
            Next inx
        End If
                
    Else
        market_index = themarketset.find_index(ul_codes(1))
    
        s = themarketset.market(market_index).s_
        the_greeks.implied_tree_delta = the_greeks.delta + the_greeks.vega * themarketset.market(market_index).sabr_surface_.get_skew(eval_date + 365 * duration, s, eval_date, 0.8, 1.1) * 100 / s
    
    End If
    
    Exit Sub
    
ErrorHandler:

    raise_err "calc_implied_tree_delta", Err.description
    

End Sub


'-----------------------------------------------
' Calculate 2 dimension ac greeks
'-----------------------------------------------
Public Sub calculate_ac_greeks_2D(ByRef greeks() As clsGreeks, deals() As clsACDealTicket, deal_count As Long _
                              , theMarket As clsMarketSet _
                              , bump_greek_sets() As clsGreekSet _
                              , Optional calc_vega As Boolean _
                              , Optional calc_skew As Boolean _
                              , Optional calc_corr As Boolean _
                              , Optional calc_rho As Boolean = False _
                              , Optional ByVal current_date As Date = -1 _
                              , Optional ByVal ignore_smoothing As Boolean = False _
                              , Optional ByVal percent_vega As Boolean = False _
                              , Optional snapshot_time As Double = 1 / 365 / 2 _
                              , Optional ByVal calc_duration As Boolean = True)




    Dim inx As Integer
    Dim jnx As Integer
    Dim market As clsMarket
    Dim last_node() As Double
    Dim tmp_greek As clsGreeks
    
    Dim status_bar_str As String
        

On Error GoTo ErrorHandler

    If current_date < 0 Then
        
        current_date = config__.current_date_
    
    Else
    
        Set theMarket.pl_currency_rate_curve_ = theMarket.pl_currency_rate_curve_.copy_obj(current_date - theMarket.pl_currency_rate_curve_.rate_dates()(0))
    
    End If
    
    status_bar_str = Application.StatusBar


    If deal_count > 0 Then
    
        ReDim Preserve greeks(LBound(deals) To UBound(deals)) As clsGreeks
        ReDim Preserve bump_greek_sets(LBound(deals) To UBound(deals)) As clsGreekSet

        For inx = LBound(deals) To UBound(deals)
        'For inx = UBound(deals) To UBound(deals)
            
            DoEvents
            
            Application.StatusBar = status_bar_str & "> AC 2D: " & inx & " / " & UBound(deals) & " : " & deals(inx).asset_code


            
            If deals(inx).no_of_ul = 2 Then
            
                Set greeks(inx) = New clsGreeks
                
                greeks(inx).redim_arrays deals(inx).no_of_ul

                deals(inx).current_date = current_date
                deals(inx).x_grid = config__.x_grid_
                deals(inx).v_grid = config__.v_grid_
                deals(inx).t_grid = -1 * Int(-1 * config__.time_step_per_day * (deals(inx).maturity_date - deals(inx).current_date))
                deals(inx).scheme_type = config__.fdm_scheme_
                deals(inx).set_term_vega_tenor config__.term_vega_tenor_array
    
    
                deals(inx).strike_at_maturity = deals(inx).autocall_schedules(deals(inx).no_of_schedule).percent_strike ' strike_values()(deals(inx).no_of_schedule) / deals(inx).reference_price
                deals(inx).coupon_at_maturity = deals(inx).coupon_on_calls()(deals(inx).no_of_schedule)
                deals(inx).maturity_date = deals(inx).call_dates()(deals(inx).no_of_schedule)
                
                run_ac_pricing_2d greeks(inx), deals(inx), theMarket, bump_greek_sets(inx), calc_vega, calc_skew, calc_corr, calc_rho, 1 / 365 / 2, ignore_smoothing
                
    
                
                greeks(inx).asset_code = deals(inx).asset_code
    
                If percent_vega Then
                    greeks(inx).set_vegas 1, greeks(inx).vegas(1) / 100
                    greeks(inx).set_vegas 2, greeks(inx).vegas(2) / 100
                End If
                
                
                If calc_duration Then
                    greeks(inx).duration = get_ac_duration_2d(deals(inx), theMarket)
                End If
            
            End If

        Next inx

    End If
    
    Application.StatusBar = status_bar_str


    Exit Sub

ErrorHandler:

    raise_err "calculate_ac_greeks", Err.description

End Sub
'-----------------------------------------------
' Calculate 2 dimension ac greeks
'-----------------------------------------------
Public Sub calculate_ac_greeks_3D(ByRef greeks() As clsGreeks, deals() As clsACDealTicket, deal_count As Long _
                              , theMarket As clsMarketSet _
                              , bump_greek_sets() As clsGreekSet _
                              , Optional calc_vega As Boolean _
                              , Optional calc_skew As Boolean _
                              , Optional calc_corr As Boolean _
                              , Optional calc_rho As Boolean = False _
                              , Optional calc_theta As Boolean = False _
                              , Optional ByVal current_date As Date = -1 _
                              , Optional ByVal ignore_smoothing As Boolean = False _
                              , Optional ByVal percent_vega As Boolean = False _
                              , Optional snapshot_time As Double = 1 / 365 / 2 _
                              , Optional ByVal calc_duration As Boolean = True)




    Dim inx As Integer
    Dim jnx As Integer
    Dim market As clsMarket
    Dim last_node() As Double
    Dim tmp_greek As clsGreeks
    Dim status_bar_str

On Error GoTo ErrorHandler

    If current_date < 0 Then
        
        current_date = config__.current_date_
    
    Else
    
        Set theMarket.pl_currency_rate_curve_ = theMarket.pl_currency_rate_curve_.copy_obj(current_date - theMarket.pl_currency_rate_curve_.rate_dates()(0))
    
    End If
    
    status_bar_str = Application.StatusBar
    

    If deal_count > 0 Then
    
        ReDim Preserve greeks(LBound(deals) To UBound(deals)) As clsGreeks
        ReDim Preserve bump_greek_sets(LBound(deals) To UBound(deals)) As clsGreekSet

        For inx = LBound(deals) To UBound(deals)
        'For inx = UBound(deals) To UBound(deals)
        
            
            
            Application.StatusBar = status_bar_str & "> AC 3D: " & inx & " / " & UBound(deals) & " : " & deals(inx).asset_code

            DoEvents
            
            If deals(inx).no_of_ul = 3 Then
                
                
                    Set greeks(inx) = New clsGreeks
                
                
                
                greeks(inx).redim_arrays deals(inx).no_of_ul
                
                deals(inx).current_date = current_date
                deals(inx).x_grid = config__.x_grid_
                deals(inx).v_grid = config__.v_grid_
                deals(inx).t_grid = -1 * Int(-1 * config__.time_step_per_day * (deals(inx).maturity_date - deals(inx).current_date))
                deals(inx).scheme_type = config__.fdm_scheme_
                deals(inx).set_term_vega_tenor config__.term_vega_tenor_array
                
                If Abs(deals(inx).notional) < 500000000 Then
                    deals(inx).no_of_trials = (config__.no_of_trials_closing_ + 1) / 2 - 1
                Else
                    deals(inx).no_of_trials = config__.no_of_trials_closing_
                End If
    
                deals(inx).strike_at_maturity = deals(inx).autocall_schedules(deals(inx).no_of_schedule).percent_strike 'deals(inx).strike_values()(deals(inx).no_of_schedule) / deals(inx).reference_price
                deals(inx).coupon_at_maturity = deals(inx).coupon_on_calls()(deals(inx).no_of_schedule)
                deals(inx).maturity_date = deals(inx).call_dates()(deals(inx).no_of_schedule)
                
                run_ac_pricing_3d greeks(inx), deals(inx), theMarket, bump_greek_sets(inx), calc_vega, calc_skew, calc_corr, calc_rho, calc_theta, 1 / 365 / 2, ignore_smoothing
    
                
                greeks(inx).asset_code = deals(inx).asset_code
    
                If percent_vega Then
                    greeks(inx).set_vegas 1, greeks(inx).vegas(1) / 100
                    greeks(inx).set_vegas 2, greeks(inx).vegas(2) / 100
                    greeks(inx).set_vegas 3, greeks(inx).vegas(3) / 100
                End If
                
                
                If calc_duration Then
                    greeks(inx).duration = get_ac_duration_3d(deals(inx), theMarket)
                End If
            
            End If

        Next inx

    End If
    
    Application.StatusBar = status_bar_str



    Exit Sub

ErrorHandler:

    raise_err "calculate_ac_greeks", Err.description

End Sub

'---------------------------------------------------------------------
' Calculate 1dimension autocallable greeks
'---------------------------------------------------------------------
Public Sub calculate_ac_greeks(ByRef greeks() As clsGreeks, ByRef intraday_greeks() As clsGreeks, deals() As clsACDealTicket, deal_count As Long _
                              , theMarket As clsMarketSet _
                              , Optional s_in As Double = -1, Optional bump_greeks As String = "Y" _
                              , Optional ByVal current_date_in As Date = -1 _
                              , Optional ByVal ignore_smoothing As Boolean = False _
                              , Optional ByVal percent_vega As Boolean = False _
                              , Optional calc_skew As Boolean = True _
                              , Optional calc_rho As Boolean = True _
                              , Optional calc_duration As Boolean = False)

    Dim inx As Integer
    Dim jnx As Integer
    Dim market As clsMarket
    Dim pl_currency_curve As clsRateCurve
    
    Dim last_node() As Double
    Dim tmp_greek As clsGreeks
    Dim current_date As Date
    
    Dim bump_greek_set As clsGreekSet
    
    Dim status_bar_str As String
        

On Error GoTo ErrorHandler

    Set market = theMarket.market_by_ul().copy_obj()

    If current_date_in < 0 Then
        current_date = config__.current_date_
    Else
        current_date = current_date_in
        Set market.rate_curve_ = theMarket.pl_currency_rate_curve_.copy_obj(current_date - theMarket.pl_currency_rate_curve_.rate_dates()(0))
    End If
    
    
    status_bar_str = Application.StatusBar



    If deal_count > 0 Then
        
        ReDim Preserve greeks(LBound(deals) To UBound(deals)) As clsGreeks

        '---------------------------------
        ' Loop for deal list
        '---------------------------------
        For inx = LBound(deals) To UBound(deals)
            
            DoEvents
        
            Application.StatusBar = status_bar_str & "> AC1D" & inx & " / " & UBound(deals) & " : " & deals(inx).asset_code
    
            Set greeks(inx) = New clsGreeks
                
            If deals(inx).no_of_ul = 1 Then
                '-------------------------------------------------------
                ' Set FDM parameters
                '-------------------------------------------------------
                deals(inx).current_date = current_date
                deals(inx).x_grid = config__.x_grid_
                deals(inx).v_grid = config__.v_grid_
                deals(inx).t_grid = -1 * Int(-1 * config__.time_step_per_day * (deals(inx).maturity_date - deals(inx).current_date))
                deals(inx).scheme_type = config__.fdm_scheme_
                deals(inx).set_term_vega_tenor config__.term_vega_tenor_array
    
    
                deals(inx).strike_at_maturity = deals(inx).autocall_schedules(deals(inx).no_of_schedule).percent_strike 'deals(inx).strike_values()(deals(inx).no_of_schedule) / deals(inx).reference_price
                deals(inx).coupon_at_maturity = deals(inx).coupon_on_calls()(deals(inx).no_of_schedule)
                deals(inx).maturity_date = deals(inx).call_dates()(deals(inx).no_of_schedule)
                
                Erase last_node
                
                '----------------------------------------
                ' RUN Pricing
                '----------------------------------------
                'run_ac_pricing greeks(inx), deals(inx), market, last_node, bump_greeks = "Y", 1, ignore_smoothing
                run_ac_pricing_1d greeks(inx), deals(inx), theMarket, bump_greek_set, bump_greeks = "Y", bump_greeks = "Y", bump_greeks = "Y", bump_greeks = "Y", 1 / 365 / 2, ignore_smoothing
                greeks(inx).asset_code = deals(inx).asset_code
    
                If percent_vega Then
                    greeks(inx).vega = greeks(inx).vega / 100
                End If
                
                
                '----------------------------------------
                ' Intraday greek
                '----------------------------------------
'                For jnx = 1 To get_array_size_double(greeks(inx).get_xAxis())
'
'                    Set tmp_greek = New clsGreeks
'
'                    tmp_greek.asset_code = deals(inx).asset_code
'                    tmp_greek.value = greeks(inx).get_snapshot_value()(inx)
'                    tmp_greek.delta = greeks(inx).get_xAxis()(inx)
'                    tmp_greek.gamma = greeks(inx).get_xAxis()(inx)
'
'    Dim rtn_array() As Double
'    Dim inx As Integer
'
'    ReDim rtn_array(1 To ul_num_index) As Double
'
'    For inx = 1 To ul_num_index
'
'        rtn_array(inx) = rtnCls.ReturnDelta(current_prices, initial_prices, nodeValue_, px_, dx_, nmin_, nmax_, inx) * notional_sign
'
'    Next inx
'
'    get_deltas = rtn_array
'
'
'                Next jnx
'                For jnx = LBound(last_node, 2) To UBound(last_node, 2)
'
'                    Set tmp_greek = New clsGreeks
'
'                    tmp_greek.asset_code = deals(inx).asset_code
'                    tmp_greek.value = last_node(2, jnx) * deals(inx).notional
'                    tmp_greek.delta = last_node(3, jnx) * deals(inx).notional
'                    tmp_greek.gamma = last_node(4, jnx) * deals(inx).notional
'                    tmp_greek.ul_price = last_node(1, jnx) ' * deals(inx).notional
'
'                    push_back_greek intraday_greeks, tmp_greek
'
'                Next jnx
                
                '----------------------------------------
                ' Calc duration
                '----------------------------------------
                If calc_duration Then
                    greeks(inx).duration = get_ac_duration(deals(inx), theMarket)
                End If
                
            End If

        Next inx

    End If

     Application.StatusBar = status_bar_str


    Exit Sub

ErrorHandler:

  '  Application.StatusBar = ""
    raise_err "calculate_ac_greeks", Err.description & ": " & deals(inx).asset_code

End Sub

'Private Sub initialize_cache_greek(greek_cache As clsGreekCache, Optional eval_date_in As Date = -1)
'
'    Dim eval_date As Date
'    Dim abscissa() As Double
'    Dim greeks() As clsGreeks
'    Dim ul_count As Integer
'
'On Error GoTo ErrorHandler
''Check evaluation date.
''Default evaluation date is the current date loaded.
'
'    If eval_date_in >= 0 Then
'
'        eval_date = eval_date_in
'
'    Else
'
'        eval_date = config__.current_date_
'
'    End If
'
'    ul_count = retrieve_cache_greek(abscissa, greeks, greek_cache.deal.asset_code, eval_date)
'
'    If ul_count > 0 Then
'
'        greek_cache.initialize abscissa, greeks
'
'    Else
'
'        greek_cache.null_greek
'       ' Err.Raise vbObjectError, , "[PRO] Cannot find greek cache:" & Chr(13) & "Asset code:" & greek_cache.deal.asset_code
'
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "initialize_cache_greek"
'
'
'End Sub
'
'
'
'
'
'Private Function retrieve_cache_greek(ByRef abscissa() As Double, ByRef greeks() As clsGreeks, asset_code As String, eval_date As Date) As Integer
'
'    Dim greek_cd() As String
'    Dim greek_value() As Double
'    Dim ul_price() As Double
'    Dim ul_count As Integer
'    Dim inx As Integer
'
'    Dim counter As Integer
'    Dim prev_ul_price As Double
'
'On Error GoTo ErrorHandler
'
'    DBConnector
'
'    If retrieve_ul_count(ul_count, eval_date, asset_code) Then
'
'        ReDim abscissa(1 To ul_count) As Double
'        ReDim greeks(1 To ul_count) As clsGreeks
'
'        retrieve_cache_greek_sql ul_price, greek_cd, greek_value, eval_date, asset_code
'
'        prev_ul_price = -1
'        counter = 0
'
'        For inx = LBound(greek_value) To UBound(greek_value)
'
'            If prev_ul_price <> ul_price(inx) Then
'
'                counter = counter + 1
'
'                abscissa(counter) = ul_price(inx)
'                Set greeks(counter) = New clsGreeks
'
'                prev_ul_price = ul_price(inx)
'
'            End If
'
'            greeks(counter).set_greek_value greek_cd(inx), greek_value(inx)
'
'        Next inx
'
'    End If
'
'    retrieve_cache_greek = ul_count
'
'    DBDisConnector
'
'    Exit Function
'
'ErrorHandler:
'
'    DBDisConnector
'    raise_err "retrieve_cache_greek"
'
'
'End Function

Public Sub get_closing_s_series(ul_code As String, start_date As String, end_date As String, ByRef result_date_array() As Long, ByRef result_price_array() As Double)

    Dim adoCon As New adoDB.Connection
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
On Error GoTo ErrorHandler

    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
'    If Left(ul_code, 3) = "KR7" Then
'        sql = "select tdate, endprice from ras.if_stock_data where tdate between '" + start_date + "' and '" + end_date + "' and code = '" + ul_code + "'  order by tdate asc"
'    Else
'        sql = "select tdate, endprice from ras.if_index_data where tdate between '" + start_date + "' and '" + end_date + "' and indexid = '" + ul_code + "'  order by tdate asc"
'    End If
    sql = "select tdate, endprice from ras.if_stock_data where tdate between '" + start_date + "' and '" + end_date + "' and code = '" + ul_code + "' union " _
        & "select tdate, endprice from ras.if_index_data where tdate between '" + start_date + "' and '" + end_date + "' and indexid = '" + ul_code + "' order by tdate asc"
    
    With oCmd
        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql
        
        oRS.Open .Execute
    End With
    
'    Dim i As Integer
'    ReDim result_date_array(0 To oRS.RecordCount - 1) As Long
'    ReDim result_price_array(0 To oRS.RecordCount - 1) As Double
    
    Do Until oRS.EOF
    
        Call push_back_long(result_date_array, CLng(str2date(oRS.Fields(0))), 0)
        Call push_back_double(result_price_array, oRS.Fields(1), 0)
    
        'result_date_array(i) = CLng(oRS.Fields(0))
        'result_price_array(i) = oRS.Fields(1)
        
        'i = i + 1
        oRS.MoveNext
    Loop
    oRS.Close
    
    Call disconnectDB(adoCon)
    

Exit Sub
    
ErrorHandler:
    
    raise_err "get_closing_s_series", Err.description

End Sub

'drift adjustment 추가: 2023.11.21
Public Function get_drift_adjustment(ByRef term_date() As Date, ByRef adjust() As Double, eval_date As Date, ul_code As String) As Integer
'Public Function get_drift_adjustme                  nt(ByRef term_date() As Date, ByRef adjust() As Double, eval_date As Date, ul_code As String) As LongPtr

On Error GoTo ErrorHandler
    
    Dim adoCon As New adoDB.Connection
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)

    Dim sql As String
    Dim oRS As New adoDB.Recordset
    Dim i As Integer
    i = 1
    
    sql = "select term_date, adjust from sps.drift_adjust where  ul_code='" + ul_code + "' and eval_date = (select max(eval_date) from sps.ul_sabr_parameter where ul_code='" + ul_code + "' and eval_date<='" + date2str(eval_date) + "')"
    
    oRS.CursorLocation = adUseClient
    oRS.CursorType = adOpenStatic
    oRS.Open sql, adoCon
    
    ReDim term_date(1 To oRS.RecordCount + 2) As Date
    ReDim adjust(1 To oRS.RecordCount + 2) As Double
    
    term_date(i) = eval_date
    adjust(i) = 1#
    i = i + 1
    
    If oRS.RecordCount > 0 Then
        oRS.MoveFirst
        Do
            term_date(i) = str2date(oRS("TERM_DATE"))
            adjust(i) = Exp(-1 * oRS("ADJUST") * (term_date(i) - eval_date) / 365)
            
            oRS.MoveNext
            i = i + 1
        Loop While Not oRS.EOF
    End If
    
    term_date(i) = DateValue(eval_date + 365 * 10)
    adjust(i) = 1#
    
    If oRS.State <> 0 Then oRS.Close
    
    get_drift_adjustment = i
    
    adoCon.Close
    
    Exit Function

ErrorHandler:
    If oRS.State <> 0 Then oRS.Close
    Err.Raise Err.number, Err.source, Err.description, Erl, "get_drift_adjustment"

End Function


'To Do(verify=true)
'deal ticket cls 구성 단계에서
'1. biz-one 테이블 별 booking 오류, 모순점 검증
'2. 각 DB source와 대사 (source: biz-one / front / risk)
'3. 발행 후 주가 경로 탐색 -> barrier touch 대사
Public Function get_ac_deal_ticket(indv_iscd As String, target_date As Date, for_frn_flag As Boolean, for_eswap_flag As Boolean, adoCon As adoDB.Connection, Optional db_source As DATA_FROM = DATA_FROM.bsys, Optional eval_date_lag As Boolean = False, Optional verify As Boolean = False, Optional market_set As clsMarketSet) As clsACDealTicket

    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim oRS2 As New adoDB.Recordset
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    
    Dim deal_ticket As clsACDealTicket
    Dim schedule_list() As clsAutocallSchedule
    Dim a_schedule As clsAutocallSchedule
    Dim no_of_schedule As Integer
    Dim call_dates() As Date
    Dim strikes() As Double
    Dim coupons() As Double
    Dim strike_shifts() As Double
    Dim early_exit_touched_flags() As Long 'data type 변경: dll(2018.7.17)
    Dim early_exit_performance_types() As Long 'dll(2018.8.8)
    Dim early_exit_barrier_types() As Long 'dll(2018.8.8)
    
    Dim inx As Integer
    Dim jnx As Integer
    Dim i As Integer
    
    Set deal_ticket = New clsACDealTicket

On Error GoTo ErrorHandler

    'Fixed Value
    deal_ticket.alive_yn = "Y"
    deal_ticket.confirmed_yn = "Y"
    deal_ticket.rate_spread = 0
    deal_ticket.hedge_cost = 0
    deal_ticket.instrument_type = INST_TYPE.note 'note
    
    'Autocall option
    deal_ticket.call_put = 0 'autocall/autoput
    deal_ticket.call_strike = 1
    deal_ticket.call_participation = 0
    deal_ticket.floor_value = 0
    
    'KI option
    deal_ticket.put_strike = 1 'ki put
    deal_ticket.put_additional_coupon = 0
    deal_ticket.put_participation = 1
    deal_ticket.ki_adj_pct = 1 'ki시 기존 coupon rate 대비 지급 승수
    
    'Range Accrual
    deal_ticket.ra_flag = 0
    deal_ticket.ra_cpn = 0
    deal_ticket.ra_tenor = 0
    deal_ticket.ra_min_percent = 0
    deal_ticket.ra_max_percent = 10000
    
    'Ejectable: TBD
'    deal_ticket.ejectable_flag = False
        
    'Floating Leg: TBD
'    no_of_schedule = shtACPricer.Range("no_of_floating_leg").value
'    If no_of_schedule >= 1 Then
'        ReDim call_dates(1 To no_of_schedule) As Date
'        ReDim coupons(1 To no_of_schedule) As Double
'
'        For inx = 1 To no_of_schedule
'            call_dates(inx) = shtACPricer.Range("floating_leg_start").Cells(inx, 1).value
'            coupons(inx) = shtACPricer.Range("floating_leg_start").Cells(inx, 2).value
'        Next inx
'    End If
'    deal_ticket.set_floating_schedule no_of_schedule, call_dates, coupons
    
    'Simulation Config
    deal_ticket.current_date = target_date
    deal_ticket.current_date_origin_ = deal_ticket.current_date
    deal_ticket.x_grid = 200
    deal_ticket.v_grid = 100
    deal_ticket.days_per_step = 0.25
    deal_ticket.scheme_type = 1 '0: Do, 1: CS, 2: MCS, 3:HV
    deal_ticket.mid_day_greek = False
    deal_ticket.vol_scheme_type = 1 '0:Stochastic, 1:Local
    deal_ticket.no_of_trials = 2 ^ 14 - 1
    
    Dim term_array() As Date
    For inx = 1 To 14
        If inx < 3 Then
            push_back_date term_array, deal_ticket.current_date + 30 * inx
        Else
            push_back_date term_array, deal_ticket.current_date + 90 * (inx - 2)
        End If
    Next inx
    deal_ticket.set_term_vega_tenor term_array
    
    'deal_ticket 미사용 필드
    'Public current_notional As Double
    'Public early_exit_touched_flag As Long
    'Public Maturity As Integer
    'Public duration As Double
    'Public ko_barrier_flag As Long
    'Public ko_touched_flag As Long
    'Public ko_barrier As Double
    'Public comment As String
        
    'biz-one -> deal_ticket
    'OTC 기본 정보, KI
    deal_ticket.asset_code = indv_iscd

    ReDim bind_variable(2) As String
    ReDim bind_value(2) As Variant
    bind_variable(1) = ":indv_iscd"
    bind_value(1) = indv_iscd
    bind_variable(2) = ":tdate"
    bind_value(2) = date2str(target_date)
    
    With oCmd

        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = getSQL(SQL_PATH_AC_DEAL, bind_variable, bind_value)

        oRS.Open .Execute

    End With
        
    Dim ki_barrier As Double
    Dim ac_every_performance_type As Integer
            
    Do Until oRS.EOF
    
        deal_ticket.fund_code_c = oRS(0)
        deal_ticket.fund_code_m = Left(deal_ticket.fund_code_c, 2)

        deal_ticket.value_date = oRS("FRST_STND_PRC_FIN_DATE")
        deal_ticket.settlement_date = oRS("PBLC_DATE")
        deal_ticket.maturity_date = oRS("MTRT_DATE")
        deal_ticket.ccy = oRS("STLM_CRCD")
    
        deal_ticket.issue_price = oRS("FUND_PBLC_UNPR") / oRS("REAL_PBLC_FCAM") '발행가
        deal_ticket.issue_cost = oRS("PERC_APLY_THPR") '장부가
        
        deal_ticket.unit_notional = oRS("REAL_PBLC_FCAM")

        If for_eswap_flag = True Then
            deal_ticket.put_additional_coupon = -1
            deal_ticket.floor_value = -1
        End If
        
        'for ELB 2024.07.12 <- temp for ELB290, ELB312, ELB316
        If oRS("PROD_CLS_CODE_D") = "13" Then 'indv_iscd = "IO322724290T" Or indv_iscd = "IO322724312T" Or indv_iscd = "IO322724316T" Then
            deal_ticket.floor_value = 1
            If for_eswap_flag = True Then
                deal_ticket.floor_value = 0
            End If
        End If
        
        'for Booster 2024.07.12 <- temp for Booster 2024.06.28
        If oRS("CLRD_TYPE_CODE") = "19" Then 'indv_iscd = "IO332724072T" Or indv_iscd = "IO332724073T" Or indv_iscd = "IO332724077T" Then
            deal_ticket.call_participation = 2 '<- 비즈원[35304] 만기수익률관리 참여율 참조로 변경 필요
        End If

        '2024.05.29
        If deal_ticket.value_date = target_date Then
            deal_ticket.qty = oRS("PBLC_STCK_QTY")
        Else
            deal_ticket.qty = oRS("RMND_QTY")
        End If

        If oRS("DEAL_CLS_CODE") = "1" Then '매도
            deal_ticket.qty = -1 * deal_ticket.qty
        End If
        
        deal_ticket.notional = deal_ticket.qty * deal_ticket.unit_notional
        
        Select Case oRS("PROD_CLS_CODE")
        Case "04": deal_ticket.instrument_type = INST_TYPE.SWAP
        Case "07": deal_ticket.instrument_type = INST_TYPE.note 'ELS
        Case "09": deal_ticket.instrument_type = INST_TYPE.note 'ELN
        End Select
        
        'KI barrier
        '만기상환 베리어 유무
        If oRS("BARR_EN") = "2" Then 'No KI barrier
            deal_ticket.ki_barrier_flag = 1
            ki_barrier = 0
            deal_ticket.dummy_coupon = 0
            deal_ticket.ki_touched_flag = 1
            deal_ticket.ki_monitoring_freq = 1
        Else
            deal_ticket.ki_barrier_flag = 1
            ki_barrier = oRS("BARR_VAL1") / 100
            deal_ticket.dummy_coupon = get_dummy(indv_iscd, ki_barrier, adoCon)
            
            '만기상환 베리어 HIT유무
            If oRS("BARR_HIT_CLS_CODE") = "2" Then
                deal_ticket.ki_touched_flag = 0
            Else
                deal_ticket.ki_touched_flag = 1
            End If
            
            If oRS("MNTG_CYCL_CODE") = "3" Then
                deal_ticket.ki_monitoring_freq = 1
            Else
                deal_ticket.comment = deal_ticket.comment + "KI 모니터링 주기(MNTG_CYCL_CODE)가 Daily로 설정되지 않음"
            End If
        End If
        
        If for_frn_flag = True Then
            deal_ticket.ki_touched_flag = 0
        End If
            
        If for_eswap_flag = True Then
            deal_ticket.dummy_coupon = deal_ticket.dummy_coupon - 1
        End If
            
        '기초자산 선정방법
        Select Case oRS("UNAS_CHOC_MTHD_CODE")
        Case "1" 'Worst
            ac_every_performance_type = -1
            deal_ticket.ki_performance_type = -1
        Case "2" 'Best
            ac_every_performance_type = 1
            deal_ticket.ki_performance_type = 1
        Case "4" 'Avg
            ac_every_performance_type = 0
            deal_ticket.ki_performance_type = 0
        Case "5" 'WP
            ac_every_performance_type = 0
            deal_ticket.ki_performance_type = -1
        End Select
        
        '조기상환 유형
        deal_ticket.monthly_coupon_flag = 0
        deal_ticket.early_exit_flag = 0
        
        Select Case oRS("CLRD_TYPE_CODE")
        Case "17" 'safe stepdown
            If oRS("BARR_EN") = "1" Then
                deal_ticket.comment = deal_ticket.comment + "조기상환유형(CLRD_TYPE_CODE)과 KI barrier 유무(BARR_EN) 불일치"
            Else
                'KI HIT한 stepdown 상품으로 간주
'                deal_ticket.ki_barrier_flag = 1
'                deal_ticket.ki_touched_flag = 1
'                deal_ticket.ki_monitoring_freq = 1
            End If
        Case "22": deal_ticket.monthly_coupon_flag = 1
        Case "29": deal_ticket.early_exit_flag = 1
        Case "38": 'five-win
        End Select
        
        If for_frn_flag = True Or for_eswap_flag = True Then '월지급식 조기상환 구조를 이용한 equity-linked frn, equity swap pricing
            deal_ticket.monthly_coupon_flag = 1
        End If
        
        oRS.MoveNext
    
    Loop
    
    oRS.Close
    
    'Simulation Config
    deal_ticket.t_grid = Round((deal_ticket.maturity_date - deal_ticket.current_date) / 4, 0)
    
    '기초자산
    Dim ref_spot As Double
    Dim enum_ua As Variant
    'Dim eval_shift_flag As Boolean
    deal_ticket.has_eval_shift_ul = False '2025.05.23 eval_shift_flag를 deal_ticket에 추가
    
    ReDim bind_variable(1) As String
    ReDim bind_value(1) As Variant
    bind_variable(1) = ":indv_iscd"
    bind_value(1) = indv_iscd
        
    With oCmd

        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = getSQL(SQL_PATH_AC_DEAL_UL, bind_variable, bind_value)

        oRS.Open .Execute

    End With

    Dim no_of_ul As Integer
    no_of_ul = 0
    
    Do Until oRS.EOF
    
        no_of_ul = no_of_ul + 1
        
        '기초자산 수 설정 및 관련 배열 초기화
        deal_ticket.set_ul_dim no_of_ul
    
        'deal_ticket.set_ul_code oRS("UNAS_ISCD"), no_of_ul
        deal_ticket.set_ul_code get_ua_code(get_ua_idx(oRS("UNAS_ISCD"))), no_of_ul '2024.04.29 ul_code 정규화
        
        '한국 영업시간 종료기준 종가 미확정 기초자산이 포함된 경우, deal_ticket.has_eval_shift_ul = true (and eval_date_lag = true 이면 평가일 1일 이연 목적)
        For Each enum_ua In eval_shift_ua
            If enum_ua = get_ua_idx(deal_ticket.ul_code(no_of_ul)) Then
                deal_ticket.has_eval_shift_ul = True
            End If
        Next
        
'        If eval_date_lag = True Then
'            For Each enum_ua In EVAL_SHIFT_UA
'                If enum_ua = get_ua_idx(deal_ticket.ul_code(no_of_ul)) Then
'                    eval_shift_flag = True
'                End If
'            Next
'        End If
        
        '평가일 <= 거래일 or Quote 모드 경우, 최초기준가 100으로 설정
        If target_date <= deal_ticket.value_date Then
            ref_spot = get_spot(deal_ticket.ul_code(no_of_ul), target_date, True, adoCon)
            '종가 입수 누락된 경우, 음수 표시
            If ref_spot = 0 Then
                ref_spot = -1
            End If
        Else
            ref_spot = oRS("UNAS_INTL_PRC")
            '평가일 = 발행일 경우, 최초기준가 검증
            If target_date = deal_ticket.settlement_date And Abs(get_spot(deal_ticket.ul_code(no_of_ul), deal_ticket.value_date, False, adoCon) - ref_spot) > 0.0001 Then
                deal_ticket.comment = deal_ticket.comment + "기초자산 최초기준가(UNAS_INTL_PRC)가 종가와 일치하지 않음"
            End If
        End If
        deal_ticket.set_reference_price ref_spot, no_of_ul
        
        '기초자산별 ejected 여부: TBD
        deal_ticket.set_ejected_ul_flag 0, no_of_ul
        
        '기초자산별 ki barreir가 모두 동일: biz-one booking 확인 필요
        deal_ticket.set_ki_barrier ki_barrier, no_of_ul
    
        oRS.MoveNext
        
    Loop
    
    oRS.Close
    

    'Front DB -> deal_ticket
    'KI barrier shift: 자체헷지만 적용
    deal_ticket.ki_barrier_shift = 0
    
    If KI_SHIFT_ENABLE = True Then '2024.07.16
    
        With oCmd
        
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = "select KIBARRIER_SHIFT_SIZE from sps.ac_deal where asset_code ='" + indv_iscd + "'"
            
            oRS.Open .Execute
        
        End With
    
        Do Until oRS.EOF
        
            deal_ticket.ki_barrier_shift = oRS(0)
            oRS.MoveNext
        
        Loop
        
        oRS.Close
        
    End If
    
    
    'biz-one, front DB -> deal_ticket
    'autocall schedules w/ strike smoothing widths
    With oCmd

        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = getSQL(SQL_PATH_AC_SCHEDULE, bind_variable, bind_value)

        oRS.Open .Execute

    End With

    'unas_choc_mthd_code과 일치 여부 검증
    Dim ac_each_performance_type As Integer
    no_of_schedule = 0
    
    Do Until oRS.EOF

        no_of_schedule = no_of_schedule + 1
        
        Set a_schedule = New clsAutocallSchedule

        If deal_ticket.has_eval_shift_ul = True And eval_date_lag = True Then
            '모두 해외지수이고 call_date가 국내 연휴기간 중인 경우, target_date > TRTH_CLRD_DTRM_DATE+1 이더라도 상환여부 판단 필요.
            '이 경우, 해당 기간 중 해외지수 종가가 동일하게 국내 비영업일 기준으로 하루 지연되어 입수되어야 함.
            'TRTH_CLRD_DTRM_DATE:조기상환결정일=T
            'CLRD_DTRM_DATE:회계반영일=T(미확정 기초자산일 경우, 조기상환결정일에 T+1로 수정)
            'CLRD_DATE:조기상환일(결제)=T+2
            a_schedule.call_date = oRS("TRTH_CLRD_DTRM_DATE") + 1
        Else
            a_schedule.call_date = oRS("TRTH_CLRD_DTRM_DATE")
        End If
        a_schedule.set_percent_strike oRS("UNAS_SDRT1") / 100
        
        If for_frn_flag = True Then
            a_schedule.set_coupon_on_call 0
            a_schedule.strike_shift = 0
        Else
            If for_eswap_flag = True Then
                a_schedule.set_coupon_on_call oRS("CLRD_ERT") / 100 - 1
            Else
                a_schedule.set_coupon_on_call oRS("CLRD_ERT") / 100
            End If
            
            If IsEmpty(oRS("strike_smoothing_width")) Then
                a_schedule.strike_shift = 0
            Else
                'If shtACPricer.Range("chkStrikeSmoothing").value = True Then
                If PAYOFF_SMOOTHING_ENABLE = True Then '2024.07.16
                    a_schedule.strike_shift = oRS("strike_smoothing_width")
                Else
                    a_schedule.strike_shift = 0
                End If
            End If
            
        End If
        
        
        If oRS("AVRG_APLY_YN") = "Y" Then
            ac_each_performance_type = 0
        Else
            ac_each_performance_type = ac_every_performance_type
        End If
        
        If IsEmpty(oRS("performance_type")) Then
            a_schedule.performance_type = -1
        Else
            If ac_every_performance_type = oRS("performance_type") And ac_every_performance_type = ac_each_performance_type Then
                a_schedule.performance_type = oRS("performance_type")
            Else
                deal_ticket.comment = deal_ticket.comment + "조기상환평가일별 performance_type이 unas_choc_mthd_code과 일치하지 않음"
            End If
        End If
        
        'for the ejectable structure : dll(2021.11.12)
        If deal_ticket.ejectable_flag = True Then
            a_schedule.ejected_event_flag = 1
            a_schedule.ejectable_order = 0 'TBD
        Else
            a_schedule.ejected_event_flag = 0
        End If

        push_back_clsAutocallSchedule schedule_list, a_schedule

        oRS.MoveNext

    Loop

    oRS.Close
    
    deal_ticket.no_of_schedule = no_of_schedule
    
    '마지막 조기상환일 = 만기일과 같은지 확인 필요
    If schedule_list(deal_ticket.no_of_schedule).call_date <> deal_ticket.maturity_date Then
        deal_ticket.comment = deal_ticket.comment + "마지막 조기상환평가일과 만기일이 일치하지 않음"
        'schedule_list(deal_ticket.no_of_schedule).call_date = deal_ticket.maturity_date '2024.02.26
        deal_ticket.maturity_date = schedule_list(deal_ticket.no_of_schedule).call_date '2024.05.16
        'Update simulation config
        deal_ticket.t_grid = Round((deal_ticket.maturity_date - deal_ticket.current_date) / 4, 0) '2024.05.16
    End If
    
    deal_ticket.set_schedule_array schedule_list
    deal_ticket.strike_at_maturity = schedule_list(deal_ticket.no_of_schedule).percent_strike
    deal_ticket.coupon_at_maturity = schedule_list(deal_ticket.no_of_schedule).coupon_on_call

    'safe stepdown인 경우, KI HIT한 stepdown 상품으로 간주하고 마지막 조기상환 베리어로 KI 설정
'    If 1 Then
'        For i = 1 To deal_ticket.no_of_ul
'            deal_ticket.set_ki_barrier deal_ticket.strike_at_maturity, i
'        Next i
'    End If

    For i = 1 To deal_ticket.no_of_schedule
        Set schedule_list(i) = Nothing
    Next i
    
    'monthly coupon schedules
    no_of_schedule = 0
    If deal_ticket.monthly_coupon_flag = 1 Then
        
        '쿠폰 지급일 관리
        With oCmd
        
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = "SELECT to_date(TRTH_CUPN_VLTN_DATE,'YYYYMMDD') TRTH_CUPN_VLTN_DATE FROM BSYS.TBSIMO213L00@GDW WHERE INDV_ISCD = '" + indv_iscd + "' ORDER BY 1 ASC"
            
            oRS.Open .Execute
        
        End With
    
    
        '쿠폰 구간 관리
        With oCmd
        
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = "SELECT SCTN_STA_CTNU, SCTN_FIN_CTNU, BONS_CUPN_STA_SDRT, BONS_CUPN_INRT FROM BSYS.TBSIMO210L00@GDW  WHERE INDV_ISCD = '" + indv_iscd + "' ORDER BY 1 ASC"
            
            oRS2.Open .Execute
        
        End With

        Do Until oRS.EOF
            
            no_of_schedule = no_of_schedule + 1
            
            ReDim Preserve call_dates(1 To no_of_schedule) As Date
            ReDim Preserve strikes(1 To no_of_schedule) As Double
            ReDim Preserve coupons(1 To no_of_schedule) As Double
            
            If deal_ticket.has_eval_shift_ul = True And eval_date_lag = True Then
                call_dates(no_of_schedule) = oRS("TRTH_CUPN_VLTN_DATE") + 1
            Else
                call_dates(no_of_schedule) = oRS("TRTH_CUPN_VLTN_DATE")
            End If
             
            Do Until oRS2.EOF
                If no_of_schedule >= oRS2("SCTN_STA_CTNU") And no_of_schedule <= oRS2("SCTN_FIN_CTNU") Then
                    strikes(no_of_schedule) = oRS2("BONS_CUPN_STA_SDRT") / 100
                    coupons(no_of_schedule) = oRS2("BONS_CUPN_INRT") / 100
                    oRS2.MoveNext
                Else
                    oRS2.MoveNext
                End If
            Loop
            oRS2.MoveFirst
            
            oRS.MoveNext
        
        Loop
   
        oRS.Close
        oRS2.Close
        
        If no_of_schedule > 0 Then
            deal_ticket.monthly_coupon_amount = coupons(no_of_schedule)
            
            '마지막 월지급평가일이 만기일 이후인 경우, 만기일로 강제 조정 2024.02.26
            If call_dates(no_of_schedule) > deal_ticket.maturity_date Then
                call_dates(no_of_schedule) = deal_ticket.maturity_date
            End If
        End If
        

        
        If for_frn_flag = True Or for_eswap_flag = True Then
            
            '변동금리 스케줄 관리
            With oCmd
            
                .ActiveConnection = adoCon
                .CommandType = adCmdText
                .CommandText = "SELECT nvl(CRCD,'KRW') CRCD, PRVS_INDT_MNRT_CODE, PRVS_DATE, INT_AMT, INDT_MNRT_ADJS_DATE, A.CCLS_MNRT, PRVS_ADTN_MNRT, VLTN_STA_DATE, VLTN_FIN_DATE FROM BSYS.TBSIMO216L00@GDW A, BSYS.TBSIMO201M00@GDW B WHERE A.INDV_ISCD=B.INDV_ISCD AND A.INDV_ISCD = '" + indv_iscd + "' AND RCVN_CLS_CODE='2' ORDER BY PRVS_DATE ASC"
                
                oRS.Open .Execute
            
            End With
            
            Do Until oRS.EOF
                
                Dim daycount As Integer
                daycount = 365
                
                If oRS("CRCD") = "USD" Then
                    daycount = 360
                End If
                
                Dim foward_rate As Double
                Dim add_rate As Double
                
                no_of_schedule = no_of_schedule + 1
                
                ReDim Preserve call_dates(1 To no_of_schedule) As Date
                ReDim Preserve strikes(1 To no_of_schedule) As Double
                ReDim Preserve coupons(1 To no_of_schedule) As Double
                
'                If deal_ticket.has_eval_shift_ul = True And eval_date_lag = True Then
'                    call_dates(no_of_schedule) = str2date(oRS("PRVS_DATE")) + 1
'                Else
                    call_dates(no_of_schedule) = str2date(oRS("PRVS_DATE"))
'                End If

                '마지막 floating leg 스케줄이 만기일 이후인 경우, 만기일로 강제 조정 2024.02.26
                If call_dates(no_of_schedule) > deal_ticket.maturity_date Then
                    call_dates(no_of_schedule) = deal_ticket.maturity_date
                End If
                
                strikes(no_of_schedule) = 0
                
                '2024.03.04
                If oRS("CCLS_MNRT") <> 0 Then
                    If call_dates(no_of_schedule) <= deal_ticket.current_date Then
                        '비즈원에 계산된 변동금리 CCLS_MNRT 사용: CD/LIBOR는 fixing 값, USD-SOFR는 현재까지 누적 관찰값 (VUS0003M의 경우 CCLS_MNRT는 누적 SOFR에 26.16bp가 더해진 값임)
                        coupons(no_of_schedule) = (oRS("CCLS_MNRT") + oRS("PRVS_ADTN_MNRT")) / 100 * (str2date(oRS("VLTN_FIN_DATE")) - str2date(oRS("VLTN_STA_DATE"))) / daycount
                    ElseIf oRS("CRCD") = "KRW" And (call_dates(max(no_of_schedule - 1, 1)) <= deal_ticket.current_date Or no_of_schedule = 1) Then
                        '비즈원에 계산된 변동금리 CCLS_MNRT 사용: CD/LIBOR는 fixing 값, USD-SOFR는 현재까지 누적 관찰값 (VUS0003M의 경우 CCLS_MNRT는 누적 SOFR에 26.16bp가 더해진 값임)
                        coupons(no_of_schedule) = (oRS("CCLS_MNRT") + oRS("PRVS_ADTN_MNRT")) / 100 * (str2date(oRS("VLTN_FIN_DATE")) - str2date(oRS("VLTN_STA_DATE"))) / daycount
                    Else
                        '변동금리 정보 없을 경우, 이전 스케줄 값 사용
                        'coupons(no_of_schedule) = Abs(coupons(no_of_schedule - 1))
                        
                        '변동금리 정보 없을 경우, forward rate 계산
                        'foward_rate = market_set.dcf_by_ccy(oRS("CRCD")).get_fwd_rate(str2date(oRS("VLTN_STA_DATE")), str2date(oRS("VLTN_FIN_DATE")))
                        '2024-03-08 dcf가 credit curve일 수도 있으므로 swap curve로 변경
                        If deal_ticket.ccy = get_dcf_ccy(DCF.KRW) Then
                            foward_rate = market_set.market_by_ul(get_ua_code(ua.KOSPI200)).rate_curve_.get_fwd_rate(str2date(oRS("VLTN_STA_DATE")), str2date(oRS("VLTN_FIN_DATE")))
                        ElseIf deal_ticket.ccy = get_dcf_ccy(DCF.USD) Then
                            foward_rate = market_set.market_by_ul(get_ua_code(ua.SPX)).rate_curve_.get_fwd_rate(str2date(oRS("VLTN_STA_DATE")), str2date(oRS("VLTN_FIN_DATE")))
                        End If
                        
                        add_rate = oRS("PRVS_ADTN_MNRT") / 100
                        
                        'VUS0003M의 경우, SOFR 보정을 위해 forward에 26.16bp 가산 필요
                        If oRS("PRVS_INDT_MNRT_CODE") = "VUS0003M" Or (oRS("PRVS_INDT_MNRT_CODE") = "LIBORUSD" And oRS("INDT_MNRT_ADJS_DATE") >= "20230701") Then
                            foward_rate = foward_rate + 0.002616
                        End If
                        
                        coupons(no_of_schedule) = (foward_rate + add_rate) * (str2date(oRS("VLTN_FIN_DATE")) - str2date(oRS("VLTN_STA_DATE"))) / daycount
                    End If
                Else
                    '변동금리 정보 없을 경우, 이전 스케줄 값 사용
                    'coupons(no_of_schedule) = Abs(coupons(no_of_schedule - 1))
                    
                    '변동금리 정보 없을 경우, forward rate 계산
                    'foward_rate = market_set.dcf_by_ccy(oRS("CRCD")).get_fwd_rate(str2date(oRS("VLTN_STA_DATE")), str2date(oRS("VLTN_FIN_DATE")))
                    '2024-03-08 dcf가 credit curve일 수도 있으므로 swap curve로 변경
                    If deal_ticket.ccy = get_dcf_ccy(DCF.KRW) Then
                        foward_rate = market_set.market_by_ul(get_ua_code(ua.KOSPI200)).rate_curve_.get_fwd_rate(str2date(oRS("VLTN_STA_DATE")), str2date(oRS("VLTN_FIN_DATE")))
                    ElseIf deal_ticket.ccy = get_dcf_ccy(DCF.USD) Then
                        foward_rate = market_set.market_by_ul(get_ua_code(ua.SPX)).rate_curve_.get_fwd_rate(str2date(oRS("VLTN_STA_DATE")), str2date(oRS("VLTN_FIN_DATE")))
                    End If
                    
                    add_rate = oRS("PRVS_ADTN_MNRT") / 100
                    
                    'VUS0003M의 경우, SOFR 보정을 위해 forward에 26.16bp 가산 필요
                    If oRS("PRVS_INDT_MNRT_CODE") = "VUS0003M" Or (oRS("PRVS_INDT_MNRT_CODE") = "LIBORUSD" And oRS("INDT_MNRT_ADJS_DATE") >= "20230701") Then
                        foward_rate = foward_rate + 0.002616
                    End If
                    
                    coupons(no_of_schedule) = (foward_rate + add_rate) * (str2date(oRS("VLTN_FIN_DATE")) - str2date(oRS("VLTN_STA_DATE"))) / daycount
                End If
            
                If for_eswap_flag = True Then
                    coupons(no_of_schedule) = -1 * coupons(no_of_schedule)
                End If
                
                oRS.MoveNext
                
               
                
                
            Loop
       
            oRS.Close
            
'''''''''

            If no_of_schedule > 0 Then
           '마지막 floating leg 스케줄 ~ 만기일 사이에 스케줄이 비어 있을 경우, 스케줄 강제 입력 2024.03.04
            Do While deal_ticket.maturity_date - call_dates(no_of_schedule) >= 90
                no_of_schedule = no_of_schedule + 1
                
                ReDim Preserve call_dates(1 To no_of_schedule) As Date
                ReDim Preserve strikes(1 To no_of_schedule) As Double
                ReDim Preserve coupons(1 To no_of_schedule) As Double
                
                call_dates(no_of_schedule) = call_dates(no_of_schedule - 1) + 90
                strikes(no_of_schedule) = 0
                
                foward_rate = market_set.dcf_by_ccy(deal_ticket.ccy).get_fwd_rate(call_dates(no_of_schedule - 1), call_dates(no_of_schedule))
                coupons(no_of_schedule) = (foward_rate + add_rate) * 90 / daycount
                
                If for_eswap_flag = True Then
                    coupons(no_of_schedule) = -1 * coupons(no_of_schedule)
                End If

            Loop
            
            '/<----- 2024.02.07
            Dim inx_target As Integer
            Dim inx_shift As Integer
            
            Dim date_target As Date
            Dim coupon_target As Double
            Dim strike_target As Double
            
            For inx = 2 To no_of_schedule
            
                'recursive function
                'input: call_dates(inx)
                'array: result(1~inx-1)
                'output: result(1~inx) -> redim result() as date

                
                'sort_dates(result, call_dates(inx))
            
                'search index
                inx_target = search_inx(call_dates, call_dates(inx))
                
                If inx_target > 0 And inx_target < inx + 1 Then
                'sorting by call_dates
                
                    'temp
                    date_target = call_dates(inx)
                    coupon_target = coupons(inx)
                    strike_target = strikes(inx)
                
                    'shift
                    For inx_shift = inx - 1 To inx_target Step -1
                        call_dates(inx_shift + 1) = call_dates(inx_shift)
                        coupons(inx_shift + 1) = coupons(inx_shift)
                        strikes(inx_shift + 1) = strikes(inx_shift)
                    Next inx_shift
                    
                    'insert
                    call_dates(inx_target) = date_target
                    coupons(inx_target) = coupon_target
                    strikes(inx_target) = strike_target
               
                End If
                
            Next inx
            '----->/
            
            '2024.03.07 조기상환결정일 이후에 equity swap의 floating leg 결제일이 발생항 경우, 해당 현금흐름이 누락됨을 방지
            For inx = 1 To deal_ticket.no_of_schedule
            
                For jnx = 1 To no_of_schedule
                
                    If call_dates(jnx) > deal_ticket.autocall_schedules(inx).call_date And call_dates(jnx) - deal_ticket.autocall_schedules(inx).call_date < 10 Then
                        call_dates(jnx) = deal_ticket.autocall_schedules(inx).call_date
                        Exit For
                    End If
                    
                Next jnx
            
            Next inx
            '----->/
            
            If deal_ticket.monthly_coupon_amount = 0 Then
            'If no_of_schedule > 0 And deal_ticket.monthly_coupon_amount = 0 Then
                deal_ticket.monthly_coupon_amount = coupons(no_of_schedule)
            End If
            End If
                
                '''''''''
            
        End If
        
    End If
    
    deal_ticket.set_coupon_schedule no_of_schedule, call_dates, strikes, coupons
    
    '상환결정일2
    'early-exit schedules: 현재까지 최대 5개 구조 발행
    'five-wins: 6회 x 연속 5일간 기회 부여 = 30개 구조 발행
    'hi-five: double strike 구조로 2nd strike 충족시 추가 쿠폰 지급. 6개 구조 발행
    no_of_schedule = 0
    If deal_ticket.early_exit_flag = 1 Then
        
        With oCmd
        
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = "SELECT INDV_ISCD, CLRD_DTRM_DATE, CALC_STA_DATE, CALC_FIN_DATE, CLRD_BARR_HIT_YN, CLRD_BARR_VAL, CLRD_INRT FROM BSYS.TBSIMO227L00@GDW WHERE INDV_ISCD = '" + indv_iscd + "' ORDER BY SEQ_SRNO"
            
            oRS.Open .Execute
        
        End With

        Do Until oRS.EOF
            
            no_of_schedule = no_of_schedule + 1
            deal_ticket.redim_early_exit_barrier no_of_schedule
        
            ReDim Preserve call_dates(1 To no_of_schedule) As Date
            ReDim Preserve coupons(1 To no_of_schedule) As Double
            ReDim Preserve strike_shifts(1 To no_of_schedule) As Double
            ReDim Preserve early_exit_touched_flags(1 To no_of_schedule) As Long  'data type 변경: dll(2018.7.17)
            ReDim Preserve early_exit_performance_types(1 To no_of_schedule) As Long  'dll(2018.8.8)
            ReDim Preserve early_exit_barrier_types(1 To no_of_schedule) As Long  'dll(2018.8.8)
            
            
            If str2date(oRS("CALC_STA_DATE")) <> deal_ticket.settlement_date Then
                deal_ticket.comment = deal_ticket.comment + "early-exit 관찰시작일(CALC_STA_DATE)이 발행일과 일치하지 않음"
            End If
            
            If deal_ticket.has_eval_shift_ul = True And eval_date_lag = True Then
                call_dates(no_of_schedule) = str2date(oRS("CALC_FIN_DATE")) + 1
            Else
                call_dates(no_of_schedule) = str2date(oRS("CALC_FIN_DATE"))
            End If
            
            If for_eswap_flag = True Then
                coupons(no_of_schedule) = oRS("CLRD_INRT") / 100 - 1
            Else
                coupons(no_of_schedule) = oRS("CLRD_INRT") / 100
            End If
            
            strike_shifts(no_of_schedule) = 0
            
            For no_of_ul = 1 To deal_ticket.no_of_ul
                deal_ticket.set_early_exit_barrier oRS("CLRD_BARR_VAL") / 100, deal_ticket.no_of_ul * (no_of_schedule - 1) + no_of_ul
            Next no_of_ul
            
            If oRS("CLRD_BARR_HIT_YN") = "Y" Then
                early_exit_touched_flags(no_of_schedule) = 1#
            Else
                early_exit_touched_flags(no_of_schedule) = 0#
            End If

            'Front db
            early_exit_performance_types(no_of_schedule) = -1  'dll(2018.8.8)
            early_exit_barrier_types(no_of_schedule) = -1  'dll(2018.8.8)
            
            oRS.MoveNext
        
        Loop
   
        oRS.Close

    End If
        
    deal_ticket.set_early_exit_schedule no_of_schedule, call_dates, coupons, strike_shifts, early_exit_touched_flags, early_exit_performance_types, early_exit_barrier_types
    
    Set get_ac_deal_ticket = deal_ticket
    
    Set deal_ticket = Nothing
    
    Exit Function

ErrorHandler:

    raise_err "get_ac_deal_ticket", Err.description

End Function

'2024.02.07
Private Function search_inx(schedules() As Date, target_date As Date) As Integer

    Dim rtn As Integer
    Dim i As Integer
    
    For i = 1 To UBound(schedules)
        
        If CLng(schedules(i)) > CLng(target_date) Or CLng(schedules(i)) = 0 Then
            rtn = i
            Exit For
        End If
        
    Next i
    
    search_inx = rtn

End Function

'당일 기초자산 가격 반영여부 검증 (미완성)
Public Sub check_barrier_hit(ac_deal_ticket, market_set)
    
    Dim ua_eval_spot As Double
    Dim ua_spot As Double
    
    '당일 KI HIT 검증 -> deal_ticket 대사
    If deal_ticket.ki_barrier_flag = 1 Then
        For i = 1 To deal_ticket.no_of_ul
            
            ua_spot = get_spot(deal_ticket.ul_code(i), target_date, adoCon) / deal_ticket.reference_price(i)
            
            If i = 1 Then
                ua_eval_spot = ua_spot
            End If
        
            Select Case deal_ticket.ki_performance_type
            Case -1: ua_eval_spot = min(ua_eval_spot, ua_spot) 'worst
            Case 0: ua_eval_spot = ua_eval_spot + ua_spot / deal_ticket.no_of_ul 'avg
            Case 1: ua_eval_spot = max(ua_eval_spot, ua_spot) 'best
            End Select
            
        Next i
        
        For i = 1 To deal_ticket.no_of_ul
            If deal_ticket.ki_barrier(i) >= ua_eval_spot And deal_ticket.ki_touched_flag = 0 Then
               Err.description = "당일 KI 터치하였으나, 원장(BARR_HIT_CLS_CODE) 반영되지 않음."
               Exit For
            End If
        Next i
    End If
    
    '당일 EE HIT 검증 -> deal_ticket 대사
    If deal_ticket.early_exit_flag = 1 Then
        
        '현재 스케줄 j 찾기
    
        For i = 1 To deal_ticket.no_of_ul
            
            ua_spot = get_spot(deal_ticket.ul_code(i), target_date, adoCon) / deal_ticket.reference_price(i)
            
            If i = 1 Then
                ua_eval_spot = ua_spot
            End If
        
            Select Case deal_ticket.early_exit_barrier_types(j)
            Case -1: ua_eval_spot = min(ua_eval_spot, ua_spot) 'worst
            Case 0: ua_eval_spot = ua_eval_spot + ua_spot / deal_ticket.no_of_ul 'avg
            Case 1: ua_eval_spot = max(ua_eval_spot, ua_spot) 'best
            End Select
            
        Next i
        
        If deal_ticket.early_exit_barrier(j) >= ua_eval_spot And deal_ticket.early_exit_touched_flags(j) = 0 Then
           Err.description = "당일 EE 터치하였으나, 원장(CLRD_BARR_HIT_YN) 반영되지 않음."
        End If
        
    End If
    
    '당일 AC 검증
    For j = 1 To deal_ticket.no_of_schedule
        
        If target_date = deal_ticket.autocall_schedules(j).call_date Then
        
            For i = 1 To deal_ticket.no_of_ul
            
                ua_spot = get_spot(deal_ticket.ul_code(i), target_date, adoCon) / deal_ticket.reference_price(i)
                
                If i = 1 Then ua_eval_spot = ua_spot
                End If
            
                Select Case deal_ticket.autocall_schedules(j).performance_type
                Case -1: ua_eval_spot = min(ua_eval_spot, ua_spot) 'worst
                Case 0: ua_eval_spot = ua_eval_spot + ua_spot / deal_ticket.no_of_ul 'avg
                Case 1: ua_eval_spot = max(ua_eval_spot, ua_spot) 'best
                End Select
            
            Next i
            
            If deal_ticket.early_exit_flag = 1 Then
            End If
        
            If deal_ticket.autocall_schedules(j).strike_value < ua_eval_spot Then
               '조기상환 성공
            '   당일 월지급 쿠폰 검증
            '   당일 EE 쿠폰 검증 -> 상환실패시, 상환일정 및 만기일에 적용된 eval_shift 취소
            ElseIf target_date = deal_ticket.maturity_date Then
                If deal_ticket.ki_touched_flag = 0 Then
                    '만기상환: dummay 지급
                Else
                    '만기상환: 원금손실
                End If
            Else
                '조기상환 이연
                '상환일정 및 만기일에 적용된 eval_shift 취소
                
            End If
        
        End If
        
    Next j
    
End Sub
            

Private Function get_dummy(indv_iscd As String, ki_barrier As Double, adoCon As adoDB.Connection) As Double

    Dim rtn As Double
    rtn = 0
    
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    
    'KI 미발생시 만기상환 조건
    ReDim bind_variable(1) As String
    ReDim bind_value(1) As Variant
    bind_variable(1) = ":indv_iscd"
    bind_value(1) = indv_iscd
    
    With oCmd

        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = getSQL(SQL_PATH_AC_DEAL_DUMMY, bind_variable, bind_value)

        oRS.Open .Execute

    End With
    
    Dim prev_sdpr_loeq As Double
        
    Do Until oRS.EOF
        
        If oRS("SRNO") = 1 Then
            prev_sdpr_loeq = oRS("SDPR_LOEQ_RATE")
        End If
        
        '구간 연속 확인
        If oRS("SDPR_LOEQ_RATE") = prev_sdpr_loeq And oRS("SDPR_EXCS_RATE") > prev_sdpr_loeq Then
        
            
            '참여율 = 100%, 상수 = 0%
            If oRS("SDPR_LOEQ_RATE") = 0 And oRS("INVL_RATE") = 100 And oRS("MTRT_ERT") = 0 Then
                
                '만기상환조건과 정합성 검증. 하단에서 마지막 ac barrier 확인
                If oRS("SDPR_EXCS_RATE") / 100 <> ki_barrier Then
                    Err.description = "만기상환조건 구간(SDPR_EXCS_RATE)과 메인 화면의 KI barrier(BARR_VAL1) 불일치"
                End If
            
            '참여율 = 0%
            ElseIf oRS("SDPR_LOEQ_RATE") > 0 And oRS("SDPR_EXCS_RATE") >= 999 And oRS("INVL_RATE") = 0 Then
                            
                '상수
                rtn = oRS("MTRT_ERT") / 100
                Exit Do
            
            End If
            
            prev_sdpr_loeq = oRS("SDPR_EXCS_RATE")
                    
        Else
            Err.description = "만기수익율 구간(SDPR_LOEQ_RATE,SDPR_EXCS_RATE)이 연속하지 않음"
        End If
        
        oRS.MoveNext

    Loop

    oRS.Close
    
    get_dummy = rtn

End Function


Public Sub insert_price(tdate As String, code As String, price_theo As Double, price_mtm As Double, adoCon As adoDB.Connection)
    
    Dim sql As String
        
    sql = "delete from ras.rm_pricing_data@rms01 where tdate = '" + tdate + "' and code = '" + code + "' "
    
    adoCon.Execute (sql)
    
    sql = "insert into ras.rm_pricing_data@rms01 values ('" + tdate + "','" + code + "', " & price_theo & ", " & price_mtm & ", SYSDATE, 'EXCEL') "
    
    adoCon.Execute (sql)


End Sub
        

Public Sub insert_greeks(greeks As clsGreeks, deal_ticket As clsACDealTicket, market_set As clsMarketSet, adoCon As adoDB.Connection)

    Dim sql As String
    
    Dim tdate As String
    tdate = date2str(deal_ticket.current_date)
            
    Dim fx As Double
    If deal_ticket.ccy = get_ccy_code(ccy.KRW) Then
        fx = 1
    Else
        fx = get_fx(get_ccy_code(ccy.KRW), deal_ticket.ccy, deal_ticket.current_date, adoCon)
    End If
    
    Dim baseasset_code As String
    
    Dim delta As Double
    Dim gamma As Double
    Dim vega As Double
    Dim theta As Double
    
    Dim delta_expsorue As Double
    Dim gamma_exposure As Double
    Dim vega_exposure As Double
    
    Dim tau As Double
    Dim dur_els As Double
    Dim close_spot As Double
    Dim wp As Double
    wp = MAX_UA_PCT_PRICE
    
    'dur_els =
    
    Call delete_dv01(tdate, deal_ticket.asset_code, adoCon)
    
    Dim j As Integer
    For j = 1 To deal_ticket.no_of_ul
        
        If deal_ticket.value_date = deal_ticket.current_date And is_eval_shift_ua(deal_ticket.ul_code(j)) = True Then
            '최초기준가설정일 기초자산이 확정되지 않은 경우 skip: delta를 잡지 않았다고 간주하고 greeks 집계하지 않음(원칙은 집계해야함)
        Else
            baseasset_code = get_ua_code(get_ua_idx(deal_ticket.ul_code(j)), UA_CODE_TYPE.ISIN)
        
            close_spot = market_set.market_by_ul(deal_ticket.ul_code(j)).s_
            
            delta = (greeks.deltas(j) * 0.75 + greeks.sticky_moneyness_deltas(j) * 0.25) * fx / deal_ticket.qty
            gamma = greeks.gammas(j) * fx / deal_ticket.qty
            vega = greeks.vegas(j) * fx / deal_ticket.qty
            
            delta_expsorue = deal_ticket.qty * delta * close_spot
            gamma_exposure = 0.5 * deal_ticket.qty * gamma * (close_spot * 0.01) ^ 2
            vega_exposure = deal_ticket.qty * vega
            
            If j = 1 Then
                theta = greeks.theta / 365 * fx / deal_ticket.qty '1day
            Else
                theta = 0
            End If
                        
            '----- rcs.pml_greek -----
            sql = "delete from rcs.pml_greek@rms01 where tdate = '" + tdate + "' and stk_code = '" + deal_ticket.asset_code + "' and baseasset_code = '" + baseasset_code + "'"

            adoCon.Execute (sql)

            sql = "insert into rcs.pml_greek@rms01 values ('" & tdate & "','" & deal_ticket.fund_code_c & "','" & deal_ticket.asset_code & "','" & baseasset_code & "','파생결합증권'," & close_spot & "," & delta & "," & gamma & "," & vega & "," & delta_expsorue & "," & gamma_exposure & "," & vega_exposure & ",SYSDATE,'EXCEL','EXCEL',null,null,null,null,null,null,null," & theta & ",'351',null)"

            adoCon.Execute (sql)

            ' 향후, oracle precedure로 수행 검토 필요: rcs.pml_greek -> ras.if_otc_template_factor, ras.if_otc_template_data
            '----- ras.if_otc_template_factor -----
            sql = "delete from ras.if_otc_template_factor where tdate = '" + tdate + "' and code = '" + deal_ticket.asset_code + "' and factorid = '" + baseasset_code + "'"

            adoCon.Execute (sql)

            sql = "insert into ras.if_otc_template_factor values ('" & tdate & "','" & deal_ticket.asset_code & "','" & baseasset_code & "'," & close_spot & "," & delta & "," & gamma & ",0,SYSDATE,'EXCEL','EXCEL')"

            adoCon.Execute (sql)

            '----- ras.if_otc_template_data -----
            sql = "delete from ras.if_otc_template_data where tdate = '" + tdate + "' and code = '" + deal_ticket.asset_code + "'"

            adoCon.Execute (sql)

            sql = "insert into ras.if_otc_template_data (tdate, code, endprice) values ('" & tdate & "','" & deal_ticket.asset_code & "'," & greeks.value / deal_ticket.notional * deal_ticket.unit_notional & ")"

            adoCon.Execute (sql)
            
            'RM StickyMoneyness델타: greeks.sticky_moneyness_deltas(j) * fx
            
            'RM 바나: greeks.vannas(j) * fx
            
            'RM 스큐: greeks.skew_s_s(j) * fx
    
            'RM 로(기초자산통화커브): greeks.rho_ul(j) * fx
            
            'tau 추정식의 분자/분모가 너무 작은 경우, 0으로 강제 조정
            If (Abs(greeks.rho_ul(j) * 10000) < 1) Or (Abs(greeks.deltas(j)) < 1) Then '0으로 강제 조정되는 종목이 너무 많음. tau의 분자와 동일하게 *10000 추가 2024.06.20
                tau = 0
            Else
                tau = greeks.rho_ul(j) / (greeks.deltas(j) * greeks.ul_prices(j)) * 10000
                
                'lower bound = 2주
                If tau < 1 / 24 Then
                    tau = 1 / 24
                End If
                'upper bound = 잔존만기 2024.06.26
                If tau > (deal_ticket.maturity_date - deal_ticket.current_date) / 365 Then
                    tau = (deal_ticket.maturity_date - deal_ticket.current_date) / 365
                End If
            End If
            
            '기초자산 통화 DV01 입력
            Call insert_dv01(tdate, deal_ticket.asset_code, market_set.market_by_ul(deal_ticket.ul_code(j)).ul_currency, tau, greeks.rho_ul(j) * fx, adoCon)
                        
            'ELS duration: wp의 greeks로 추정. 이전에는 max tau로 추정했으나 max값이 기초자산 wp가 아닌 경우가 발생하여 로직 수정.
            If close_spot / deal_ticket.reference_price(j) < wp Then
                dur_els = tau
            End If
        
        End If

    Next j
    
    'RM 로(할인커브): 명목금액 통화 DV01 입력
    Call insert_dv01(tdate, deal_ticket.asset_code, deal_ticket.ccy, dur_els, greeks.rho * fx, adoCon)
    
    'CrossGamma
'        If deal_ticket.no_of_ul > 1 Then
'            shtELS.Range("close_xgamma_col").Cells(i + 9, calc_xgamma_ofs(deal_ticket.ul_code(1), deal_ticket.ul_code(2))) = greeks.cross_gamma12 * fx
'            If deal_ticket.no_of_ul = 3 Then
'                shtELS.Range("close_xgamma_col").Cells(i + 9, calc_xgamma_ofs(deal_ticket.ul_code(1), deal_ticket.ul_code(3))) = greeks.cross_gamma13 * fx
'                shtELS.Range("close_xgamma_col").Cells(i + 9, calc_xgamma_ofs(deal_ticket.ul_code(2), deal_ticket.ul_code(3))) = greeks.cross_gamma23 * fx
'            End If
'        End If

End Sub

Public Sub insert_scenario_result(greeks As clsGreeks, deal_ticket As clsACDealTicket, market_set As clsMarketSet, adoCon As adoDB.Connection, scenario_id As String)

    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    
    Dim sql As String
    
    Dim tdate As String
    tdate = date2str(deal_ticket.current_date)
    
    Dim fx0 As Double
    Dim fx As Double
    If deal_ticket.ccy = get_ccy_code(ccy.KRW) Then
        fx0 = 1
        fx = 1
    Else
        fx0 = get_fx(get_ccy_code(ccy.KRW), deal_ticket.ccy, deal_ticket.current_date, adoCon)
        fx = get_fx(get_ccy_code(ccy.KRW), deal_ticket.ccy, deal_ticket.current_date, adoCon, scenario_id)
    End If
    
    Dim baseasset_code As String
    
    Dim delta As Double
    Dim gamma As Double
    Dim vega As Double
    Dim theta As Double
    
    Dim delta_exposure As Double
    Dim gamma_exposure As Double
    Dim vega_exposure As Double
    
    Dim close_spot As Double

    Dim gds_tp As String
    gds_tp = "파생결합증권"
    
    sql = "delete from rcs.pml_position_st where tdate = '" + tdate + "' and stk_code = '" + deal_ticket.asset_code + "' and scenarioid = '" + scenario_id + "'"
    
    adoCon.Execute (sql)
    
    sql = "delete from rcs.pml_greek_st where tdate = '" + tdate + "' and stk_code = '" + deal_ticket.asset_code + "' and scenarioid = '" + scenario_id + "'"
    
    adoCon.Execute (sql)
    
    sql = "delete from rms.mr_scenario_detail_data where tdate = '" + tdate + "' and code = '" + deal_ticket.asset_code + "' and scenarioid = '" + scenario_id + "'"
    
    adoCon.Execute (sql)

    Dim j As Integer
    For j = 1 To deal_ticket.no_of_ul
        
        If deal_ticket.value_date = deal_ticket.current_date And is_eval_shift_ua(deal_ticket.ul_code(j)) = True Then
            '최초기준가설정일 기초자산이 확정되지 않은 경우 skip: delta를 잡지 않았다고 간주하고 greeks 집계하지 않음(원칙은 집계해야함)
        Else
            baseasset_code = get_ua_code(get_ua_idx(deal_ticket.ul_code(j)), UA_CODE_TYPE.ISIN)
        
            close_spot = market_set.market_by_ul(deal_ticket.ul_code(j)).s_
            
            delta = (greeks.deltas(j) * 0.75 + greeks.sticky_moneyness_deltas(j) * 0.25) * fx / deal_ticket.qty
            gamma = greeks.gammas(j) * fx / deal_ticket.qty
            vega = greeks.vegas(j) * fx / deal_ticket.qty
            
            delta_exposure = deal_ticket.qty * delta * close_spot
            gamma_exposure = 0.5 * deal_ticket.qty * gamma * (close_spot * 0.01) ^ 2
            vega_exposure = deal_ticket.qty * vega
            
            If j = 1 Then
                theta = greeks.theta / 365 * fx / deal_ticket.qty '1day
            Else
                theta = 0
            End If
                        
            '----- rcs.pml_greek -----
'            sql = "delete from rcs.pml_greek_st@rms01 where tdate = '" + tdate + "' and stk_code = '" + deal_ticket.asset_code + "' and baseasset_code = '" + baseasset_code + "'"
'
'            adoCon.Execute (sql)

'            sql = "insert into rcs.pml_greek_st@rms01 values ('" & tdate & "','" & deal_ticket.fund_code_c & "','" & deal_ticket.asset_code & "','" & baseasset_code & "','파생결합증권'," & close_spot & "," & delta & "," & gamma & "," & vega & "," & delta_exposure & "," & gamma_exposure & "," & vega_exposure & ",SYSDATE,'EXCEL','EXCEL',null,null,null,null,null,null,null," & theta & ",'351',null)"
            sql = "insert into rcs.pml_greek_st (tdate, scenarioid, fund_code, stk_code, baseasset_code, gds_tp, close_amt, " _
                & " delta, gamma, vega, delta_exposure, gamma_exposure, vega_exposure, work_time, work_trm, work_memb, dept_code) values (" _
                & "'" + tdate + "','" + scenario_id + "','" + deal_ticket.fund_code_c + "','" + deal_ticket.asset_code + "','" + baseasset_code + "', " _
                & "'" + gds_tp + "', " & close_spot & "," & delta & "," & gamma & "," & vega & "," & delta_exposure & ", " & gamma_exposure & "," & vega_exposure & ",SYSDATE, 'EXCEL','EXCEL','351') "
                    
            adoCon.Execute (sql)


        
        End If

    Next j

    Dim unitPrice As Double
    Dim unitPrice0 As Double
    
    unitPrice = Round(greeks.value / deal_ticket.notional * deal_ticket.unit_notional, 5)

    sql = "insert into rcs.pml_position_st (tdate, scenarioid, fund_code, stk_code, gds_tp, book_qty, " _
        & " work_time, work_trm, work_memb, pure_unit_price, pure_evlt_amt) values (" _
        & "'" + tdate + "','" + scenario_id + "','" + deal_ticket.fund_code_c + "','" + deal_ticket.asset_code + "'," _
        & "'" + gds_tp + "'," & deal_ticket.qty & ",SYSDATE, 'EXCEL','EXCEL', " & unitPrice & ", " & unitPrice * deal_ticket.qty & ") "
    
    adoCon.Execute (sql)
    
    sql = "select theory_price from ras.rm_pricing_data where tdate='" + tdate + "' and code='" + deal_ticket.asset_code + "'" '2024.06.26 rm_els_data -> rm_pricing_data, rm_els_info 삭제
    
    With oCmd
        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql
        oRS.Open .Execute
    End With
    
    Do Until oRS.EOF
        unitPrice0 = oRS(0)
        oRS.MoveNext
    Loop
    oRS.Close
    
    sql = "insert into rms.mr_scenario_detail_data (tdate, portfolio_id, scenarioid, code, asset_type, modul_type, fund_cd, qty, assetstartprice, assetendprice, assetscenariopl, currency)" _
        & " values ('" + tdate + "','351','" + scenario_id + "','" + deal_ticket.asset_code + "','" + gds_tp + "','EXCEL','" + deal_ticket.fund_code_c + "'," & deal_ticket.qty & ", " _
        & "" & unitPrice0 & ", " & unitPrice & "," & (unitPrice * fx - unitPrice0 * fx0) * deal_ticket.qty & ",'" + deal_ticket.ccy + "') "
        
    adoCon.Execute (sql)
    
    
End Sub


Public Sub delete_dv01(tdate As String, els_t_code As String, adoCon As adoDB.Connection)

    Dim sql As String
    sql = "delete from rcs.pml_dv01@rms01 where tdate = '" + tdate + "' and code = '" + els_t_code + "' "

    adoCon.Execute (sql)

End Sub

Public Sub insert_dv01(tdate As String, els_t_code As String, ccy As String, t As Double, dv01_tmp As Double, adoCon As adoDB.Connection)
    
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    
    Dim sql As String
    Dim sub_sql As String
    Dim insert_sql As String
    Dim update_sql As String
    
    Dim dv01 As Double
    Dim bcktDV01(14) As Double
    Dim bcktDV01_tmp(14) As Double
        
    Call map_bcktDV01(t, dv01_tmp, bcktDV01_tmp)
    
    'Select
    sql = "select dv01, BCKT_DV01_3M, BCKT_DV01_6M, BCKT_DV01_9M, BCKT_DV01_1Y, BCKT_DV01_18M, BCKT_DV01_2Y, BCKT_DV01_3Y, BCKT_DV01_4Y, BCKT_DV01_5Y, BCKT_DV01_7Y, BCKT_DV01_10Y, BCKT_DV01_12Y, BCKT_DV01_15Y, BCKT_DV01_20Y from rcs.pml_dv01@rms01 where tdate = '" + tdate + "' and code = '" + els_t_code + "' and currency = '" + ccy + "'"
    
    With oCmd

        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql

        oRS.Open .Execute

    End With
    
    '같은 통화 값이 이미 있는 경우, DV01과 bcktDV01에 읽어온 후 dv01_tmp와 bcktDV01_tmp를 누적
    Dim i As Integer
    
    Do Until oRS.EOF
        dv01 = dv01 + oRS(0)
        For i = 1 To 14
            bcktDV01(i - 1) = bcktDV01(i - 1) + oRS(i)
        Next i

        oRS.MoveNext
    Loop
    
    dv01 = dv01 + dv01_tmp
    Call acc_bcktDV01(bcktDV01, bcktDV01_tmp)
    
    'Merge
    update_sql = "UPDATE SET DURATION=" & t & ", DV01=" & dv01 & ", BCKT_DV01_3M=" & bcktDV01(0) & ", BCKT_DV01_6M=" & bcktDV01(1) & ", BCKT_DV01_9M=" & bcktDV01(2) & ", BCKT_DV01_1Y=" & bcktDV01(3) & ", BCKT_DV01_18M=" & bcktDV01(4) & ", BCKT_DV01_2Y=" & bcktDV01(5) & ", BCKT_DV01_3Y=" & bcktDV01(6) & ", BCKT_DV01_4Y=" & bcktDV01(7) & ", BCKT_DV01_5Y=" & bcktDV01(8) & ", BCKT_DV01_7Y=" & bcktDV01(9) & ", BCKT_DV01_10Y=" & bcktDV01(10) & ", BCKT_DV01_12Y=" & bcktDV01(11) & ", BCKT_DV01_15Y=" & bcktDV01(12) & ", BCKT_DV01_20Y=" & bcktDV01(13)
    'update_sql = "UPDATE SET DV01=?, BCKT_DV01_3M=?, BCKT_DV01_6M=?, BCKT_DV01_9M=?, BCKT_DV01_1Y=?, BCKT_DV01_18M=?, BCKT_DV01_2Y=?, BCKT_DV01_3Y=?, BCKT_DV01_4Y=?, BCKT_DV01_5Y=?, BCKT_DV01_7Y=?, BCKT_DV01_10Y=?, BCKT_DV01_12Y=?, BCKT_DV01_15Y=?, BCKT_DV01_20Y=?"
    
    insert_sql = "INSERT (TDATE, CODE, DEPTCODE, GDS_TP, NAME, DURATION, POSITION, DV01, WORK_TIME, WORK_TRM, WORK_MEMB, CURRENCY, FUNDCODE, NOTIONAL, BCKT_DV01_3M, BCKT_DV01_6M, BCKT_DV01_9M, BCKT_DV01_1Y, BCKT_DV01_18M, BCKT_DV01_2Y, BCKT_DV01_3Y, BCKT_DV01_4Y, BCKT_DV01_5Y, BCKT_DV01_7Y, BCKT_DV01_10Y, BCKT_DV01_12Y, BCKT_DV01_15Y, BCKT_DV01_20Y) " _
                  + " VALUES (b.TDATE, b.CODE, b.DEPTCODE, b.GDS_TP, b.NAME, b.DURATION, b.POSITION, b.DV01, b.WORK_TIME, b.WORK_TRM, b.WORK_MEMB, b.CURRENCY, b.FUNDCODE, b.NOTIONAL, b.BCKT_DV01_3M, b.BCKT_DV01_6M, b.BCKT_DV01_9M, b.BCKT_DV01_1Y, b.BCKT_DV01_18M, b.BCKT_DV01_2Y, b.BCKT_DV01_3Y, b.BCKT_DV01_4Y, b.BCKT_DV01_5Y, b.BCKT_DV01_7Y, b.BCKT_DV01_10Y, b.BCKT_DV01_12Y, b.BCKT_DV01_15Y, b.BCKT_DV01_20Y)"
    
'    sub_sql = "SELECT C.STND_DATE TDATE, A.ISCD CODE, ? DEPTCODE, ? GDS_TP, A.KOR_ISNM NAME, ? DURATION, C.RMND_QTY POSITION, ? DV01, SYSDATE WORK_TIME, null WORK_TRM, null WORK_MEMB, " _
'        & "           ? CURRENCY, B.PROD_FNCD FUNDCODE, D.REAL_PBLC_FCAM*C.RMND_QTY NOTIONAL, ? BCKT_DV01_3M,? BCKT_DV01_6M,? BCKT_DV01_9M,? BCKT_DV01_1Y,? BCKT_DV01_18M,? BCKT_DV01_2Y,? BCKT_DV01_3Y,? BCKT_DV01_4Y,? BCKT_DV01_5Y,? BCKT_DV01_7Y,? BCKT_DV01_10Y,? BCKT_DV01_12Y,? BCKT_DV01_15Y,? BCKT_DV01_20Y " _
'        & "     FROM   BSYS.TBSIMM100M00@GDW A," _
'        & "            BSYS.TBSIMO201M00@GDW B," _
'        & "            BSYS.TBFNOM021L00@GDW C," _
'        & "            BSYS.TBSIMO100M00@GDW D " _
'        & "    WHERE  A.ISCD='" + els_t_code + "' " _
'        & "    AND    A.ISCD=B.INDV_ISCD " _
'        & "    AND    A.ISCD=C.ISCD " _
'        & "    AND    C.STND_DATE='" + tdate + "' " _
'        & "    AND    B.OTC_FUND_ISCD=D.OTC_FUND_ISCD "
    sub_sql = "SELECT C.STND_DATE TDATE, A.ISCD CODE, '351' DEPTCODE, DECODE(B.PROD_CLS_CODE,'04','EquitySwap','09','ELN',B.PROD_CLS_CODE) GDS_TP, A.KOR_ISNM NAME, " & t & " DURATION, C.RMND_QTY POSITION, " & dv01 & " DV01, SYSDATE WORK_TIME, 'EXCEL' WORK_TRM, 'EXCEL' WORK_MEMB, " _
        & "           '" + ccy + "' CURRENCY, B.PROD_FNCD FUNDCODE, D.REAL_PBLC_FCAM*C.RMND_QTY NOTIONAL, " & bcktDV01(0) & " BCKT_DV01_3M," & bcktDV01(1) & " BCKT_DV01_6M," & bcktDV01(2) & " BCKT_DV01_9M," & bcktDV01(3) & " BCKT_DV01_1Y," & bcktDV01(4) & " BCKT_DV01_18M," & bcktDV01(5) & " BCKT_DV01_2Y," & bcktDV01(6) & " BCKT_DV01_3Y," & bcktDV01(7) & " BCKT_DV01_4Y," & bcktDV01(8) & " BCKT_DV01_5Y," & bcktDV01(9) & " BCKT_DV01_7Y," & bcktDV01(10) & " BCKT_DV01_10Y," & bcktDV01(11) & " BCKT_DV01_12Y," & bcktDV01(12) & " BCKT_DV01_15Y," & bcktDV01(13) & " BCKT_DV01_20Y " _
        & "     FROM   BSYS.TBSIMM100M00@GDW A," _
        & "            BSYS.TBSIMO201M00@GDW B," _
        & "            BSYS.TBFNOM021L00@GDW C," _
        & "            BSYS.TBSIMO100M00@GDW D " _
        & "    WHERE  A.ISCD='" + els_t_code + "' " _
        & "    AND    A.ISCD=B.INDV_ISCD " _
        & "    AND    A.ISCD=C.ISCD " _
        & "    AND    C.STND_DATE='" + tdate + "' " _
        & "    AND    B.OTC_FUND_ISCD=D.OTC_FUND_ISCD "
        
    sql = "MERGE INTO RCS.PML_DV01 a USING (" + sub_sql + ") b ON (a.tdate = b.tdate and a.code = b.code and a.currency = b.currency) WHEN MATCHED THEN " + update_sql + " WHEN NOT MATCHED THEN " + insert_sql
    


    With oCmd

        .CommandText = sql

'        Call .Parameters.append(oCmd.CreateParameter("TDATE", adoDB.DataTypeEnum.adVarChar, adoDB.ParameterDirectionEnum.adParamInput, 8, tdate))
'        Call .Parameters.append(oCmd.CreateParameter("CODE", adoDB.DataTypeEnum.adVarChar, adoDB.ParameterDirectionEnum.adParamInput, 30, els_t_code))
'        Call .Parameters.append(oCmd.CreateParameter("CURRENCY", adoDB.DataTypeEnum.adVarChar, adoDB.ParameterDirectionEnum.adParamInput, 12, ccy))
'        Call .Parameters.append(oCmd.CreateParameter("DURATION", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=t))
'        Call .Parameters.append(oCmd.CreateParameter("DV01", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=dv01))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_3M", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(0)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_6M", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(1)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_9M", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(2)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_1Y", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(3)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_18M", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(4)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_2Y", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(5)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_3Y", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(6)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_4Y", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(7)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_5Y", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(8)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_7Y", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(9)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_10Y", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(10)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_12Y", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(11)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_15Y", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(12)))
'        Call .Parameters.append(oCmd.CreateParameter("BCKT_DV01_20Y", adoDB.DataTypeEnum.adNumeric, adoDB.ParameterDirectionEnum.adParamInput, value:=bcktDV01(13)))
'        Call .Parameters.append(oCmd.CreateParameter("DEPTCODE", adoDB.DataTypeEnum.adVarChar, adoDB.ParameterDirectionEnum.adParamInput, 12, "351"))
'        Call .Parameters.append(oCmd.CreateParameter("GDS_TP", adoDB.DataTypeEnum.adVarChar, adoDB.ParameterDirectionEnum.adParamInput, 20, "EquitySwap"))

        .ActiveConnection = adoCon
        Call .Execute

    End With
    
    Set oCmd = Nothing
    
End Sub


Public Sub map_bcktDV01(t As Double, dv01 As Double, bcktDV01() As Double)

    Dim t_lower As Double
    Dim t_upper As Double
    Dim w As Double
    Dim i As Integer
    
    If t < 0.25 Then
        i = 0
        w = 1
    ElseIf t < 0.5 Then
        i = 0
        t_lower = 0.25
        t_upper = 0.5
    ElseIf t < 0.75 Then
        i = 1
        t_lower = 0.5
        t_upper = 0.75
    ElseIf t < 1# Then
        i = 2
        t_lower = 0.75
        t_upper = 1#
    ElseIf t < 1.5 Then
        i = 3
        t_lower = 1#
        t_upper = 1.5
    ElseIf t < 2# Then
        i = 4
        t_lower = 1.5
        t_upper = 2#
    ElseIf t < 3# Then
        i = 5
        t_lower = 2#
        t_upper = 3#
    ElseIf t < 4# Then
        i = 6
        t_lower = 3#
        t_upper = 4#
    ElseIf t < 5# Then
        i = 7
        t_lower = 4#
        t_upper = 5#
    ElseIf t < 7# Then
        i = 8
        t_lower = 5#
        t_upper = 7#
    ElseIf t < 10# Then
        i = 9
        t_lower = 7#
        t_upper = 10#
    ElseIf t < 12# Then
        i = 10
        t_lower = 10#
        t_upper = 12#
    ElseIf t < 15# Then
        i = 11
        t_lower = 12#
        t_upper = 15#
    ElseIf t < 20# Then
        i = 12
        t_lower = 15#
        t_upper = 20#
    ElseIf t >= 20# Then
        i = 13
        w = 1
    End If
    
    If t_lower * t_lower <> 0 Then
        w = (t_upper - t) / (t_upper - t_lower)
    End If

    bcktDV01(i) = dv01 * w
    If w <> 1 Then
        bcktDV01(i + 1) = dv01 * (1 - w)
    End If
        
End Sub
                
Public Sub acc_bcktDV01(bcktDV01() As Double, bcktDV01_tmp() As Double)
    
    Dim i As Integer
    For i = 0 To UBound(bcktDV01)
        bcktDV01(i) = bcktDV01(i) + bcktDV01_tmp(i)
    Next i
    
End Sub