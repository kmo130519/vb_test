Option Explicit

Public Sub btnReadMarketData_Click()

    '전역변수 설정
    Call SET_GLOBAL
    
    If POPUP_WARNING_ENABLE = True Then
        If GREEKS_ENABLE = True Then
            If MsgBox("Greeks 계산 옵션이 선택되어 있습니다. 기준가 산출을 하려면 <AC Pricer>의 계산 옵션을 모두 해제해야 합니다. 계속하겠습니까?", vbYesNo) = vbYes Then
                MsgBox ("Discount Rate를 IRS Curve로 설정합니다.")
            Else
                Exit Sub
            End If
        Else
            If MsgBox("기준가를 산출합니다. Greeks를 계산하려면 <AC Pricer>의 계산 옵션을 설정해야 합니다. 계속하겠습니까?", vbYesNo) = vbYes Then
                MsgBox ("Discount Rate는 신영증권 신용등급 Curve로 설정합니다.")
            Else
                Exit Sub
            End If
        End If
    End If
    
    Dim scenario_id As String
    If SCENARIO_ENABLE = False Then
        scenario_id = "0"
    Else
        scenario_id = Range("SCENARIO_ID").value
    End If
            
    Dim adoCon As New adoDB.Connection
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
    Dim market_date As Date
    market_date = shtMarket.Range("market_date").value

    Dim market_prevdate As Date
    market_prevdate = shtMarket.Range("market_prevdate").value

    Dim market_source As String
    market_source = shtMarket.Range("market_source").value
    
    'USD/KRW
    Range("USDKRW").Cells(1, 1) = get_fx("KRW", "USD", market_date, adoCon)
    Range("USDKRW").Cells(2, 1) = get_fx("KRW", "USD", market_prevdate, adoCon)
    
    Dim the_range As Range

    With Application
        .ScreenUpdating = False
        .StatusBar = False
        .DisplayStatusBar = True
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    Dim i_ua As Integer
    Dim ua_code As String

    For i_ua = 1 To NUM_UA
        
        ua_code = get_ua_code(i_ua)
        Set the_range = shtMarket.Range(ua_code)
        Application.StatusBar = "Importing market data... " & ua_code
        Call fill_ua_data(the_range, ua_code, market_date, market_source, scenario_id, adoCon)
        Call fill_ua_data(the_range.Offset(0, 5), ua_code, market_prevdate, market_source, scenario_id, adoCon)
    
    Next i_ua
    
    'Correlation
    Application.StatusBar = "Importing historical correlation data..."
    Set the_range = shtMarket.Range("correlation_table")
    Call fill_corr_matrix(the_range, False, market_date, market_source, scenario_id, adoCon)
    Call fill_corr_matrix(the_range.Offset(the_range.Rows.count + 3, 0), False, market_prevdate, market_source, scenario_id, adoCon)

    'Minimum Correlation 2019.3.27
    Application.StatusBar = "Importing minimum correlation data..."
    Set the_range = shtMarket.Range("min_correlation_table")
    Call fill_corr_matrix(the_range, True, market_date, market_source, scenario_id, adoCon)
    Call fill_corr_matrix(the_range.Offset(the_range.Rows.count + 3, 0), True, market_prevdate, market_source, scenario_id, adoCon)

    'Discount Curve
    '금리시나리오는 DV01 계산 중 반영되므로 시나리오를 적용하지 않음
    Dim i_dcf As Integer
    Dim dcf_ccy As String

    For i_dcf = 1 To NUM_DCF

        dcf_ccy = get_dcf_ccy(i_dcf)
        Application.StatusBar = "Importing discount curves..." & dcf_ccy
        Set the_range = shtMarket.Range("DCF_" & dcf_ccy)
        Call fill_dcf_data(the_range, dcf_ccy, market_date, market_source, scenario_id, adoCon)
        Call fill_dcf_data(the_range.Offset(0, 5), dcf_ccy, market_prevdate, market_source, scenario_id, adoCon)

    Next i_dcf

    With Application
        .ScreenUpdating = True
        .StatusBar = False
        .DisplayStatusBar = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    
    Call disconnectDB(adoCon)
    
    Set adoCon = Nothing
    Set the_range = Nothing
    
End Sub

Private Sub fill_dcf_data(the_range As Range, dcf_ccy As String, tdate As Date, source As String, scenario_id As String, adoDB As adoDB.Connection)

    Dim tdate_str As String
    tdate_str = date2str(tdate)
    
    Dim rate_id As String
    Dim dcf_date As Date
    Dim dcf_curve_name As String
    Dim dcf_curve As Variant
    
    Range(the_range.Cells(8, 1), the_range.Cells(22, 4)).ClearContents
    dcf_curve = Range(the_range.Cells(8, 1), the_range.Cells(22, 4)).value
    
    'Discount Rate Curve
    '/<----- 델타 계산시 swap curve 사용 2019.4.5
    If GREEKS_ENABLE = True Then
            
        'Select Case get_ccy_idx(dcf_ccy)
        'Case CCY.USD
        
            '/--- tmp --- 2021.1.4 USD swap curve 불안정 -> 신용등급 BBB+ 커브사용 -> 제거 2024.01.30
            'rate_id = DCF_USD_RATEID
            
            'Call get_dcf_curve(rate_id, dcf_ccy, tdate_str, source, scenario_id, dcf_date, dcf_curve_name, dcf_curve, adoDB)
            
        'Case CCY.KRW '---/
        
            rate_id = "IRS"
            dcf_curve_name = "IRS"
            dcf_date = tdate
            
            Call get_rf_curve(dcf_ccy, get_rf_num_term(dcf_ccy), tdate_str, source, scenario_id, dcf_curve, adoDB)
        'End Select

    Else '----->/
    
        Select Case get_ccy_idx(dcf_ccy)
        Case ccy.USD: rate_id = DCF_USD_RATEID
        Case ccy.KRW: rate_id = DCF_KRW_RATEID
        End Select
    
        Call get_dcf_curve(rate_id, dcf_ccy, tdate_str, source, scenario_id, dcf_date, dcf_curve_name, dcf_curve, adoDB)

    End If
    
    the_range.Cells(1, 3) = tdate
    the_range.Cells(3, 3) = rate_id
    the_range.Cells(4, 3) = dcf_curve_name
    the_range.Cells(5, 3) = dcf_date
    Range(the_range.Cells(8, 1), the_range.Cells(22, 4)).value = dcf_curve
        
End Sub

Private Sub fill_ua_data(the_range As Range, ua_code As String, tdate As Date, source As String, scenario_id As String, adoCon As adoDB.Connection)

    the_range.Cells(1, 3).ClearContents
    the_range.Cells(3, 3).ClearContents
    the_range.Cells(4, 3).ClearContents 'Spot Price
    the_range.Cells(5, 3).ClearContents
    Range(the_range.Cells(8, 1), the_range.Cells(22, 4)).ClearContents 'Riskfree Rate Curve
    the_range.Cells(24, 3).ClearContents 'Lambda Neutral 2019.3.27
    the_range.Cells(27, 3).ClearContents '3m ATM Implied Volatility
    the_range.Cells(32, 3).ClearContents '평균 Dividend Yield
    the_range.Cells(33, 3).ClearContents '1년 Dividend Yield
    the_range.Cells(34, 3).ClearContents '1~2년 Dividend Yield
    the_range.Cells(35, 3).ClearContents '2~3년 Dividend Yield
    the_range.Cells(36, 3).ClearContents 'Ex-dividend Dates 수
    Range(the_range.Cells(38, 2), the_range.Cells(37 + NUM_DIV_DATES, 3)).ClearContents 'Dividend Schedule
    
    Dim i_dcf As Integer
    For i_dcf = 1 To NUM_DCF
        the_range.Cells(24 + i_dcf, 3).ClearContents ' FX Vol,
        the_range.Cells(27 + i_dcf, 3).ClearContents 'Quanto Correlation
        the_range.Cells(29 + i_dcf, 3).ClearContents 'Quanto Dividend
    Next i_dcf
        
    If is_active_ua(ua_code) = True Then
    
        Dim tdate_str As String
        tdate_str = date2str(tdate)
        
        Dim ua_ccy As String
        ua_ccy = get_ua_currency(ua_code)
        
        the_range.Cells(1, 3) = tdate
        the_range.Cells(3, 3) = ua_code
        the_range.Cells(5, 3) = ua_ccy
        
        'Spot Price
        the_range.Cells(4, 3) = get_spot_price(ua_code, tdate_str, scenario_id, adoCon)
    
        'Riskfree Rate Curve
        Dim rf_curve As Variant
        rf_curve = Range(the_range.Cells(8, 1), the_range.Cells(22, 4)).value
        Call get_rf_curve(ua_ccy, get_rf_num_term(ua_ccy), tdate_str, source, scenario_id, rf_curve, adoCon)
        Range(the_range.Cells(8, 1), the_range.Cells(22, 4)).value = rf_curve
        
        'Lambda Neutral 2019.3.27
        the_range.Cells(24, 3) = get_lambda_neutral(ua_code, tdate_str, source, scenario_id, adoCon)
        
        '3m ATM Implied Volatility
        the_range.Cells(27, 3) = get_vol_on_surface(tdate_str, ua_code, the_range.Cells(4, 3).value, 0.25, 1, "Implied", "FRONT", adoCon)
        
        'Dividend Yield
        the_range.Cells(32, 3) = get_div_yield(ua_code, tdate_str, source, scenario_id, adoCon)
        
        'Dividend Schedule
        Dim div_schedule As Variant
        div_schedule = Range(the_range.Cells(38, 2), the_range.Cells(37 + NUM_DIV_DATES, 3)).value
        Call get_div_schedule(ua_code, tdate_str, source, scenario_id, div_schedule, adoCon)
        Range(the_range.Cells(38, 2), the_range.Cells(37 + NUM_DIV_DATES, 3)).value = div_schedule
        
        '1년 단위 ex-dividend dates 범위
        Dim num_ex_div_dates_1y As Integer
        Dim num_ex_div_dates_2y As Integer
        Dim num_ex_div_dates_3y As Integer
        num_ex_div_dates_1y = WorksheetFunction.CountIf(Range(the_range.Cells(38, 2), the_range.Cells(37 + NUM_DIV_DATES, 2)), "<=" & WorksheetFunction.EDate(tdate, 12))
        num_ex_div_dates_2y = WorksheetFunction.CountIf(Range(the_range.Cells(38, 2), the_range.Cells(37 + NUM_DIV_DATES, 2)), "<=" & WorksheetFunction.EDate(tdate, 24))
        num_ex_div_dates_3y = WorksheetFunction.CountIf(Range(the_range.Cells(38, 2), the_range.Cells(37 + NUM_DIV_DATES, 2)), "<=" & WorksheetFunction.EDate(tdate, 36))
        
        '1년 단위 dividend yield
        the_range.Cells(33, 3).Formula = "=SUM(" & the_range.Cells(38, 3).Address() & ":" & the_range.Cells(37 + num_ex_div_dates_1y, 3).Address() & ")/" & the_range.Cells(4, 3).Address()
        the_range.Cells(34, 3).Formula = "=SUM(" & the_range.Cells(37 + num_ex_div_dates_1y + 1, 3).Address() & ":" & the_range.Cells(37 + num_ex_div_dates_2y, 3).Address() & ")/" & the_range.Cells(4, 3).Address()
        the_range.Cells(35, 3).Formula = "=SUM(" & the_range.Cells(37 + num_ex_div_dates_2y + 1, 3).Address() & ":" & the_range.Cells(37 + num_ex_div_dates_3y, 3).Address() & ")/" & the_range.Cells(4, 3).Address()
        the_range.Cells(36, 3).Formula = "=COUNT(" & the_range.Cells(38, 2).Address() & ":" & the_range.Cells(37 + NUM_DIV_DATES, 2).Address() & ")"
        
        'Quanto
        Dim dcf_ccy As String
        Dim quanto_pair_code(1) As String
        quanto_pair_code(0) = ua_code
        
        For i_dcf = 1 To NUM_DCF
        
            dcf_ccy = get_dcf_ccy(i_dcf)
            quanto_pair_code(1) = ua_ccy & dcf_ccy
            
            'FX Vol, Quanto Correlation
            If get_ccy_idx(ua_ccy) = get_ccy_idx(dcf_ccy) Then
                the_range.Cells(24 + i_dcf, 3) = 0
                the_range.Cells(27 + i_dcf, 3) = 0
            Else
                the_range.Cells(24 + i_dcf, 3) = get_fx_vol(ua_ccy, dcf_ccy, tdate_str, source, scenario_id, adoCon)
                the_range.Cells(27 + i_dcf, 3) = get_corr(quanto_pair_code, False, tdate_str, source, scenario_id, adoCon)
            End If
            
            'Quanto Dividend
            the_range.Cells(29 + i_dcf, 3).Formula = "=" & the_range.Cells(27, 3).Address() & "*" & the_range.Cells(27 + i_dcf, 3).Address() & "*" & the_range.Cells(24 + i_dcf, 3).Address()
            
        Next i_dcf
        
        With Range(the_range.Cells(1, 1), the_range.Cells(1, 5)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
    Else
    
        the_range.Cells(1, 3) = "미사용"
     
        With Range(the_range.Cells(1, 1), the_range.Cells(1, 5)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.499984740745262
            .PatternTintAndShade = 0
        End With
    End If
    
End Sub

Private Sub fill_corr_matrix(corr_matrix As Range, isLocMinCorr As Boolean, tdate As Date, source As String, scenario_id As String, oDB As adoDB.Connection)
'local correlation 추가 2019.3.27

    Dim tdate_str As String
    tdate_str = date2str(tdate)
            
    corr_matrix.ClearContents
    corr_matrix.Cells(0, 0).value = tdate
    
    Dim no_of_asset As Integer
    no_of_asset = corr_matrix.Columns.count
    
    Dim corr As Variant
    corr = corr_matrix.value
    
    Dim corr_pool As Variant
    corr_pool = Range(corr_matrix.Cells(1, 0), corr_matrix.Cells(no_of_asset, 0)).value
    
    Dim pair_code(1) As String
    
    Dim i, j As Integer
    Dim target_point As Boolean
    target_point = False
    
    For i = 1 To no_of_asset
    
        pair_code(0) = corr_pool(i, 1)
       
        For j = 1 To no_of_asset
        
            pair_code(1) = corr_pool(j, 1)
            
            If is_active_ua(pair_code(0)) = True Then
                If is_active_ua(pair_code(1)) = True Then
                    target_point = True
                Else
                    If get_ccy_idx(get_ua_currency(pair_code(0))) = get_ccy_idx(Left(pair_code(1), 3)) And is_active_ccy(Left(pair_code(1), 3)) = True Then
                        target_point = True
                    Else
                        target_point = False
                    End If
                End If
            Else
                If is_active_ua(pair_code(1)) = True Then
                    If get_ccy_idx(get_ua_currency(pair_code(1))) = get_ccy_idx(Left(pair_code(0), 3)) And is_active_ccy(Left(pair_code(0), 3)) = True Then
                        target_point = True
                    Else
                        target_point = False
                    End If
                Else
                    target_point = False
                End If
            End If
            
            If target_point = True Then
            
                If i = j Then
                    corr(i, j) = 1
                ElseIf i > j Then
                    corr(i, j) = corr(j, i)
                Else
                    corr(i, j) = get_corr(pair_code, isLocMinCorr, tdate_str, source, scenario_id, oDB)
                End If

            End If
        
        Next j
    
    Next i
    
    corr_matrix.value = corr
    
    Set corr = Nothing
    Set corr_pool = Nothing
    
End Sub