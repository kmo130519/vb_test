Option Explicit
Option Base 1
' test

'asdkjlfhlkasdfhlasjkd
Public Enum ELS_COL
    SQ = 1
    HEDG_TR_TYPE_CODE = 2
    PROD_CLS_CODE = 3
    ISIN = 4
    FNCD_ELS = 5
    FO_ELS_D = 6
    IO_ELS_D = 7
    FO_ELS_H = 8
    PROD_CLS_CODE_H = 9
    IO_ESWAP_H = 10
    IO_FRN_H = 11
    FNCD_ESWAP = 12
    FO_ESWAP_T = 13
    IO_ESWAP_T = 14
    FO_FRN_T = 15
    IO_FRN_T = 16
    KOR_ISNM = 17
    STLM_CRCD = 18
    REAL_PBLC_FCAM = 19
    DEAL_CLS_CODE = 20
    PBLC_STCK_QTY = 21
    RMND_QTY = 22
    NOTIONAL_CCY = 23
    NOTIONAL_KRW = 24
    FRST_STND_PRC_FIN_DATE = 25
    PBLC_DATE = 26
    MTRT_DATE = 27
    RMND_DAYS = 28
    FUND_PBLC_UNPR = 29
    FUND_BOOK_UNPR = 30
    HEDG_BUF_D = 31
    HEDG_BUF_H = 32
    clrd_type_code = 33
    unas_choc_mthd_code = 34
    NUM_UNAS = 35
    UNAS_ISCD1 = 40
    UNAS_INTL_PRC1 = 44
    AC_UNAS_SDRT1 = 48
    AC_CLRD_DTRM_DATE1 = 60
    AC_CLRD_ERT1 = 72
    KI_BARR_VAL = 84
    KI_HIT_DATE = 85
    M_BONS_CUPN_STA_SDRT = 86
    M_BONS_CUPN_INRT = 87
    EE_CLRD_BARR_VAL1 = 88
    EE_CALC_FIN_DATE1 = 91
    EE_CLRD_INRT1 = 94
    EE_CLRD_BARR_HIT_YN1 = 97
    PRC_APLY_CLS_CODE = 102
End Enum

Private Enum ELS_PRICING_COL
    SQ = 1
    CALC_TIME = 2
    TODAY_EVENT = 3
    NO = 4
    HEDG_TR_TYPE_CODE = 5
    
    PROD_CLS_CODE = 6
    FO_ELS_D = 7
    PROD_CLS_CODE_H = 8
    IO_ESWAP_H = 9
    IO_ESWAP_T = 10
    KOR_ISNM = 11
    STLM_CRCD = 12
    NOTIONAL_CCY = 13
    DAY1_PL_D = 14
    DAY1_PL_H = 15
    ua_pct_price = 16
    UA_WORST = 22
    UA_WORST_PCT_PRICE = 23
    REDEEMABLE_NO = 24
    REDEEMABLE_DATE = 25
    REDEEMABLE_DAYS_LEFT = 26
    REDEEMABLE_BARRIER = 27
    REDEEMABLE_RATIO = 28
    REDEEMABLE_MARGIN = 29
    IS_MATURITY = 30
    next_ac_no = 31
    NEXT_AC_DATE = 32
    NEXT_AC_DAYS_LEFT = 33
    NEXT_AC_BARRIER = 34
    NEXT_AC_REDEEM_RATIO = 35
    NEXT_AC_MARGIN = 36
    NEXT_AC_AVAILABLE = 37
'    ELS_PRICE_EST = 34
'    ELS_PRICE_EST_ERR = 35
'    els_price_theo = 36
'    ESWAP_PRICE_THEO = 37
    PRICE_THEO_D = 38
    PRICE_THEO_H = 39
'    FRN_PRICE_THEO = 38
'    els_price_mtm = 39
'    ESWAP_PRICE_MTM = 40
    PRICE_MTM_D = 40
    PRICE_MTM_H = 41
    ELS_PRICE_THEO_FR = 42
    ESWAP_PRICE_THEO_FR = 43
    ELS_PRICE_THEO_DIFF = 44
    ELS_PRICE_THEO_DIFF_CHG = 45
    PL_REAL = 46
    PL_EST = 47
    PL_EST_ERR = 48
    PL_FR = 49
    PL_FR_DIFF = 50
    PL_DELTA = 51
    PL_GAMMA = 52
    PL_XGAMMA = 53
    PL_VEGA = 55
    PL_THETA = 57
    PL_RF = 58
    PL_CVA = 59
    PL_FX = 60
    PL_DAY1 = 61
End Enum

Private Enum TODAY_EVENT
    EVAL_AC = 1 '상환평가일: 상환확정
    EVAL_TBD = 2 '상환평가일: 상환가능
    EVAL_FAIL = 3 '상환평가일: 상환실패
    DEAL_TRADE = 4 '신규 거래
    DEAL_ISSUE = 5 '신규 발행
    ESWAP_RCV_FLT = 6 '스왑 변동금리 수취
End Enum

Public Sub btnCalcPrice_Click()

    '전역변수 설정
    Call SET_GLOBAL
    
    Dim rng As Range
    Set rng = Range("TBL_ELS_PRICING_LIST")
    
    Dim rng_els_list As Range
    Set rng_els_list = shtELSList.Range("TBL_ELS_LIST")
    
    If Range("clear_price").value = True Then
        rng.Columns(CALC_TIME).ClearContents
        rng.Columns(TODAY_EVENT).ClearContents
        rng.Columns(PRICE_THEO_D).ClearContents
        rng.Columns(PRICE_THEO_H).ClearContents
        rng.Columns(PRICE_MTM_D).ClearContents
        rng.Columns(PRICE_MTM_H).ClearContents
    End If
    
    Dim bScreenUpdating As Boolean
    Dim btatusBar As Boolean
    Dim bDisplayStatusBar As Boolean
    Dim vCalculation As Variant
    Dim bEnableEvents As Boolean
    bScreenUpdating = Application.ScreenUpdating
    btatusBar = Application.StatusBar
    bDisplayStatusBar = Application.DisplayStatusBar
    vCalculation = Application.Calculation
    bEnableEvents = Application.EnableEvents
    
    With Application
        .ScreenUpdating = True
        .StatusBar = False
        .DisplayStatusBar = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With

    Dim target_date As Date
    target_date = Range("target_date").value

    Dim latest_call_date As Date
    Dim latest_bus_date As Date
    Dim spot_date As Date
    'latest_bus_date = get_latest_bus_date(target_date)
    latest_bus_date = shtMarket.Range("market_prevdate").value 'temp
    
    Dim str_target_date As String
    str_target_date = date2str(target_date)
    
    Dim adoCon As New adoDB.Connection
    
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim sql As String

    Dim start_time As Date
    Dim elapsed_time As Date
    
    Dim prod_code As String
    
    Dim for_frn_flag As Boolean
    Dim for_eswap_flag As Boolean

    Dim ac_deal_ticket As clsACDealTicket
    Dim eswap_deal_ticket As clsACDealTicket

    Dim ac_greeks As clsGreeks
    Dim eswap_greeks As clsGreeks

    Dim riskfree_dcf_enable As Boolean
    Dim prev_riskfree_dcf_enable As Boolean
    
    Dim market_set As clsMarketSet
    Dim ua_pct_price() As Double
    Dim is_delayed_spot As Boolean
    
    Dim rtn(4) As Variant
    
    Dim calc_delta_gamma As Boolean
    Dim calc_sticky_moneyness_delta As Boolean
    Dim calc_sticky_strike_delta As Boolean
    Dim calc_cross_gamma As Boolean
    Dim calc_vega As Boolean
    Dim calc_term_vega As Boolean
    Dim calc_skew As Boolean
    Dim calc_corr As Boolean
    Dim calc_rho As Boolean
    Dim calc_theta As Boolean
    'Dim calc_local_corr As Boolean
    
    If GREEKS_ENABLE = True Then
        calc_delta_gamma = True
        calc_sticky_moneyness_delta = True
        calc_sticky_strike_delta = False
        calc_cross_gamma = False
        calc_vega = True
        calc_term_vega = False
        calc_skew = False
        calc_corr = False
        calc_rho = True
        calc_theta = True
        'calc_local_corr = True
    Else
        calc_delta_gamma = False
        calc_sticky_moneyness_delta = False
        calc_sticky_strike_delta = False
        calc_cross_gamma = False
        calc_vega = False
        calc_term_vega = False
        calc_skew = False
        calc_corr = False
        calc_rho = False
        calc_theta = False
        'calc_local_corr = True
    End If

    Dim seq_from_els As Integer
    Dim seq_to_els As Integer
    Dim seq_cur_els As Integer
    seq_from_els = Range("from_seq").value
    seq_to_els = Range("to_seq").value
        
    Dim note_price As Double
    Dim swap_price As Double
    Dim price As Double
    Dim day1_pl_price As Double
    
    Dim i_dh As Integer
    Dim prod_cls As String
    
    'i_dh = 1 ... 대상 pricing loop
    'i_dh = 2 ... 수단 pricing loop
    Dim i_dh_from As Integer
    Dim i_dh_to As Integer
    Select Case Range("pricing_target").value
    Case "전체"
        i_dh_from = 1
        i_dh_to = 2
    Case "대상"
        i_dh_from = 1
        i_dh_to = 1
    Case "수단"
        i_dh_from = 2
        i_dh_to = 2
    End Select
    
    For i_dh = i_dh_from To i_dh_to
    
'2024.07.29 미래금융팀 테스트용으로 Range("insert_greeks") 삭제 및 주석처리
'        If Range("insert_greeks").value = True And GREEKS_ENABLE = True Then
'            i_dh = 2
'        End If
        
        For i = seq_from_els To seq_to_els
    
            start_time = Now
            Range("cur_seq") = i
            rng.Cells(i, CALC_TIME) = "Calcuating..."
    
            If i_dh = 1 Then
                prod_cls = rng(i, ELS_PRICING_COL.PROD_CLS_CODE).value
            Else
                prod_cls = rng(i, ELS_PRICING_COL.PROD_CLS_CODE_H).value
            End If
            
            If rng(i, ELS_PRICING_COL.IO_ESWAP_T).value <> "" Then
                prod_code = rng(i, ELS_PRICING_COL.IO_ESWAP_T).value '자체 E.swap IO~T
            Else
                prod_code = rng(i, ELS_PRICING_COL.IO_ESWAP_H).value '백투백 E.swap IO~H
            End If
            
            'market_set
            If prod_cls = "스왑" Then
                riskfree_dcf_enable = True '무위험 커브 할인
                for_frn_flag = False
                for_eswap_flag = True
            Else
                riskfree_dcf_enable = False '신용등급 커브 할인
                for_frn_flag = False
                for_eswap_flag = False
            End If
            
            If i = seq_from_els Or prev_riskfree_dcf_enable <> riskfree_dcf_enable Then
                Set market_set = read_market_set(riskfree_dcf_enable)
                prev_riskfree_dcf_enable = riskfree_dcf_enable
            End If
            
            '시나리오테스트 중 미완료 종목만 계산시 이미 완료된 종목이면 skip '2024.07.18
            If Not (SCENARIO_ENABLE = True And Range("missed_prod_only").value = True And is_completed(prod_code, target_date, shtMarket.Range("SCENARIO_ID").value, adoCon) = True) Then
            
                'evaluate autocall
                Set ac_deal_ticket = get_ac_deal_ticket(prod_code, target_date, False, False, adoCon) '회계처리일이 아닌 계약 상 실제 조기상환결정일 기준: eval_date_lag = false
                
                'target_date 포함 이전 최직근 call_date
                latest_call_date = CDate(0) 'ac_deal_ticket.value_date : 2024.07.22 dummy 초기값 삭제. date형 기본값으로 적용.
                For k = 1 To ac_deal_ticket.no_of_schedule
                    If target_date >= ac_deal_ticket.autocall_schedules(k).call_date Then
                        latest_call_date = ac_deal_ticket.autocall_schedules(k).call_date
                    End If
                Next k
                
                spot_date = target_date
                is_delayed_spot = True
                    
                If ac_deal_ticket.has_eval_shift_ul = True Then
                    
                        '1 거래일: latest_bus_date < value_date = latest_call_date = target_date                  -> spot_date = target_date, is_delayed_spot = true
                        '2 발행일: latest_bus_date <= value_date = latest_call_date <= target_date                -> spot_date = target_date, is_delayed_spot = true
                        '1/2: latest_bus_date <= latest_call_date <= target_date & value_date = latest_call_date
                        
                        '3 평상시: latest_call_date < latest_bus_date < target_date                               -> spot_date = target_date, is_delayed_spot = true
                        '4 조기상환평가일 도래했으나 일부 종가 미형성: latest_bus_date < latest_call_date = target_date       -> spot_date = target_date, is_delayed_spot = true
                                            
                        '5 조기상환평가 회계적용일: latest_bus_date = latest_call_date < target_date                  -> spot_date = latest_call_date, is_delayed_spot = false
                        '6 한국 연휴기간 중 조기상환평가일 도래: latest_bus_date < latest_call_date < target_date         -> spot_date = latest_call_date, is_delayed_spot = false
                        '5/6: latest_bus_date <= latest_call_date < target_date & value_date < latest_call_date
                        
                    If latest_bus_date <= latest_call_date And latest_call_date < target_date And ac_deal_ticket.value_date < latest_call_date Then
                        spot_date = latest_call_date
                        is_delayed_spot = False
                    End If
                    
                End If
                
                '기초자산 현재가
                ReDim ua_pct_price(ac_deal_ticket.no_of_ul) As Double
                For k = 1 To ac_deal_ticket.no_of_ul
                    ua_pct_price(k) = get_spot(ac_deal_ticket.ul_code(k), spot_date, is_delayed_spot, adoCon) / ac_deal_ticket.reference_price(k)
                Next k
                
                '상환 확정시 : 2024.05.21
                If eval_next_autocall(ac_deal_ticket, spot_date, ua_pct_price, rtn) = True And spot_date = latest_call_date And (ac_deal_ticket.has_eval_shift_ul = False Or is_delayed_spot = False) Then
                
                    day1_pl_price = 0
                    
                    If prod_cls = "스왑" Then
                        price = (CDbl(rtn(3)) - 1) * ac_deal_ticket.unit_notional
                    Else
                        price = CDbl(rtn(3)) * ac_deal_ticket.unit_notional
                    End If
                    
                    rng.Cells(i, ELS_PRICING_COL.TODAY_EVENT) = "상환 확정"
          
                ElseIf ac_deal_ticket.notional <> 0 Then
                
                    Set ac_deal_ticket = get_ac_deal_ticket(prod_code, target_date, False, for_eswap_flag, adoCon, , True, , market_set) 'eval_date_lag = true (1 calendar day)
                    
                    'modify ac_deal_ticket for leverage indices : 2018.7.9
                    Set ac_deal_ticket = modify_ac_deal_ticket(ac_deal_ticket, market_set)
                
                    Set ac_greeks = New clsGreeks
                
                    run_ac_pricing_mc ac_greeks, ac_deal_ticket, market_set _
                                        , calc_delta_gamma _
                                        , calc_sticky_moneyness_delta _
                                        , calc_sticky_strike_delta _
                                        , calc_cross_gamma _
                                        , calc_vega _
                                        , calc_term_vega _
                                        , calc_skew _
                                        , calc_corr _
                                        , calc_rho _
                                        , calc_theta _
                                        , CORR_SKEW_ENABLE
                                       ' , calc_local_corr
                
                    price = ac_greeks.value / ac_deal_ticket.notional * ac_deal_ticket.unit_notional
    
                    If i_dh = 1 Then
                        day1_pl_price = rng.Cells(i, ELS_PRICING_COL.DAY1_PL_D) / Abs(ac_deal_ticket.qty)
                        
    '                    rng.Cells(i, ELS_PRICING_COL.PRICE_THEO_D) = Round(price / ac_deal_ticket.unit_notional, 7) * ac_deal_ticket.unit_notional
    '                    rng.Cells(i, ELS_PRICING_COL.PRICE_MTM_D) = Round((price + day1_pl_price) / ac_deal_ticket.unit_notional, 7) * ac_deal_ticket.unit_notional
                    Else
                        day1_pl_price = rng.Cells(i, ELS_PRICING_COL.DAY1_PL_H) / Abs(ac_deal_ticket.qty)
                        
                        '수단 Day1 PL 관리가 안 된 과거 기간은 대상과 동일한 값 사용
                        If date2str(ac_deal_ticket.settlement_date) < "20240403" Then
                            day1_pl_price = rng.Cells(i, ELS_PRICING_COL.DAY1_PL_D) / Abs(ac_deal_ticket.qty)
                        End If
                        
    '                    rng.Cells(i, ELS_PRICING_COL.PRICE_THEO_H) = Round(price / ac_deal_ticket.unit_notional, 7) * ac_deal_ticket.unit_notional
    '                    rng.Cells(i, ELS_PRICING_COL.PRICE_MTM_H) = Round((price + day1_pl_price) / ac_deal_ticket.unit_notional, 7) * ac_deal_ticket.unit_notional
                        
'2024.07.29 미래금융팀 테스트용으로 Range("insert_greeks") 삭제 및 주석처리
'                        If Range("insert_greeks").value = True And GREEKS_ENABLE = True Then
'                            If SCENARIO_ENABLE = True Then
'                                Call insert_scenario_result(ac_greeks, ac_deal_ticket, market_set, adoCon, shtMarket.Range("SCENARIO_ID").value)
'                            Else
'                                Call insert_greeks(ac_greeks, ac_deal_ticket, market_set, adoCon)
'                            End If
'                        End If
                        
                    End If
                    
                End If
                
                If i_dh = 1 Then
                    rng.Cells(i, ELS_PRICING_COL.PRICE_THEO_D) = Round(price / ac_deal_ticket.unit_notional, 8) * ac_deal_ticket.unit_notional
                    
                    '발행일 기준가 = 장부가
                    If ac_deal_ticket.settlement_date = target_date Then
                        If prod_cls = "스왑" Then
                            rng.Cells(i, ELS_PRICING_COL.PRICE_MTM_D) = rng_els_list.Cells(rng.Cells(i, ELS_PRICING_COL.NO), ELS_COL.FUND_PBLC_UNPR) - ac_deal_ticket.unit_notional 'temp
                        Else
                            rng.Cells(i, ELS_PRICING_COL.PRICE_MTM_D) = rng_els_list.Cells(rng.Cells(i, ELS_PRICING_COL.NO), ELS_COL.FUND_BOOK_UNPR)
                        End If
                    Else
                        rng.Cells(i, ELS_PRICING_COL.PRICE_MTM_D) = Round((price + day1_pl_price) / ac_deal_ticket.unit_notional, 8) * ac_deal_ticket.unit_notional
                    End If
                Else
                    rng.Cells(i, ELS_PRICING_COL.PRICE_THEO_H) = Round(price / ac_deal_ticket.unit_notional, 8) * ac_deal_ticket.unit_notional
                    
                    '발행일 기준가 = 장부가
                    If ac_deal_ticket.settlement_date = target_date Then 'temp : ac_deal_ticket.issue_cost = #35304 장부가(PERC_APLY_THPR)
                        If prod_cls = "스왑" Then
                            rng.Cells(i, ELS_PRICING_COL.PRICE_MTM_H) = -1 * ac_deal_ticket.issue_cost * ac_deal_ticket.unit_notional '스왑인 경우, 장부가에 (-) 빠져있음
                        Else
                            rng.Cells(i, ELS_PRICING_COL.PRICE_MTM_H) = ac_deal_ticket.issue_cost * ac_deal_ticket.unit_notional
                        End If
                    Else
                        rng.Cells(i, ELS_PRICING_COL.PRICE_MTM_H) = Round((price + day1_pl_price) / ac_deal_ticket.unit_notional, 8) * ac_deal_ticket.unit_notional
                    End If
                End If
                        
                'Greeks로 가격 추정
        
                'Numerical 손익 분해: 델타, 감마, 베가, 세타, 기타(크로스감마, 스큐, 금리, 상관계수)
        
                'Impled Vol 변동 추정치(%p) ~ 실제 베가 PL / 전일 베가
        
                Set ac_deal_ticket = Nothing
                Set eswap_deal_ticket = Nothing
    
                Set ac_greeks = Nothing
                Set eswap_greeks = Nothing
                        
                elapsed_time = elapsed_time + (Now - start_time)
        
                rng.Cells(i, CALC_TIME) = Now - start_time
                
                Range("elapsed_time") = elapsed_time
        
                'AutoSave: every [auto_save_freq] products
                If Range("auto_save").value = True And (i Mod Range("auto_save_freq").value) = 0 Then
                    ThisWorkbook.Save
                End If
            
            End If
    
        Next i
    
    Next i_dh

    'Reset SEQ
    Range("from_seq") = 1
    
    'AutoSave
    If Range("auto_save").value = True Then
        ThisWorkbook.Save
        ThisWorkbook.Close
    End If
    
    Call disconnectDB(adoCon)
    
    Set market_set = Nothing
    
    Set ac_deal_ticket = Nothing
    Set eswap_deal_ticket = Nothing
    
    Set ac_greeks = Nothing
    Set eswap_greeks = Nothing

    With Application
        .ScreenUpdating = bScreenUpdating
        .StatusBar = btatusBar
        .DisplayStatusBar = bDisplayStatusBar
        .Calculation = vCalculation
        .EnableEvents = bEnableEvents
    End With

End Sub

Private Sub btnComparePrice_Click()
'Equity파생운용부 이론가 및 기준가 조회

    Dim target_date As Date
    target_date = Range("target_date").value
    
    Dim rng As Range
    Set rng = Range("TBL_ELS_PRICING_LIST")

    Dim adoCon As New adoDB.Connection
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(2) As String
    ReDim bind_value(2) As Variant
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":code"
    bind_value(1) = date2str(target_date)
    
    Dim i As Integer
    Dim sql As String
        
    For i = 1 To rng.Rows.count
    
        If IsEmpty(rng(i, ELS_COL.SQ)) = False Then
            
            bind_value(2) = rng(i, ELS_PRICING_COL.IO_ESWAP_T).value
            
            sql = getSQL(SQL_PATH_ELS_PRICE_FR, bind_variable, bind_value)
    
            With oCmd
            
                .ActiveConnection = adoCon
                .CommandType = adCmdText
                .CommandText = sql
                
                oRS.Open .Execute
            
            End With
            
            Do Until oRS.EOF
                rng(i, ELS_PRICING_COL.ELS_PRICE_THEO_FR) = oRS(1)
                oRS.MoveNext
            Loop
            
            oRS.Close
            
        End If
        
    Next i
    
    Call disconnectDB(adoCon)
    
End Sub

Private Sub btnDecomposePL_Click()
    
    '실제 회계 손익
    
    '손익 추정 오차
    
    
    
    Dim rng As Range
    Set rng = Range("TBL_ELS_PRICING_LIST")
    rng.Columns(CALC_TIME).ClearContents
    rng.Columns(els_price_theo).ClearContents
    
    Dim bScreenUpdating As Boolean
    Dim btatusBar As Boolean
    Dim bDisplayStatusBar As Boolean
    Dim vCalculation As Variant
    Dim bEnableEvents As Boolean
    bScreenUpdating = Application.ScreenUpdating
    btatusBar = Application.StatusBar
    bDisplayStatusBar = Application.DisplayStatusBar
    vCalculation = Application.Calculation
    bEnableEvents = Application.EnableEvents
    
    With Application
        .ScreenUpdating = True
        .StatusBar = False
        .DisplayStatusBar = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With

    Dim target_date As Date
    target_date = Range("target_date").value
    
    Dim str_target_date As String
    str_target_date = date2str(target_date)
    
    Dim adoCon As New adoDB.Connection
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(2) As String
    ReDim bind_value(2) As Variant
    bind_variable(1) = ":tdate"
    bind_value(1) = str_target_date
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim sql As String
    
'    Dim ofs As Integer
'    Dim unit_notional As Double
'    Dim next_cpn As Double
'    Dim hedge_buff As Double
'    Dim book_price As Double
'    Dim qty As Double
'    Dim ds As Double
'    Dim delta As Double
'    Dim gamma As Double
'    Dim vega As Double
'    Dim theta As Double
'    Dim delta_pl As Double
'    Dim gamma_pl As Double
'    Dim vega_pl As Double
'    Dim theta_pl As Double
'    Dim tau As Double
'    Dim dur_els As Double
'    Dim iv_110 As Double
'    Dim iv_90 As Double
'    Dim price_frn As Double
'    Dim dv01_frn As Double
'    Dim frn As Double

    Dim seq_from_els As Integer
    Dim seq_to_els As Integer
    seq_from_els = Range("from_seq").value
    seq_to_els = Range("to_seq").value
    
    Dim note_price As Double
    Dim swap_price As Double
    Dim frn_price As Double
    Dim plain_frn_price As Double
    Dim day1_pl_price As Double
    
    For i = seq_from_els To seq_to_els
    
        sql = getSQL(SQL_PATH_ELS, bind_variable, bind_value)
        
        With oCmd
    
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = sql
    
            oRS.Open .Execute
    
        End With
    
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        
        i = 1
    
        Do Until oRS.EOF
            
            If selected_hedge_type_code = HEDG_TR_TYPE_CODE.all Or selected_hedge_type_code = oRS(0) Then
        
                els_list(i, 1) = i
                
                For j = 0 To oRS.Fields.count - 1
                    els_list(i, j + 2) = oRS(j)
                Next j
                
                i = i + 1
            End If
    
            oRS.MoveNext
    
        Loop
    
    Next i

    With Application
        .ScreenUpdating = bScreenUpdating
        .StatusBar = btatusBar
        .DisplayStatusBar = bDisplayStatusBar
        .Calculation = vCalculation
        .EnableEvents = bEnableEvents
    End With

End Sub

Private Sub btnEstimatePrice_Click()

End Sub

Private Sub btnReadPrice_Click()
'RM 이론가 및 기준가 조회

    Dim target_date As Date
    target_date = Range("target_date").value
    
    Dim rng As Range
    Set rng = Range("TBL_ELS_PRICING_LIST")

    Dim adoCon As New adoDB.Connection
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(2) As String
    ReDim bind_value(2) As Variant
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":code"
    bind_value(1) = date2str(target_date)
    
    Dim i As Integer
    Dim sql As String
        
    For i = 1 To rng.Rows.count
    
        If IsEmpty(rng(i, ELS_COL.SQ)) = False Then
            
            bind_value(2) = rng(i, ELS_PRICING_COL.FO_ELS_D).value
            
            'sql = getSQL(SQL_PATH_ELS_PRICE, bind_variable, bind_value)
            sql = "select code, theory_price, market_price from ras.rm_els_data where tdate = '" + bind_value(1) + "' and code = '" + bind_value(2) + "'"
    
            With oCmd
            
                .ActiveConnection = adoCon
                .CommandType = adCmdText
                .CommandText = sql
                
                oRS.Open .Execute
            
            End With
            
            Do Until oRS.EOF
                rng(i, ELS_PRICING_COL.els_price_theo) = oRS(1)
                rng(i, ELS_PRICING_COL.els_price_mtm) = oRS(2)
                oRS.MoveNext
            Loop
            
            oRS.Close
            
        End If
        
    Next i
    
    Call disconnectDB(adoCon)
    
End Sub

Public Sub btnReadTargetELS_Click()

    SET_GLOBAL

    Dim target_date As Date
    target_date = Range("target_date").value
    
    Dim tdate_str As String
    tdate_str = date2str(target_date)
                
    Dim tbl As Range
    Set tbl = shtELSList.Range("TBL_ELS_LIST")
    
    Dim rng As Range
    Set rng = Range("TBL_ELS_PRICING_LIST")
    rng.ClearContents
    
    Range("cur_seq").ClearContents
    Range("elapsed_time").ClearContents
    
    Dim els_pricing_list As Variant
    els_pricing_list = rng.value

    Dim adoCon As New adoDB.Connection
'    Dim oCmd As New adoDB.Command
'    Dim oRS As New adoDB.Recordset
    
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
'    Dim bind_variable() As String
'    Dim bind_value() As Variant
'    ReDim bind_variable(2) As String
'    ReDim bind_value(2) As Variant
'    bind_variable(1) = ":tdate"
'    bind_value(1) = tdate_str
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim sql As String
    
    Dim ua_pct_price() As Double
    Dim ua_wp As String
    Dim ua_wp_pct_price As Double
    Dim ac_deal_ticket As clsACDealTicket
    Dim latest_call_date As Date
    Dim latest_bus_date As Date
    Dim spot_date As Date
    Dim is_delayed_spot As Boolean
    
    'latest_bus_date = get_latest_bus_date(target_date)
    latest_bus_date = shtMarket.Range("market_prevdate").value 'temp
    
    Dim rtn(4) As Variant
    
    j = 1
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    With rng.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Dim prod_code As String
    
    For i = 1 To tbl.Rows.count
    
        'If tbl(i, ELS_COL.HEDG_TR_TYPE_CODE).value = 1 Then
        'If tbl(i, ELS_COL.HEDG_TR_TYPE_CODE).value = 2 Then
        If tbl(i, ELS_COL.HEDG_TR_TYPE_CODE).value = 1 Or (tbl(i, ELS_COL.HEDG_TR_TYPE_CODE).value = 2 And tbl(i, ELS_COL.IO_FRN_H) <> "" And tbl(i, ELS_COL.PROD_CLS_CODE) <> "ELB" And tbl(i, ELS_COL.PRC_APLY_CLS_CODE) = "D") Then '자체헤지 + unfunded BTB(ex.ELB) & 자체평가
            
            els_pricing_list(j, ELS_PRICING_COL.SQ) = j
            els_pricing_list(j, ELS_PRICING_COL.NO) = i
            els_pricing_list(j, ELS_PRICING_COL.HEDG_TR_TYPE_CODE) = tbl(i, ELS_COL.HEDG_TR_TYPE_CODE).value
            
            els_pricing_list(j, ELS_PRICING_COL.PROD_CLS_CODE) = tbl(i, ELS_COL.PROD_CLS_CODE).value
            els_pricing_list(j, ELS_PRICING_COL.FO_ELS_D) = tbl(i, ELS_COL.FO_ELS_D).value
            
            els_pricing_list(j, ELS_PRICING_COL.PROD_CLS_CODE_H) = tbl(i, ELS_COL.PROD_CLS_CODE_H).value
            els_pricing_list(j, ELS_PRICING_COL.IO_ESWAP_H) = tbl(i, ELS_COL.IO_ESWAP_H).value
            els_pricing_list(j, ELS_PRICING_COL.IO_ESWAP_T) = tbl(i, ELS_COL.IO_ESWAP_T).value
            
            els_pricing_list(j, ELS_PRICING_COL.KOR_ISNM) = tbl(i, ELS_COL.KOR_ISNM).value
            els_pricing_list(j, ELS_PRICING_COL.STLM_CRCD) = tbl(i, ELS_COL.STLM_CRCD).value
            els_pricing_list(j, ELS_PRICING_COL.NOTIONAL_CCY) = tbl(i, ELS_COL.NOTIONAL_CCY).value
            
            'els_pricing_list(j, ELS_PRICING_COL.HEDG_BUF) = tbl(i, ELS_COL.HEDG_BUF).value
            els_pricing_list(j, ELS_PRICING_COL.DAY1_PL_D) = tbl(i, ELS_COL.HEDG_BUF_D).value * tbl(i, ELS_COL.NOTIONAL_CCY).value * (tbl(i, ELS_COL.MTRT_DATE).value - target_date) / (tbl(i, ELS_COL.MTRT_DATE).value - tbl(i, ELS_COL.PBLC_DATE).value)
            els_pricing_list(j, ELS_PRICING_COL.DAY1_PL_H) = tbl(i, ELS_COL.HEDG_BUF_H).value * tbl(i, ELS_COL.NOTIONAL_CCY).value * (tbl(i, ELS_COL.MTRT_DATE).value - target_date) / (tbl(i, ELS_COL.MTRT_DATE).value - tbl(i, ELS_COL.PBLC_DATE).value)

            '/<------ pricing 가능 기초자산 판단 목적
            If tbl(i, ELS_COL.HEDG_TR_TYPE_CODE).value = 1 Then
                prod_code = tbl(i, ELS_COL.IO_ESWAP_T).value '자체 E.swap IO~T
            ElseIf tbl(i, ELS_COL.HEDG_TR_TYPE_CODE).value = 2 Then
                prod_code = tbl(i, ELS_COL.IO_ESWAP_H).value '백투백 ELS IO~H
            End If

            'deal ticket 생성
            Set ac_deal_ticket = get_ac_deal_ticket(prod_code, target_date, False, False, adoCon) '회계처리일이 아닌 계약 상 실제 조기상환결정일 기준: eval_date_lag = false
            'Set ac_deal_ticket = get_ac_deal_ticket(prod_code, target_date, False, False, adoCon, , True)

            'pricing 불가능 기초자산 종목은 제외
            'For k = 1 To tbl(i, ELS_COL.NUM_UNAS)
            For k = 1 To ac_deal_ticket.no_of_ul
                If is_active_ua(ac_deal_ticket.ul_code(k)) = False Then
                    GoTo JUMP_TO_NEXT_PROD
                End If
            Next k
            '------->/

            If Range("check_autocall").value = True Then
                
                'target_date 포함 이전 최직근 call_date
                latest_call_date = ac_deal_ticket.value_date
                For k = 1 To ac_deal_ticket.no_of_schedule
                    If target_date >= ac_deal_ticket.autocall_schedules(k).call_date Then
                        latest_call_date = ac_deal_ticket.autocall_schedules(k).call_date
                    End If
                Next k
                
                spot_date = target_date
                is_delayed_spot = True
                    
                If ac_deal_ticket.has_eval_shift_ul = True Then
                
                    '1 거래일: latest_bus_date < value_date = latest_call_date = target_date                  -> spot_date = target_date, is_delayed_spot = true
                    '2 발행일: latest_bus_date <= value_date = latest_call_date <= target_date                -> spot_date = target_date, is_delayed_spot = true
                    '1/2: latest_bus_date <= latest_call_date <= target_date & value_date = latest_call_date
                    
                    '3 평상시: latest_call_date < latest_bus_date < target_date                               -> spot_date = target_date, is_delayed_spot = true
                    '4 조기상환평가일 도래했으나 일부 종가 미형성: latest_bus_date < latest_call_date = target_date       -> spot_date = target_date, is_delayed_spot = true
                                        
                    '5 조기상환평가 회계적용일: latest_bus_date = latest_call_date < target_date                  -> spot_date = latest_call_date, is_delayed_spot = false
                    '6 한국 연휴기간 중 조기상환평가일 도래: latest_bus_date < latest_call_date < target_date         -> spot_date = latest_call_date, is_delayed_spot = false
                    '5/6: latest_bus_date <= latest_call_date < target_date & value_date < latest_call_date
                    
                    If latest_bus_date <= latest_call_date And latest_call_date < target_date And ac_deal_ticket.value_date < latest_call_date Then
                        spot_date = latest_call_date
                        is_delayed_spot = False
                    End If
                
                End If
                
                '기초자산 현재가
                'ReDim ua_pct_price(tbl(i, ELS_COL.NUM_UNAS)) As Double
                ReDim ua_pct_price(ac_deal_ticket.no_of_ul) As Double
                ua_wp_pct_price = MAX_UA_PCT_PRICE
                
                'For k = 1 To tbl(i, ELS_COL.NUM_UNAS)
                For k = 1 To ac_deal_ticket.no_of_ul
                               
                    'ua_pct_price(k) = get_spot(ac_deal_ticket.ul_code(k), target_date, True, adoCon) / ac_deal_ticket.reference_price(k)
                    ua_pct_price(k) = get_spot(ac_deal_ticket.ul_code(k), spot_date, is_delayed_spot, adoCon) / ac_deal_ticket.reference_price(k)
    
                    If ua_pct_price(k) < ua_wp_pct_price Then
                        ua_wp = ac_deal_ticket.ul_code(k)
                        ua_wp_pct_price = ua_pct_price(k)
                    End If
    
                    els_pricing_list(j, ELS_PRICING_COL.ua_pct_price + get_ua_ofs(ac_deal_ticket.ul_code(k))) = ua_pct_price(k)
                Next k
    
                'Worst performer
                els_pricing_list(j, ELS_PRICING_COL.UA_WORST) = ua_wp
                els_pricing_list(j, ELS_PRICING_COL.UA_WORST_PCT_PRICE) = ua_wp_pct_price
    
                '기준일 이벤트
                'els_pricing_list(j, ELS_PRICING_COL.TODAY_EVENT) = ac_deal_ticket.comment
                'els_pricing_list(j, ELS_PRICING_COL.TODAY_EVENT) = check_event(ac_deal_ticket, target_date, ua_pct_price)
    
                For k = 1 To ac_deal_ticket.no_of_schedule
    
                    'If ac_deal_ticket.autocall_schedules(k).call_date >= target_date Then
                    If ac_deal_ticket.autocall_schedules(k).call_date >= spot_date Then
                    
                        If eval_next_autocall(ac_deal_ticket, ac_deal_ticket.autocall_schedules(k).call_date, ua_pct_price, rtn) = True Then
    
                            els_pricing_list(j, ELS_PRICING_COL.REDEEMABLE_NO) = rtn(1)
                            els_pricing_list(j, ELS_PRICING_COL.REDEEMABLE_DATE) = ac_deal_ticket.autocall_schedules(CInt(rtn(1))).call_date
                            els_pricing_list(j, ELS_PRICING_COL.REDEEMABLE_DAYS_LEFT) = ac_deal_ticket.autocall_schedules(CInt(rtn(1))).call_date - target_date
                            els_pricing_list(j, ELS_PRICING_COL.REDEEMABLE_BARRIER) = rtn(2)
                            els_pricing_list(j, ELS_PRICING_COL.REDEEMABLE_RATIO) = rtn(3)
                            els_pricing_list(j, ELS_PRICING_COL.REDEEMABLE_MARGIN) = rtn(4) - rtn(2) '베리어까지 버퍼(+: 조기상환, -: 조기상환 실패 또는 만기손실상환)
    
                            '만기 판단 조건
                            If els_pricing_list(j, ELS_PRICING_COL.REDEEMABLE_DATE) = ac_deal_ticket.autocall_schedules(ac_deal_ticket.no_of_schedule).call_date Then
                                els_pricing_list(j, ELS_PRICING_COL.IS_MATURITY) = "Y"
                            End If
    
                            Exit For
                        End If
    
                    End If
    
                Next k
    
                '다음 조기상환 회차
                'Call eval_next_autocall(ac_deal_ticket, target_date, ua_pct_price, rtn)
                Call eval_next_autocall(ac_deal_ticket, spot_date, ua_pct_price, rtn)
                
                If rtn(1) > 0 Then
                    els_pricing_list(j, ELS_PRICING_COL.next_ac_no) = rtn(1)
                    els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_DATE) = ac_deal_ticket.autocall_schedules(CInt(rtn(1))).call_date
                    els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_DAYS_LEFT) = ac_deal_ticket.autocall_schedules(CInt(rtn(1))).call_date - target_date
                    els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_BARRIER) = rtn(2)
                    els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_REDEEM_RATIO) = rtn(3)
                    els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_MARGIN) = rtn(4) - rtn(2) '베리어까지 버퍼(+: 조기상환, -: 조기상환 실패 또는 만기손실상환)
                End If
    
                'Check event
                'check_event ac_deal_ticket, target_date, ua_pct_price, next_ac_no
    
                '상환 / PIN 판단 조건
                If els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_DAYS_LEFT) < 14 And Abs(els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_MARGIN)) < 0.03 Then
                    els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_AVAILABLE) = "PIN"
                ElseIf els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_DATE) = ac_deal_ticket.autocall_schedules(ac_deal_ticket.no_of_schedule).call_date Then
                    els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_AVAILABLE) = "만기"
                ElseIf els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_MARGIN) > 0 Then
                    els_pricing_list(j, ELS_PRICING_COL.NEXT_AC_AVAILABLE) = "Y"
                End If
            
            End If
            
            If j Mod 2 = 0 Then
                With rng.Rows(j).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            End If
            
            j = j + 1
        
        End If
        
JUMP_TO_NEXT_PROD:
        
    Next i
    
    rng.value = els_pricing_list
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Call disconnectDB(adoCon)
    
    Range("from_seq") = 1
    Range("to_seq") = WorksheetFunction.CountA(Range("TBL_ELS_PRICING_LIST").Columns(1))
    
    
End Sub

Private Function get_ua_ofs(ua_code As String) As Integer

    Dim rtn As Integer
    
    Select Case get_ua_idx(ua_code)
    Case ua.KOSPI200: rtn = 0
    Case ua.NKY: rtn = 1
    Case ua.HSCEI: rtn = 2
    Case ua.SX5E: rtn = 3
    Case ua.SPX: rtn = 4
    Case ua.TSLA: rtn = 5
    End Select

    get_ua_ofs = rtn
    
End Function

Private Function check_event(ac_deal_ticket As clsACDealTicket, target_date As Date, ua_pct_price() As Double, next_ac_no As Integer) As String
    
    Dim rtn As String

    '[발행관련]
    
    '신규거래
    If target_date = ac_deal_ticket.value_date Then
        rtn = "신규 거래"
        '발행취소
'        If ac_deal_ticket.alive_yn Then
'            rtn = "발행 취소"
'        End If
    End If
    
    '신규발행
    If target_date = ac_deal_ticket.settlement_date Then
        rtn = "신규 발행"
    End If

    '[상환관련]
    
    '상환평가일: 상환실패
    '상환평가일: 종가 미확정
    '상환확정: 상환확정(#차)
    '상환확정: 만기
    '상환확정: 만기손실
    
    'ac_deal_ticket에는 eval_shift가 적용되어 있지 않다고 가정.
    
    '상환평가일: 종가 미확정, 상환 실패
    If target_date = ac_deal_ticket.autocall_schedules(next_ac_no).call_date Then

        Select Case ac_deal_ticket.autocall_schedules(next_ac_no).performance_type
        Case 0: rtn = "상환평가일: 종가 미확정" 'avg -> 무조건 미확정
        Case -1, 1: rtn = "상환평가일: 종가 미확정" 'min, max -> EVAL_SHIFT_UA이 아닌 기초지산 종가 체크 -> 1) 하나라도 만족 못하면 상환실패 2) 모두 만족하면, 미확정

            Dim u As Integer
            For u = 1 To ac_deal_ticket.no_of_ul
                If is_eval_shift_ua(ac_deal_ticket.ul_code(u)) = False Then
                
                    eval_next_autocall
                
                    '29 EarlyExit
                    Dim ee_touched_array() As Long
                    ee_touched_array = ac_deal_ticket.early_exit_touched_flags
                    
                    If (ee_touched_array(ee_no) = 0) And (ua_pct_price(u) < ac_deal_ticket.early_exit_barrier(ee_no)) Then
                        rtn = "상환평가일: 상환실패"
                    End If
                    
                    '01 HiFive, 02 StepDown, 17 SafeStepDown, 22 CouponStepDown, 29 EarlyExit, 36 A-win, 38 Five Wins
                    If ua_pct_price(u) < ac_deal_ticket.autocall_schedules(next_ac_no).percent_strike Then
                        rtn = "상환평가일: 상환실패"
                    End If
                    
                    Exit For
                End If
            Next u
            
        End Select

    End If
    
    '상환평가일: 상환확정
    If target_date > ac_deal_ticket.autocall_schedules(next_ac_no).call_date Then
    
    End If
    
    '모두 해외지수이고 call_date가 국내 연휴기간 중인 경우, target_date > 조기상환결정일(TRTH_CLRD_DTRM_DATE)+1 이더라도 상환여부 판단 필요.
    '이 경우, 해당 기간 중 해외지수 종가가 동일하게 국내 비영업일 기준으로 하루 지연되어 입수되어야 함.
    

    
    '[베리어 관련]
    'KI touch
        Select Case ac_deal_ticket.ki_performance_type
        Case -1: ki_det_price = WorksheetFunction.min(ua_pct_price)
        Case 0: ki_det_price = WorksheetFunction.avg(ua_pct_price)
        Case 1: ki_det_price = WorksheetFunction.max(ua_pct_price)
        End Select
    'EE touch
    
    '[현금흐름 관련]
    '월 쿠폰지급
    
    '변동금리 수취
    '변동금리 스케줄 관리 -> floating_settlement_date
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(1) As String
    ReDim bind_value(1) As Variant
    bind_variable(1) = ":indv_iscd"
    bind_value(1) = ac_deal_ticket.asset_code
    
    Dim sql As String
    sql = getSQL(SQL_PATH_FLOATING_LEG, bind_variable, bind_value)
    
    With oCmd
    
        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql
        
        oRS.Open .Execute
    
    End With
    
    Do Until oRS.EOF
        
        If oRS("PRVS_DATE") = date2str(target_date) Then
            rtn = rtn + "변동금리 수취"
            Exit Do
        End If
        
        oRS.MoveNext
    Loop
    
    oRS.Close
    
    check_event = rtn
    
End Function

Private Function eval_next_autocall(ac_deal_ticket As clsACDealTicket, target_date As Date, ua_pct_price() As Double, rtn() As Variant) As Boolean
    
    '조기상환유형
    '01 HiFive
    '02 StepDown
    '17 SafeStepDown
    '22 CouponStepDown
    '29 EarlyExit
    '36 A-win
    '38 Five Wins
    
    'out
    'rtn(1) = 조기상환 회차
    'rtn(2) = barrier(%)
    'rtn(3) = 상환 지급율
    'rtn(4) = 기초자산 평가기준가
    
    eval_next_autocall = False
    
    'check barrier
    Dim autocall_no As Integer
    Dim barrier As Double
    Dim redeem_ratio As Double
    
    Dim i As Integer
    
    autocall_no = 0 '기본값
    
    For i = 1 To ac_deal_ticket.no_of_schedule
        If ac_deal_ticket.autocall_schedules(i).call_date >= target_date Then
            autocall_no = i
            Exit For
        End If
    Next i
    
    If autocall_no > 0 Then
    
    '기초자산 평가방법
    Dim ua_det_price As Double
    Dim ki_det_price As Double
    
    Select Case ac_deal_ticket.ki_performance_type
    Case -1: ki_det_price = WorksheetFunction.min(ua_pct_price)
    Case 0: ki_det_price = WorksheetFunction.avg(ua_pct_price)
    Case 1: ki_det_price = WorksheetFunction.max(ua_pct_price)
    End Select
    
    Select Case ac_deal_ticket.autocall_schedules(autocall_no).performance_type
    Case -1: ua_det_price = WorksheetFunction.min(ua_pct_price)
    Case 0: ua_det_price = WorksheetFunction.avg(ua_pct_price)
    Case 1: ua_det_price = WorksheetFunction.max(ua_pct_price)
    End Select
    
    Dim ac_success As Boolean
    ac_success = (ac_deal_ticket.autocall_schedules(autocall_no).percent_strike < ua_det_price)
    
    If ac_deal_ticket.early_exit_flag = 1 Then

        Dim ee_no As Integer
        For i = 1 To ac_deal_ticket.no_of_early_exit_schedule
            If ac_deal_ticket.early_exit_date(ee_no) >= target_date Then
                ee_no = i
                Exit For
            End If
        Next i
        
        Dim ee_touched_array() As Long
        ee_touched_array = ac_deal_ticket.early_exit_touched_flags
        
        Dim ee_success As Boolean
        If ee_no > 0 Then
            ee_success = (ee_touched_array(ee_no) = 0) And (ac_deal_ticket.early_exit_barrier(ee_no) < ua_det_price)
        Else
            ee_success = False
        End If
        
    End If
    
    If ac_deal_ticket.monthly_coupon_flag = 1 Then

        Dim m_cpn_no As Integer
        For i = 1 To ac_deal_ticket.no_of_coupon_schedule
            If ac_deal_ticket.monthly_coupon_schedules(i).call_date >= target_date Then
                m_cpn_no = i
                Exit For
            End If
        Next i
            
        Dim m_cpn_success As Boolean
        If m_cpn_no > 0 Then '2024.02.26
            m_cpn_success = (ac_deal_ticket.monthly_coupon_schedules(m_cpn_no).percent_strike < ua_det_price)
        End If
        
    End If
    
    'update barrier touch
    If ac_deal_ticket.ki_barrier_flag = 1 Then
        Dim ki_touched As Boolean
        ki_touched = (ac_deal_ticket.ki_touched_flag = 1) Or (ac_deal_ticket.ki_barrier > ki_det_price)
    End If
    
    '조건 적용 순서
    '29 EarlyExit
    If ee_success Then
    
        eval_next_autocall = True
    
        barrier = ac_deal_ticket.early_exit_barrier(ee_no)
        redeem_ratio = 1 + ac_deal_ticket.early_exit_coupon(ee_no)
    
    Else
        barrier = ac_deal_ticket.autocall_schedules(autocall_no).percent_strike
        
        '01 HiFive, 02 StepDown, 17 SafeStepDown, 22 CouponStepDown, 29 EarlyExit, 36 A-win, 38 Five Wins
        If ac_success Then
            
            eval_next_autocall = True

            redeem_ratio = 1 + ac_deal_ticket.autocall_schedules(autocall_no).coupon_on_call
            
            If m_cpn_success Then
                redeem_ratio = redeem_ratio + ac_deal_ticket.monthly_coupon_schedules(m_cpn_no).coupon_on_call
            End If
        
        Else
            redeem_ratio = 0
            
            '만기상환
            'If ac_deal_ticket.autocall_schedules(autocall_no).call_date = ac_deal_ticket.maturity_date Then
            If autocall_no = ac_deal_ticket.no_of_schedule Then
            
                eval_next_autocall = True
                
                If ac_deal_ticket.ki_barrier_flag = 0 Then
                    redeem_ratio = ua_det_price
                Else
                    If ki_touched Then
                        redeem_ratio = ua_det_price
                    Else
                        barrier = ac_deal_ticket.ki_barrier
                        redeem_ratio = 1 + ac_deal_ticket.autocall_schedules(autocall_no).coupon_on_call
                    End If
                End If
                
            End If
            
        End If
        
    End If
    
    End If
    
    rtn(1) = autocall_no
    rtn(2) = barrier
    rtn(3) = redeem_ratio
    rtn(4) = ua_det_price

End Function

Private Sub btnInsertPrice_Click()

    Dim from_seq As Integer
    Dim to_seq As Integer
    from_seq = Range("from_seq").value
    to_seq = Range("to_seq").value
    
    If vbYes = MsgBox("지금 화면에 출력된 SEQ=" & from_seq & "~" & to_seq & " ELS 가격을 DB에 입력하겠습니까?", vbYesNo) Then
    
        Dim rng As Range
        Set rng = Range("TBL_ELS_PRICING_LIST")
    
        Dim rng2 As Range
        Set rng2 = shtELSList.Range("TBL_ELS_LIST")
        
        Dim sql As String
        
        Dim i, j As Integer
        Dim hedge_buf As Double
        Dim hedge_buf_h As Double
        Dim str_db As String
        
        Dim oCmd As New adoDB.Command
        Dim oRS As New adoDB.Recordset
        
        Dim target_date As Date
        target_date = Range("target_date").value
        
        Dim tdate_str As String
        tdate_str = date2str(target_date)
   
        Dim adoCon As New adoDB.Connection
        Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
        'Call connectDB(adoCon, TNS_SERVICE_NAME_TEST, USER_ID, PASSWORD) 'for TEST DB
        
        Dim prod_cls_name As String
        Dim prod_seq As String
        Dim deal_name As String
        
        Dim unit_notional As Double
        Dim ccy As String
        Dim payoff_desc As String
        Dim value_date As String
        Dim issue_date As String
        Dim deal_price As Double
        
'        Dim els_code As String
'        Dim eswap_code As String
'        Dim eswap_code_t As String
'
'        Dim els_theo_price As Double
'        Dim els_market_price As Double
'        Dim eswap_theo_price As Double
'        Dim eswap_market_price As Double

        Dim io_code As String

        Dim fo_code_d As String
        Dim io_code_h As String
        Dim io_code_t As String
        
        Dim theo_price_d As Double
        Dim theo_price_h As Double
        Dim mtm_price_d As Double
        Dim mtm_price_h As Double
        
        Dim tag As String
        
        For i = from_seq To to_seq
            
            fo_code_d = rng(i, ELS_PRICING_COL.FO_ELS_D)
            io_code_h = rng(i, ELS_PRICING_COL.IO_ESWAP_H)
            io_code_t = rng(i, ELS_PRICING_COL.IO_ESWAP_T)
            
            sql = "delete from ras.rm_pricing_data where tdate = '" + tdate_str + "' and code in ('" + fo_code_d + "','" + io_code_h + "','" + io_code_t + "') "
            adoCon.Execute (sql)
            
            theo_price_d = rng(i, ELS_PRICING_COL.PRICE_THEO_D)
            theo_price_h = rng(i, ELS_PRICING_COL.PRICE_THEO_H)
            mtm_price_d = rng(i, ELS_PRICING_COL.PRICE_MTM_D)
            mtm_price_h = rng(i, ELS_PRICING_COL.PRICE_MTM_H)
            
            If rng(i, ELS_PRICING_COL.PROD_CLS_CODE) = "스왑" Then
                tag = "eq_ac_swap"
            Else
                tag = "eq_ac_note"
            End If
            
            sql = "insert into ras.rm_pricing_data values ('" + tdate_str + "','" + fo_code_d + "', " & theo_price_d & ", " & mtm_price_d & ", SYSDATE, '" & tag & "') "
            adoCon.Execute (sql)
            
            If rng(i, ELS_PRICING_COL.PROD_CLS_CODE_H) = "스왑" Then
                tag = "eq_ac_swap"
            Else
                tag = "eq_ac_note"
            End If
            
            'SP 종목
            sql = "insert into ras.rm_pricing_data values ('" + tdate_str + "','" + io_code_h + "', " & theo_price_h & ", " & mtm_price_h & ", SYSDATE, '" & tag & "') "
            adoCon.Execute (sql)
            io_code = io_code_h
                            
            'E파 종목
            If io_code_t <> "" Then
                sql = "insert into ras.rm_pricing_data values ('" + tdate_str + "','" + io_code_t + "', " & theo_price_h & ", " & mtm_price_h & ", SYSDATE, '" & tag & "') "
                adoCon.Execute (sql)
                io_code = io_code_t
            End If
                        
            '신규종목 입력
            If target_date = rng2(rng(i, ELS_PRICING_COL.NO), ELS_COL.PBLC_DATE) Then

                sql = "select a.PROD_CLS_CODE, a.FUND_PBLC_TIMS, a.REAL_PBLC_FCAM, a.STLM_CRCD, a.BRIF_EXPL, a.FRST_STND_PRC_FIN_DATE, a.PBLC_DATE, c.PERC_APLY_THPR, replace(b.ENG_ISNM,'SY'), a.FUND_PBLC_UNPR/a.REAL_PBLC_FCAM from BSYS.TBSIMO100M00@GDW a, BSYS.TBSIMM100M00@GDW b, BSYS.TBSIMO201M00@GDW c where a.OTC_FUND_ISCD = '" + fo_code_d + "' and a.OTC_FUND_ISCD = c.OTC_FUND_ISCD and a.OTC_FUND_ISCD = b.ISCD"

                With oCmd

                    .ActiveConnection = adoCon
                    .CommandType = adCmdText
                    .CommandText = sql

                    oRS.Open .Execute

                End With

                Do Until oRS.EOF
'                    Select Case oRS(0)
'                    Case "04", "09": prod_cls_name = "OTC"
'                    Case "07": prod_cls_name = "ELS"
'                    Case "13": prod_cls_name = "ELB"
'                    End Select
'
'                    prod_seq = oRS(1)
                    unit_notional = oRS(2)
                    ccy = oRS(3)
                    payoff_desc = oRS(4)
                    value_date = oRS(5)
                    issue_date = oRS(6)
                    If oRS(0) = "04" Then 'swap인 경우, note 기준 장부가
                        deal_price = oRS(9)
                    Else
                        deal_price = oRS(7)
                    End If
                    deal_name = oRS(8)

                    oRS.MoveNext
                Loop

                oRS.Close

                'deal_name = prod_cls_name + prod_seq

                hedge_buf = Round((mtm_price_d - theo_price_d) / unit_notional, 6)
                hedge_buf_h = Round((mtm_price_h - theo_price_h) / unit_notional, 6)

                str_db = ", ras.rm_els_info"

                sql = "MERGE INTO RAS.RM_ELS_INFO USING dual ON (deal_name = '" + deal_name + "') WHEN MATCHED THEN Update SET code = '" + fo_code_d + "', hedge_buf = " & hedge_buf & ", hedge_buf_h = " & hedge_buf_h & ", inserttime = sysdate, indv_iscd = '" + io_code + "', status = '발행 완료' "
                sql = sql & " WHEN NOT MATCHED THEN INSERT (code, hedge_buf, hedge_buf_h, inserttime, indv_iscd, payoff_desc, ccy, value_date, issue_date, deal_price, status, deal_name) VALUES ('" + fo_code_d + "'," & hedge_buf & "," & hedge_buf_h & ", SYSDATE, '" + io_code + "', '" + payoff_desc + "','" + ccy + "','" + value_date + "','" + issue_date + "'," & deal_price & ",'발행 완료','" + deal_name + "')"

                adoCon.Execute (sql)

            End If
        
        Next i
        
        Call disconnectDB(adoCon)
        
        Set oCmd = Nothing
        Set oRS = Nothing
        Set adoCon = Nothing
        
        'MsgBox ("insert 완료: excel -> ras.rm_els_data" & str_db)
        MsgBox ("insert 완료: excel -> ras.rm_pricing_data" & str_db)
        
    End If
    
End Sub

Private Function is_completed(prod_code As String, target_date As Date, scenario_id As String, adoCon As adoDB.Connection) As Boolean
    
    Dim rtn As Boolean
    
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    
    Dim sql As String
    
    sql = "select count(a.stk_code) from rcs.pml_position_st a, rcs.pml_greek_st b, rms.mr_scenario_detail_data c where a.tdate = b.tdate" _
        & " and a.tdate=c.tdate and a.stk_code=b.stk_code and a.stk_code=c.code and a.scenarioid=b.scenarioid and a.scenarioid=c.scenarioid and a.tdate='" + date2str(target_date) + "' and a.scenarioid='" + scenario_id + "' and a.stk_code ='" + prod_code + "'"

    With oCmd
        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql
        oRS.Open .Execute
    End With
    
    Do Until oRS.EOF
        If oRS(0) = 0 Then
            rtn = False
        Else
            rtn = True
        End If
        oRS.MoveNext
    Loop
    
    oRS.Close
    
    is_completed = rtn
            
End Function