Option Explicit

'Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Sub checkboxQuote_Click()

    If shtACPricer.Range("isQuote") = True Then
        shtACPricer.Range("asset_code").Cells(1, 2).ClearContents
        shtACPricer.Range("asset_code").ClearContents
    Else
        shtACPricer.Range("asset_code").Cells(1, 2).Formula = "=ELS!B6"
        shtACPricer.Range("asset_code").Formula = "=VLOOKUP(ELS!B6,ELS!A10:AL1000,38,FALSE)"
    End If
    
End Sub


Public Sub checkboxOpenGreeks_Click()
    
    If shtACPricer.Range("chkOpenGreeks") = True Then
        shtACPricer.Range("chkVega") = True
        shtACPricer.Range("chkSkew") = False
        shtACPricer.Range("chkConv") = False
        shtACPricer.Range("chkCorr") = False
        shtACPricer.Range("chkRho") = False
        shtACPricer.Range("chkStickyStrikeDelta") = False
        shtACPricer.Range("chkTermVega") = False
        shtACPricer.Range("chkTermSkew") = False
        shtACPricer.Range("chkTermConv") = False
        shtACPricer.Range("chkDeltaGamma") = True
        shtACPricer.Range("chkSnapshot") = False
        shtACPricer.Range("chkDuration") = False
        shtACPricer.Range("chkTheta") = False
        shtACPricer.Range("chkVegaPL") = False
    Else
        shtACPricer.Range("chkVega") = False
        shtACPricer.Range("chkSkew") = False
        shtACPricer.Range("chkConv") = False
        shtACPricer.Range("chkCorr") = False
        shtACPricer.Range("chkRho") = False
        shtACPricer.Range("chkStickyStrikeDelta") = False
        shtACPricer.Range("chkTermVega") = False
        shtACPricer.Range("chkTermSkew") = False
        shtACPricer.Range("chkTermConv") = False
        shtACPricer.Range("chkDeltaGamma") = False
        shtACPricer.Range("chkSnapshot") = False
        shtACPricer.Range("chkDuration") = False
        shtACPricer.Range("chkTheta") = False
        shtACPricer.Range("chkVegaPL") = False
    End If
    
End Sub


'----------------------------------------------
' Sub: run
' Called by User
' Calculate price and greeks for the deal loaded on the screen.
'-----------------------------------------------
Public Sub cmdCalcPrice_Click()
 
'Dim tick

On Error GoTo ErrorHandler
Call SET_GLOBAL
    'Application.Calculation = xlCalculationManual
    
    'tick = GetTickCount
    
    Dim deal_ticket As clsACDealTicket
    Set deal_ticket = read_ac_deal_ticket()

    Dim market_set As clsMarketSet
    Set market_set = read_market_set
'    Select Case deal_ticket.CCY
'    Case "KRW": Set market_set = read_market_set("KRW")
'    Case "USD": Set market_set = read_market_set("USD")
'    End Select
    
    'modify deal_ticket for leverage indices : 2018.7.9
    'modify deal_ticket for the ejectable structure : 2021.11.12
    Set deal_ticket = modify_ac_deal_ticket(deal_ticket, market_set)
                       
    'Call cmd_calc_ac_price(deal_ticket, market_set)
    
    Dim greeks As New clsGreeks
        
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
                                
    clear_ac_greeks
        
    display_greeks_nd greeks, deal_ticket.no_of_ul
        
    Set greeks = Nothing
    
    'tick = GetTickCount - tick
    'Application.Range("A1").value = tick
    
    'Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    show_error
    Application.Calculation = xlCalculationAutomatic

End Sub

Private Function find_ua(ary_ua() As String, str_ua As String) As Boolean
    
    Dim i As Integer
    
    For i = 1 To UBound(ary_ua)
    
        If ary_ua(i) = str_ua Then
            find_ua = True
            Exit For
        End If
        
    Next i
    
    
End Function

'----------------------------------------------
' Sub: run
' Called by User
' Calculate price and greeks for the deal loaded on the screen.
'-----------------------------------------------
'분산처리
Sub cmdCalcPriceDist_Click()
 
Dim tick

On Error GoTo 0 'On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual
    
    Application.Range("A1").value = GetTickCount
    
    cmd_calc_ac_price_distributed
    
'    If Me.Range("chkDuration").value Then
'        cmd_duration
'    End If
    

    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    show_error
    Application.Calculation = xlCalculationAutomatic

End Sub

Public Sub cmdCallProbability_Click()

On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual
      
    cmd_call_probability
    
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    show_error
    Application.Calculation = xlCalculationAutomatic



End Sub

Public Sub cmdDuration_Click()


On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual

    cmd_duration
    
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    show_error
    Application.Calculation = xlCalculationAutomatic

End Sub

Public Sub cmdExpiry_Click()
On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual

    cmd_expiry_ac
    
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    show_error
    Application.Calculation = xlCalculationAutomatic
End Sub
Public Sub cmdChangeReference_Click()

    Dim inx As Integer
    Dim jnx As Integer
    Dim Maturity As Double
    Dim frequency As Integer
    Dim coupon_pa As Double
    Dim cpn_frequency As Integer
    Dim no_of_ul As Integer
    
    Dim date_shift As Integer
        
On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual
    
'=====================================
' Clear Schedule
'=====================================
'    If Me.Range("no_of_schedules").value > 0 Then
'        Me.Range("schedule_start").Range("A1:F" & Me.Range("no_of_schedules").value).ClearContents
'    End If
'
'    If Me.Range("cpn_schedule_start").value > 0 Then
'        Me.Range("cpn_schedule_start").Range("A1:F" & Me.Range("no_of_coupon_schedule").value).ClearContents
'    End If
    
    no_of_ul = Me.Range("No_of_Underlying").value

''======================================
'' KI Barrier
''======================================
'    Me.Range("ki_barrier").Cells(1, 2).value = "=" & Me.Range("ki_barrier").Cells(1, 1).Address
'    Me.Range("ki_barrier").Cells(1, 3).value = "=" & Me.Range("ki_barrier").Cells(1, 1).Address
    

'======================================
' Date
'======================================

'    date_shift = shtMarket.Range("market_date").value - Me.Range("value_date").value
'    Me.Range("current_date").value = "=market_date"
'    Me.Range("value_date").value = "=market_date"
    
    Me.Range("dummy").value = "=max(" & Me.Range("schedule_start").Cells(1, 2).Address & ":" & Me.Range("schedule_start").Cells(20, 2).Address & ")"
    
'======================================
' Reference
'======================================
    For inx = 1 To no_of_ul
        If Left(Me.Range("ul_code").Cells(1, inx), 6) = "KOSPI2" Then
            Me.Range("reference").Cells(1, inx) = "=S"
        Else
            Me.Range("reference").Cells(1, inx) = "=S_" & Me.Range("ul_code").Cells(1, inx)
        End If
    Next inx
    
'======================================
' Fill Schedule
'======================================

'
'    For inx = 1 To Me.Range("no_of_schedules").value
'
'        Me.Range("schedule_start").Cells(inx, 0).value = inx
'        Me.Range("schedule_start").Cells(inx, 1).value = Me.Range("schedule_start").Cells(inx, 1).value + date_shift
'
'        If Me.Range("Monthly_Coupon_Flag").Cells(1, 2) = 0 And Me.Range("Range_Accrual_Flag").Cells(1, 2) = 0 Then
'
'                Me.Range("schedule_start").Cells(inx, 2).value = "=Coupon_P.A./12*Frequency_M*" & Me.Range("schedule_start").Cells(inx, 0).Address
'
'        Else
'
'            Me.Range("schedule_start").Cells(inx, 2).value = 0
'
'        End If
'
'    Next inx
'

'======================================
' Fill Coupon Schedule
'======================================
'    cpn_frequency = Me.Range("cpn_frequency").value
'
'    If Me.Range("Monthly_Coupon_Flag").Cells(1, 2) = 1 Then
'
'        For inx = 1 To Me.Range("no_of_coupon_schedule").value
'
'            Me.Range("cpn_schedule_start").Cells(inx, 1).value = Me.Range("cpn_schedule_start").Cells(inx, 1).value + date_shift
'            Me.Range("cpn_schedule_start").Cells(inx, 3).value = "=Coupon_P.A./12*cpn_frequency"
'
'        Next inx
'
'    End If
'
'    Me.Range("mid_day_greek").value = "N"
    
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:

    show_error
    Application.Calculation = xlCalculationAutomatic
    

End Sub
Public Sub cmdFillFormula_Click()

    Dim inx As Integer
    Dim jnx As Integer
    Dim Maturity As Double
    Dim frequency As Integer
    Dim coupon_pa As Double
    Dim cpn_frequency As Integer
    Dim no_of_ul As Integer
        
On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual
    
'=====================================
' Clear Schedule
'=====================================
    If Me.Range("no_of_schedules").value > 0 Then
        Me.Range("schedule_start").Range("A1:F" & Me.Range("no_of_schedules").value).ClearContents
    End If
    
    If Me.Range("cpn_schedule_start").value > 0 Then
        Me.Range("cpn_schedule_start").Range("A1:F" & Me.Range("no_of_coupon_schedule").value).ClearContents
    End If
    
    no_of_ul = Me.Range("No_of_Underlying").value

'======================================
' KI Barrier
'======================================
    Me.Range("ki_barrier").Cells(1, 2).value = "=" & Me.Range("ki_barrier").Cells(1, 1).Address
    Me.Range("ki_barrier").Cells(1, 3).value = "=" & Me.Range("ki_barrier").Cells(1, 1).Address
    
    Me.Range("put_strike").value = 1
    Me.Range("put_participation").value = 1
    Me.Range("Put_Add_CPN").value = 0
    Me.Range("ki_adj_pct").value = 1
    Me.Range("Call_Strike").value = 1
    Me.Range("Call_PR").value = 0
    Me.Range("floor_value").value = 0
    

'======================================
' Date
'======================================
    Me.Range("current_date").value = "=market_date"
    Me.Range("value_date").value = "=market_date"
    Me.Range("dummy").value = "=max(" & Me.Range("schedule_start").Cells(1, 2).Address & ":" & Me.Range("schedule_start").Cells(20, 2).Address & ")"
    
'======================================
' Reference
'======================================
    For inx = 1 To no_of_ul
        If Left(Me.Range("ul_code").Cells(1, inx), 6) = "KOSPI2" Then
            Me.Range("reference").Cells(1, inx) = "=S"
        Else
            Me.Range("reference").Cells(1, inx) = "=S_" & Me.Range("ul_code").Cells(1, inx)
        End If
    Next inx
    
'======================================
' Fill Schedule
'======================================
    Maturity = Me.Range("Maturity_y").value
    frequency = Me.Range("Frequency_M").value
    coupon_pa = Me.Range("Coupon_P.A.").value
    
    For inx = 1 To Int(Maturity * 12# / frequency)
    
        Me.Range("schedule_start").Cells(inx, 0).value = inx
        Me.Range("schedule_start").Cells(inx, 1).value = "=add_month(value_date," & Me.Range("schedule_start").Cells(inx, 0).Address & "*Frequency_M)-3"
        
        If Me.Range("Monthly_Coupon_Flag").Cells(1, 2) = 0 And Me.Range("Range_Accrual_Flag").Cells(1, 2) = 0 Then
            Me.Range("schedule_start").Cells(inx, 2).value = "=Coupon_P.A./12*Frequency_M*" & Me.Range("schedule_start").Cells(inx, 0).Address
            Me.Range("schedule_start").Cells(inx, 3).value = "=Coupon_P.A./12*Frequency_M*" & Me.Range("schedule_start").Cells(inx, 0).Address
            Me.Range("schedule_start").Cells(inx, 4).value = "=Coupon_P.A./12*Frequency_M*" & Me.Range("schedule_start").Cells(inx, 0).Address
        Else
            Me.Range("schedule_start").Cells(inx, 2).value = 0
            Me.Range("schedule_start").Cells(inx, 3).value = 0
            Me.Range("schedule_start").Cells(inx, 4).value = 0
        End If
    
    Next inx
    

'======================================
' Fill Coupon Schedule
'======================================
    cpn_frequency = Me.Range("cpn_frequency").value
    
    If Me.Range("Monthly_Coupon_Flag").Cells(1, 2) = 1 Then
    
        For inx = 1 To Int(Maturity * 12 / cpn_frequency)
        
            Me.Range("cpn_schedule_start").Cells(inx, 0).value = inx
            Me.Range("cpn_schedule_start").Cells(inx, 1).value = "=add_month(value_date," & Me.Range("cpn_schedule_start").Cells(inx, 0).Address & "*cpn_frequency) - 3"
            
            Me.Range("cpn_schedule_start").Cells(inx, 2).value = "=Monthly_Cpn"
            
            Me.Range("cpn_schedule_start").Cells(inx, 3).value = "=Coupon_P.A./12*cpn_frequency"
          
            For jnx = 1 To no_of_ul
            
                Me.Range("cpn_schedule_start").Cells(inx, 3 + jnx).value = "=" & Me.Range("reference").Cells(1, jnx).Address & " * " & Me.Range("cpn_schedule_start").Cells(inx, 2).Address
                
            Next jnx
        
        Next inx
        
    End If
    
    
    If Me.Range("Range_Accrual_Flag").Cells(1, 2) = 1 Then
    
        Me.Range("ra_cpn") = "=Coupon_P.A. / 12 * Frequency_M"
    
    End If
    
    Me.Range("mid_day_greek").value = "N"
    
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:

    show_error
    Application.Calculation = xlCalculationAutomatic
    

End Sub

Public Sub cmdLoadDeal_Click()


On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual
    
    clear_ac_deal
    
    cmd_load_ac_deal
    
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    show_error
    Application.Calculation = xlCalculationAutomatic


End Sub

'
'Private Sub cmdExpiry_Click()
'
'On Error GoTo ErrorHandler
'
'    Application.Calculation = xlCalculationManual
'
'    cmd_expiry
'
'    Application.Calculation = xlCalculationAutomatic
'
'    Exit Sub
'
'ErrorHandler:
'
'    show_error
'    Application.Calculation = xlCalculationAutomatic
'
'End Sub
'
'Private Sub cmdLoadDeal_Click()
'
'On Error GoTo ErrorHandler
'
'    Application.Calculation = xlCalculationManual
'
'    cmd_load_deal
'
'    Application.Calculation = xlCalculationAutomatic
'
'    Exit Sub
'
'ErrorHandler:
'
'    show_error
'    Application.Calculation = xlCalculationAutomatic
'
'End Sub
'
'
'
'Private Sub cmdNewDeal_Click()
'
'On Error GoTo ErrorHandler
'
'    Application.Calculation = xlCalculationManual
'
'    cmd_new_deal
'
'    Application.Calculation = xlCalculationAutomatic
'
'    Exit Sub
'
'ErrorHandler:
'
'    show_error
'    Application.Calculation = xlCalculationAutomatic
'
'End Sub
'
'
'Private Sub cmdRunMultiDeal_Click()
'
'
'On Error GoTo ErrorHandler
'
'    cmd_cal_price_multi_deals
'
'    Exit Sub
'
'ErrorHandler:
'
'    show_error
'
'End Sub
'
'Private Sub CommandButton1_Click()
'
'End Sub
'
'Private Sub cmdUnconfirm_Click()
'
'On Error GoTo ErrorHandler
'
'    cmd_unconfirm
'
'    Exit Sub
'
'ErrorHandler:
'
'    show_error
'
'
'End Sub
'
'Private Sub cmdUpdateDeal_Click()
'
'On Error GoTo ErrorHandler
'
'    cmd_update_deal
'
'    Exit Sub
'
'ErrorHandler:
'
'    show_error
'
'
'End Sub
'
'Private Sub Label1_Click()
'
'End Sub
Public Sub cmdNewDeal_Click()


On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual
      
    cmd_new_ac_deal
    
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    show_error
    Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmdShowFlushNode_Click()

On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual

   ' cmd_show_flush_node

     Application.Calculation = xlCalculationAutomatic

    Exit Sub
    
ErrorHandler:
    show_error
    Application.Calculation = xlCalculationAutomatic

End Sub

'Private Sub cmdSpotVol_Click()
'
'On Error GoTo ErrorHandler
'
'    Application.Calculation = xlCalculationManual
'
'    cmd_spot_vol_ac
'
'    Application.Calculation = xlCalculationAutomatic
'    Exit Sub
'
'ErrorHandler:
'    show_error
'    Application.Calculation = xlCalculationAutomatic
'
'
'End Sub

Public Sub cmdUnconfirm_Click()

On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual

    cmd_unconfirm_ac
    
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    show_error
    Application.Calculation = xlCalculationAutomatic

End Sub

Public Sub cmdUpdateDeal_Click()

On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationManual
      
    cmd_update_ac_deal
    
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    show_error
    Application.Calculation = xlCalculationAutomatic

End Sub
'
'Private Sub cmdLocalVega_Click()
'
'On Error GoTo ErrorHandler
'
'    Application.Calculation = xlCalculationManual
'
'    cmd_local_vega
'
'    Application.Calculation = xlCalculationAutomatic
'
'    Exit Sub
'
'ErrorHandler:
'    show_error
'    Application.Calculation = xlCalculationAutomatic
'
'End Sub

Private Sub CommandButton1_Click()

cmd_3d_closing_test

End Sub

Public Sub btnLoadDeal_Click()
    
    'Clear ranges : deal ticket
    shtACPricer.Range("fund_code_m").ClearContents
    shtACPricer.Range("fund_code_c").ClearContents
    shtACPricer.Range("No_of_Underlying").ClearContents
    shtACPricer.Range("ul_code").ClearContents
    shtACPricer.Range("ejected_ul_flag").ClearContents
    shtACPricer.Range("ejectable_flag").ClearContents
    
    shtACPricer.Range("value_date").ClearContents
    shtACPricer.Range("maturity_date").ClearContents
    shtACPricer.Range("settlement_date").ClearContents

    shtACPricer.Range("dummy").ClearContents
    
    shtACPricer.Range("ki_flag").ClearContents
    shtACPricer.Range("ki_touch_flag").ClearContents
    shtACPricer.Range("ki_barrier").ClearContents
    
    shtACPricer.Range("KI_Monitoring_Freq").ClearContents
    
    shtACPricer.Range("monthly_coupon_flag").ClearContents
    shtACPricer.Range("Early_Exit_Flag").ClearContents
    'shtACPricer.Range("Early_Exit_Touched").ClearContents
    
    shtACPricer.Range("notional").ClearContents
    shtACPricer.Range("CCY").ClearContents
    shtACPricer.Range("reference").ClearContents
    
    shtACPricer.Range("issue_cost").ClearContents
    shtACPricer.Range("strike_shift").ClearContents
    shtACPricer.Range("KI_Barrier_Shift").ClearContents
    
    shtACPricer.Range("performance_type").ClearContents
    shtACPricer.Range("ki_performance_type").ClearContents
    
    'Clear ranges: call schedule
    shtACPricer.Range("schedule_start").Range(Cells(1, 1), Cells(36, 1)).ClearContents
    shtACPricer.Range("cpn_on_call").Range(Cells(1, 1), Cells(36, 1)).ClearContents
    shtACPricer.Range("strike_rate").Range(Cells(1, 1), Cells(36, 1)).ClearContents
    shtACPricer.Range("performance_type_on_call").Range(Cells(1, 1), Cells(36, 1)).ClearContents
    shtACPricer.Range("strike_smoothing_width").Range(Cells(1, 1), Cells(36, 1)).ClearContents
    shtACPricer.Range("ejectable_order_on_call").Range(Cells(1, 1), Cells(36, 1)).ClearContents
    
    'Clear ranges: monthly coupon schedule
    shtACPricer.Range("cpn_schedule_start").Range(Cells(1, 1), Cells(100, 3)).ClearContents
    
    'Clear ranges: early exit schedule
    shtACPricer.Range("early_exit_schedule_start").Range(Cells(1, 1), Cells(22, 8)).ClearContents
    
    'Reset Autocall YN
    shtACPricer.Range("Autocall_YN").ClearContents
    With shtACPricer.Range("Autocall_YN").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
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
        .ScreenUpdating = False
        .StatusBar = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    'Set defulat values
    shtACPricer.Range("call_put") = "CALL"
    shtACPricer.Range("put_strike") = 1
    shtACPricer.Range("Put_Add_CPN") = 0
    shtACPricer.Range("ki_adj_pct") = 1
    shtACPricer.Range("Call_Strike") = 1
    shtACPricer.Range("put_participation") = 1
    'floor: 원금보장율
    If shtACPricer.Range("asset_code").value = "IO322714A77T" Then
        shtACPricer.Range("floor_value") = 1 'ELB
    Else
        shtACPricer.Range("floor_value") = 0
    End If
    shtACPricer.Range("Call_PR") = 0
    shtACPricer.Range("Range_Accrual_Flag") = False
    shtACPricer.Range("alive_yn") = "Y"
    shtACPricer.Range("confirmed_yn") = "Y"
    shtACPricer.Range("buy_sell") = "SELL"
    '발행사 신용스프레드
    shtACPricer.Range("Rate_Spread") = 0
    shtACPricer.Range("Hedge_Cost") = 0
    shtACPricer.Range("instrument_type") = "Note"
    shtACPricer.Range("performance_type") = "Worst"
    shtACPricer.Range("ki_performance_type") = "Worst"
    
    Dim scenario_test As Boolean
    scenario_test = shtMarket.Range("SCENARIO_ENABLE")
    
    Dim isQuote As Boolean
    isQuote = shtACPricer.Range("isQuote").value
    
    Dim i, j As Integer
    
    Dim sql As String
    
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    
    Dim today As Date
    today = shtACPricer.Range("current_date").value
    
    Dim tDayStr As String
    tDayStr = date2str(today)

    Dim adoCon As New adoDB.Connection
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
    Dim productCode As String
    productCode = shtACPricer.Range("asset_code").value
    
    If productCode = "IO322720A08T" Then
        productCode = "IO322720508T"
    End If
    Dim seq As Integer
    seq = shtACPricer.Range("D4")
        
    '---------- From Front DB
    sql = " select ASSET_CODE, ORIGIN_ASSET_CODE, FUND_CODE_M, FUND_CODE_C, NOTIONAL, CURRENT_NOTIONAL, VALUE_DATE, EXPIRY_DATE, SETTLEMENT_DATE, OPTION_TYPE, DUMMY_COUPON, KI_BARRIER_YN, KI_TOUCHED_YN, KI_BARRIER, KO_BARRIER_YN, KO_TOUCHED_YN, KO_BARRIER, PUT_STRIKE, PUT_PARTICIPATION, ISSUE_PRICE, RATE_SPREAD, STRIKE_SMOOTHING, KI_MONITORING_FREQ, ISSUE_COST, NOTE, LIVE_YN, CONFIRM_YN, UPDATE_DT, COST, KIBARRIER_SHIFT_SIZE, INSTRUMENT_TYPE, PERFORMANCE_TYPE, RA_YN, RA_CPN_PER_PERIOD, RA_TENOR, RA_MIN_PCT, RA_MAX_PCT, DEAD_DATE, ORIGIN_NOTIONAL, MINIMUM_PAYOFF, KI_PERFORMANCE_TYPE, NOTE_CURRENCY, COND_FEE, NOTE_CURVE, MAX_NOTIONAL from sps.ac_deal where asset_code = '" + productCode + "' "
    
    With oCmd
        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql
        
        oRS.Open .Execute
    End With
        
    i = 1
    Dim value_date As Date
    Dim maturity_date As Date
    Dim settlement_date As Date
    Dim dummy_coupon As Double
    Dim ki_flag As Boolean
    Dim ki_touch_flag As Boolean
    Dim ki_monitoring_freq As Integer
    Dim notional As Double
    Dim issue_cost As Double
    Dim strike_smoothing As Double
    Dim ki_barrier_shift As Double
    
    Do Until oRS.EOF
    
        If isQuote = True Then
            shtACPricer.Range("fund_code_m") = oRS("FUND_CODE_M")
            shtACPricer.Range("fund_code_c") = oRS("FUND_CODE_C")
            shtACPricer.Range("CCY") = oRS("NOTE_CURRENCY")
        End If
    
        'value date
        value_date = str2date(oRS("VALUE_DATE"))
        shtACPricer.Range("value_date").Cells(1, i) = value_date
        'maturity date
        maturity_date = str2date(oRS("EXPIRY_DATE"))
        shtACPricer.Range("maturity_date").Cells(1, i) = maturity_date
        'issue date
        settlement_date = str2date(oRS("SETTLEMENT_DATE"))
        shtACPricer.Range("settlement_date").Cells(1, i) = settlement_date
        
        'dummy coupon
        dummy_coupon = oRS("DUMMY_COUPON")
        shtACPricer.Range("dummy").Cells(1, i) = dummy_coupon
        
        'KI barrier flag
        If (oRS("KI_BARRIER_YN") = "Y") Then
            ki_flag = True
        Else
            ki_flag = False
        End If
        shtACPricer.Range("ki_flag").Cells(1, i) = ki_flag
            
        'KI touched flag
        If (oRS("KI_TOUCHED_YN") = "Y") Then
            ki_touch_flag = True
        Else
            ki_touch_flag = False
        End If
        shtACPricer.Range("ki_touch_flag").Cells(1, i) = ki_touch_flag
        
        'KI monitoring freq.
        ki_monitoring_freq = oRS("KI_MONITORING_FREQ")
        shtACPricer.Range("KI_Monitoring_Freq").Cells(1, i) = ki_monitoring_freq
        
        'notional
        If oRS("NOTIONAL") = 0 Then
            notional = 1
        Else
            notional = Abs(oRS("NOTIONAL"))
        End If
        
'        If productCode = "IO322720A08T" Or productCode = "IO322720508T" Then
'            notional = -1 * shtELS.Range("eswap_qty_col").Cells(seq + 9, 1) * shtELS.Range("els_facevalue_col").Cells(seq + 9, 1)
'        End If
        
        shtACPricer.Range("notional").Cells(1, i) = notional
        
        'issue cost
        issue_cost = oRS("ISSUE_COST")
        shtACPricer.Range("issue_cost").Cells(1, i) = issue_cost
        
        'strike smoothing
        If shtACPricer.Range("chkStrikeSmoothing") = True Then
            If IsNull(oRS("STRIKE_SMOOTHING")) Then
                strike_smoothing = 0.02
            Else
                strike_smoothing = oRS("STRIKE_SMOOTHING")
            End If
            shtACPricer.Range("strike_shift").Cells(1, i) = strike_smoothing
        Else
            shtACPricer.Range("strike_shift").Cells(1, i) = 0
        End If
        
        'KI barrier shift
        ki_barrier_shift = oRS("KIBARRIER_SHIFT_SIZE")
        shtACPricer.Range("KI_Barrier_Shift").Cells(1, i) = ki_barrier_shift
        
        i = i + 1
        oRS.MoveNext
    Loop
    
    oRS.Close
    
    '---------- From BizOne
    '기초자산 코드, 최초기준가, KI수준
    If isQuote = True Then
        sql = " SELECT ASSET_CODE, UL_CODE, REFERENCE_PRICE, UPDATE_DT, KI_BARRIER, SORT_SEQ, UL_COST_YIELD from sps.ac_underlying where asset_code = '" + productCode + "' order by 6"
    Else
        sql = " SELECT indv_iscd, decode(unas_iscd,'NIKKEI225','NKY',unas_iscd), unas_intl_prc, clrd_sdrt/100, barr_val/100 from BSYS.TBSIMO202D00@GDW where indv_iscd = '" + productCode + "' order by 2"
    End If
        
    With oCmd
        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql
       
        oRS.Open .Execute
    End With
        
    i = 1
    Dim redeem_shift As Boolean
    redeem_shift = False
    
    Dim ua_code As String
    Dim ua_ref_spot() As Double
    Dim ua_ki_barr() As Double
    Dim ua_close_spot() As Double
    ReDim ua_ref_spot(0) As Double
    ReDim ua_ki_barr(0) As Double
    ReDim ua_close_spot(0) As Double
    
    Do Until oRS.EOF
    
        '단축코드 -> KR코드 변환 2019.3.5
        If isQuote = True Then
            Select Case oRS(1)
            Case "005380": ua_code = "KR7005380001" '현대차
            Case "005930": ua_code = "KR7005930003" '삼성전자
            Case "000030": ua_code = "KR7000030007" '우리은행
            Case "028260": ua_code = "KR7028260008" '삼성물산
            Case "105560": ua_code = "KR7105560007" 'KB금융
            Case "035420": ua_code = "KR7035420009" 'NAVER
            Case "018260": ua_code = "KR7018260000" '삼성SDS
            Case "005490": ua_code = "KR7005490008" 'POSCO
            Case "034220": ua_code = "KR7034220004" 'LG디스플레이
            Case "D02002": ua_code = "KRD020021147" 'KOSPI200 레버리지
            Case Else: ua_code = oRS(1)
            End Select
        Else
            ua_code = oRS(1)
        End If
        
        'underlying code
        shtACPricer.Range("ul_code").Cells(1, i) = ua_code
        'ejectable
        shtACPricer.Range("ejected_ul_flag").Cells(1, i) = False '기본값. 실제 상품 발행시 수정 필요.
        
        ReDim Preserve ua_ref_spot(UBound(ua_ref_spot) + 1) As Double
        ReDim Preserve ua_ki_barr(UBound(ua_ki_barr) + 1) As Double
        ReDim Preserve ua_close_spot(UBound(ua_close_spot) + 1) As Double
        
        ua_close_spot(i) = WorksheetFunction.VLookup(ua_code, Range("ua_close_spot"), 2, False)
        
        'ref. spot: 신규발행종목 최초기준가 100으로 설정
        '스트레스테스트 일 경우는 제외 2018.12.20
        'If (scenario_test = False) And (today = value_date Or isQuote = True) Then '평가일이 초초기준가설정일일 때
        If (scenario_test = False) And (today <= value_date Or isQuote = True) Then '평가일이 최초기준가설정일 이전일 때 포함(예. 발행풀가 기간을 피하기 위해 발행일 < 최초기준가평가일 인 경우)
            ua_ref_spot(i) = ua_close_spot(i)
        Else
            ua_ref_spot(i) = oRS(2)
            '최초기준가 검증(발행일에만)
            If today = settlement_date And check_spot(oRS(1), date2str(value_date), oRS(2), adoCon) = False Then
                MsgBox productCode & ": " & oRS(1) & " 기초자산최초기준가 오류"
            End If
        End If
        shtACPricer.Range("reference").Cells(1, i) = ua_ref_spot(i)
        
        'KI barrier
        ua_ki_barr(i) = oRS(4)
        shtACPricer.Range("KI_barrier").Cells(1, i) = ua_ki_barr(i)
        
        If oRS(1) = "SX5E" Then redeem_shift = True
        If oRS(1) = "SPX" Then redeem_shift = True
        
        i = i + 1
        oRS.MoveNext
    Loop
    
    'no of underlying
    Dim no_of_underlying As Integer
    no_of_underlying = i - 1
    shtACPricer.Range("no_of_underlying") = no_of_underlying
    
    If redeem_shift Then
        maturity_date = maturity_date + 1
        shtACPricer.Range("maturity_date") = maturity_date
    End If
    
    oRS.Close
    
    '----------------------------------------
    'performance type, fund_code, ccy
    
    Dim early_exit_flag As Boolean
    Dim isFiveWins As Boolean
    Dim ejectable_flag As Boolean
        
    If isQuote = True Then
    

    Else
        sql = " SELECT UNAS_CHOC_MTHD_CODE, CLRD_TYPE_CODE, substr(PROD_FNCD,1,2), PROD_FNCD, STLM_CRCD, DEAL_CLS_CODE " _
            & " FROM    bsys.TBSIMO201M00@gdw " _
            & " WHERE   INDV_ISCD = '" + productCode + "' "
        
        With oCmd
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = sql
            
            oRS.Open .Execute
        End With

        isFiveWins = False
        
        Do Until oRS.EOF
            
            Select Case oRS(0)
            Case "1"
                shtACPricer.Range("performance_type") = "Worst"
                shtACPricer.Range("ki_performance_type") = "Worst"
            Case "2"
                shtACPricer.Range("performance_type") = "Best"
                shtACPricer.Range("ki_performance_type") = "Best"
            Case "4"
                shtACPricer.Range("performance_type") = "Average"
                shtACPricer.Range("ki_performance_type") = "Average"
            Case "5"
                shtACPricer.Range("performance_type") = "Average"
                shtACPricer.Range("ki_performance_type") = "Worst"
            End Select
            
            If oRS(1) = "29" Then
                early_exit_flag = True
                shtACPricer.Range("early_exit_flag") = True
            Else
                early_exit_flag = False
                shtACPricer.Range("early_exit_flag") = False
            End If
            
            If oRS(1) = "38" Then
                isFiveWins = True
            Else
                isFiveWins = False
            End If
            
            If False Then '기본값. 실제 상품 발생시 수정 필요.
                ejectable_flag = True
                shtACPricer.Range("ejectable_flag") = True
            Else
                ejectable_flag = False
                shtACPricer.Range("ejectable_flag") = False
            End If
            
            shtACPricer.Range("fund_code_m") = oRS(2)
            shtACPricer.Range("fund_code_c") = oRS(3)
            shtACPricer.Range("CCY") = oRS(4)
            If oRS(5) = "2" Then
                shtACPricer.Range("buy_sell") = "BUY"
            Else
                shtACPricer.Range("buy_sell") = "SELL"
            End If
            
            oRS.MoveNext
        Loop
        
        oRS.Close
    End If
        
    '----------------------------------------
    'Call schedule
    Dim no_of_ac_schedule As Integer
    Dim call_date() As Date
    Dim coupon_on_call() As Double
    Dim barr_strike() As Double
    Dim barr_perform_type() As Integer
    ReDim call_date(0) As Date
    ReDim coupon_on_call(0) As Double
    ReDim barr_strike(0) As Double
    ReDim barr_perform_type(0) As Integer
        
    If isQuote = True Then
        sql = " SELECT  CALL_DATE   " _
            & "        ,null        " _
            & "    ,NVL(coupon_on_call,0)  " _
            & "    ,NVL(strike,0)      " _
            & "    ,NVL(performance_type,-1)      " _
            & "    ,NVL(strike_smoothing_width,0) " _
            & "FROM    sps.ac_schedule " _
            & "WHERE   asset_code = '" + productCode + "'" _
            & "ORDER BY 2 "
            
        With oCmd
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = sql
            
            oRS.Open .Execute
        End With
        
        i = 1
        Do Until oRS.EOF
            
            ReDim Preserve call_date(UBound(call_date) + 1) As Date
            ReDim Preserve coupon_on_call(UBound(coupon_on_call) + 1) As Double
            ReDim Preserve barr_strike(UBound(barr_strike) + 1) As Double
            ReDim Preserve barr_perform_type(UBound(barr_perform_type) + 1) As Integer
            
            If redeem_shift Then
                '기초자산 지역에 따라 상환스케줄 1차 조정: 실제 평가일 + 1달력일
                call_date(i) = str2date(oRS(0)) + 1
            Else
                call_date(i) = str2date(oRS(0))
            End If
            shtACPricer.Range("schedule_start").Cells(i, 1) = call_date(i)
            
            coupon_on_call(i) = oRS(2)
            barr_strike(i) = oRS(3)
            
            shtACPricer.Range("cpn_on_call").Cells(i, 1) = coupon_on_call(i)
            shtACPricer.Range("strike_rate").Cells(i, 1) = barr_strike(i)
            
            'performance_type_on_call
            barr_perform_type(i) = oRS(4)
            shtACPricer.Range("performance_type_on_call").Cells(i, 1) = barr_perform_type(i)
            
            'strike smoothing
            If shtACPricer.Range("chkStrikeSmoothing") = True Then
                shtACPricer.Range("strike_smoothing_width").Cells(i, 1) = oRS(5)
            Else
                shtACPricer.Range("strike_smoothing_width").Cells(i, 1) = 0
            End If
            
            'ejectable
            If shtACPricer.Range("ejectable_flag") = True Then
                shtACPricer.Range("ejectable_order_on_call").Cells(i, 1) = 0 '기본값. 실제 상품 발행시 수정 필요.
            Else
                shtACPricer.Range("ejectable_order_on_call").Cells(i, 1) = 0
            End If
            
            i = i + 1
            oRS.MoveNext
        Loop
        
        oRS.Close
            
    Else
    
        If isFiveWins Then
            sql = " SELECT CLRD_DTRM_DATE, null, "
            sql = sql & "                 NVL(CLRD_INRT,0)/100 CLRD_ERT,"
            sql = sql & "                 NVL(CLRD_BARR_VAL,0)/100 UNAS_SDRT,"
            sql = sql & "                'N' AVRG_APLY_YN"
            sql = sql & "         FROM   BSYS.TBSIMO227L00@GDW"
            sql = sql & "         WHERE  INDV_ISCD = '" + productCode + "'"
            sql = sql & " ORDER BY CLRD_DTRM_DATE"
    '        SQL = " SELECT CLRD_DTRM_DATE, null, "
    '        SQL = SQL & "                NVL(CLRD_INRT,0)/100 CLRD_ERT, "
    '        SQL = SQL & "                NVL(CLRD_BARR_VAL,0)/100 UNAS_SDRT, "
    '        SQL = SQL & "                'N' AVRG_APLY_YN, "
    '        SQL = SQL & "                NVL(b.strike_smoothing_width,0) "
    '        SQL = SQL & "         FROM   BSYS.TBSIMO227L00@GDW a, sps.ac_schedule b "
    '        SQL = SQL & "         WHERE  a.INDV_ISCD = '" + productCode + "' and a.INDV_ISCD = b.ASSET_CODE(+) "
    '        SQL = SQL & " and a.CLRD_DTRM_DATE = b.call_date(+) "
    '        SQL = SQL & " ORDER BY 1 "
        Else
            sql = " SELECT  TRTH_CLRD_DTRM_DATE   DT  " _
                & "        ,RMS.GET_WORKDATE(TRTH_CLRD_DTRM_DATE,1)   BFRPY_BASE_DT " _
                & "    ,NVL(CLRD_ERT,0) / 100       BFRPY_BASE_RT " _
                & "    ,NVL(UNAS_SDRT1,0)/100     BASERT1 " _
                & "    ,NVL(AVRG_APLY_YN,'N')     AVRG_APLY_YN " _
                & "FROM    BSYS.TBSIMO203D00@gdw " _
                & "WHERE   INDV_ISCD = '" + productCode + "' " _
                & "ORDER BY CLRD_DTRM_DATE "
    '        SQL = " SELECT  TRTH_CLRD_DTRM_DATE   DT  " _
    '            & "        ,RMS.GET_WORKDATE(TRTH_CLRD_DTRM_DATE,1)   BFRPY_BASE_DT " _
    '            & "    ,NVL(CLRD_ERT,0) / 100       BFRPY_BASE_RT " _
    '            & "    ,NVL(UNAS_SDRT1,0)/100     BASERT1 " _
    '            & "    ,NVL(AVRG_APLY_YN,'N')     AVRG_APLY_YN " _
    '            & "    ,NVL(b.strike_smoothing_width,0) " _
    '            & "FROM    bsys.TBSIMO203d00@gdw a, sps.ac_schedule b " _
    '            & "WHERE   a.INDV_ISCD = '" + productCode + "' and a.INDV_ISCD = b.ASSET_CODE(+) " _
    '            & "and TRTH_CLRD_DTRM_DATE = b.call_date(+) " _
    '            & "ORDER BY 1 "
        End If
    
        With oCmd
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = sql
            
            oRS.Open .Execute
        End With
        
        i = 1
        Do Until oRS.EOF
            
            ReDim Preserve call_date(UBound(call_date) + 1) As Date
            ReDim Preserve coupon_on_call(UBound(coupon_on_call) + 1) As Double
            ReDim Preserve barr_strike(UBound(barr_strike) + 1) As Double
            ReDim Preserve barr_perform_type(UBound(barr_perform_type) + 1) As Integer
            
            If redeem_shift Then
                '2019.4.18
                '오늘이 실제 조기상환평가일 후 첫 영업일이면 call date를 오늘로 설정: 상환여부 판단 때문
                If today = WorksheetFunction.WorkDay(str2date(oRS(0)), 1, shtHolidays.Range("Holidays")) Then
                    call_date(i) = today
                Else
                    '기초자산 지역에 따라 상환스케줄 1차 조정: 실제 평가일 + 1달력일
                    call_date(i) = str2date(oRS(0)) + 1
                End If
            Else
                call_date(i) = str2date(oRS(0))
            End If
            shtACPricer.Range("schedule_start").Cells(i, 1) = call_date(i)
            coupon_on_call(i) = oRS(2)
            barr_strike(i) = oRS(3)
            shtACPricer.Range("cpn_on_call").Cells(i, 1) = coupon_on_call(i)
            shtACPricer.Range("strike_rate").Cells(i, 1) = barr_strike(i)
            
            'performance_type_on_call
            If oRS(4) = "Y" Then
                barr_perform_type(i) = 0
            Else
                barr_perform_type(i) = -1
            End If
            shtACPricer.Range("performance_type_on_call").Cells(i, 1) = barr_perform_type(i)
            
            ''strike smoothing
            'shtACPricer.Range("strike_smoothing_width").Cells(i, 1) = oRS(5)
                        
            'ejectable
            If shtACPricer.Range("ejectable_flag") = True Then
                shtACPricer.Range("ejectable_order_on_call").Cells(i, 1) = 0 '기본값. 실제 상품 발행시 수정 필요.
            Else
                shtACPricer.Range("ejectable_order_on_call").Cells(i, 1) = 0
            End If
            
            i = i + 1
            oRS.MoveNext
        Loop
        
        oRS.Close
    
        '----------------------------------------
        'Strike smoothing factor: sps DB의 call date를 shift할 경우 비즈원 call date와 join되지 않는 문제가 있어 부득이 독립적으로 입수함
        sql = "SELECT call_date, strike_smoothing_width from sps.ac_schedule where ASSET_CODE = '" + productCode + "' and call_date >='" + date2str(call_date(1) - 1) + "' ORDER BY 1 "
            
        With oCmd
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = sql
            
            oRS.Open .Execute
        End With
        
        i = 1
        Do Until oRS.EOF
            
            'strike smoothing
            If shtACPricer.Range("chkStrikeSmoothing") = True Then
                shtACPricer.Range("strike_smoothing_width").Cells(i, 1) = oRS(1)
            Else
                shtACPricer.Range("strike_smoothing_width").Cells(i, 1) = 0
            End If
            
            i = i + 1
            oRS.MoveNext
        Loop
        
        oRS.Close
    
    End If
    no_of_ac_schedule = i - 1
    
    '----------------------------------------
    'Monthly coupon schedule
    Dim monthly_coupon_flag As Boolean
    monthly_coupon_flag = False
    
    Dim coupon_date() As Date
    Dim montly_coupon() As Double
    Dim coupon_barr() As Double
    ReDim coupon_date(0) As Date
    ReDim montly_coupon(0) As Double
    ReDim coupon_barr(0) As Double
    
    Dim no_of_monthly_coupon_schedule As Integer
    
    If isQuote = True Then
    
    Else
        
        '쿠폰구간관리
        sql = " SELECT NVL(BONS_CUPN_STA_SDRT/100, 0) CouponBarrier, " _
            & "         NVL(BONS_CUPN_FIN_SDRT/100, 100) CouponUpperBarrier, " _
            & "         NVL(BONS_CUPN_INRT/100, 0) CouponRate, SCTN_STA_CTNU, SCTN_FIN_CTNU " _
            & " FROM    bsys.TBSIMO210L00@gdw " _
            & " WHERE   INDV_ISCD = '" + productCode + "' ORDER BY SCTN_STA_CTNU ASC"
        
        With oCmd
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = sql
            
            oRS.Open .Execute
        End With
            
        i = 1
        Do Until oRS.EOF

'            ReDim Preserve coupon_barr(UBound(coupon_barr) + 1) As Double
'            ReDim Preserve montly_coupon(UBound(montly_coupon) + 1) As Double
            
            For i = oRS(3) To oRS(4)
                        ReDim Preserve coupon_barr(UBound(coupon_barr) + 1) As Double
            ReDim Preserve montly_coupon(UBound(montly_coupon) + 1) As Double
                'coupon barrier
                coupon_barr(i) = oRS(0)
                shtACPricer.Range("cpn_schedule_start").Cells(i, 2) = coupon_barr(i)
                'coupon on call
                montly_coupon(i) = oRS(2)
                shtACPricer.Range("cpn_schedule_start").Cells(i, 3) = montly_coupon(i)
            Next i
            
            oRS.MoveNext
        
        Loop
        
        oRS.Close
        
        '쿠폰지급일관리
        sql = " SELECT TRTH_CUPN_VLTN_DATE COUPON_DATE " _
            & " FROM    BSYS.TBSIMO213L00@GDW   " _
            & " WHERE   INDV_ISCD = '" + productCode + "' ORDER BY 1 ASC"
            
        With oCmd
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = sql
            
            oRS.Open .Execute
        End With
            
        i = 1
        'coupon date
        Do Until oRS.EOF
            monthly_coupon_flag = True
            ReDim coupon_date(UBound(coupon_date) + 1) As Date
            
            If redeem_shift Then
                '2019.4.18
                '오늘이 실제 조기상환평가일 후 첫 영업일이면 call date를 오늘로 설정: 상환여부 판단 때문
                If today = WorksheetFunction.WorkDay(str2date(oRS(0)), 1, shtHolidays.Range("Holidays")) Then
                    coupon_date(i) = today
                Else
                    '기초자산 지역에 따라 상환스케줄 1차 조정: 실제 평가일 + 1달력일
                    coupon_date(i) = str2date(oRS(0)) + 1
                End If
            Else
                coupon_date(i) = str2date(oRS(0))
            End If
            shtACPricer.Range("cpn_schedule_start").Cells(i, 1) = coupon_date(i)
            
            i = i + 1
            oRS.MoveNext
        
        Loop
        
        oRS.Close
    
    End If
    
    'Monthly coupon flag
    shtACPricer.Range("monthly_coupon_flag") = monthly_coupon_flag
    no_of_monthly_coupon_schedule = i - 1

    '----------------------------------------
    'Early Exit(Lizard)
'    Dim isEEchecked As Boolean
'    isEEchecked = False

    Dim no_of_ee_schedule As Integer
    
    Dim ee_call_date() As Date
    Dim ee_coupon_on_call() As Double
    Dim ee_barr_strike() As Double
    Dim ee_barr_perform_type() As Integer
    Dim ee_barr_type() As Integer
    Dim ee_barr_touched() As Boolean
    ReDim ee_call_date(0) As Date
    ReDim ee_coupon_on_call(0) As Double
    ReDim ee_barr_strike(0) As Double
    ReDim ee_barr_perform_type(0) As Integer
    ReDim ee_barr_type(0) As Integer
    ReDim ee_barr_touched(0) As Boolean
    
    If isQuote = True Then
    
        early_exit_flag = False
    
        sql = " SELECT * FROM ("
        sql = sql & " SELECT CALL_DATE,"
        sql = sql & "        COUPON_ON_CALL,"
        sql = sql & "        STRIKE,"
        sql = sql & "        STRIKE_SMOOTHING_WIDTH,"
        sql = sql & "        EE_TOUCHED_YN,"
        sql = sql & "        PERFORMANCE_TYPE,"
        sql = sql & "        BARRIER_TYPE,"
        sql = sql & "        ROW_NUMBER() OVER(PARTITION BY CALL_DATE ORDER BY UL_CODE) RNUM"
        sql = sql & " FROM SPS.AC_EE_SCHEDULE"
        sql = sql & " WHERE  ASSET_CODE='" + productCode + "')"
        sql = sql & " WHERE RNUM=1 "
        
        With oCmd
            .ActiveConnection = adoCon
            .CommandType = adCmdText
            .CommandText = sql
            
            oRS.Open .Execute
        End With
        
        i = 1
        Do Until oRS.EOF
        
            early_exit_flag = True
            ReDim Preserve ee_call_date(UBound(ee_call_date) + 1) As Date
            ReDim Preserve ee_coupon_on_call(UBound(ee_coupon_on_call) + 1) As Double
            ReDim Preserve ee_barr_strike(UBound(ee_barr_strike) + 1) As Double
            ReDim Preserve ee_barr_perform_type(UBound(ee_barr_perform_type) + 1) As Integer
            ReDim Preserve ee_barr_type(UBound(ee_barr_type) + 1) As Integer
            ReDim Preserve ee_barr_touched(UBound(ee_barr_touched) + 1) As Boolean
            
            'early exit schedule: call date, end date
            If redeem_shift Then
               '2019.4.18
                '오늘이 실제 조기상환평가일 후 첫 영업일이면 call date를 오늘로 설정: 상환여부 판단 때문
                If today = WorksheetFunction.WorkDay(str2date(oRS(0)), 1, shtHolidays.Range("Holidays")) Then
                    ee_call_date(i) = today
                Else
                    '기초자산 지역에 따라 상환스케줄 1차 조정: 실제 평가일 + 1달력일
                    ee_call_date(i) = str2date(oRS(0)) + 1
                End If
            Else
                ee_call_date(i) = str2date(oRS(0))
            End If
            shtACPricer.Range("early_exit_schedule_start").Cells(i, 1) = ee_call_date(i)
            shtACPricer.Range("early_exit_schedule_start").Cells(i, 4) = ee_call_date(i)
                
            'early exit schedule: coupon on call
            ee_coupon_on_call(i) = oRS(1)
            shtACPricer.Range("early_exit_schedule_start").Cells(i, 2) = ee_coupon_on_call(i)
            'early exit schedule: early exit barrier
            ee_barr_strike(i) = oRS(2)
            shtACPricer.Range("early_exit_schedule_start").Cells(i, 5) = ee_barr_strike(i)
            'early exit schedule: start date
            shtACPricer.Range("early_exit_schedule_start").Cells(i, 3) = settlement_date
            
            'early exit schedule: early exit touched
            If oRS(4) = 0 Then
                ee_barr_touched(i) = False
                shtACPricer.Range("early_exit_schedule_start").Cells(i, 6) = "N"
            Else
                ee_barr_touched(i) = True
                shtACPricer.Range("early_exit_schedule_start").Cells(i, 6) = "Y"
            End If
            ee_barr_perform_type(i) = oRS(5)
            ee_barr_type(i) = oRS(6)
            shtACPricer.Range("early_exit_schedule_start").Cells(i, 7) = ee_barr_perform_type(i)
            shtACPricer.Range("early_exit_schedule_start").Cells(i, 8) = ee_barr_type(i)

            '다음 early exit barrier touch 여부: BizOne 입력값 기준
'            If isEEchecked = False And today <= shtACPricer.Range("early_exit_schedule_start").Cells(i, 1) Then
'
'                For j = 1 To shtACPricer.Range("no_of_underlying")
'                    shtACPricer.Range("EE_Barrier").Cells(1, j) = oRS(2)
'                Next j
'
'                If oRS(4) = 0 Then
'                    shtACPricer.Range("Early_Exit_Touched") = False
'                Else
'                    shtACPricer.Range("Early_Exit_Touched") = True
'                End If
'
'                isEEchecked = True
'            End If
            
            i = i + 1
            oRS.MoveNext
        Loop
        
        oRS.Close
        
        shtACPricer.Range("early_exit_flag") = early_exit_flag
        
    Else
        
        'Early Exit Touched
        If early_exit_flag = True Then
            sql = " SELECT CALC_FIN_DATE EarlyExitDate, " _
                & "         CLRD_INRT/100 EarlyExitCouponRate, " _
                & "         CLRD_BARR_VAL/100 EarlyExitBarrier, " _
                & "         CALC_STA_DATE, " _
                & "         CALC_FIN_DATE, " _
                & "         CLRD_BARR_HIT_YN " _
                & " FROM    bsys.TBSIMO227L00@gdw " _
                & " WHERE   INDV_ISCD = '" + productCode + "' "
            
            With oCmd
                .ActiveConnection = adoCon
                .CommandType = adCmdText
                .CommandText = sql
                
                oRS.Open .Execute
            End With
            
            i = 1
            Do Until oRS.EOF
                
                ReDim Preserve ee_call_date(UBound(ee_call_date) + 1) As Date
                ReDim Preserve ee_coupon_on_call(UBound(ee_coupon_on_call) + 1) As Double
                ReDim Preserve ee_barr_strike(UBound(ee_barr_strike) + 1) As Double
                ReDim Preserve ee_barr_perform_type(UBound(ee_barr_perform_type) + 1) As Integer
                ReDim Preserve ee_barr_type(UBound(ee_barr_type) + 1) As Integer
                ReDim Preserve ee_barr_touched(UBound(ee_barr_touched) + 1) As Boolean
                'early exit schedule: call date, end date
                If redeem_shift Then
                    '2019.4.18
                    '오늘이 실제 조기상환평가일 후 첫 영업일이면 call date를 오늘로 설정: 상환여부 판단 때문
                    If today = WorksheetFunction.WorkDay(str2date(oRS(0)), 1, shtHolidays.Range("Holidays")) Then
                        ee_call_date(i) = today
                    Else
                        '기초자산 지역에 따라 상환스케줄 1차 조정: 실제 평가일 + 1달력일
                        ee_call_date(i) = str2date(oRS(0)) + 1
                    End If
                Else
                    ee_call_date(i) = str2date(oRS(0))
                End If
                shtACPricer.Range("early_exit_schedule_start").Cells(i, 1) = ee_call_date(i)
                shtACPricer.Range("early_exit_schedule_start").Cells(i, 4) = ee_call_date(i)
            
                'early exit schedule: coupon on call
                ee_coupon_on_call(i) = oRS(1)
                shtACPricer.Range("early_exit_schedule_start").Cells(i, 2) = ee_coupon_on_call(i)
                'early exit schedule: early exit barrier
                ee_barr_strike(i) = oRS(2)
                shtACPricer.Range("early_exit_schedule_start").Cells(i, 5) = ee_barr_strike(i)
                'early exit schedule: start date
                shtACPricer.Range("early_exit_schedule_start").Cells(i, 3) = str2date(oRS(3))
                'early exit schedule: early exit touched
                If oRS(5) = "N" Then
                    ee_barr_touched(i) = False
                Else
                    ee_barr_touched(i) = True
                End If
                shtACPricer.Range("early_exit_schedule_start").Cells(i, 6) = oRS(5)
                
                'Hero 구조 원장 완성 전까지 기존 Early-exit구조 처리를 위한 임시 하드 코딩: dll(2018.8.8)
                ee_barr_perform_type(i) = -1
                ee_barr_type(i) = -1
                shtACPricer.Range("early_exit_schedule_start").Cells(i, 7) = ee_barr_perform_type(i)
                shtACPricer.Range("early_exit_schedule_start").Cells(i, 8) = ee_barr_type(i)
    
                '다음 early exit barrier touch 여부: BizOne 입력값 기준
'                If isEEchecked = False And today <= shtACPricer.Range("early_exit_schedule_start").Cells(i, 1) Then
'
'                    For j = 1 To shtACPricer.Range("no_of_underlying")
'                        shtACPricer.Range("EE_Barrier").Cells(1, j) = oRS(2)
'                    Next j
'
'                    If oRS(5) = "N" Then
'                        shtACPricer.Range("Early_Exit_Touched") = False
'                    Else
'                        shtACPricer.Range("Early_Exit_Touched") = True
'                    End If
'
'                    isEEchecked = True
'                End If
                
                i = i + 1
                oRS.MoveNext
            Loop
            
            oRS.Close
            
        End If
    
    End If

    no_of_ee_schedule = i - 1
            
    Call disconnectDB(adoCon)
    
    'Clear results
    clear_ac_greeks
    shtACPricer.Range("price").Cells(2, 1) = 0
    
    '----------------------------------------
    'Check KI for today
    If ki_flag = True Then
        For j = 1 To no_of_underlying
            If ua_close_spot(j) / ua_ref_spot(j) < ua_ki_barr(j) Then
            
                If ki_touch_flag = False Then
                    If scenario_test = False Then
                        MsgBox productCode & ": " & "당일 KI 터지했으나 원장에 반영 안됨"
                    Else
                        ki_touch_flag = True
                        shtACPricer.Range("ki_touch_flag") = True
                    End If
                End If
        
                Exit For
            End If
        Next j
    End If
    
    'Check EE KI for today
    If early_exit_flag = True Then
        For i = 1 To no_of_ee_schedule
            If today <= ee_call_date(i) Then
                For j = 1 To no_of_underlying
                    
                    If ua_close_spot(j) / ua_ref_spot(j) < ee_barr_strike(i) Then
                    
                        If ee_barr_touched(i) = False Then
                            If scenario_test = False Then
                                MsgBox productCode & ": " & "당일 early exit barrer 터지했으나 원장에 반영 안됨"
                            Else
                                ee_barr_touched(i) = True
                                shtACPricer.Range("early_exit_schedule_start").Cells(i, 6) = "Y"
                            End If
                        End If
    
                        Exit For
                    End If
                Next j
            End If
        Next i
    End If
    
    Dim redeem_rate As Double
    redeem_rate = 0
    
    'Check monthly coupon
    Dim monthly_cpn As Double
    monthly_cpn = 0
    If monthly_coupon_flag = True Then
        For i = 1 To no_of_monthly_coupon_schedule
            If today = coupon_date(i) Then
            
                If check_monthly_cpn(coupon_barr(i)) = True Then
                    monthly_cpn = montly_coupon(i)
                End If
                
            End If
        Next i
    End If
      
    'Check early-exit coupon
    Dim ee_success As Boolean
    ee_success = False
    Dim ee_cpn As Double
    ee_cpn = 0
    If early_exit_flag = True Then
        For i = 1 To no_of_ee_schedule
            If today = ee_call_date(i) And ee_barr_touched(i) = False Then
                ee_success = True
                ee_cpn = ee_coupon_on_call(i)
                Exit For
            End If
        Next i
    End If

    'Check autocall dates for today
    Dim ac_success As Boolean
    ac_success = False
    For i = 1 To no_of_ac_schedule
            
        If today = call_date(i) Then
            
            ac_success = check_autocall(barr_strike(i), barr_perform_type(i))
            
            If ac_success = True Then
                'autocall condition
                redeem_rate = monthly_cpn + coupon_on_call(i)
                Exit For
            Else
                'early-exit condition
                If ee_success = True Then
                    ac_success = True
                    redeem_rate = monthly_cpn + ee_cpn
                    Exit For
                End If
            
                'dummy coupon condition
                If ki_touch_flag = False And today = maturity_date Then
                    ac_success = True
                    redeem_rate = dummy_coupon
                    Exit For
                End If
            End If

        End If

    Next i
           
    If redeem_shift Then
    
        Dim ac_fail_confirm As Boolean
        ac_fail_confirm = False
        Dim ee_fail_confirm As Boolean
        ee_fail_confirm = False
        
        'call schedule
        For i = 1 To no_of_ac_schedule
            If today = call_date(i) - 1 Then
            
                ac_fail_confirm = check_autocall_fail(barr_strike(i), barr_perform_type(i))

                Exit For
            End If
        Next i
        
        'early-exit schedule
        If early_exit_flag = True Then
            Dim isEEdate As Boolean
            isEEdate = False
        
            For i = 1 To no_of_ee_schedule
                If today = ee_call_date(i) - 1 Then
                
                    isEEdate = True
                    ee_fail_confirm = ee_barr_touched(i) = True Or _
                                      check_autocall_fail(ee_barr_strike(i), ee_barr_perform_type(i))
                    Exit For
                End If
            Next i
            If isEEdate = False Then
                ee_fail_confirm = True
            End If
        Else
            ee_fail_confirm = True
        End If
        
       
        '당일 redeem_shift 취소
        If ac_fail_confirm = True And ee_fail_confirm = True Then
            'call schedule
            For i = 1 To no_of_ac_schedule
                If today = call_date(i) - 1 Then
                    call_date(i) = today
                    shtACPricer.Range("schedule_start").Cells(i, 1) = today
                    Exit For
                End If
            Next i
        
            'early-exit schedule
            For i = 1 To no_of_ee_schedule
                If today = ee_call_date(i) - 1 Then
                    ee_call_date(i) = today
                    shtACPricer.Range("early_exit_schedule_start").Cells(i, 1) = today
                    shtACPricer.Range("early_exit_schedule_start").Cells(i, 4) = today
                    Exit For
                End If
            Next i
            'monthly coupon schedule
            For i = 1 To no_of_monthly_coupon_schedule
                If today = coupon_date(i) - 1 Then
                    coupon_date(i) = today
                    shtACPricer.Range("cpn_schedule_start").Cells(i, 1) = today
                    Exit For
                End If
            Next i
            'maturity date
            If today = maturity_date - 1 Then
                maturity_date = today
                shtACPricer.Range("maturity_date") = today
            End If
        End If
    End If
    
    shtACPricer.Range("Autocall_YN") = ac_success
    If ac_success = True Then
        shtACPricer.Range("price").Cells(2, 1) = (1 + redeem_rate) * notional
    End If

    If ac_success = True Then
        With shtACPricer.Range("Autocall_YN").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Else
        With shtACPricer.Range("Autocall_YN").Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
    
    With Application
        .ScreenUpdating = bScreenUpdating
        .StatusBar = btatusBar
        .DisplayStatusBar = bDisplayStatusBar
        .Calculation = vCalculation
        .EnableEvents = bEnableEvents
    End With
    
    Set oRS = Nothing
    Set oCmd = Nothing
    Set adoCon = Nothing
    
End Sub


Private Function check_monthly_cpn(strike As Double) As Boolean

    check_monthly_cpn = check_autocall(strike, -1)

End Function

Public Function check_autocall(strike As Double, performance_type_on_call As Integer) As Boolean

    Dim i As Integer
    Dim tmp_autocall As Boolean
    
    Dim ref_spot As Double
    Dim close_spot As Double
    Dim idx As Integer
    Dim ofs As Integer
    
    Dim min_ As Double
    Dim avg_ As Double
    min_ = 2
    avg_ = 0
    
    For i = 1 To shtACPricer.Range("no_of_underlying")
    
        ref_spot = shtACPricer.Range("reference").Cells(1, i)
        
        'SX5E, SPX를 기초로 한 종목은 call date가 shift되어 있으므로 다른 기초자산 종가는 전 영업일 값을 참조한다.
        idx = WorksheetFunction.Match(shtACPricer.Range("ul_code").Cells(1, i), shtACPricer.Range("ul_code_list"), 0)
        If shtACPricer.Range("ul_code").Cells(1, i) = "SX5E" Or shtACPricer.Range("ul_code").Cells(1, i) = "SPX" Then
            ofs = 0
        Else
            ofs = 1
        End If
        close_spot = shtACPricer.Range("ua_close_spot").Cells(idx, 2 + ofs)
        
        min_ = WorksheetFunction.min(min_, close_spot / ref_spot)
        avg_ = avg_ + close_spot / ref_spot
    Next i
    avg_ = avg_ / shtACPricer.Range("no_of_underlying")
    
    If performance_type_on_call = -1 Then 'worst performer
    
        If min_ < strike Then
            tmp_autocall = False
        Else
            tmp_autocall = True
        End If
    
    ElseIf performance_type_on_call = 0 Then 'average
    
        If avg_ < strike Then
            tmp_autocall = False
        Else
            tmp_autocall = True
        End If
            
    End If
    
    check_autocall = tmp_autocall

End Function

Private Function check_autocall_fail(strike As Double, performance_type_on_call As Integer) As Boolean

    Dim tmp_autocall_fail As Boolean
    
    If performance_type_on_call = -1 Then 'worst performer
    
        Dim i As Integer
        Dim ref_spot As Double
        Dim close_spot As Double
        
        Dim idx As Integer
        Dim min_ As Double
        min_ = 2
        
        For i = 1 To shtACPricer.Range("no_of_underlying")
        
            idx = WorksheetFunction.Match(shtACPricer.Range("ul_code").Cells(1, i), shtACPricer.Range("ul_code_list"), 0)
            
            If shtACPricer.Range("ul_code").Cells(1, i) <> "SX5E" And shtACPricer.Range("ul_code").Cells(1, i) <> "SPX" Then
                ref_spot = shtACPricer.Range("reference").Cells(1, i)
                close_spot = shtACPricer.Range("ua_close_spot").Cells(idx, 2)
                min_ = WorksheetFunction.min(min_, close_spot / ref_spot)
            End If
    
        Next i
    
        If min_ < strike Then
            tmp_autocall_fail = True
        Else
            tmp_autocall_fail = False
        End If
    
    ElseIf performance_type_on_call = 0 Then 'average

        tmp_autocall_fail = False
            
    End If
    
    check_autocall_fail = tmp_autocall_fail

End Function

Private Function check_spot(ua_code As String, tdate As String, price As Double, oDB As adoDB.Connection) As Boolean

    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    sql = " select code, endprice from ras.if_stock_data where code= '" + ua_code + "' and tdate='" + tdate + "' "
    sql = sql & " union "
    sql = sql & " select code, price from ras.bl_quanto_base where code='" + ua_code + "' and tdate='" + tdate + "' "
    
    With oCmd
        .ActiveConnection = oDB
        .CommandType = adCmdText
        .CommandText = sql
       
        oRS.Open .Execute
    End With

    check_spot = False
    
    Do Until oRS.EOF
        If Abs(oRS(1) - price) < 0.0001 Then
            check_spot = True
        Else
            check_spot = False
        End If
        oRS.MoveNext
    Loop
    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing

End Function