'-------------------------------
' Modified on
' 2013-10-16
'-------------------------------
Option Explicit



'Option Explicit



'#######################################################
' Class: clsACDealTicket
' Modified by Lee, Seok on 2012/07/12
'         - added members to support 2 index pricing.
'#######################################################

Public ccy As String
Public no_of_ul As Integer '< --- 2d

Public asset_code As String
Public fund_code_m As String
Public fund_code_c As String
Private ul_code_() As String

'for leverage indices: 2018.7.9
'Private ratioLeverage_() As Double
'ratioLeverage_ -> codeLeverage:2020.6.19
Private codeLeverage_() As Integer

Public isLeveraged As Boolean
Public has_eval_shift_ul As Boolean '2024.05.23

Public current_date As Date
Public current_date_origin_ As Date

Public value_date As Date
Public Maturity As Integer
Public maturity_date As Date
Public settlement_date As Date

Public confirmed_yn As String
Public alive_yn As String

Public call_put As Long
Public strike_at_maturity As Double
Public coupon_at_maturity As Double
Public dummy_coupon As Double

Public ki_barrier_flag As Long
Public ki_touched_flag As Long
Private ki_barrier_() As Double
'Public ki_barrier_2 As Double '< --- 2d
Private early_exit_barrier_() As Double

Public ko_barrier_flag As Long
Public ko_touched_flag As Long
Public ko_barrier As Double

Public put_strike As Double
Public put_participation As Double
Public put_additional_coupon As Double
Public ki_monitoring_freq As Integer

Public ki_adj_pct As Double ' 2015-01-19

Public call_strike As Double
Public call_participation As Double


Public monthly_coupon_flag As Long
Public monthly_coupon_amount As Double '<-- Monthly coupon
Public monthly_coupon_memory_flag As Long '2024.06.10 memory 구조 추가
Public unpaid_coupon As Double '2024.06.10 memory 구조 추가

Public early_exit_flag As Long
Public early_exit_touched_flag As Long
'Private early_exit_touched_flag_() As Double
Private early_exit_touched_flag_() As Long 'data type 변경: dll(2018.7.17)

Private early_exit_performance_types_() As Long 'dll(2018.8.8)
Private early_exit_barrier_types_() As Long 'dll(2018.8.8)

Public ejectable_flag As Boolean 'dll(2021.11.12)
Private ejected_ul_flag_() As Long 'dll(2021.11.12)

Public x_grid As Long
Public v_grid As Long
Public t_grid As Long
Public days_per_step As Double 'dll(2023.05.30)

Public mid_day_greek As Boolean

Public scheme_type As Long

Public no_of_trials As Long

Public vol_scheme_type As Long

Public comment As String
Public notional As Double
Public unit_notional As Double '2023.10.23
Public qty As Double '2023.10.23
Public instrument_type As String '2023.10.23
Public current_notional As Double

Private reference_price_() As Double
'Public reference_price_2 As Double '< --- 2d
'Public strike_shift As Double
Public ki_barrier_shift As Double

Public no_of_schedule As Integer
'Public no_of_schedule_2 As Integer '< --- 2d
Public no_of_coupon_schedule As Integer
Public no_of_floating_coupon_schedule As Integer
Public no_of_early_exit_schedule As Integer


Public rate_spread As Double

Public issue_price As Double
Public issue_cost As Double

Public hedge_cost As Double ' expected hedge cost as percentage of notional

Public duration As Double



Private autocall_schedules_() As clsAutocallSchedule
'Private autocall_schedules_2_() As clsAutocallSchedule '< --- 2d
Private monthly_coupon_schedules_() As clsAutocallSchedule '<-- Monthly Coupon
Private floating_schedules_() As clsAutocallSchedule
Private early_exit_schedules_() As clsAutocallSchedule

'Public instrument_type As Long
'Public performance_type As Long
Public ki_performance_type As Long

Public ra_flag As Integer
Public ra_tenor As Integer
Public ra_cpn As Double
Public ra_min_percent As Double
Public ra_max_percent As Double

Public floor_value As Double

Private term_vega_tenor_() As Date '2015/10/05
 

Private func As New Functions

Public Function term_vega_tenor_array() As Date() '2015/10/05

    term_vega_tenor_array = term_vega_tenor_
    
End Function

Public Function term_vega_tenor(inx As Integer) As Date       '2015/10/05
    
    term_vega_tenor = term_vega_tenor_(inx)
    
End Function



Public Sub set_term_vega_tenor(term_vega_tenor() As Date)

    Dim inx As Integer
    
    ReDim term_vega_tenor_(1 To get_array_size_date(term_vega_tenor))
    
    
    For inx = 1 To get_array_size_date(term_vega_tenor)
        term_vega_tenor_(inx) = term_vega_tenor(inx)
    Next inx

End Sub

Public Function floating_schedules(inx As Integer) As clsAutocallSchedule

    Set floating_schedules = floating_schedules_(inx)

End Function

Public Sub set_floating_schedule(no_of_schedule_in As Integer, coupon_date() As Date, fixing_value() As Double)

    Dim a_schedule As clsAutocallSchedule
    Dim inx As Integer
    
On Error GoTo ErrorHandler
    
    If no_of_schedule_in >= 1 Then

        ReDim floating_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
        no_of_floating_coupon_schedule = no_of_schedule_in
'    Else
'
'        raise_err "clsACDealTicket.setSchedule", "No schedule found"

    
'
    
        For inx = 1 To no_of_schedule_in
            
            Set floating_schedules_(inx) = New clsAutocallSchedule
            
            floating_schedules_(inx).call_date = coupon_date(inx)
            floating_schedules_(inx).set_coupon_on_call fixing_value(inx)
        
        Next inx
    
    End If
    
    Exit Sub


ErrorHandler:

    raise_err "clsACDealTicket.set_floating_schedule"


End Sub



Public Function floating_coupon_dates(Optional base As Integer = 1) As Long()

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(floating_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number <> 9 And no_of_schedule_in > 0 Then
        
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Long
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = CLng(floating_schedules_(inx + 1 - base).call_date)
        
        Next inx
        
    Else
        
        ReDim rtn_array(base To base) As Long
        
        rtn_array(base) = CLng(#12/31/2100#) ' Dummy
    
    End If
    
    floating_coupon_dates = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsBarrierDealTicket.call_dates"

End Function


Public Function floating_fixing_values(Optional base As Integer = 1) As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(floating_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number <> 9 And no_of_schedule_in > 0 Then
                
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Double
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = floating_schedules_(inx + 1 - base).coupon_on_call
        
        Next inx
        
    Else
        ReDim rtn_array(base To base) As Double
        
        rtn_array(base) = 0
        
    End If
    
    floating_fixing_values = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsBarrierDealTicket.coupon_on_calls", Err.description


End Function

Public Sub set_ul_dim(no_of_ul_in As Integer)

    no_of_ul = no_of_ul_in
    
    'Preserve 추가 (2023.09.27)
    ReDim Preserve ul_code_(1 To no_of_ul_in) As String
    ReDim Preserve reference_price_(1 To no_of_ul_in) As Double
    ReDim Preserve ki_barrier_(1 To no_of_ul_in) As Double
    'ReDim early_exit_barrier_(1 To no_of_ul_in) As Double
    
    'for leverage indices: 2018.7.9
    'ReDim ratioLeverage_(1 To no_of_ul_in) As Double
    ReDim Preserve codeLeverage_(1 To no_of_ul_in) As Integer
    
    'for ejectable structure: dll(2021.11.12)
    ReDim Preserve ejected_ul_flag_(1 To no_of_ul_in) As Long
    
End Sub

Public Sub redim_early_exit_barrier(no_of_schedule_in As Integer)

    'Preserve 추가 (2023.09.27)
    ReDim Preserve early_exit_barrier_(1 To no_of_ul * no_of_schedule_in) As Double
    
End Sub
 
 
Public Function get_reference_price() As Double()

    get_reference_price = reference_price_

End Function

Public Function reference_price(Optional inx As Integer = 1) As Double

On Error Resume Next

    reference_price = reference_price_(inx)
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.ki_barrier", "Underlying Asset is not specified."
    
    End If

End Function
Public Function reference_prices(Optional base As Integer = 1) As Double()
    
    Dim rtn_array() As Double
    Dim inx As Integer
    
    ReDim rtn_array(base To base + no_of_ul - 1) As Double
    
On Error Resume Next


    For inx = base To base + no_of_ul - 1
        
        rtn_array(inx) = reference_price_(inx + 1 - base)
    
    Next inx
    
    reference_prices = rtn_array

End Function

Public Sub set_reference_price(in_reference_price As Double, Optional inx As Integer = 1)

On Error Resume Next

    reference_price_(inx) = in_reference_price
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.set_reference_price", "Underlying Asset is not specified."
    
    End If
    

End Sub

Public Sub set_reference_prices(in_reference_price() As Double)

    reference_price_ = in_reference_price
    
End Sub

Public Sub set_ejected_ul_flag(in_ejected_ul_flag As Integer, Optional inx As Integer = 1)
    
    On Error Resume Next
    
    ejected_ul_flag_(inx) = in_ejected_ul_flag
    
    If Err.number = 9 Then
        
        raise_err "clsACDealTicket.set_ejected_ul_flag", "Underlying asset is not specified."
        
    End If
    
End Sub

Public Sub set_ejected_ul_flags(in_ejected_ul_flag() As Integer)

    ejected_ul_flag_ = in_ejected_ul_flag
    
End Sub

Public Function percent_ki_barriers(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer

    ReDim rtn_array(base To base + no_of_ul - 1) As Double

    For inx = base To base + no_of_ul - 1
    
        rtn_array(inx) = ki_barrier_(inx + 1 - base)
    
    Next inx
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.ki_barrier", "Underlying Asset is not specified."
    
    End If
    
    percent_ki_barriers = rtn_array

End Function

Public Function abs_ki_barriers(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer

    ReDim rtn_array(base To base + no_of_ul - 1) As Double

    For inx = base To base + no_of_ul - 1
    
        rtn_array(inx) = ki_barrier_(inx + 1 - base) * reference_price(inx + 1 - base)
    
    Next inx
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.abs_ki_barriers", "Underlying Asset is not specified."
    
    End If
    
    abs_ki_barriers = rtn_array

End Function
 
Public Function ki_barrier(Optional inx As Integer = 1) As Double

On Error Resume Next

    ki_barrier = ki_barrier_(inx)
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.ki_barrier", "Underlying Asset is not specified."
    
    End If

End Function
Public Function early_exit_barrier(Optional inx As Integer = 1) As Double

On Error Resume Next

    early_exit_barrier = early_exit_barrier_(inx)

    If Err.number = 9 Then

        raise_err "clsACDealTicket.early_exit_barrier", "Underlying Asset is not specified."

    End If

End Function

Public Function abs_early_exit_barriers(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    Dim jnx As Integer
    
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(early_exit_schedules_)
    
On Error GoTo ErrorHandler
    
    If early_exit_flag = 0 Then
    
        ReDim rtn_array(base To no_of_ul + base - 1) As Double
        
        For inx = base To no_of_ul + base - 1
            rtn_array(inx) = 1 ' Dummy
        Next inx
        
    Else
    
        If Err.number = 9 Then
            raise_err "clsACDealTicket.coupon_barriers", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in * no_of_ul + base - 1) As Double
        
        For inx = 1 To no_of_schedule_in
            For jnx = 1 To no_of_ul
                
                'early_exit_schedules_(inx).percent_strike를 set 해주지 않아 항상 0으로 나옴. 대신 early_exit_barrier 함수를 사용
                'rtn_array(no_of_ul * (inx - 1) + jnx - 1) = early_exit_schedules_(inx).percent_strike() * reference_price(jnx)
                'rtn_array(no_of_ul * (inx - 1) + jnx - 1) = early_exit_barrier(jnx) * reference_price(jnx)
                rtn_array(no_of_ul * (inx - 1) + jnx - 1) = early_exit_barrier_(no_of_ul * (inx - 1) + jnx) * reference_price(jnx)
                
            Next jnx
        Next inx
        
    End If
    
    abs_early_exit_barriers = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.abs_early_exit_barriers"
'
'    Dim rtn_array() As Double
'    Dim inx As Integer
'
'    Dim no_of_ee_schedule As Integer
'
'    no_of_ee_schedule = UBound(early_exit_barrier)
'
'On Error Resume Next
'
'    ReDim rtn_array(base To base)
'
'    early_exit_barriers = early_exit_barrier_
'
'    If Err.number = 9 Then
'
'        raise_err "clsACDealTicket.early_exit_barrier", "Underlying Asset is not specified."
'
'    End If

End Function

Public Function early_exit_barriers(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(early_exit_schedules_)
    
On Error GoTo ErrorHandler
    
    If early_exit_flag = 0 Then
    
        ReDim rtn_array(base To base) As Double
        rtn_array(base) = 1 ' Dummy
        
    Else
    
        If Err.number = 9 Then
            raise_err "clsACDealTicket.coupon_barriers", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Double
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = early_exit_schedules_(inx + 1 - base).percent_strike()
        
        Next inx
        
    End If
    
    early_exit_barriers = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.early_exit_barriers"
'
'    Dim rtn_array() As Double
'    Dim inx As Integer
'
'    Dim no_of_ee_schedule As Integer
'
'    no_of_ee_schedule = UBound(early_exit_barrier)
'
'On Error Resume Next
'
'    ReDim rtn_array(base To base)
'
'    early_exit_barriers = early_exit_barrier_
'
'    If Err.number = 9 Then
'
'        raise_err "clsACDealTicket.early_exit_barrier", "Underlying Asset is not specified."
'
'    End If

End Function

Public Sub set_ki_barrier(in_ki_barrier As Double, Optional inx As Integer = 1)

On Error Resume Next

    ki_barrier_(inx) = in_ki_barrier
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.set_ki_barrier", "Underlying Asset is not specified."
    
    End If
    

End Sub

Public Sub set_ki_barriers(in_ki_barrier() As Double)

    ki_barrier_ = in_ki_barrier
    
End Sub
Public Sub set_early_exit_barrier(in_early_exit_barrier As Double, Optional inx As Integer = 1)

On Error Resume Next

    early_exit_barrier_(inx) = in_early_exit_barrier
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.set_early_exit_barrier", "Underlying Asset is not specified."
    
    End If
    

End Sub

Public Sub set_early_exit_barriers(in_early_exit_barrier() As Double)

    early_exit_barrier_ = in_early_exit_barrier
    
End Sub
 
 
Public Function get_ul_codes() As String()

On Error Resume Next

    get_ul_codes = ul_code_

End Function
 
Public Function ul_code(Optional inx As Integer = 1) As String

On Error Resume Next

    ul_code = ul_code_(inx)
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.ul_code", "Underlying Asset is not specified."
    
    End If

End Function

'for leverage indices: 2018.7.9
Public Function ratioLeverage(Optional inx As Integer = 1) As Double

On Error Resume Next

    ratioLeverage = ratioLeverage_(inx)
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.ratioLeverage", "Leverage ratio is not specified."
    
    End If

End Function

Public Function codeLeverage(Optional inx As Integer = 1) As Integer

On Error Resume Next

    codeLeverage = codeLeverage_(inx)
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.codeLeverage", "Leverage ratio is not specified."
    
    End If

End Function

Public Sub set_ratioLeverage(ratioLeverage As Double, Optional inx As Integer = 1)

On Error Resume Next

    ratioLeverage_(inx) = ratioLeverage
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.set_ratioLeverage", "Leverage ratio is not specified."
    
    End If

End Sub


Public Sub set_codeLeverage(codeLeverage As Integer, Optional inx As Integer = 1)

On Error Resume Next

    codeLeverage_(inx) = codeLeverage
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.set_codeLeverage", "Leverage ratio is not specified."
    
    End If

End Sub

Public Sub set_ul_code(in_ul_code As String, Optional inx As Integer = 1)

On Error Resume Next

    ul_code_(inx) = in_ul_code
    
    If Err.number = 9 Then
    
        raise_err "clsACDealTicket.set_ul_code", "Underlying Asset is not specified."
    
    End If
    

End Sub

Public Sub set_ul_codes(in_ul_codes() As String)

    ul_code_ = in_ul_codes
    
End Sub
 
'Public current_date_origin_ As Date

Public Function copy_obj() As clsACDealTicket

    Dim obj As clsACDealTicket
    
    Set obj = New clsACDealTicket
    
    obj.no_of_ul = no_of_ul
    
    obj.fund_code_m = fund_code_m
    obj.fund_code_c = fund_code_c
    obj.asset_code = asset_code
    obj.set_ul_codes ul_code_
    

    obj.current_date = current_date
    obj.current_date_origin_ = current_date_origin_
    
    obj.value_date = value_date
    obj.Maturity = Maturity
    obj.maturity_date = maturity_date
    obj.settlement_date = settlement_date

    obj.confirmed_yn = confirmed_yn
    obj.alive_yn = alive_yn
    
    
    obj.call_put = call_put
    obj.strike_at_maturity = strike_at_maturity
    obj.coupon_at_maturity = coupon_at_maturity
    obj.dummy_coupon = dummy_coupon
    
    obj.ki_barrier_flag = ki_barrier_flag
    obj.ki_touched_flag = ki_touched_flag
    obj.set_ki_barriers ki_barrier_
    
    
    obj.ki_monitoring_freq = ki_monitoring_freq
    
    obj.x_grid = x_grid
    obj.v_grid = v_grid
    obj.t_grid = t_grid
    
    obj.scheme_type = scheme_type
    
    obj.comment = comment
    obj.notional = notional
    
    obj.no_of_trials = no_of_trials
    
    obj.set_reference_prices reference_price_
    
    obj.issue_cost = issue_cost
    
    obj.copy_schedule no_of_schedule, autocall_schedules_
    
    
'
'    If obj.no_of_ul >= 2 Then
'        obj.copy_schedule_2 no_of_schedule_2, autocall_schedules_2_
'    End If
    
    obj.early_exit_flag = early_exit_flag
    obj.set_early_exit_barriers early_exit_barrier_


    obj.monthly_coupon_flag = monthly_coupon_flag
    obj.monthly_coupon_amount = monthly_coupon_amount
    
    obj.copy_coupon_schedule no_of_coupon_schedule, monthly_coupon_schedules_
 '   obj.copy_floating_schedule UBound(floating_schedules_), floating_schedules_
    obj.copy_early_exit_schedule no_of_early_exit_schedule, early_exit_schedules_
    
    obj.instrument_type = instrument_type
    obj.performance_type = performance_type
    
    obj.ra_cpn = ra_cpn
    obj.ra_flag = ra_flag
    obj.ra_max_percent = ra_max_percent
    obj.ra_min_percent = ra_min_percent
    obj.ra_tenor = ra_tenor
    
    obj.rate_spread = rate_spread
    
    obj.ki_adj_pct = ki_adj_pct
    
    obj.duration = duration
    
    obj.issue_price = issue_price
    
    obj.ki_barrier_shift = ki_barrier_shift
    obj.strike_shift = strike_shift
    obj.hedge_cost = hedge_cost
    
    obj.put_strike = put_strike
    obj.put_participation = put_participation
    
    If get_array_size_date(term_vega_tenor_) > 0 Then
        obj.set_term_vega_tenor term_vega_tenor_
    End If

    Set copy_obj = obj

End Function

Public Function autocall_schedules(inx As Integer) As clsAutocallSchedule

    Set autocall_schedules = autocall_schedules_(inx)

End Function

Public Function monthly_coupon_schedules(inx As Integer) As clsAutocallSchedule

    Set monthly_coupon_schedules = monthly_coupon_schedules_(inx)

End Function
Public Function coupon_dates(Optional base As Integer = 1) As Long()

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(monthly_coupon_schedules_)
    
On Error GoTo ErrorHandler

    If monthly_coupon_flag = 0 Then
    
        ReDim rtn_array(base To base) As Long
        rtn_array(base) = 1 ' Dummy
        
    Else
        
        If Err.number = 9 Then
            raise_err "clsACDealTicket.coupon_dates", "Not initialized array"
        End If
        
        ReDim rtn_array(base To base + no_of_schedule_in - 1) As Long
        
        For inx = base To base + no_of_schedule_in - 1
        
            rtn_array(inx) = CLng(monthly_coupon_schedules_(inx + 1 - base).call_date)
        
        Next inx
    
    End If
        
    coupon_dates = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.coupon_dates"

End Function
'Public Function early_exit_touched_flags(Optional base As Integer = 1) As Double()
Public Function early_exit_touched_flags(Optional base As Integer = 1) As Long() 'data type 변경: dll(2018.7.17)

    'Dim rtn_array() As Double
    Dim rtn_array() As Long 'data type 변경: dll(2018.7.17)
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(early_exit_schedules_)
    
On Error GoTo ErrorHandler

    If early_exit_flag = 0 Then
    
        'ReDim rtn_array(base To base) As Double
        ReDim rtn_array(base To base) As Long 'data type 변경: dll(2018.7.17)
        rtn_array(base) = 1#  ' Dummy
        
    Else
        
        If Err.number = 9 Then
            raise_err "clsACDealTicket.early_exit_touched_flags", "Not initialized array"
        End If
        
        'ReDim rtn_array(base To no_of_schedule_in + base - 1) As Double
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Long 'data type 변경: dll(2018.7.17)
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = early_exit_touched_flag_(inx + 1 - base)
        
        Next inx
    
    End If
        
    early_exit_touched_flags = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.early_exit_touched_flags"

End Function

Public Function early_exit_performance_types(Optional base As Integer = 1) As Long() 'dll(2018.8.8)

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(early_exit_schedules_)
    
On Error GoTo ErrorHandler

    If early_exit_flag = 0 Then
    
        ReDim rtn_array(base To base) As Long
        rtn_array(base) = 1#  ' Dummy
        
    Else
        
        If Err.number = 9 Then
            raise_err "clsACDealTicket.early_exit_performance_types_", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Long
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = early_exit_performance_types_(inx + 1 - base)
        
        Next inx
    
    End If
        
    early_exit_performance_types = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.early_exit_performance_types_"

End Function

Public Function early_exit_barrier_types(Optional base As Integer = 1) As Long() 'dll(2018.8.8)

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(early_exit_schedules_)
    
On Error GoTo ErrorHandler

    If early_exit_flag = 0 Then
    
        ReDim rtn_array(base To base) As Long
        rtn_array(base) = 1#  ' Dummy
        
    Else
        
        If Err.number = 9 Then
            raise_err "clsACDealTicket.early_exit_barrier_types_", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Long
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = early_exit_barrier_types_(inx + 1 - base)
        
        Next inx
    
    End If
        
    early_exit_barrier_types = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.early_exit_barrier_types_"

End Function

Public Function early_exit_dates(Optional base As Integer = 1) As Long()

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(early_exit_schedules_)
    
On Error GoTo ErrorHandler

    If early_exit_flag = 0 Then
    
        ReDim rtn_array(base To base) As Long
        rtn_array(base) = 1 ' Dummy
        
    Else
        
        If Err.number = 9 Then
            raise_err "clsACDealTicket.coupon_dates", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Long
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = CLng(early_exit_schedules_(inx + 1 - base).call_date)
        
        Next inx
    
    End If
        
    early_exit_dates = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.early_exit_dates"

End Function

Public Function early_exit_date(inx As Integer, Optional base As Integer = 1) As Date

On Error Resume Next

    early_exit_date = early_exit_schedules_(inx + 1 - base).call_date

    If Err.number = 9 Then

        raise_err "clsACDealTicket.early_exit_date"

    End If

End Function

Public Function ejected_ul_flag(Optional base As Integer = 1) As Long() 'dll(2021.11.12)

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_ul_in As Integer
    
    On Error Resume Next
    
    no_of_ul_in = UBound(ejected_ul_flag_)
    
On Error GoTo ErrorHandler
    
    If ejectable_flag = False Then
    
        ReDim rtn_array(base To base) As Long
        rtn_array(base) = 1 ' Dummy
        
    Else
    
        If Err.number = 9 Then
            raise_err "clsACDealTicket.ejected_ul_flag", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_ul_in + base - 1) As Long
        
        For inx = base To no_of_ul_in + base - 1
        
            rtn_array(inx) = ejected_ul_flag_(inx + 1 - base)
        
        Next inx
        
    End If
    
    ejected_ul_flag = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.ejected_ul_flag"

End Function

Public Function ejected_event_flag(Optional base As Integer = 1) As Long() 'dll(2021.11.12)

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If ejectable_flag = False Then
    
        ReDim rtn_array(base To base) As Long
        rtn_array(base) = 1 ' Dummy
        
    Else
    
        If Err.number = 9 Then
            raise_err "clsACDealTicket.ejected_event_flag", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Long
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = autocall_schedules_(inx + 1 - base).ejected_event_flag
        
        Next inx
        
    End If
    
    ejected_event_flag = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.ejected_event_flag"

End Function

Public Function ejectable_order(Optional base As Integer = 1) As Long() 'dll(2021.11.12)

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If ejectable_flag = False Then
    
        ReDim rtn_array(base To base) As Long
        rtn_array(base) = 1 ' Dummy
        
    Else
    
        If Err.number = 9 Then
            raise_err "clsACDealTicket.ejectable_order", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Long
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = autocall_schedules_(inx + 1 - base).ejectable_order
        
        Next inx
        
    End If
    
    ejectable_order = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.ejectable_order"

End Function

Public Function early_exit_coupon_amt(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(early_exit_schedules_)
    
On Error GoTo ErrorHandler
    
    If early_exit_flag = 0 Then
    
        ReDim rtn_array(base To base) As Double
        rtn_array(base) = 1 ' Dummy
        
    Else
    
        If Err.number = 9 Then
            raise_err "clsACDealTicket.coupon_barriers", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Double
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = early_exit_schedules_(inx + 1 - base).coupon_on_call
        
        Next inx
        
    End If
    
    early_exit_coupon_amt = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.early_exit_coupon_amt"


End Function

Public Function early_exit_coupon(inx As Integer, Optional base As Integer = 1) As Double

On Error Resume Next

    early_exit_coupon = early_exit_schedules_(inx + 1 - base).coupon_on_call

    If Err.number = 9 Then

        raise_err "clsACDealTicket.early_exit_coupon"

    End If

End Function

Public Function monthly_coupon_amt(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(monthly_coupon_schedules_)
    
On Error GoTo ErrorHandler
    
    If monthly_coupon_flag = 0 Then
    
        ReDim rtn_array(base To base) As Double
        rtn_array(base) = 1 ' Dummy
        
    Else
    
        If Err.number = 9 Then
            raise_err "clsACDealTicket.coupon_barriers", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Double
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = monthly_coupon_schedules_(inx + 1 - base).coupon_on_call
        
        Next inx
        
    End If
    
    monthly_coupon_amt = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.coupon_barriers"


End Function
'Public Function coupon_barriers() As Double()
'
'    Dim rtn_array() As Double
'    Dim inx As Integer
'    Dim no_of_schedule_in As Integer
'
'    On Error Resume Next
'
'    no_of_schedule_in = UBound(monthly_coupon_schedules_)
'
'On Error GoTo ErrorHandler
'
'    If monthly_coupon_flag = 0 Then
'
'        ReDim rtn_array(1 To 1) As Double
'        rtn_array(1) = 1 ' Dummy
'
'    Else
'
'        If Err.number = 9 Then
'            raise_err "clsACDealTicket.coupon_barriers", "Not initialized array"
'        End If
'
'        ReDim rtn_array(1 To no_of_schedule_in) As Double
'
'        For inx = 1 To no_of_schedule_in
'
'            rtn_array(inx) = monthly_coupon_schedules_(inx).strike_value
'
'        Next inx
'
'    End If
'
'    coupon_barriers = rtn_array
'
'    Exit Function
'
'ErrorHandler:
'
'    raise_err "clsACDealTicket.coupon_barriers"
'
'
'End Function
Public Function abs_coupon_barriers(Optional base As Integer = 1) As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim jnx As Integer
    
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(monthly_coupon_schedules_)
    
On Error GoTo ErrorHandler
    If monthly_coupon_flag = 0 Then
    
        ReDim rtn_array(base To no_of_ul + base - 1) As Double
        
        For inx = base To no_of_ul + base - 1
            rtn_array(inx) = 1 ' Dummy
        Next inx
    
    Else
        If Err.number = 9 Then
            raise_err "clsACDealTicket.percent_coupon_barriers", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in * no_of_ul + base - 1) As Double
        
        For inx = 1 To no_of_schedule_in
            For jnx = 1 To no_of_ul
        
                rtn_array(no_of_ul * (inx - 1) + jnx - 1) = monthly_coupon_schedules_(inx).percent_strike * reference_price(jnx)
                
            Next jnx
        Next inx
    End If
    
    abs_coupon_barriers = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.percent_coupon_barriers"


End Function
Public Function percent_coupon_barriers(Optional base As Integer = 1) As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(monthly_coupon_schedules_)
    
On Error GoTo ErrorHandler
    If monthly_coupon_flag = 0 Then
    
        ReDim rtn_array(base To base) As Double
        rtn_array(base) = 1 ' Dummy
    
    Else
        If Err.number = 9 Then
            raise_err "clsACDealTicket.percent_coupon_barriers", "Not initialized array"
        End If
        
        ReDim rtn_array(base To no_of_schedule_in + base - 1) As Double
        
        For inx = base To no_of_schedule_in + base - 1
        
            rtn_array(inx) = monthly_coupon_schedules_(inx + 1 - base).percent_strike
        
        Next inx
    End If
    
    percent_coupon_barriers = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.percent_coupon_barriers"


End Function


Public Function call_dates(Optional base As Integer = 1) As Long()

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.call_dates", "Not initialized array"
    End If
    
    ReDim rtn_array(base To base + no_of_schedule_in - 1) As Long
    
    For inx = base To base + no_of_schedule_in - 1
    
        rtn_array(inx) = CLng(autocall_schedules_(inx + 1 - base).call_date)
    
    Next inx
        
    call_dates = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.call_dates"

End Function
'
'Public Function strike_values() As Double()
'
'    Dim rtn_array() As Double
'    Dim inx As Integer
'    Dim no_of_schedule_in As Integer
'
'    On Error Resume Next
'
'    no_of_schedule_in = UBound(autocall_schedules_)
'
'On Error GoTo ErrorHandler
'
'    If Err.number = 9 Then
'        raise_err "clsACDealTicket.strike_values", "Not initialized array"
'    End If
'
'    ReDim rtn_array(1 To no_of_schedule_in) As Double
'
'    For inx = 1 To no_of_schedule_in
'
'        rtn_array(inx) = autocall_schedules_(inx).strike_value
'
'    Next inx
'
'    strike_values = rtn_array
'
'    Exit Function
'
'ErrorHandler:
'
'    raise_err "clsACDealTicket.call_dates"
'
'
'End Function

Public Function coupon_on_call_array(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double

    Dim inx As Integer
    Dim jnx As Integer
    
    Dim no_of_schedule_in As Integer
    Dim no_of_jumps As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    no_of_jumps = autocall_schedules_(1).no_of_jumps()
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.coupon_on_calls", "Not initialized array"
    End If
    
    ReDim rtn_array(base To no_of_schedule_in * no_of_jumps + base - 1) As Double
        
    For inx = 1 To no_of_jumps
        For jnx = 1 To no_of_schedule_in
    
                rtn_array((inx - 1) * no_of_schedule_in + jnx - 1 + base) = autocall_schedules_(jnx).coupon_on_call(inx)
            
        Next jnx
    Next inx
        
    coupon_on_call_array = rtn_array
    

    Exit Function
    
ErrorHandler:

    raise_err "clsAcDealTicket.coupon_on_call_array", Err.description

End Function


'-------------------------------------------------------------------------
'This will be Deperciated Soon. 2014-03-26
'-------------------------------------------------------------------------
Public Function coupon_on_calls(Optional base As Integer = 1) As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    Dim no_of_jumps As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.coupon_on_calls", "Not initialized array"
    End If
    
    ReDim rtn_array(base To no_of_schedule_in + base - 1) As Double
    
    no_of_jumps = autocall_schedules_(1).no_of_jumps()
    
    For inx = base To no_of_schedule_in * no_of_jumps + base - 1
    
        rtn_array(inx) = autocall_schedules_(inx + 1 - base).coupon_on_call
    
    Next inx
        
    coupon_on_calls = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.call_dates"


End Function

'-------------------------------------------------------------------------
'This will be Deperciated Soon. 2014-03-26
'-------------------------------------------------------------------------
Public Function coupon_on_calls_multi_jump() As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim jnx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.coupon_on_calls", "Not initialized array"
    End If
    
    ReDim rtn_array(1 To no_of_schedule_in, 1 To autocall_schedules_(1).no_of_jumps()) As Double
    
    For inx = 1 To no_of_schedule_in
        For jnx = 1 To autocall_schedules_(inx).no_of_jumps()
            rtn_array(inx, jnx) = autocall_schedules_(inx).coupon_on_call(jnx)
        Next jnx
    Next inx
        
    coupon_on_calls_multi_jump = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.call_dates"


End Function


Public Function percent_strikes_multi_jump() As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim jnx As Integer
    
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.coupon_on_calls", "Not initialized array"
    End If
    
    ReDim rtn_array(1 To no_of_schedule_in, 1 To autocall_schedules_(1).no_of_jumps()) As Double
    
    For inx = 1 To no_of_schedule_in
        For jnx = 1 To autocall_schedules_(inx).no_of_jumps
            rtn_array(inx, jnx) = autocall_schedules_(inx).percent_strike(jnx)
        Next jnx
    Next inx
        
    percent_strikes_multi_jump = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.percent_strikes"


End Function
Public Function abs_ra_max(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    
    ReDim rtn_array(base To base + no_of_ul - 1) As Double
    
    For inx = 1 To no_of_ul
        
        rtn_array((inx - 1) + base) = ra_max_percent * reference_price_(inx)
    
    Next inx
    
    abs_ra_max = rtn_array

End Function
Public Function abs_ra_min(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    
    ReDim rtn_array(base To base + no_of_ul - 1) As Double
    
    For inx = 1 To no_of_ul
        
        rtn_array((inx - 1) + base) = ra_min_percent * reference_price_(inx)
    
    Next inx
    
    abs_ra_min = rtn_array

End Function

'-------------------------------------------------------------------
' Returns 1dim array
'
'-------------------------------------------------------------------
Public Function abs_strikes(Optional base As Integer = 1) As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim jnx As Integer
    Dim knx As Integer
    
    Dim no_of_schedule_in As Integer
    Dim no_of_jumps As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    no_of_jumps = autocall_schedules_(1).no_of_jumps()
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.coupon_on_calls", "Not initialized array"
    End If
    
    ReDim rtn_array(base To no_of_schedule_in * no_of_ul * no_of_jumps + base - 1) As Double
        
    For inx = 1 To no_of_jumps
        For jnx = 1 To no_of_schedule_in
            For knx = 1 To no_of_ul
    
                rtn_array((inx - 1) * no_of_schedule_in * no_of_ul + (jnx - 1) * no_of_ul + knx - 1 + base) = autocall_schedules_(jnx).percent_strike(inx) * reference_price(knx)
            
            Next knx
        Next jnx
    Next inx
        
    abs_strikes = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.abs_strikes", Err.description


End Function
Public Function percent_strikes(Optional base As Integer = 1) As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.coupon_on_calls", "Not initialized array"
    End If
    
    ReDim rtn_array(base To no_of_schedule_in + base - 1) As Double
    
    For inx = base To no_of_schedule_in + base - 1
    
        rtn_array(inx) = autocall_schedules_(inx + 1 - base).percent_strike
    
    Next inx
        
    percent_strikes = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.percent_strikes"


End Function

'<--added
Public Function performance_type(Optional base As Integer = 1) As Long()

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.performance_type", "Not initialized array"
    End If
    
    ReDim rtn_array(base To no_of_schedule_in + base - 1) As Long
    
    For inx = base To no_of_schedule_in + base - 1
    
        rtn_array(inx) = autocall_schedules_(inx + 1 - base).performance_type
    
    Next inx
        
    performance_type = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.performance_type"

End Function

Public Function strike_shift(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.strike_shift", "Not initialized array"
    End If
    
    ReDim rtn_array(base To no_of_schedule_in + base - 1) As Double
    
    For inx = base To no_of_schedule_in + base - 1
    
        rtn_array(inx) = autocall_schedules_(inx + 1 - base).strike_shift
    
    Next inx
        
    strike_shift = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.strike_shift"

End Function

Public Function early_exit_strike_shift(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(early_exit_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.early_exit_strike_shift", "Not initialized array"
    End If
    
    ReDim rtn_array(base To no_of_schedule_in + base - 1) As Double
    
    For inx = base To no_of_schedule_in + base - 1
    
        rtn_array(inx) = early_exit_schedules_(inx + 1 - base).strike_shift
    
    Next inx
        
    early_exit_strike_shift = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.early_exit_strike_shift"

End Function

Public Function strike_shift_at_maturity(Optional base As Integer = 1) As Double

    strike_shift_at_maturity = autocall_schedules_(UBound(autocall_schedules_)).strike_shift
    
End Function

Public Function performance_type_at_maturity() As Long

    performance_type_at_maturity = autocall_schedules_(UBound(autocall_schedules_)).performance_type

End Function
'-->added

Public Function percent_strikes_at_maturity(Optional base As Integer = 1) As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule As Integer
    
    On Error Resume Next
    
    no_of_schedule = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        raise_err "clsACDealTicket.coupon_on_calls", "Not initialized array"
    End If
    
    ReDim rtn_array(base To base + autocall_schedules_(no_of_schedule).no_of_jumps - 1) As Double
    
    For inx = base To base + autocall_schedules_(no_of_schedule).no_of_jumps - 1
    
        rtn_array(inx) = autocall_schedules_(no_of_schedule).percent_strike(inx + 1 - base)
    
    Next inx
        
    percent_strikes_at_maturity = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.percent_strikes_at_maturity"


End Function



Public Sub copy_coupon_schedule(no_of_schedule_in As Integer, schs() As clsAutocallSchedule)

    Dim inx As Integer
    
    If no_of_schedule_in > 0 Then
    
        ReDim monthly_coupon_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
        no_of_coupon_schedule = no_of_schedule_in
        
        For inx = 1 To no_of_coupon_schedule
                
            Set monthly_coupon_schedules_(inx) = New clsAutocallSchedule
            
            monthly_coupon_schedules_(inx).call_date = schs(inx).call_date
            monthly_coupon_schedules_(inx).set_percent_strike schs(inx).percent_strike
            monthly_coupon_schedules_(inx).set_coupon_on_call schs(inx).coupon_on_call
            monthly_coupon_schedules_(inx).set_strike_value schs(inx).strike_value
'            monthly_coupon_schedules_(inx).set_strike_value monthly_coupon_schedules_(inx).percent_strike * reference_price
        
        Next inx
    
    End If
    

End Sub

'Public Sub set_early_exit_schedule(no_of_schedule_in As Integer, calldate() As Date, coupon_on_call() As Double)
'
'    Dim a_schedule As clsAutocallSchedule
'    Dim inx As Integer
'
'On Error GoTo ErrorHandler
'
'    If no_of_schedule_in >= 1 Then
'
'        ReDim early_exit_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
'        no_of_early_exit_schedule = no_of_schedule_in
'
'        For inx = 1 To no_of_early_exit_schedule
'
'            Set early_exit_schedules_(inx) = New clsAutocallSchedule
'
'            early_exit_schedules_(inx).call_date = calldate(inx)
'            early_exit_schedules_(inx).set_coupon_on_call coupon_on_call(inx)
'            '<---added
'            early_exit_schedules_(inx).strike_shift = 0
'            '--->
'
'        Next inx
'    Else
'
'        ReDim early_exit_schedules_(1 To 1) As clsAutocallSchedule 'Dummy schedule
'        Set early_exit_schedules_(1) = New clsAutocallSchedule
'        no_of_early_exit_schedule = 1
'
'    End If
'
'
'
'
'    Exit Sub
'
'
'ErrorHandler:
'
'    raise_err "clsACDealTicket.set_early_exit_schedule", Err.description
'
'
'End Sub

'Public Sub set_early_exit_schedule(no_of_schedule_in As Integer, calldate() As Date, coupon_on_call() As Double, strike_shift() As Double, early_exit_touched_flag() As Double)
'Public Sub set_early_exit_schedule(no_of_schedule_in As Integer, calldate() As Date, coupon_on_call() As Double, strike_shift() As Double, early_exit_touched_flag() As Long)
'data type 변경: dll(2018.7.17)
Public Sub set_early_exit_schedule(no_of_schedule_in As Integer, calldate() As Date, coupon_on_call() As Double, strike_shift() As Double, early_exit_touched_flag() As Long, early_exit_performance_types() As Long, early_exit_barrier_types() As Long)
'dll(2018.8.8)

    'Dim a_schedule As clsAutocallSchedule
    Dim inx As Integer

On Error GoTo ErrorHandler

    'Preserve 추가 (2023.09.27)
    If no_of_schedule_in >= 1 Then

        ReDim Preserve early_exit_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
        'ReDim early_exit_touched_flag_(1 To no_of_schedule_in) As Double
        ReDim Preserve early_exit_touched_flag_(1 To no_of_schedule_in) As Long 'data type 변경: dll(2018.7.17)
        ReDim Preserve early_exit_performance_types_(1 To no_of_schedule_in) As Long 'dll(2018.8.8)
        ReDim Preserve early_exit_barrier_types_(1 To no_of_schedule_in) As Long 'dll(2018.8.8)
        no_of_early_exit_schedule = no_of_schedule_in

        For inx = 1 To no_of_early_exit_schedule

            Set early_exit_schedules_(inx) = New clsAutocallSchedule

            early_exit_schedules_(inx).call_date = calldate(inx)
            early_exit_schedules_(inx).set_coupon_on_call coupon_on_call(inx)
            '<---added
            early_exit_schedules_(inx).strike_shift = strike_shift(inx)
            early_exit_touched_flag_(inx) = early_exit_touched_flag(inx)
            early_exit_performance_types_(inx) = early_exit_performance_types(inx) 'dll(2018.8.8)
            early_exit_barrier_types_(inx) = early_exit_barrier_types(inx) 'dll(2018.8.8)
            '--->

        Next inx
    Else

        ReDim Preserve early_exit_schedules_(1 To 1) As clsAutocallSchedule 'Dummy schedule
        Set early_exit_schedules_(1) = New clsAutocallSchedule
        no_of_early_exit_schedule = 1

    End If




    Exit Sub


ErrorHandler:

    raise_err "clsACDealTicket.set_early_exit_schedule", Err.description


End Sub

Public Sub set_coupon_schedule(no_of_schedule_in As Integer, calldate() As Date, strike_percent() As Double, coupon_on_call() As Double)

    Dim a_schedule As clsAutocallSchedule
    Dim inx As Integer
    
On Error GoTo ErrorHandler
    
    If no_of_schedule_in >= 1 Then
    
        ReDim monthly_coupon_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
        no_of_coupon_schedule = no_of_schedule_in
                
        For inx = 1 To no_of_coupon_schedule
            
            Set monthly_coupon_schedules_(inx) = New clsAutocallSchedule
         
            monthly_coupon_schedules_(inx).call_date = calldate(inx)
            monthly_coupon_schedules_(inx).set_coupon_on_call coupon_on_call(inx)
            monthly_coupon_schedules_(inx).set_percent_strike strike_percent(inx)
    '        monthly_coupon_schedules_(inx).set_strike_value monthly_coupon_schedules_(inx).percent_strike * reference_price
        
        Next inx
    Else
    
        ReDim monthly_coupon_schedules_(1 To 1) As clsAutocallSchedule 'Dummy schedule
        Set monthly_coupon_schedules_(1) = New clsAutocallSchedule
        no_of_coupon_schedule = 1
        
    End If
        

    
    Exit Sub


ErrorHandler:

    raise_err "clsACDealTicket.set_coupon_schedule"


End Sub


Public Sub copy_schedule(no_of_schedule_in As Integer, schs() As clsAutocallSchedule)

    Dim inx As Integer
    Dim jnx As Integer
    
    ReDim autocall_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
    no_of_schedule = no_of_schedule_in
    
    For inx = 1 To no_of_schedule
            
        Set autocall_schedules_(inx) = schs(inx).copy_obj ' clsAutocallSchedule
'
'        autocall_schedules_(inx).call_date = schs(inx).call_date
'
'        For jnx = 1 To autocall_schedules_(inx).coupon_on_call
'        autocall_schedules_(inx).set_coupon_on_call schs(inx).coupon_on_call
'        autocall_schedules_(inx).set_percent_strike schs(inx).percent_strike
'        autocall_schedules_(inx).set_strike_value autocall_schedules_(inx).percent_strike * reference_price
    
    Next inx
    
    

End Sub

''------------------------------------
'' To support 2 index ELS
''------------------------------------
'Public Sub copy_schedule_2(no_of_schedule_in As Integer, schs() As clsAutocallSchedule)
'
'    Dim inx As Integer
'
'    ReDim autocall_schedules_2_(1 To no_of_schedule_in) As clsAutocallSchedule
'    no_of_schedule_2 = no_of_schedule_in
'
'    For inx = 1 To no_of_schedule_2
'
'        Set autocall_schedules_(inx) = New clsAutocallSchedule
'
'        autocall_schedules_2_(inx).call_date = schs(inx).call_date
'        autocall_schedules_2_(inx).coupon_on_call = schs(inx).coupon_on_call
'        autocall_schedules_2_(inx).percent_strike = schs(inx).percent_strike
'        autocall_schedules_2_(inx).strike_value = autocall_schedules_2_(inx).percent_strike * reference_price_2
'
'    Next inx
'
'
'
'End Sub



Public Sub set_schedule_array(schedules() As clsAutocallSchedule)

On Error GoTo ErrorHandler
    
    autocall_schedules_ = schedules
    
    Exit Sub


ErrorHandler:

    raise_err "clsACDealTicket.set_schedule_array"


End Sub

Public Sub set_schedule(no_of_schedule_in As Integer, calldate() As Date, strike_percent() As Double, coupon_on_call() As Double)

    Dim a_schedule As clsAutocallSchedule
    Dim inx As Integer
    Dim sch_index As Integer
    Dim jump_index As Integer
    Dim pre_call_date As Date
    
On Error GoTo ErrorHandler
    
    If no_of_schedule_in >= 1 Then
    
        Erase autocall_schedules_
    
        'ReDim autocall_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
        
        
        pre_call_date = 0
        
        
        For inx = 1 To no_of_schedule_in
        
            If pre_call_date <> calldate(inx) Then
                jump_index = 0
                sch_index = sch_index + 1
                
                
                
                Set a_schedule = New clsAutocallSchedule
            
                a_schedule.call_date = calldate(inx)
                push_back_clsAutocallSchedule autocall_schedules_, a_schedule
                
                pre_call_date = calldate(inx)
                
            End If
            
            jump_index = jump_index + 1
            
            autocall_schedules_(sch_index).set_coupon_on_call coupon_on_call(inx), jump_index
            autocall_schedules_(sch_index).set_percent_strike strike_percent(inx), jump_index
'            autocall_schedules_(sch_index).set_strike_value autocall_schedules_(inx).percent_strike * reference_price(inx), jump_index
        
        Next inx
        
        
        no_of_schedule = sch_index
        
    Else
    
        raise_err "clsACDealTicket.setSchedule", "No schedule found"
        
    End If
        
    
    Exit Sub


ErrorHandler:

    raise_err "clsACDealTicket.setSchedule"


End Sub

Public Function get_next_call_date(ByVal current_date As Date) As Date

    Dim tmp_schedule As clsAutocallSchedule
    Dim rtn_value As Date
    
    If get_next_schedule(tmp_schedule, current_date) Then
        rtn_value = tmp_schedule.call_date
    Else
        rtn_value = 0
    End If
    
    get_next_call_date = rtn_value

End Function


Public Function get_next_schedule(next_schedule As clsAutocallSchedule, ByVal current_date As Date) As Boolean

    Dim schedule_exist As Boolean
    Dim inx As Integer
    
    schedule_exist = False
    
On Error GoTo ErrorHandler
    
    If no_of_schedule >= 1 Then
    
        For inx = 1 To no_of_schedule
        
            If autocall_schedules_(inx).call_date >= current_date Then
        
                Set next_schedule = autocall_schedules_(inx)
                schedule_exist = True
                Exit For
                
            End If
        
        Next inx
            
    Else
    
        schedule_exist = False
    
    End If
    
    
    get_next_schedule = schedule_exist


    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.get_next_schedule"

End Function
Public Function get_prev_schedule(prev_schedule As clsAutocallSchedule, ByVal current_date As Date) As Boolean

    Dim schedule_exist As Boolean
    Dim inx As Integer
    
    schedule_exist = False
    
On Error GoTo ErrorHandler
    
    If no_of_schedule >= 1 Then
    
        For inx = 1 To no_of_schedule
        
            If autocall_schedules_(inx).call_date >= current_date Then
            
                If inx > 1 Then
        
                    Set prev_schedule = autocall_schedules_(inx - 1)
                    schedule_exist = True
                Else
                    schedule_exist = False
                End If
                
                Exit For
                
            End If
        
        Next inx
            
    Else
    
        schedule_exist = False
    
    End If
    
    
    get_prev_schedule = schedule_exist


    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.get_next_schedule"

End Function

Public Function get_next_coupon_schedule(next_schedule As clsAutocallSchedule, ByVal current_date As Date) As Boolean

    Dim schedule_exist As Boolean
    Dim inx As Integer
    
    schedule_exist = False
    
On Error GoTo ErrorHandler
    
    If no_of_coupon_schedule >= 1 Then
    
        For inx = 1 To no_of_coupon_schedule
        
            If monthly_coupon_schedules_(inx).call_date >= current_date Then
        
                Set next_schedule = monthly_coupon_schedules_(inx)
                schedule_exist = True
                Exit For
                
            End If
        
        Next inx
            
    Else
    
        schedule_exist = False
    
    End If
    
    
    get_next_coupon_schedule = schedule_exist


    Exit Function
    
ErrorHandler:

    raise_err "clsACDealTicket.get_next_coupon_schedule", Err.description

End Function

Public Function get_call_frequency() As Integer

    Dim call_frequency As Integer
    
    If no_of_schedule >= 2 Then
    
        call_frequency = Round((autocall_schedules_(no_of_schedule).call_date - autocall_schedules_(no_of_schedule - 1).call_date) / 30.4, 0)

    End If
    
    get_call_frequency = call_frequency

End Function


'------------------------------------------------------------
Public Sub set_to_call_prob_mode(rate As Double, seq As Integer)

    Dim inx As Integer
    
    Me.x_grid = 50
    Me.t_grid = 50
        
    autocall_schedules_(no_of_schedule).set_percent_strike 0.001
    
    monthly_coupon_flag = 0
    ra_flag = 0
    
    For inx = 1 To no_of_schedule
    
        If autocall_schedules_(inx).call_date >= current_date Then
    
            autocall_schedules_(inx).set_coupon_on_call -1
            
        End If
    
    Next inx
    
    If seq <= no_of_schedule Then
        autocall_schedules_(seq).set_coupon_on_call Exp(rate * (autocall_schedules_(seq).call_date - current_date) / 365) - 1
    Else
        raise_err "set_to_call_prob_mode", "Invalid call sequence"
    End If

End Sub
'------------------------------------------------------------
Public Sub set_to_duration_mode(rate As Double)

    Dim inx As Integer
    
    'Me.x_grid = 200
    Me.t_grid = 50
    
    autocall_schedules_(no_of_schedule).set_percent_strike 0.001
    
    monthly_coupon_flag = 0
    ra_flag = 0
    Me.ki_barrier_flag = 0
    ki_adj_pct = 1
    
    For inx = 1 To no_of_schedule
    
        If autocall_schedules_(inx).call_date >= current_date Then
    
            autocall_schedules_(inx).set_coupon_on_call (autocall_schedules_(inx).call_date - current_date) / 365 * Exp(rate * (autocall_schedules_(inx).call_date - current_date) / 365) - 1
            
        End If
    
    Next inx

    Me.coupon_at_maturity = autocall_schedules_(no_of_schedule).coupon_on_call()
End Sub


Public Sub copy_early_exit_schedule(no_of_schedule_in As Integer, schs() As clsAutocallSchedule)

    Dim inx As Integer
    
    If no_of_schedule_in > 0 Then
        
        ReDim early_exit_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
        no_of_early_exit_schedule = no_of_schedule_in
        
        For inx = 1 To no_of_early_exit_schedule
                
            Set early_exit_schedules_(inx) = New clsAutocallSchedule
            
            early_exit_schedules_(inx).call_date = schs(inx).call_date
            early_exit_schedules_(inx).set_coupon_on_call schs(inx).coupon_on_call
            early_exit_schedules_(inx).set_percent_strike schs(inx).percent_strike
    '        early_exit_schedules_(inx).set_strike_value schs(inx).strike_value  ' * reference_price
        
        Next inx
        
    End If

End Sub

Public Sub copy_floating_schedule(no_of_schedule_in As Integer, schs() As clsAutocallSchedule)

    Dim inx As Integer
    
    ReDim floating_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
    no_of_floating_coupon_schedule = no_of_schedule_in
    
    For inx = 1 To no_of_floating_coupon_schedule
            
        Set floating_schedules_(inx) = New clsAutocallSchedule
        
        floating_schedules_(inx).call_date = schs(inx).call_date
        floating_schedules_(inx).set_coupon_on_call schs(inx).coupon_on_call
        floating_schedules_(inx).set_percent_strike schs(inx).percent_strike
'        floating_schedules_(inx).set_strike_value schs(inx).strike_value  ' * reference_price
    
    Next inx
    
    

End Sub