Option Explicit


Public call_date As Date


Private percent_strike_() As Double
Private coupon_on_call_() As Double
Private strike_value_() As Double
'<---added
Public performance_type As Long
Public strike_shift As Double
Public ejectable_order As Long 'dll(2021.11.12)
Public ejected_event_flag As Long 'dll(2021.11.12)
'--->


Public Function no_of_jumps() As Integer
    
    Dim rtn_value As Integer
    
    
    rtn_value = get_array_size_double(percent_strike_) '<------------------------------
    
    no_of_jumps = rtn_value

End Function

Public Function copy_obj() As clsAutocallSchedule

    Dim rtn_obj As clsAutocallSchedule
    Dim inx As Integer
    
On Error GoTo ErrorHandler

    Set rtn_obj = New clsAutocallSchedule
    
    rtn_obj.call_date = call_date
    rtn_obj.performance_type = performance_type
    rtn_obj.strike_shift = strike_shift
    rtn_obj.ejectable_order = ejectable_order
    rtn_obj.ejected_event_flag = ejected_event_flag
    
    For inx = 1 To get_array_size_double(percent_strike_)
        rtn_obj.set_percent_strike percent_strike_(inx), inx
    Next inx
    
    For inx = 1 To get_array_size_double(coupon_on_call_)
        rtn_obj.set_coupon_on_call coupon_on_call_(inx), inx
    Next inx
    
    For inx = 1 To get_array_size_double(strike_value_)
        rtn_obj.set_strike_value strike_value_(inx), inx
    Next inx
    
    Set copy_obj = rtn_obj
    
    Exit Function
    
ErrorHandler:

    raise_err "clsAutocallSchedule.copy_obj", Err.description

End Function

Public Function percent_strike(Optional jump_no As Integer = 1) As Double

    percent_strike = percent_strike_(jump_no)

End Function

Public Function coupon_on_call(Optional jump_no As Integer = 1) As Double

    coupon_on_call = coupon_on_call_(jump_no)

End Function

Public Function strike_value(Optional jump_no As Integer = 1) As Double

    strike_value = strike_value_(jump_no)

End Function

Public Sub set_coupon_on_call(cpn_value As Double, Optional jump_no As Integer = 1)

    If UBound(coupon_on_call_) < jump_no Then
        push_back_double coupon_on_call_, cpn_value
    Else
        coupon_on_call_(jump_no) = cpn_value
    End If

End Sub

Public Sub set_percent_strike(pct_strike As Double, Optional jump_no As Integer = 1)

    If UBound(percent_strike_) < jump_no Then
        push_back_double percent_strike_, pct_strike
    Else
        percent_strike_(jump_no) = pct_strike
    End If

End Sub

Public Sub set_strike_value(strike As Double, Optional jump_no As Integer = 1)

    If UBound(strike_value_) < jump_no Then
        push_back_double strike_value_, strike
    Else
        strike_value_(jump_no) = strike
    End If

End Sub

Private Sub Class_Initialize()

    ReDim percent_strike_(1 To 1) As Double
    ReDim coupon_on_call_(1 To 1) As Double
    ReDim strike_value_(1 To 1) As Double

End Sub