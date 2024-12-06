Option Explicit


Public atm_s As Double

Private strikes_() As Double
Private dates_() As Date

Private tau_() As Double

Private Const min_strike_step As Double = 5

'Public Function tau(eval_date As Date, date_index As Integer) As Double
'
'    tau = (dates_(date_index) - eval_date) / 365
'
'End Function

Public Function tau(date_index As Integer) As Double

    tau = tau_(date_index) '(dates_(date_index) - eval_date) / 365

End Function

Public Function get_all_tau() As Double()

    get_all_tau = tau_

End Function

Public Sub set_tau(tau_in() As Double)

    tau_ = tau_in

End Sub


Public Function copy_obj(Optional min_date As Date = 0) As clsPillarGrid

    Dim rtn_obj As clsPillarGrid
    
    Set rtn_obj = New clsPillarGrid
    
    rtn_obj.atm_s = atm_s
    
    Dim tmp_strikes() As Double
    Dim tmp_dates() As Date
    Dim tmp_tau() As Double
    
    Dim inx As Integer
    
    For inx = LBound(strikes_) To UBound(strikes_)
        push_back_double tmp_strikes, strikes_(inx)
    Next inx
    
    For inx = LBound(dates_) To UBound(dates_)
        If dates_(inx) > min_date Then
            push_back_date tmp_dates, dates_(inx)
        End If
    Next inx
    
    For inx = 1 To get_array_size_double(tau_)
    
        push_back_double tmp_tau, tau_(inx)
        
    Next inx
    
    rtn_obj.set_strikes tmp_strikes
    rtn_obj.set_dates tmp_dates
    rtn_obj.set_tau tmp_tau
    
    
    Set copy_obj = rtn_obj


End Function

Public Sub redim_arrays(no_of_strikes As Integer, no_of_dates As Integer)

    ReDim strikes_(1 To no_of_strikes) As Double
    ReDim dates_(1 To no_of_dates) As Date

End Sub

Public Function no_of_strikes() As Integer

    Dim rtn_value As Integer
    
On Error Resume Next

    rtn_value = UBound(strikes_)
    
    If Err.number = 9 Then
    
        rtn_value = 0
        
    End If
    
    no_of_strikes = rtn_value

End Function

Public Function no_of_dates() As Integer

    Dim rtn_value As Integer
    
On Error Resume Next

    rtn_value = UBound(dates_)
    
    If Err.number = 9 Then
    
        rtn_value = 0
        
    End If
    
    no_of_dates = rtn_value

End Function

Public Function make_strikes(spot As Double, no_of_grid As Integer, width As Double) As Double()
    
    Dim Ratio As Double
    Dim rtn_array() As Double
    Dim inx As Integer
    
    ReDim rtn_array(1 To no_of_grid) As Double
    
    Ratio = width ^ (1 / no_of_grid)
    
    For inx = 1 To no_of_grid
    
        rtn_array(inx) = spot * Ratio ^ (inx - 1 - Int(no_of_grid / 2))
    
    Next inx
    
    make_strikes = rtn_array


End Function
Public Sub initialize(current_s As Double, dates() As Date, eval_date As Date)


    Dim no_of_dates As Integer
    Dim inx As Integer
    
On Error Resume Next

    no_of_dates = UBound(dates)
    
On Error GoTo ErrorHandler
    If Err.number = 8 Then
    
        no_of_dates = 0
        raise_err "Invalid volatility surface"
        
    End If
'--------------------------------------
' Improve
'--------------------------------------

    For inx = 1 To no_of_dates
    
        'If dates(inx) >= eval_date Then
        'day_shift 적용으로 이 부분 삭제:2018.12.5
        'If dates(inx) > eval_date Then
        
            push_back_date dates_, dates(inx)
            push_back_double tau_, (dates(inx) - eval_date) / 365
            
        'End If
    
    Next inx
'
'    strikes_ = make_strikes(current_s, config__.no_of_strike_grid, config__.width_of_strike)
'
'    atm_s = Int(current_s / 5 + 0.5) * 5
'
'
'
''    dates_ = dates
'
''    push_back_double strikes_, atm_s - 10 * min_strike_step
''    push_back_double strikes_, atm_s - 4 * min_strike_step
''    push_back_double strikes_, atm_s - 2 * min_strike_step
''    push_back_double strikes_, atm_s - 1 * min_strike_step
''    push_back_double strikes_, atm_s
''    push_back_double strikes_, atm_s + 1 * min_strike_step
''    push_back_double strikes_, atm_s + 2 * min_strike_step
''    push_back_double strikes_, atm_s + 4 * min_strike_step
''    push_back_double strikes_, atm_s + 10 * min_strike_step
'
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsPillarGrid.initialize"
    

End Sub


Public Sub set_strikes(strikes_in() As Double)

    strikes_ = strikes_in

End Sub

Public Sub set_dates(dates_in() As Date, Optional shift As Integer = 0)
    
    Dim inx As Integer
    
    dates_ = dates_in
    
    If shift <> 0 Then
        For inx = 1 To get_array_size_date(dates_)
        
            dates_(inx) = dates_(inx) + shift
        
        Next inx
    End If

End Sub

Public Function get_all_strikes(Optional base As Integer = 1) As Double()

    Dim rtn_array() As Double
    
    Dim inx As Integer
    
    If no_of_strikes() > 0 Then
    
        ReDim rtn_array(base To base - 1 + no_of_strikes()) As Double
        
        
        For inx = base To base - 1 + no_of_strikes
        
            rtn_array(inx) = strikes_(inx + 1 - base)
            
        Next inx
    
    End If
    
    get_all_strikes = rtn_array


End Function


Public Function get_all_dates(Optional base As Integer = 1, Optional day_shift As Long = 0, Optional current_date As Date = 0) As Date()

    Dim dates_as_long() As Date
    Dim inx As Integer
    
    If no_of_dates() > 0 Then
    
        ReDim dates_as_long(base To base - 1 + no_of_dates()) As Date
        
        For inx = base To base - 1 + no_of_dates()
        
            dates_as_long(inx) = CLng(dates_(inx + 1 - base)) + day_shift
                        
            If current_date >= dates_as_long(inx) Then
                dates_as_long(inx) = dates_as_long(inx) + current_date - dates_as_long(inx) + 1
            End If
            
            If inx > base Then
                If dates_as_long(inx) <= dates_as_long(inx - 1) Then
                    dates_as_long(inx) = dates_as_long(inx - 1) + 1
                End If
            End If
        
        Next inx
        
    End If
    
    get_all_dates = dates_as_long

    'get_all_dates = dates_

End Function
Public Function get_all_dates_as_long(Optional base As Integer = 1, Optional day_shift As Long = 0, Optional current_date As Date = 0) As Long()

    Dim dates_as_long() As Long
    Dim inx As Integer
    
    If no_of_dates() > 0 Then
    
        ReDim dates_as_long(base To base - 1 + no_of_dates()) As Long
        
        For inx = base To base - 1 + no_of_dates()
        
            dates_as_long(inx) = CLng(dates_(inx + 1 - base)) + day_shift
            
            If current_date >= dates_as_long(inx) Then
                dates_as_long(inx) = dates_as_long(inx) + (current_date - dates_as_long(base)) + 1
            End If
            
            If inx > base Then
                If dates_as_long(inx) <= dates_as_long(inx - 1) Then
                    dates_as_long(inx) = dates_as_long(inx - 1) + 1
                End If
            End If
                    
        
        Next inx
        
    End If
    
    get_all_dates_as_long = dates_as_long

End Function


Public Function strikes(ByVal inx As Integer) As Double

    strikes = strikes_(inx)


End Function

Public Function dates(ByVal inx As Integer) As Date

    dates = dates_(inx)

End Function

Public Function get_last_date() As Date
    
    Dim rtn_date As Date
    
    rtn_date = dates_(UBound(dates_))
    
    get_last_date = rtn_date


End Function