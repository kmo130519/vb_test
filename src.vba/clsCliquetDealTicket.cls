Option Explicit

'#######################################################
' Class: clsCliquetDealTicket
' This class contains information on the cliquet deal.
' Fixing schedules and fixing values, if any, are also included.
'#######################################################

Public asset_code As String
Public fund_code_m As String
Public fund_code_c As String
Public ul_code As String

Public fixing_index As Integer '0 if no fixing has occured. Indicates next fixing index.

Private fixing_schedule_() As clsCliquetFixing

Public current_date As Date
Public value_date As Date
Public Maturity As Integer
Public maturity_date As Date
Public settlement_date As Date
Public reset_freq As Integer

Public ini_fixing_inx As Integer

Public global_cap As Double
Public global_floor As Double

Public confirmed_yn As String
Public alive_yn As String


Public time_step As Integer
Public no_of_trials As Long
Public required_tolerance As Double

Public last_fixing As Double
Public previous_performance As Double

Public comment As String

Public notional As Double

Public spread As Double
'---------------------------
Public issue_cost As Double
'---------------------------
Public bump_vega As String
Public bump_theta As String
Public bump_skew As String

Public day_fraction_ As Double

Private period_performances_() As Double

Public cliquet_type As String

Public r_cliquet_cap As Double
Public replace_period_no As Integer


Private func As New Functions
 
Public current_date_origin_ As Date

Public Function copy_obj() As clsCliquetDealTicket

    Dim obj As clsCliquetDealTicket
    Dim inx As Integer
    
    Set obj = New clsCliquetDealTicket
    
    obj.fixing_index = fixing_index
    
    For inx = LBound(fixing_schedule_) To UBound(fixing_schedule_)
    
        obj.add_fixing_schedule fixing_schedule_(inx).copy_obj()
        'obj.add_fixing_schedule fixing_schedule_(inx)
    
    Next inx
    
    obj.fund_code_m = fund_code_m
    obj.fund_code_c = fund_code_c
    obj.asset_code = asset_code
    obj.notional = notional
    
    
    obj.current_date = current_date
    obj.current_date_origin_ = current_date_origin_
    obj.value_date = value_date
    obj.Maturity = Maturity
    obj.maturity_date = maturity_date
    obj.settlement_date = settlement_date
    obj.reset_freq = reset_freq
    

    
    
    
    obj.ini_fixing_inx = ini_fixing_inx
    
    obj.global_cap = global_cap
    obj.global_floor = global_floor
    
    obj.time_step = time_step
    obj.no_of_trials = no_of_trials
    obj.required_tolerance = required_tolerance
    
    obj.last_fixing = last_fixing
    obj.previous_performance = previous_performance
    
    
    obj.bump_vega = bump_vega
    obj.bump_skew = bump_skew
    obj.bump_theta = bump_theta
    
    obj.cliquet_type = cliquet_type
    obj.r_cliquet_cap = r_cliquet_cap
    obj.replace_period_no = replace_period_no
    
    obj.spread = spread
    '--------------------------
    obj.issue_cost = issue_cost
    '--------------------------
    obj.initialize
    
    obj.day_fraction_ = day_fraction_
        
'    For inx = get_min_period_performance() To get_max_period_performance()
'
'        obj.add_period_performance get_preriod_performance(inx)
'
'
'    Next inx
    
    'obj
        
    Set copy_obj = obj

End Function

Public Sub add_period_performance(aPerformance As Double)

    If IsNull(period_performances_) Or UBound(period_performances_) <= 0 Then
        ReDim period_performances_(1 To 1) As Double
        period_performances_(UBound(period_performances_)) = aPerformance
    Else
        ReDim Preserve period_performances_(LBound(period_performances_) To UBound(period_performances_) + 1) As Double
        period_performances_(UBound(period_performances_)) = aPerformance
    End If
    

End Sub



Public Function get_preriod_performance(ByVal inx As Integer) As Double


    get_preriod_performance = period_performances_(inx)


End Function

Public Function get_max_period_performance() As Integer

    get_max_period_performance = UBound(period_performances_)

End Function
Public Function get_min_period_performance() As Integer

    get_min_period_performance = LBound(period_performances_)

End Function

'set the fixing index to be the first fixing date which is not fixed yet.
Public Function initialize(Optional default_fixing_value As Double) As Boolean
    
    Dim inx As Integer
    
    inx = 0
    
On Error GoTo ErrorHandler
    
    If IsNull(fixing_schedule_) Then
    
        Err.Raise vbObjectError + 1002, "clsCliquetDealTicket", "Not initialized"
    
    End If
    
'    Do While fixing_schedule_(inx).fixing_date < current_date
'
'        ' If the fixing value is not specified for a past fixing date...
'        If fixing_schedule_(inx).fixing_value <= 0 Then
'
'            fixing_schedule_(inx).fixing_value = default_fixing_value
'
'        End If
'
'        inx = inx + 1
'
'    Loop
    
'    ' If today is a fixing date and the fixing value for today is already specified, increase fixing index to indicate the next one
'    If fixing_schedule_(inx).fixing_date = current_date And fixing_schedule_(inx).fixing_value > 0 Then
'
'        inx = inx + 1
'
'    End If
    
    For inx = 0 To get_max_fixing_index
    
        If fixing_schedule_(inx).fixing_date > current_date Then
        
            Exit For
            
        ElseIf fixing_schedule_(inx).fixing_date = current_date Then
        
            If fixing_schedule_(inx).fixing_value > 0 Then
            
                inx = inx + 1
                
            End If
            
            Exit For
            
        End If
        
        
        If fixing_schedule_(inx).fixing_value <= 0 Then
        
            fixing_schedule_(inx).fixing_value = default_fixing_value
            
        End If
    
    Next inx
                
    
    fixing_index = inx
    ini_fixing_inx = inx
    
    ReDim period_performances_(1 To get_max_fixing_index) As Double
    
    calculate_previous_performance
    
    Exit Function
    
ErrorHandler:

    MsgBox "Initialization failure in clsCliquetDealTicket" & Err.description
    Exit Function
            
End Function

Public Sub move_current_date(difference As Integer)

    current_date = current_date + difference

End Sub

Public Function get_local_floors() As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_fixing_schedule As Integer
        
    
On Error GoTo ErrorHandler
    
    no_of_fixing_schedule = UBound(fixing_schedule_)
    
    If fixing_index <= no_of_fixing_schedule Then
    
        ReDim rtn_array(0 To no_of_fixing_schedule - fixing_index) As Double
    
        For inx = fixing_index To no_of_fixing_schedule
        
            rtn_array(inx - fixing_index) = fixing_schedule_(inx).local_floor
        
        Next inx
    
    End If
    
    
    get_local_floors = rtn_array
    

    Exit Function
    
ErrorHandler:

    raise_err "clsCliquetDealTicket.get_local_floors"

End Function

Public Function get_local_caps() As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_fixing_schedule As Integer
        
    
On Error GoTo ErrorHandler
    
    no_of_fixing_schedule = UBound(fixing_schedule_)
    
    If fixing_index <= no_of_fixing_schedule Then
    
        ReDim rtn_array(0 To no_of_fixing_schedule - fixing_index) As Double
    
        For inx = fixing_index To no_of_fixing_schedule
        
            rtn_array(inx - fixing_index) = fixing_schedule_(inx).local_cap
        
        Next inx
    
    End If
    
    
    get_local_caps = rtn_array
    

    Exit Function
    
ErrorHandler:

    raise_err "clsCliquetDealTicket.get_local_caps"


End Function

Public Function get_reset_dates() As Long()

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_fixing_schedule As Integer
        
    
On Error GoTo ErrorHandler
    
    no_of_fixing_schedule = UBound(fixing_schedule_)
    
    If fixing_index <= no_of_fixing_schedule Then
    
        ReDim rtn_array(0 To no_of_fixing_schedule - fixing_index) As Long
    
        For inx = fixing_index To no_of_fixing_schedule
        
            rtn_array(inx - fixing_index) = fixing_schedule_(inx).fixing_date
        
        Next inx
    
    End If
    
    
    get_reset_dates = rtn_array
    

    Exit Function
    
ErrorHandler:

    raise_err "clsCliquetDealTicket.get_reset_dates"


End Function

Public Sub rewind_date()

    current_date = current_date_origin_

End Sub

Public Sub rewind_fixing_inx()

    fixing_index = ini_fixing_inx

End Sub

Private Sub calculate_previous_performance()

    Dim inx As Integer
    
    previous_performance = 0
    inx = 1
    
    'If fixing ever happened
    If fixing_index > 0 Then
    
        ' Find last fixing and claculate previous performance...
        Do While inx <= fixing_index - 1
        
            If fixing_schedule_(inx).fixing_value > 0 Then
            
                period_performances_(inx) = func.min(func.max(fixing_schedule_(inx).fixing_value / fixing_schedule_(inx - 1).fixing_value - 1 _
                                                                          , fixing_schedule_(inx).local_floor) _
                                                                    , fixing_schedule_(inx).local_cap)
        
                previous_performance = previous_performance + func.min(func.max(fixing_schedule_(inx).fixing_value / fixing_schedule_(inx - 1).fixing_value - 1 _
                                                                          , fixing_schedule_(inx).local_floor) _
                                                                    , fixing_schedule_(inx).local_cap)
                                                                    
                inx = inx + 1
                
            Else
            
                Exit Do
                
            End If
        
        Loop
        
        ' Set the last fixing value
        last_fixing = fixing_schedule_(fixing_index - 1).fixing_value
        
    End If

End Sub

Public Function get_fixing_date(Optional index As Integer = -1) As clsCliquetFixing
    
    If index = -1 Then
    
        index = fixing_index
        
    End If
    
    If index > UBound(fixing_schedule_) Then
    
        Set get_fixing_date = Nothing
        
    Else
    
        Set get_fixing_date = fixing_schedule_(index)
        
    End If

End Function

Public Function get_max_fixing_index() As Integer

    get_max_fixing_index = UBound(fixing_schedule_)

End Function


Public Sub move_next()

    fixing_index = fixing_index + 1

End Sub


Public Sub set_fixing_schedule(no_of_schedule As Integer, schedule() As Date, floors() As Double, caps() As Double, fixing_values() As Double)

    Dim inx As Integer
    Dim temp_schedule As clsCliquetFixing
    
    For inx = 0 To no_of_schedule - 1
    
        Set temp_schedule = New clsCliquetFixing
        
        temp_schedule.fixing_date = schedule(inx)
        temp_schedule.local_floor = floors(inx)
        temp_schedule.local_cap = caps(inx)
        temp_schedule.fixing_value = fixing_values(inx)
        
        add_fixing_schedule temp_schedule
        
    Next inx
        
    current_date_origin_ = current_date
    
End Sub

Public Sub add_fixing_schedule(a_fixing_schedule As clsCliquetFixing)

    If IsNull(fixing_schedule_) Then
        Set fixing_schedule_(UBound(fixing_schedule_)) = a_fixing_schedule
    Else
        ReDim Preserve fixing_schedule_(LBound(fixing_schedule_) To UBound(fixing_schedule_) + 1) As clsCliquetFixing
        Set fixing_schedule_(UBound(fixing_schedule_)) = a_fixing_schedule
    End If
    

End Sub

Public Sub set_fixing_value(inx As Integer, fixing_value As Double)

    fixing_schedule_(inx).fixing_value = fixing_value

End Sub


Public Function get_local_floor_cap(ByRef local_floor As Double, ByRef local_cap As Double) As Boolean

On Error GoTo ErrorHandler

    Dim inx As Integer
    Dim homogeneous As Boolean
    Dim first_cap As Double
    Dim first_floor As Double
    
    homogeneous = True
    
    local_floor = 0
    local_cap = 0
    
    first_floor = fixing_schedule_(1).local_floor
    first_cap = fixing_schedule_(1).local_cap
    
    For inx = 1 To get_max_fixing_index
    
'        local_floor = local_floor + fixing_schedule_(inx).local_floor
'        local_cap = local_cap + fixing_schedule_(inx).local_cap
'
    
        If fixing_schedule_(inx).local_cap <> first_cap Then
        
            homogeneous = False
        
        End If
        
        If fixing_schedule_(inx).local_floor <> first_floor Then
        
            homogeneous = False
        
        End If
    
    Next inx
    
'    local_floor = local_floor / get_max_fixing_index
'    local_cap = local_cap / get_max_fixing_index
    
    If get_max_fixing_index >= fixing_index Then
        local_floor = fixing_schedule_(fixing_index).local_floor
        local_cap = fixing_schedule_(fixing_index).local_cap
    End If
        
    get_local_floor_cap = homogeneous

    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 10000, "get_local_floor_cap : " & Chr(13) & Err.source, Err.description

End Function


Public Sub get_fixing_status(ByRef prev_fixing_value As Double, ByRef next_fixing_date As Date, ByRef remaining_fixing_period As Integer, _
                                  Optional ByVal cur_date As Date = -1)

    Dim inx As Integer
    
On Error GoTo ErrorHandler

    inx = 0

    If cur_date < 0 Then
    
        cur_date = current_date
        
    End If
    
    If fixing_schedule_(0).fixing_date > cur_date Then
    
        next_fixing_date = fixing_schedule_(0).fixing_date
        prev_fixing_value = -1000
        remaining_fixing_period = get_max_fixing_index + 1
        
    Else
    
    
        For inx = 1 To get_max_fixing_index
        
            If fixing_schedule_(inx).fixing_date = cur_date And fixing_schedule_(inx).fixing_value <= 0 Then
            
                Exit For
                
            ElseIf fixing_schedule_(inx).fixing_date > cur_date Then
            
                Exit For
                
            End If
            
        Next inx
        
        
        If inx > get_max_fixing_index Then
        
            next_fixing_date = -1
            
        Else
        
            next_fixing_date = fixing_schedule_(inx).fixing_date
            
        End If
        
        prev_fixing_value = fixing_schedule_(inx - 1).fixing_value
        remaining_fixing_period = get_max_fixing_index - inx + 1
        
    End If

    Exit Sub
    
ErrorHandler:
    
    Err.Raise vbObjectError + 10000, "get_fixing_status : " & Chr(13) & Err.source, Err.description

End Sub