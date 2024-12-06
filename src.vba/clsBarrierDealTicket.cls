Option Explicit



'#######################################################
' Class: clsBarrierDealTicket
'#######################################################

Public asset_code As String
Public fund_code_m As String
Public fund_code_c As String
Public ul_code As String


Public current_date As Date
Public current_date_origin_ As Date

Public value_date As Date
Public Maturity As Integer
Public maturity_date As Date
Public settlement_date As Date

Public confirmed_yn As String
Public alive_yn As String

Public issue_cost As Double
Public reference As Double

Public call_put As Long
Public strike As Double
Public barrier_type As Long
Public barrier As Double
Public rebate As Double
Public barrier_shift As Double
Public barrier_monitoring As Integer

Public x_grid As Long
Public v_grid As Long
Public t_grid As Long

Public scheme_type As Long

Public comment As String

Public quantity As Double

Public funding_spread As Double

Public rebate_only As Long  '1: one touch option, 0:Barrier option

Public participation_rate As Double

Public instrument_type As Long

Private func As New Functions

Private autocall_schedules_() As clsAutocallSchedule
Private floating_schedules_() As clsAutocallSchedule '<--- floating coupon schedule. Should include first fixing date which is equal to value date

Public no_of_schedule As Integer
 
Public Function floating_schedules(inx As Integer) As clsAutocallSchedule

    Set floating_schedules = floating_schedules_(inx)

End Function
 Public Function autocall_schedules(inx As Integer) As clsAutocallSchedule

    Set autocall_schedules = autocall_schedules_(inx)

End Function


Public Sub set_floating_schedule(no_of_schedule_in As Integer, coupon_date() As Date, fixing_value() As Double)

    Dim a_schedule As clsAutocallSchedule
    Dim inx As Integer
    
On Error GoTo ErrorHandler
    
    If no_of_schedule_in >= 1 Then

        ReDim floating_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
        no_of_schedule = no_of_schedule_in
'    Else
'
'        raise_err "clsACDealTicket.setSchedule", "No schedule found"

    
'
    
        For inx = 1 To no_of_schedule
            
            Set floating_schedules_(inx) = New clsAutocallSchedule
            
            floating_schedules_(inx).call_date = coupon_date(inx)
            floating_schedules_(inx).set_coupon_on_call fixing_value(inx)
        
        Next inx
    
    End If
    
    Exit Sub


ErrorHandler:

    raise_err "clsACDealTicket.set_floating_schedule"


End Sub

Public Sub set_schedule(no_of_schedule_in As Integer, calldate() As Date, strike_values() As Double, coupon_on_call() As Double)

    Dim a_schedule As clsAutocallSchedule
    Dim inx As Integer
    
On Error GoTo ErrorHandler
    
    If no_of_schedule_in >= 1 Then

        ReDim autocall_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
        no_of_schedule = no_of_schedule_in
'    Else
'
'        raise_err "clsACDealTicket.setSchedule", "No schedule found"

    
'
    
        For inx = 1 To no_of_schedule
            
            Set autocall_schedules_(inx) = New clsAutocallSchedule
            
            autocall_schedules_(inx).call_date = calldate(inx)
            autocall_schedules_(inx).set_coupon_on_call coupon_on_call(inx)
            autocall_schedules_(inx).set_strike_value strike_values(inx)
        
        Next inx
    
    End If
    
    Exit Sub


ErrorHandler:

    raise_err "clsACDealTicket.setSchedule"


End Sub



Public Function floating_coupon_dates() As Long()

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(floating_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number <> 9 And no_of_schedule_in > 0 Then
        
        ReDim rtn_array(1 To no_of_schedule_in) As Long
        
        For inx = 1 To no_of_schedule_in
        
            rtn_array(inx) = CLng(floating_schedules_(inx).call_date)
        
        Next inx
        
    Else
        
        ReDim rtn_array(1 To 1) As Long
        
        rtn_array(1) = CLng(#12/31/2100#) ' Dummy
    
    End If
    
    floating_coupon_dates = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsBarrierDealTicket.call_dates"

End Function
Public Function call_dates() As Long()

    Dim rtn_array() As Long
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number <> 9 And no_of_schedule_in > 0 Then
        
        ReDim rtn_array(1 To no_of_schedule_in) As Long
        
        For inx = 1 To no_of_schedule_in
        
            rtn_array(inx) = CLng(autocall_schedules_(inx).call_date)
        
        Next inx
        
    Else
        
        ReDim rtn_array(1 To 1) As Long
        
        rtn_array(1) = CLng(#12/31/2100#) ' Dummy
    
    End If
    
    call_dates = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsBarrierDealTicket.call_dates"

End Function

Public Function strike_values() As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number <> 9 And no_of_schedule_in > 0 Then
            
        
        ReDim rtn_array(1 To no_of_schedule_in) As Double
        
        For inx = 1 To no_of_schedule_in
        
            rtn_array(inx) = autocall_schedules_(inx).strike_value
        
        Next inx
    
    Else
        ReDim rtn_array(1 To 1) As Double
        
        rtn_array(1) = 0
        
    End If
        
    strike_values = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsBarrierDealTicket.strike_values"


End Function

Public Function floating_fixing_values() As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(floating_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number <> 9 And no_of_schedule_in > 0 Then
                
        ReDim rtn_array(1 To no_of_schedule_in) As Double
        
        For inx = 1 To no_of_schedule_in
        
            rtn_array(inx) = floating_schedules_(inx).coupon_on_call
        
        Next inx
        
    Else
        ReDim rtn_array(1 To 1) As Double
        
        rtn_array(1) = 0
        
    End If
    
    floating_fixing_values = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsBarrierDealTicket.coupon_on_calls"


End Function
'Public current_date_orig

Public Function coupon_on_calls() As Double()


    Dim rtn_array() As Double
    Dim inx As Integer
    Dim no_of_schedule_in As Integer
    
    On Error Resume Next
    
    no_of_schedule_in = UBound(autocall_schedules_)
    
On Error GoTo ErrorHandler
    
    If Err.number <> 9 And no_of_schedule_in > 0 Then
                
        ReDim rtn_array(1 To no_of_schedule_in) As Double
        
        For inx = 1 To no_of_schedule_in
        
            rtn_array(inx) = autocall_schedules_(inx).coupon_on_call
        
        Next inx
        
    Else
        ReDim rtn_array(1 To 1) As Double
        
        rtn_array(1) = 0
        
    End If
    
    coupon_on_calls = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsBarrierDealTicket.coupon_on_calls"


End Function
'Public current_date_origin_ As Date

Public Function copy_obj() As clsCliquetDealTicket

    Dim obj As clsBarrierDealTicket
    
    Set obj = New clsCliquetDealTicket
    
    obj.fund_code_m = fund_code_m
    obj.fund_code_c = fund_code_c
    obj.asset_code = asset_code
    obj.ul_code = ul_code
    

    obj.current_date = current_date
    obj.current_date_origin_ = current_date_origin_
    
    obj.value_date = value_date
    obj.Maturity = Maturity
    obj.maturity_date = maturity_date
    obj.settlement_date = settlement_date
    

    obj.confirmed_yn = confirmed_yn
    obj.alive_yn = alive_yn
    
    obj.issue_cost = issue_cost
    obj.reference = reference
    
    obj.call_put = call_put
    obj.strike = strike
    obj.barrier_type = barrier_type
    obj.barrier = barrier
    obj.rebate = rebate
    obj.barrier_monitoring = barrier_monitoring
    
    obj.x_grid = x_grid
    obj.v_grid = v_grid
    obj.t_grid = t_grid
    
    obj.comment = comment
    
    obj.quantity = quantity
    
   ' obj.set_floating_schedule UBound(floating_schedules_)
   
   obj.copy_schedule UBound(autocall_schedules_), autocall_schedules_
   obj.copy_floating_schedule UBound(floating_schedules_), floating_schedules_
        
    Set copy_obj = obj

End Function


Public Sub copy_schedule(no_of_schedule_in As Integer, schs() As clsAutocallSchedule)

    Dim inx As Integer
    
    ReDim autocall_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
    no_of_schedule = no_of_schedule_in
    
    For inx = 1 To no_of_schedule
            
        Set autocall_schedules_(inx) = New clsAutocallSchedule
        
        autocall_schedules_(inx).call_date = schs(inx).call_date
        autocall_schedules_(inx).set_coupon_on_call schs(inx).coupon_on_call
        autocall_schedules_(inx).set_percent_strike schs(inx).percent_strike
        autocall_schedules_(inx).set_strike_value autocall_schedules_(inx).strike_value  ' * reference_price
    
    Next inx
    
    

End Sub


Public Sub copy_floating_schedule(no_of_schedule_in As Integer, schs() As clsAutocallSchedule)

    Dim inx As Integer
    
    ReDim floating_schedules_(1 To no_of_schedule_in) As clsAutocallSchedule
    no_of_schedule = no_of_schedule_in
    
    For inx = 1 To no_of_schedule
            
        Set floating_schedules_(inx) = New clsAutocallSchedule
        
        floating_schedules_(inx).call_date = schs(inx).call_date
        floating_schedules_(inx).set_coupon_on_call schs(inx).coupon_on_call
        floating_schedules_(inx).set_percent_strike schs(inx).percent_strike
        floating_schedules_(inx).set_strike_value floating_schedules_(inx).strike_value  ' * reference_price
    
    Next inx
    
    

End Sub