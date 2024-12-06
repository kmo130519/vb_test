Option Explicit

'#########################################################
' Class: clsTimeStep
' DESC : Contains time steps. Basically time steps are days preset by the user.
'        But, if any the fixing date or dividend date is included in a period,
'        it will be seperated by the initializer.
'#########################################################

Private date_() As Date
Private dt_() As Double
Private is_fixing_date_() As Boolean
Private is_first_fixing_date_() As Boolean

Private is_dividend_date_() As Boolean

Private time_step_ As Integer

Public no_of_steps As Integer


'------------------------------------------------
' Initialize
'------------------------------------------------
Public Sub initialize(deal_ticket As clsCliquetDealTicket, div_schedule As clsDividendSchedule)

    Dim inx As Integer
    Dim first_fixing_checked As Boolean
    Dim fixing_inx As Integer
    
    Dim div_inx As Integer
    
On Error GoTo ErrorHandler
    
    ReDim date_(0 To 0) As Date
    ReDim inx_(0 To 0) As Integer
    ReDim dt_(0 To 0) As Double
    ReDim is_fixing_date_(0 To 0) As Boolean
    ReDim is_dividend_date_(0 To 0) As Boolean
    ReDim is_first_fixing_date_(0 To 0) As Boolean

        
    inx = 0
    first_fixing_checked = False
    fixing_inx = 0
    
    time_step_ = deal_ticket.time_step
    
    date_(inx) = deal_ticket.current_date
    
    fixing_inx = deal_ticket.fixing_index

    If fixing_inx > 0 Then
    
        first_fixing_checked = True
        
    End If
        
    div_inx = div_schedule.ini_div_inx
    
    Do While date_(inx) < deal_ticket.maturity_date
        
        fixing_check inx, fixing_inx, first_fixing_checked, deal_ticket
        div_check inx, div_inx, div_schedule
        
        make_next_step deal_ticket, fixing_inx, div_schedule, div_inx, inx
        
    Loop
    
    '---------------------------------------------------
    'check if the last date is fixing date
    fixing_check inx, fixing_inx, first_fixing_checked, deal_ticket
    
    no_of_steps = UBound(date_)
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsTimeStep.initialize :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
        

End Sub



Public Function get_date(inx As Integer) As Date
    
On Error GoTo ErrorHandler
        
    get_date = date_(inx)
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsTimeStep.get_date :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Function
        
End Function

Public Function dt(inx As Integer) As Double
    
On Error GoTo ErrorHandler
        
    dt = dt_(inx)
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsTimeStep.dt :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Function
    
End Function
Public Function is_first_fixing_date(inx As Integer) As Boolean
    
On Error GoTo ErrorHandler
        
    is_first_fixing_date = is_first_fixing_date_(inx)

    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsTimeStep.is_first_fixing_date :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Function
    
End Function

Public Function is_fixing_date(inx As Integer) As Boolean
    
On Error GoTo ErrorHandler
    
    is_fixing_date = is_fixing_date_(inx)
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsTimeStep.is_fixing_date :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Function
    
End Function

Public Function is_dividend_date(inx As Integer) As Boolean
    
On Error GoTo ErrorHandler
    
    is_dividend_date = is_dividend_date_(inx)
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsTimeStep.is_dividend_date :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Function
    
End Function

'------------------------------------------------
' called by Initialize
' calculate this step's values and create next step
'------------------------------------------------
Private Sub make_next_step(deal_ticket As clsCliquetDealTicket, fixing_inx As Integer, div_schedule As clsDividendSchedule, div_inx As Integer, inx As Integer)


    Dim days_per_year As Integer
    
On Error GoTo ErrorHandler
    
    '-------------------------------
    'subject to improve
    '-------------------------------
    days_per_year = DateAdd("Yyyy", 1, deal_ticket.current_date) - deal_ticket.current_date
    
        
    If deal_ticket.get_max_fixing_index < fixing_inx Then
    
        time_step_ = deal_ticket.maturity_date - date_(inx)
        dt_(inx) = time_step_ / days_per_year
        
    Else
    
        If date_(inx) + deal_ticket.time_step > deal_ticket.get_fixing_date(fixing_inx).fixing_date Then
        
            
            If deal_ticket.get_fixing_date(fixing_inx).fixing_date > div_schedule.get_dividend(div_inx).dividend_date Then
                time_step_ = div_schedule.get_dividend(div_inx).dividend_date - date_(inx)
                dt_(inx) = time_step_ / days_per_year
            Else
                time_step_ = deal_ticket.get_fixing_date(fixing_inx).fixing_date - date_(inx)
                dt_(inx) = time_step_ / days_per_year
            End If
            
        Else
            
            If date_(inx) + deal_ticket.time_step > div_schedule.get_dividend(div_inx).dividend_date Then
                time_step_ = div_schedule.get_dividend(div_inx).dividend_date - date_(inx)
                dt_(inx) = time_step_ / days_per_year
            Else
                time_step_ = deal_ticket.time_step
                dt_(inx) = time_step_ / days_per_year
            End If
            
        End If
        
    End If
    
    inx = inx + 1
    
    ReDim Preserve date_(0 To inx) As Date
    ReDim Preserve dt_(0 To inx) As Double
    ReDim Preserve is_fixing_date_(0 To inx) As Boolean
    ReDim Preserve is_first_fixing_date_(0 To inx) As Boolean
    ReDim Preserve is_dividend_date_(0 To inx) As Boolean
    
    date_(inx) = date_(inx - 1) + time_step_
    
        
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsTimeStep.make_next_step :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
    

End Sub

'------------------------------------------------
' called by Initialize
'------------------------------------------------
Private Sub fixing_check(inx As Integer, ByRef fixing_inx As Integer, ByRef first_fixing_checked As Boolean, deal_ticket As clsCliquetDealTicket)

     
On Error GoTo ErrorHandler
    
    '---------------------------------------------------------------------------
    ' If fixing index is larger than the boundary, which is the case when the current date is later than last fixing date
    
    If fixing_inx <= deal_ticket.get_max_fixing_index Then
    
            
        
        If date_(inx) = deal_ticket.get_fixing_date(fixing_inx).fixing_date Then
        
            is_fixing_date_(inx) = True
            fixing_inx = fixing_inx + 1
            
            If first_fixing_checked Then
            
                is_first_fixing_date_(inx) = False
                
            Else
                
                is_first_fixing_date_(inx) = True
                first_fixing_checked = True
                
                
            End If
            
        Else
        
            is_fixing_date_(inx) = False
            is_first_fixing_date_(inx) = False
            
        End If
        
    End If
        
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsTimeStep.fixing_check :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
    

End Sub

'------------------------------------------------
' called by Initialize
'------------------------------------------------
Private Sub div_check(inx As Integer, ByRef div_inx As Integer, div_schedule As clsDividendSchedule)
    
On Error GoTo ErrorHandler
 
    If date_(inx) = div_schedule.get_dividend(div_inx).dividend_date Then
    
        is_dividend_date_(inx) = True
        div_inx = div_inx + 1
               
    Else
    
        is_dividend_date_(inx) = False
        
    End If
        
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsTimeStep.div_check :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
    

End Sub