Option Explicit

Private dividend_schedule_() As clsDividend
Public ratioDividend As Double '2018.7.19

Public ini_div_inx As Integer
Public div_inx As Integer

Private func As New Functions

Public Function get_sum_of_div(maturity_date As Date, current_date As Date) As Double

    Dim inx As Integer
    Dim rtn_value As Double
    inx = 0
    
    Do While inx <= UBound(dividend_schedule_)
    
        If dividend_schedule_(inx).dividend_date <= maturity_date Then
            
            If dividend_schedule_(inx).dividend_date > current_date Then
                
                rtn_value = rtn_value + dividend_schedule_(inx).dividend
            
            End If
            
        Else
        
            Exit Do

        End If
        
        inx = inx + 1
    
    Loop
    
    get_sum_of_div = rtn_value

End Function


Public Function get_sum_of_div_pvs(maturity_date As Date, current_date As Date, rate_curve As clsRateCurve) As Double

    Dim inx As Integer
    Dim rtn_value As Double
    inx = 0
    
    Do While inx <= UBound(dividend_schedule_)
    
        If dividend_schedule_(inx).dividend_date <= maturity_date Then
            
            If dividend_schedule_(inx).dividend_date > current_date Then
                
                rtn_value = rtn_value + rate_curve.get_discount_factor(current_date, dividend_schedule_(inx).dividend_date) * dividend_schedule_(inx).dividend
            
            End If
            
        Else
        
            Exit Do

        End If
        
        inx = inx + 1
    
    Loop
    
    get_sum_of_div_pvs = rtn_value

End Function

'Public Function get_div_dates(Optional base As Integer = 1, Optional max_date As Date = #12/31/2100#, Optional ul_code As String) As Long()
'day_shift 추가 : 2018.12.05
Public Function get_div_dates(Optional base As Integer = 1, Optional max_date As Date = #12/31/2100#, Optional ul_code As String, Optional day_shift As Long = 0) As Long()

    Dim inx As Integer
    Dim rtn_array() As Long
    
    'ReDim rtn_array(base To UBound(dividend_schedule_) - LBound(dividend_schedule_) + base) As Long
    
    For inx = base To UBound(dividend_schedule_) - LBound(dividend_schedule_) + base
        
        If dividend_schedule_(inx + LBound(dividend_schedule_) - base).dividend_date < max_date Then
            'push_back_long rtn_array, dividend_schedule_(inx + LBound(dividend_schedule_) - base).dividend_date, 0
            'day_shift 추가 : 2018.12.05
            push_back_long rtn_array, CLng(dividend_schedule_(inx + LBound(dividend_schedule_) - base).dividend_date) + day_shift, 0
            'rtn_array(inx) = dividend_schedule_(inx + LBound(dividend_schedule_) - base).dividend_date
        End If
    
    Next inx
    
'    For inx = 0 To Sheet2.Range(ul_code + "_DIV").Cells(1, 5) - 1
'
'    If Sheet2.Range(ul_code + "_DIV").Cells(inx + 1, 3) < max_date Then
'    push_back_long rtn_array, Sheet2.Range(ul_code + "_DIV").Cells(inx + 1, 3), 0
'    End If
'
'    Next inx
    
    get_div_dates = rtn_array

End Function

Public Function get_divs(Optional base As Integer = 1, Optional max_date As Date = #12/31/2100#, Optional ul_code As String) As Double()

    Dim inx As Integer
    Dim rtn_array() As Double
    
    'ReDim rtn_array(base To UBound(dividend_schedule_) - LBound(dividend_schedule_) + base) As Double
    
    For inx = base To UBound(dividend_schedule_) - LBound(dividend_schedule_) + base
        If dividend_schedule_(inx + LBound(dividend_schedule_) - base).dividend_date < max_date Then
        
            push_back_double rtn_array, dividend_schedule_(inx + LBound(dividend_schedule_) - base).dividend, 0
            'rtn_array(inx) = dividend_schedule_(inx + LBound(dividend_schedule_) - base).dividend
        End If
    Next inx
    
'    For inx = 0 To Sheet2.Range(ul_code + "_DIV").Cells(1, 5) - 1
'
'    If Sheet2.Range(ul_code + "_DIV").Cells(inx + 1, 3) < max_date Then
'        push_back_double rtn_array, Val(Sheet2.Range(ul_code + "_DIV").Cells(inx + 1, 4)), 0
'    End If
'
'    Next inx
    
    get_divs = rtn_array

End Function

'set the div index to be the first div date.
Public Function initialize_div_inx(current_date As Date) As Boolean
    
    Dim inx As Integer
    
On Error GoTo ErrorHandler

    inx = 0
    
    If Not IsNull(dividend_schedule_) Then
    
        Do While dividend_schedule_(inx).dividend_date <= current_date
        
            inx = inx + 1
        
        Loop
        
    End If
        
    ini_div_inx = inx
    div_inx = inx
        
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsDividendSchedule.initialize_div_inx :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Function
        
End Function

Public Sub rewind_fixing_inx()

    div_inx = ini_div_inx

End Sub

Public Function copy_obj() As clsDividendSchedule

    Dim obj As clsDividendSchedule
    Dim inx As Integer
    
    Set obj = New clsDividendSchedule
    
    obj.div_inx = div_inx
    obj.ini_div_inx = ini_div_inx
    
    For inx = LBound(dividend_schedule_) To UBound(dividend_schedule_)
    
        obj.add_dividend_schedule dividend_schedule_(inx)
    
    Next inx

    Set copy_obj = obj

End Function

Public Function get_dividend(Optional index As Integer = -1) As clsDividend
    
    Dim temp_div As New clsDividend
    
On Error GoTo ErrorHandler
    
    If index = -1 Then
    
        index = div_inx
        
    End If
    
    If index > UBound(dividend_schedule_) Then
        
        temp_div.dividend = 0
        temp_div.dividend_date = #12/31/2100#
        
        Set get_dividend = temp_div
        
    Else
            
        Set get_dividend = dividend_schedule_(index)
        
    End If
    
       
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsDividendSchedule.get_dividend :" & Chr(13) & Err.source, Err.description '1002 <-- run error
    Exit Function

End Function



Public Function move_next() As Boolean

    div_inx = div_inx + 1
    
    If div_inx > UBound(dividend_schedule_) Then
    
        move_next = False
        
    Else
        
        move_next = True
        
    End If

End Function


Public Sub initialize_div(no_of_schedule As Integer, schedule() As Date, dividend_amount() As Double)

    Dim inx As Integer
    Dim temp_schedule As clsDividend
    
    
On Error GoTo ErrorHandler
        
    
    For inx = 1 To no_of_schedule
    
        Set temp_schedule = New clsDividend
        
        temp_schedule.dividend_date = schedule(inx)
        temp_schedule.dividend = dividend_amount(inx)
        
        add_dividend_schedule temp_schedule
        
    Next inx
    
    Set temp_schedule = New clsDividend
    
    temp_schedule.dividend = 0
    temp_schedule.dividend_date = #12/31/2100#
    
    add_dividend_schedule temp_schedule
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsDividendSchedule.initialize_div :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
            
    
End Sub

Public Sub add_dividend_schedule(a_dividend As clsDividend)

    
    
On Error GoTo ErrorHandler
        

    If IsNull(dividend_schedule_) Then
        ReDim dividend_schedule_(1 To 1) As clsDividend
        Set dividend_schedule_(UBound(dividend_schedule_)) = a_dividend
    Else
        ReDim Preserve dividend_schedule_(LBound(dividend_schedule_) To UBound(dividend_schedule_) + 1) As clsDividend
        Set dividend_schedule_(UBound(dividend_schedule_)) = a_dividend
    End If
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsDividendSchedule.add_dividend_schedule :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
                

End Sub