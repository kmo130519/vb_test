Option Explicit

Public ErrObject As New clsError

Public business_days_list() As Date




'Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Function get_process_id()
    
    get_process_id = GetCurrentProcessId

End Function

Public Function yyyymmdd_to_date(yyyymmdd As String) As Date

    Dim rtn_date As Date
    
On Error GoTo ErrorHandler

    If CDec(yyyymmdd) < 0 Then
    
        Err.Raise vbObjectError + 1000, , "Invalid date type"
        
    End If
    
    If Len(yyyymmdd) <> 8 Then
    
        Err.Raise vbObjectError + 1000, , "Invalid date type"
        
    End If
        
    
    rtn_date = CDate(Left(yyyymmdd, 4) & "/" & Mid(yyyymmdd, 5, 2) & "/" & Right(yyyymmdd, 2))
    
    yyyymmdd_to_date = rtn_date
    
    Exit Function
    
ErrorHandler:

   raise_err "yyyymmdd_to_date:", Err.description

End Function

Public Function linear_interpolation_form(x1 As Double, x2 As Double, y1 As Double, y2 As Double, x As Double) As Double

    Dim y As Double
    
    y = ((x1 - x) * y2 + (x - x2) * y1) / (x1 - x2)
    
    linear_interpolation_form = y

End Function

Public Function tqbilin(x() As Double, y() As Double, z() As Double, p() As Double) As Double
    
    Dim zp(1) As Double

    zp(0) = lin(x(0), z(0), x(2), z(2), p(0))
    zp(1) = lin(x(1), z(1), x(3), z(3), p(0))
    tqbilin = Sqr(lin(y(0), zp(0) ^ 2, y(1), zp(1) ^ 2, p(1)))

End Function

Public Function bilin(x() As Double, y() As Double, z() As Double, p() As Double) As Double
    
    Dim zp(1) As Double

    zp(0) = lin(x(0), z(0), x(2), z(2), p(0))
    zp(1) = lin(x(1), z(1), x(3), z(3), p(0))
    bilin = lin(y(0), zp(0), y(1), zp(1), p(1))

End Function

Public Function lin(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x As Double) As Double

    lin = (y2 - y1) / (x2 - x1) * (x - x1) + y1

End Function

Public Function find_location_date(abscissa() As Date, x As Date) As Integer

    Dim inx_up As Integer
    Dim inx_low As Integer
    Dim inx_mid As Integer

    Dim rtn_inx As Integer

On Error GoTo ErrorHandler

    inx_low = LBound(abscissa) - 1
    inx_up = UBound(abscissa) + 1

    Do While inx_up - inx_low > 1

        inx_mid = Int((inx_up + inx_low) / 2)

        If x >= abscissa(inx_mid) Then

            inx_low = inx_mid

        Else

            inx_up = inx_mid

        End If

    Loop


    If x = abscissa(LBound(abscissa)) Then

        rtn_inx = LBound(abscissa)

    ElseIf x = abscissa(UBound(abscissa)) Then

        rtn_inx = UBound(abscissa) - 1

    Else

        rtn_inx = inx_low

    End If


    find_location_date = rtn_inx

    Exit Function

ErrorHandler:

    Err.source = "find_location" & Err.source
'    DBDisConnector
    Err.Raise Err.number, Err.source

 End Function
 
 


Public Function find_location(abscissa() As Double, x As Double) As Integer

    Dim inx_up As Integer
    Dim inx_low As Integer
    Dim inx_mid As Integer
    
    Dim rtn_inx As Integer
    
On Error GoTo ErrorHandler
    
    inx_low = LBound(abscissa) - 1
    inx_up = UBound(abscissa) + 1
    
    Do While inx_up - inx_low > 1
    
        inx_mid = Int((inx_up + inx_low) / 2)
        
        If x >= abscissa(inx_mid) Then
        
            inx_low = inx_mid
        
        Else
            
            inx_up = inx_mid
            
        End If

    Loop
    
    
    If x = abscissa(LBound(abscissa)) Then
    
        rtn_inx = LBound(abscissa)
        
    ElseIf x = abscissa(UBound(abscissa)) Then
        
        rtn_inx = UBound(abscissa) - 1
        
    Else
    
        rtn_inx = inx_low
    
    End If
    
    
    find_location = rtn_inx
    
    Exit Function
    
ErrorHandler:

    Err.source = "find_location" & Err.source
'    DBDisConnector
    Err.Raise Err.number, Err.source
    
 End Function
 Public Function find_location_long(abscissa() As Long, x As Long) As Integer

    Dim inx_up As Integer
    Dim inx_low As Integer
    Dim inx_mid As Integer
    
    Dim rtn_inx As Integer
    
On Error GoTo ErrorHandler
    
    inx_low = LBound(abscissa) - 1
    inx_up = UBound(abscissa) + 1
    
    Do While inx_up - inx_low > 1
    
        inx_mid = Int((inx_up + inx_low) / 2)
        
        If x >= abscissa(inx_mid) Then
        
            inx_low = inx_mid
        
        Else
            
            inx_up = inx_mid
            
        End If

    Loop
    
    
    If x = abscissa(LBound(abscissa)) Then
    
        rtn_inx = LBound(abscissa)
        
    ElseIf x = abscissa(UBound(abscissa)) Then
        
        rtn_inx = UBound(abscissa) - 1
        
    Else
    
        rtn_inx = inx_low
    
    End If
    
    
    find_location_long = rtn_inx
    
    Exit Function
    
ErrorHandler:

    Err.source = "find_location" & Err.source
'    DBDisConnector
    Err.Raise Err.number, Err.source
    
 End Function
Public Function find_location_wksheet(abscissa As Range, x As Double) As Integer

    Dim abscissa_arr() As Double
    
    abscissa_arr = range_to_array(abscissa, 1)
    
    Dim rtn_inx As Integer
    
On Error GoTo ErrorHandler
    
    rtn_inx = find_location(abscissa_arr, x)
    
    
    find_location_wksheet = rtn_inx
    
    Exit Function
    
ErrorHandler:

    Err.source = "find_location" & Err.source
'    DBDisConnector
    Err.Raise Err.number, Err.source
    
 End Function

'#####################
' Error Message
'#####################

Public Function raise_err(function_name As String, Optional description As String = "") As ErrObject

    If UCase(Left(Err.description, 5)) = "[BIZ]" Then
    
        MsgBox ErrObject.description
        Resume Next
    
    End If
    
    If ErrObject Is Nothing Then
        ErrObject = New clsError
    End If

   ErrObject.setError ErrObject.number, function_name & Chr(13) & ErrObject.source, description & Chr(13) & ErrObject.description
       Err.Raise ErrObject.number
    
End Function
'
'Public Sub temp()
'
'    Dim a As Date
'
'    a = 0
'
'
'
'End Sub

    
Public Function add_month(in_date As Date, Optional months As Integer = 1) As Date

On Error GoTo ErrorHandler

    Dim rtn_date As Date
    
    rtn_date = DateAdd("m", months, in_date)
    add_month = rtn_date
    
    Exit Function
    
ErrorHandler:
    
    raise_err "add_month"

End Function

Public Function check_dynamic_array_double(an_array() As Double) As Boolean

    Dim rtn_value As Boolean
    Dim inx As Integer
    
    rtn_value = True
    
On Error Resume Next
    
    inx = LBound(an_array)
    
    If Err.number = 9 Then
    
        rtn_value = False
        
    Else
    
        rtn_value = True
        
    End If
    
    check_dynamic_array_double = rtn_value


End Function

Public Function check_dynamic_array_long(an_array() As Long) As Boolean

    Dim rtn_value As Boolean
    Dim inx As Integer
    
    rtn_value = True
    
On Error Resume Next
    
    inx = LBound(an_array)
    
    If Err.number = 9 Then
    
        rtn_value = False
        
    Else
    
        rtn_value = True
        
    End If
    
    check_dynamic_array_long = rtn_value


End Function

Public Function check_dynamic_array_int(an_array() As Integer) As Boolean

    Dim rtn_value As Boolean
    Dim inx As Integer
    
    rtn_value = True
    
On Error Resume Next
    
    inx = LBound(an_array)
    
    If Err.number = 9 Then
    
        rtn_value = False
        
    Else
    
        rtn_value = True
        
    End If
    
    check_dynamic_array_int = rtn_value


End Function


Public Sub show_error()

    MsgBox ErrObject.description & Chr(13) & Chr(13) & ErrObject.source
    ErrObject.clear


End Sub

Public Function initialized_dynamic_array(an_array() As Double) As Boolean

    Dim rtn_value As Boolean
    Dim temp As Variant
    
On Error Resume Next

    temp = an_array(LBound(an_array))
    
    If Err.number = 9 Then
    
        rtn_value = False
        
    Else
    
        rtn_value = True
        
    End If
    
    
    initialized_dynamic_array = rtn_value

End Function



Public Sub test___()

Dim l_floor(3) As Double
Dim l_cap(3) As Double
Dim fixing(3) As Double

l_floor(0) = -0.7
l_floor(1) = -0.7
l_floor(2) = -0.7

l_cap(0) = 0.2
l_cap(1) = 0.2
l_cap(2) = 0.2

fixing(0) = 241.5
fixing(1) = 250
fixing(2) = 255.5

'cliquet_performance l_floor, l_cap, fixing


End Sub

Public Function date_to_array_long(in_array() As Date, Optional base_index As Integer = 0) As Long()
    
    Dim rtn_array() As Long
    Dim inx As Long
    
    ReDim rtn_array(base_index To UBound(in_array) + base_index - 1) As Long
       
    For inx = LBound(rtn_array) To UBound(rtn_array)
    
        rtn_array(inx) = CLng(in_array(inx + 1 - base_index))
    
    Next inx
    
    date_to_array_long = rtn_array

End Function

Public Function double_to_array(in_array() As Double, Optional base_index As Integer = 0) As Double()
    
    Dim rtn_array() As Double
    Dim inx As Long
    
    ReDim rtn_array(base_index To UBound(in_array) + base_index - 1) As Double
       
    For inx = LBound(rtn_array) To UBound(rtn_array)
    
        rtn_array(inx) = in_array(inx + 1 - base_index)
    
    Next inx
    
    double_to_array = rtn_array

End Function


Public Function range_to_array_long(in_range As Range, Optional base_index As Integer = 0) As Long()
    
    Dim Direction As Boolean ' true = row, false= column
    Dim rtn_array() As Long
    Dim inx As Long
    
    If in_range.Columns.count > in_range.Rows.count Then
    
        Direction = False
        
    Else
        Direction = True
        
    End If
    
    If Direction Then
    
        ReDim rtn_array(base_index To in_range.Rows.count + base_index - 1) As Long
        
    Else
    
        ReDim rtn_array(base_index To in_range.Columns.count + base_index - 1) As Long
        
    End If
        
    For inx = LBound(rtn_array) To UBound(rtn_array)
    
        If Direction Then
         
            rtn_array(inx) = in_range.Cells(inx + 1 - base_index, 1)
            
        Else
        
            rtn_array(inx) = in_range.Cells(1, inx + 1 - base_index)
            
        End If
    
    
    Next inx
    
    range_to_array_long = rtn_array

End Function

Public Function range_to_array_date(in_range As Range, Optional base_index As Integer = 0) As Date()
    
    Dim Direction As Boolean ' true = row, false= column
    Dim rtn_array() As Date
    Dim inx As Long
    
    If in_range.Columns.count > in_range.Rows.count Then
    
        Direction = False
        
    Else
        Direction = True
        
    End If
    
    If Direction Then
    
        ReDim rtn_array(base_index To in_range.Rows.count + base_index - 1) As Date
        
    Else
    
        ReDim rtn_array(base_index To in_range.Columns.count + base_index - 1) As Date
        
    End If
        
    For inx = LBound(rtn_array) To UBound(rtn_array)
    
        If Direction Then
         
            rtn_array(inx) = in_range.Cells(inx + 1 - base_index, 1)
            
        Else
        
            rtn_array(inx) = in_range.Cells(1, inx + 1 - base_index)
            
        End If
    
    
    Next inx
    
    range_to_array_date = rtn_array

End Function
Public Function range_to_array_2d(in_range As Range, Optional initial_index As Integer = 0, Optional transpose As Boolean = False) As Double()
    

    Dim rtn_array() As Double
    Dim inx As Long
    Dim jnx As Integer
    
    
    If transpose Then
    
        ReDim rtn_array(initial_index To in_range.Columns.count + initial_index - 1 _
                      , initial_index To in_range.Rows.count + initial_index - 1) As Double
        
    Else
    
        ReDim rtn_array(initial_index To in_range.Rows.count + initial_index - 1 _
                       , initial_index To in_range.Columns.count + initial_index - 1) As Double
        
    End If
        
    For inx = LBound(rtn_array, 1) To UBound(rtn_array, 1)
        For jnx = LBound(rtn_array, 2) To UBound(rtn_array, 2)
    
            If transpose Then
             
                rtn_array(inx, jnx) = in_range.Cells(jnx + 1 - initial_index, inx + 1 - initial_index)
                
            Else
            
                rtn_array(inx, jnx) = in_range.Cells(inx + 1 - initial_index, jnx + 1 - initial_index)
                
            End If
            
        Next jnx
    
    
    Next inx
    
    range_to_array_2d = rtn_array

End Function

Public Function range_to_array(in_range As Range, Optional base_index As Integer = 0) As Double()
    
    Dim Direction As Boolean ' true = row, false= column
    Dim rtn_array() As Double
    Dim inx As Long
    
    If in_range.Columns.count > in_range.Rows.count Then
    
        Direction = False
        
    Else
        Direction = True
        
    End If
    
    If Direction Then
    
        ReDim rtn_array(base_index To in_range.Rows.count + base_index - 1) As Double
        
    Else
    
        ReDim rtn_array(base_index To in_range.Columns.count + base_index - 1) As Double
        
    End If
        
    For inx = LBound(rtn_array) To UBound(rtn_array)
    
        If Direction Then
         
            rtn_array(inx) = in_range.Cells(inx + 1 - base_index, 1)
            
        Else
        
            rtn_array(inx) = in_range.Cells(1, inx - base_index + 1)
            
        End If
    
    
    Next inx
    
    range_to_array = rtn_array

End Function






'###################################################################################
' Author: YK Jeon
'###################################################################################

Public Function ToBinary(ByVal n As Double) As String

'This function return the binary expansion of N where N is natural number - ykjeon

Dim temp As String

On Error GoTo ErrorHandler

    If Int(n) <> n Or n < 0 Then
        MsgBox "INPUT MUST BE A NONNEGATIVE INTEGER"
        Exit Function
        
    End If
    
    temp = ""
    
    Do
    
        temp = CStr(n Mod 2) & temp

        n = n \ 2

    Loop While n > 0
    
    ToBinary = temp
    
Exit Function

ErrorHandler:
    
    MsgBox "FAIL IN BINARY CONVERSION"
    Exit Function

End Function

Public Function DigitConversion( _
ByVal n As Double, _
ByVal BaseNum As Long) As String

Dim temp As String

On Error GoTo ErrorHandler

    If Int(n) <> n Or n < 0 Then
    
        MsgBox "INPUT MUST BE A NONNEGATIVE INTEGER"
        Exit Function
        
    Else
    
        If BaseNum = 1 Then
        
            DigitConversion = "-1"
            
        Else
    
            temp = ""
            
            Do
            
                temp = CStr(n Mod BaseNum) & temp
        
                n = n \ BaseNum
        
            Loop While n > 0
            
            DigitConversion = temp
        
        End If
    
    End If
    
Exit Function

ErrorHandler:
    
    MsgBox "FAIL IN DIGIT CONVERSION"
    Exit Function

End Function

Public Sub ReturnPermutator( _
ByVal Dimension As Long, _
ByVal BaseNum As Long, _
ByRef PermutArray() As Long)

Dim i As Long
Dim j As Long

Dim temp As Long
Dim Strtemp As String

    temp = BaseNum ^ (Dimension) - 1

    ReDim PermutArray(0 To temp, 1 To Dimension) As Long
    
    For i = 0 To temp
    
        Strtemp = StrReverse(DigitConversion(i, BaseNum))
        
        For j = 1 To Len(Strtemp)
        
            PermutArray(i, j) = CLng(Mid(Strtemp, j, 1))
        
        Next j
        
        For j = Len(Strtemp) + 1 To Dimension
        
            PermutArray(i, j) = 0
        
        Next j

    Next i

End Sub


Public Function Summation(ByRef values() As Double) As Double

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Double
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = 0
    
    For i = IndexMin To IndexMax
    
        temp = temp + values(i)
    
    Next i
    
    Summation = temp

End Function

Public Function max(ByVal x As Double, ByVal y As Double) As Double

    max = (x + y + Abs(x - y)) * 0.5

End Function

Public Function min(ByVal x As Double, ByVal y As Double) As Double

    'Min = x + y - Max(x, y)
    min = (x + y - Abs(x - y)) * 0.5

End Function

Public Function max_date(ByVal x As Date, ByVal y As Date) As Date

    max_date = (x + y + Abs(x - y)) * 0.5

End Function
Public Function min_date(ByVal x As Date, ByVal y As Date) As Date

    min_date = (x + y - Abs(x - y)) * 0.5

End Function

Public Function ReturnMax(ByRef values() As Double) As Double

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Double
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = values(IndexMin)
    
    If IndexMax = IndexMin Then
    
        ReturnMax = values(IndexMin)
        
    Else
    
        For i = IndexMin + 1 To IndexMax
        
            temp = max(temp, values(i))
        
        Next i
        
        ReturnMax = temp
        
    End If

End Function

Public Function ReturnMin(ByRef values() As Double) As Double

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Double
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = values(IndexMin)
    
    If IndexMax = IndexMin Then
    
        ReturnMin = values(IndexMin)
        
    Else
    
        For i = IndexMin + 1 To IndexMax
        
            temp = min(temp, values(i))
        
        Next i
        
        ReturnMin = temp
        
    End If

End Function

Public Function ReturnMinDate(ByRef values() As Date) As Date

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Date
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = values(IndexMin)
    
    If IndexMax = IndexMin Then
    
        ReturnMinDate = values(IndexMin)
        
    Else
    
        For i = IndexMin + 1 To IndexMax
        
            temp = min(temp, values(i))
        
        Next i
        
        ReturnMinDate = temp
        
    End If

End Function

Public Function lReturnMax(ByRef values() As Long) As Long

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Double
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = values(IndexMin)
    
    If IndexMax = IndexMin Then
    
        lReturnMax = values(IndexMin)
        
    Else
    
        For i = IndexMin + 1 To IndexMax
        
            temp = max(temp, values(i))
        
        Next i
        
        lReturnMax = temp
        
    End If

End Function

Public Function lReturnMin(ByRef values() As Long) As Long

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Double
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = values(IndexMin)
    
    If IndexMax = IndexMin Then
    
        lReturnMin = values(IndexMin)
        
    Else
    
        For i = IndexMin + 1 To IndexMax
        
            temp = min(temp, values(i))
        
        Next i
        
        lReturnMin = temp
        
    End If

End Function

Public Function Combin(ByVal n As Long, ByVal m As Long) As Double

Dim i As Long
Dim k As Long
Dim temp As Double
    
    temp = 1

    If Int(n) <> n Or Int(m) <> m Or n < m Then
    
        MsgBox "Input Values for Combination function are not valid"
        Exit Function
        
    Else
    
        If m = 0 Or m = n Then
        
            Combin = 1
        
        ElseIf m = 1 Or n - m = 1 Then
        
            Combin = n
            
        Else
        
            k = max(n - m, m)
            
            For i = 0 To n - k - 1
                
                temp = temp * (n - i) / (n - k - i)
            
            Next i
            
            Combin = temp
                        
        End If
        
    End If
        
End Function

Public Sub ReturnTangentVector( _
ByRef p1() As Double, _
ByRef p2() As Double, _
ByRef TangentVector() As Double)

Dim i As Long
Dim j As Long
Dim Dimension As Long

    Dimension = UBound(p1(), 1) - LBound(p1(), 1) + 1
    
    ReDim TangentVector(1 To Dimension) As Double
    
    For i = LBound(p1(), 1) To UBound(p1(), 1)
    
        TangentVector(i - LBound(p1(), 1) + 1) = _
        (p2(i) - p1(i))
    
    Next i

End Sub

Public Function ReturnVectorNorm( _
ByRef p1() As Double, _
ByRef p2() As Double, _
Optional ByVal NormIndex As Long = 2) As Double

Dim i As Long
Dim v() As Double

Dim temp As Double

Dim Dimension As Long

    Dimension = UBound(p1(), 1) - LBound(p1(), 1) + 1
    
    ReDim v(1 To Dimension) As Double
    
    For i = LBound(p1(), 1) To UBound(p1(), 1)
    
        v(i - LBound(p1(), 1) + 1) = Abs(p2(i) - p1(i))
    
    Next i

    If NormIndex = 0 Then
    
        temp = ReturnMax(v())
        
        ReturnVectorNorm = temp
        
    Else
    
        For i = 1 To Dimension
        
            temp = temp + (v(i)) ^ (NormIndex)
        
        Next i
        
        ReturnVectorNorm = (temp) ^ (1 / NormIndex)
        
    End If

End Function

Public Function Sorting( _
ByRef values() As Variant, _
ByRef OrderedValues() As Variant, _
Optional ByVal DataOdering As String = "ASC") As Double

Dim i As Long
Dim j As Long

Dim IndexMax As Long
Dim IndexMin As Long
Dim SwapArray() As Variant
Dim temp As Variant

On Error GoTo ErrorHandler

    IndexMax = UBound(values)
    IndexMin = LBound(values)
    
    ReDim OrderedValues(IndexMin To IndexMax) As Variant
    ReDim SwapArray(IndexMin To IndexMax) As Variant
    
    For i = IndexMin To IndexMax
    
        SwapArray(i) = values(i)
    
    Next i
    
    For i = IndexMin + 1 To IndexMax
    
        temp = values(i)
        
        For j = i - 1 To IndexMin Step -1
    
            If SwapArray(j) > values(i) Then
            
                SwapArray(j + 1) = SwapArray(j)
                SwapArray(j) = temp
                
            End If
    
        Next j
        
    Next i
    
    If LCase(DataOdering) = LCase("ASC") Then
        
        For i = IndexMin To IndexMax
        
            OrderedValues(i) = SwapArray(i)
        
        Next i
        
    Else
    
        For i = IndexMin To IndexMax
        
            OrderedValues(i) = SwapArray(IndexMax - i + IndexMin)
        
        Next i
    
    End If
    
    Sorting = 1
    
    Exit Function
    
ErrorHandler:

    Sorting = -1
    
    Exit Function

End Function

Public Function Moment( _
ByRef values() As Double, _
ByVal MomentOrder As Long) As Double

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Long
Dim temp As Double
    
    temp = 0
    IndexMax = UBound(values)
    IndexMin = LBound(values)
    
    If IndexMax >= IndexMin Then
    
        For i = IndexMin To IndexMax
    
            temp = temp + values(i) ^ MomentOrder / (IndexMax - IndexMin + 1)
    
        Next i
        
        Moment = temp
    
    Else
    
        Moment = -999
        
    End If

End Function

Public Function CMoment( _
ByRef values() As Double, _
ByVal MomentOrder As Long) As Double

Dim i As Long
Dim IndexMax As Long
Dim IndexMin As Long
Dim temp As Double
Dim Average As Double
    
    temp = 0
    IndexMax = UBound(values)
    IndexMin = LBound(values)
    
    Average = Moment(values, 1)
    
    If MomentOrder = 1 Then
    
        CMoment = Average
    
    Else
    
        If IndexMax > IndexMin Then
    
            For i = IndexMin To IndexMax
        
                temp = temp + (values(i) - Average) ^ MomentOrder / (IndexMax - IndexMin)
        
            Next i
            
            CMoment = temp
            
        ElseIf IndexMax = IndexMin Then
        
            CMoment = temp
            
        Else
        
            CMoment = -999
            
        End If
        
    End If

End Function

Public Function ReturnAVG( _
ByRef values() As Double) As Double
    
    ReturnAVG = Moment(values(), 1)

End Function

Public Function ReturnSTDEV( _
ByRef values() As Double) As Double

Dim temp As Double

    temp = Sqr(CMoment(values(), 2))

    If temp > 0 Then

        ReturnSTDEV = Sqr(CMoment(values(), 2))
        
    Else
    
        ReturnSTDEV = -1
        
    End If

End Function

Public Function ReturnNCSTDEV( _
ByRef values() As Double) As Double

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Long
Dim temp As Double

    temp = 0
    
    IndexMax = UBound(values)
    IndexMin = LBound(values)
    
    If IndexMax - IndexMin >= 0 Then
        
        For i = IndexMin To IndexMax
    
            temp = temp + values(i) * values(i) / (IndexMax - IndexMin + 1)
    
        Next i
    
        ReturnNCSTDEV = Sqr(temp)

    Else
    
        ReturnNCSTDEV = -1
        
    End If

End Function

Public Function ReturnABSSTDEV( _
ByRef values() As Double) As Double

Dim i As Long
Dim IndexMax As Long
Dim IndexMin As Long
Dim Average As Double
Dim temp As Double

    temp = 0
    
    IndexMax = UBound(values)
    IndexMin = LBound(values)

    If IndexMax > IndexMin Then
    
        Average = ReturnAVG(values())
    
        For i = IndexMin To IndexMax
    
            temp = temp + Abs(values(i) - Average) / (IndexMax - IndexMin)
    
        Next i
        
        ReturnABSSTDEV = temp
    
    ElseIf IndexMax = IndexMin Then
    
        ReturnABSSTDEV = 0
        
    Else
    
        ReturnABSSTDEV = -999
        
    End If

End Function

Function LogMoment( _
ByRef PriceArray() As Double, _
ByVal MomentOrder As Long, _
Optional ByVal DateOrdering As String = "ASC") As Double

Dim i As Long
Dim IndexMax As Long
Dim IndexMin As Long
Dim LogReturn() As Double

    IndexMax = UBound(PriceArray)
    IndexMin = LBound(PriceArray)
        
    If IndexMax <= IndexMin + 1 Then
        
        LogMoment = -999
        
        Exit Function
        
    Else
    
        ReDim LogReturn(IndexMin To IndexMax - 1) As Double
        
        If LCase(DateOrdering) = LCase("ASC") Then
        
            For i = IndexMin To IndexMax - 1
                
                LogReturn(i) = Log(PriceArray(i + 1) / PriceArray(i))
            
            Next i

        Else
        
            For i = IndexMin To IndexMax - 1
            
                LogReturn(i) = Log(PriceArray(i) / PriceArray(i + 1))
            
            Next i
        
        End If
            
        LogMoment = Moment(LogReturn(), MomentOrder)
        
    End If

End Function

Function CLogMoment( _
ByRef PriceArray() As Double, _
ByVal MomentOrder As Long, _
Optional ByVal DateOrdering As String = "ASC") As Double

Dim i As Long
Dim IndexMax As Long
Dim IndexMin As Long
Dim LogReturn() As Double

    IndexMax = UBound(PriceArray)
    IndexMin = LBound(PriceArray)
        
    If IndexMax <= IndexMin + 1 Then
        
        CLogMoment = -999
        
        Exit Function
        
    Else
    
        ReDim LogReturn(IndexMin To IndexMax - 1) As Double
        
        If LCase(DateOrdering) = LCase("ASC") Then
        
            For i = IndexMin To IndexMax - 1
                
                LogReturn(i) = Log(PriceArray(i + 1) / PriceArray(i))
            
            Next i

        Else
        
            For i = IndexMin To IndexMax - 1
            
                LogReturn(i) = Log(PriceArray(i) / PriceArray(i + 1))
            
            Next i
        
        End If
            
        CLogMoment = CMoment(LogReturn(), MomentOrder)
        
    End If

End Function

Function GetDrift( _
ByRef PriceArray() As Double, _
Optional ByVal AnnualizeFactor As Long = 252, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetDrift = LogMoment(PriceArray(), 1, DateOrdering) * AnnualizeFactor

End Function

Function GetVol( _
ByRef PriceArray() As Double, _
Optional ByVal AnnualizeFactor As Long = 252, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetVol = Sqr(LogMoment(PriceArray(), 2, DateOrdering) * AnnualizeFactor)

End Function

Function GetCVol( _
ByRef PriceArray() As Double, _
Optional ByVal AnnualizeFactor As Long = 252, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetCVol = Sqr(CLogMoment(PriceArray(), 2, DateOrdering) * AnnualizeFactor)

End Function

Function GetSkew( _
ByRef PriceArray() As Double, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetSkew = LogMoment(PriceArray(), 3, DateOrdering) / (GetVol(PriceArray(), 1, DateOrdering) ^ 3)

End Function

Function GetCSkew( _
ByRef PriceArray() As Double, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetCSkew = CLogMoment(PriceArray(), 3, DateOrdering) / (GetCVol(PriceArray(), 1, DateOrdering) ^ 3)

End Function

Function GetKurtosis( _
ByRef PriceArray() As Double, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetKurtosis = LogMoment(PriceArray(), 4, DateOrdering) / (GetVol(PriceArray(), 1, DateOrdering) ^ 4)

End Function

Function GetCKurtosis( _
ByRef PriceArray() As Double, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetCKurtosis = CLogMoment(PriceArray(), 4, DateOrdering) / (GetCVol(PriceArray(), 1, DateOrdering) ^ 4)

End Function

Function GetCCORR( _
ByRef PriceArrayX() As Double, _
ByRef PriceArrayY() As Double, _
Optional ByVal DateOrdering As String = "ASC" _
) As Double

Dim i As Long
Dim temp As Double
Dim IndexMax As Long
Dim IndexMin As Long
Dim AverageX As Double
Dim AverageY As Double

    temp = 0
    IndexMax = UBound(PriceArrayX)
    IndexMin = LBound(PriceArrayX)
    
    If (IndexMax <= IndexMin) Or ((IndexMax - IndexMin) <> (UBound(PriceArrayY(), 1) - LBound(PriceArrayY(), 1))) Then
    
        GetCCORR = -999
            
        Exit Function
        
    Else
            
        AverageX = GetDrift(PriceArrayX(), 1, DateOrdering)
        AverageY = GetDrift(PriceArrayY(), 1, DateOrdering)
        
        If LCase(DateOrdering) = LCase("ASC") Then
        
            For i = IndexMin To IndexMax - 1
            
                temp = temp + _
                (Log(PriceArrayX(i + 1) / PriceArrayX(i)) - AverageX) * _
                (Log(PriceArrayY(i + 1) / PriceArrayY(i)) - AverageY) _
                / (IndexMax - IndexMin - 1)
            
            Next i
        
        Else
        
            For i = IndexMin To IndexMax - 1
            
                temp = temp + _
                (Log(PriceArrayX(i) / PriceArrayX(i + 1)) - AverageX) * _
                (Log(PriceArrayY(i) / PriceArrayY(i + 1)) - AverageY) _
                / (IndexMax - IndexMin - 1)
            
            Next i
        
        End If
        
        GetCCORR = temp / (GetCVol(PriceArrayX(), 1, DateOrdering) * GetCVol(PriceArrayY(), 1, DateOrdering))
        
    End If

End Function

Function GetCORR( _
ByRef PriceArrayX() As Double, _
ByRef PriceArrayY() As Double _
) As Double

Dim i As Long
Dim temp As Double
Dim IndexMax As Long
Dim IndexMin As Long

    temp = 0
    IndexMax = UBound(PriceArrayX)
    IndexMin = LBound(PriceArrayX)
    
    If (IndexMax <= IndexMin) Or ((IndexMax - IndexMin) <> (UBound(PriceArrayY(), 1) - LBound(PriceArrayY(), 1))) Then
    
        GetCORR = -999
            
        Exit Function
        
    Else
        
        For i = IndexMin To IndexMax - 1
        
            temp = temp + _
            (Log(PriceArrayX(i + 1) / PriceArrayX(i)) * Log(PriceArrayY(i + 1) / PriceArrayY(i))) _
            / (IndexMax - IndexMin)
        
        Next i
    
        GetCORR = temp / (GetVol(PriceArrayX(), 1, "ASC") * GetVol(PriceArrayY(), 1, "ASC"))
        
    End If

End Function


Public Function indicator(condition As Boolean) As Double

    If condition Then
    
        indicator = 1
        
    Else
        
        indicator = 0
        
    End If

End Function
Public Sub push_back_long(an_array() As Long, an_obj As Long, Optional base As Integer = 1)


      Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = base
        initial_ubound = base - 1
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As Long
    
    an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_long", Err.description

End Sub
Public Sub push_back_integer(an_array() As Integer, an_obj As Integer)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As Integer
    
    an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    Err.Raise Err.number, "push_back_integer", Err.description

End Sub
Public Sub push_back_string(an_array() As String, an_obj As String)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As String
    
    an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_string"

End Sub
Public Sub push_back_boolean(an_array() As Boolean, an_obj As Boolean)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As Boolean
    
    an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_boolean"

End Sub

Public Sub push_back_double(an_array() As Double, an_obj As Double, Optional base As Integer = 1)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = base
        initial_ubound = base - 1
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As Double
    
    an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_double", Err.description

End Sub

Public Sub push_back_date(an_array() As Date, an_obj As Date)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As Date
    
    an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_double"

End Sub
Public Function get_array_size_long(an_array() As Long) As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next
    
    rtn_value = UBound(an_array) - LBound(an_array) + 1
    If Err.number = 9 Or rtn_value < 0 Then
        rtn_value = 0
    End If
    
    
    get_array_size_long = rtn_value
    

End Function
Public Function get_array_size_integer(an_array() As Integer) As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next
    
    rtn_value = UBound(an_array) - LBound(an_array) + 1
    If Err.number = 9 Or rtn_value < 0 Then
        rtn_value = 0
    End If
    
    
    get_array_size_integer = rtn_value
    

End Function

Public Function get_array_size_date(an_array() As Date) As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next
    
    rtn_value = UBound(an_array) - LBound(an_array) + 1
    If Err.number = 9 Or rtn_value < 0 Then
        rtn_value = 0
    End If
    
    
    get_array_size_date = rtn_value
    

End Function

Public Function get_array_size_double(an_array() As Double) As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next
    
    rtn_value = UBound(an_array) - LBound(an_array) + 1
    If Err.number = 9 Or rtn_value < 0 Then
        rtn_value = 0
    End If
    
    
    get_array_size_double = rtn_value
    

End Function

Public Function get_array_size_string(an_array() As String) As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next
    
    rtn_value = UBound(an_array) - LBound(an_array) + 1
    If Err.number = 9 Or rtn_value < 0 Then
        rtn_value = 0
    End If
    
    
    get_array_size_string = rtn_value
    

End Function


'---------------------------------------
' Modified on
' 2013-10-23
'---------------------------------------


Public Function array_base_zero(in_array() As Double) As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    
    ReDim rtn_array(0 To get_array_size_double(in_array) - 1) As Double
    
    For inx = 0 To get_array_size_double(in_array) - 1
        rtn_array(inx) = in_array(inx + LBound(in_array))
    Next inx
    
    array_base_zero = rtn_array

End Function

'------------------------------------
' Subject to improve
'------------------------------------
Public Function get_business_days(from_date As Date, to_date As Date) As Integer

    Dim inx As Integer
    
    Dim index_from As Integer
    Dim index_to As Integer
    
    index_from = find_location_long(holiday_list__, CLng(from_date))
    index_to = find_location_long(holiday_list__, CLng(to_date))
    
    get_business_days = to_date - from_date - (index_to - index_from) ', holidays__) ', date_initialized__)
        
End Function

'Public Function get_business_days_list(from_date As Date, to_date As Date) As Date()
'
'    Dim inx As Integer
'
'    Dim rtn_array() As Date
'
'    Dim index_from As Integer
'    Dim index_to As Integer
'
'    index_from = find_location_date(holiday_list__, from_date)
'    index_to = find_location_date(holiday_list__, to_date)
'
'    get_business_days = to_date - from_date - (index_to - index_from) ', holidays__) ', date_initialized__)
'
'End Function


' From DB
Public Function business_days_between(from_date As Date, to_date As Date) As Integer

    Dim rtn_value As Integer
    
On Error GoTo ErrorHandler

    DBConnector

    rtn_value = retrieve_business_days_between(from_date, to_date)
    
    DBDisConnector
    
    business_days_between = rtn_value

    Exit Function
    
ErrorHandler:

    DBDisConnector

End Function


Public Function polynomial_value_sht(coeff As Range, x As Double) As Double
    
    polynomial_value_sht = polynomial_value(range_to_array(coeff, 1), x)
    

End Function

Public Function polynomial_value(coeff() As Double, x As Double) As Double

    Dim inx As Integer
    Dim rtn_value As Double
    
    rtn_value = 0
    
    For inx = 1 To get_array_size_double(coeff)
    
        rtn_value = rtn_value + coeff(inx) * x ^ (inx - 1)
    
    Next inx
    
    polynomial_value = rtn_value

End Function


Public Function date_to_long_array(date_array() As Date) As Long()

    Dim inx As Integer
    Dim rtn_array() As Long
    
    
    ReDim rtn_array(LBound(date_array) To UBound(date_array)) As Long
    
    For inx = LBound(date_array) To UBound(date_array)
        
        rtn_array(inx) = CLng(date_array(inx))
    
    Next inx
    
    date_to_long_array = rtn_array
    

End Function



Public Sub push_back_clsSABRSurface(an_array() As clsSABRSurface, an_obj As clsSABRSurface)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If
    

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsSABRSurface
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_clsSABRSurface"

End Sub

Public Sub push_back_clsDoubleArray(an_array() As clsDoubleArray, an_obj As clsDoubleArray)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsDoubleArray
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_clsDblArray"

End Sub
Public Sub push_back_clsPlExplainComponent(an_array() As clsPlExplainComponent, an_obj As clsPlExplainComponent)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsPlExplainComponent
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_clsPlExplainComponent"

End Sub

'=====================================
' Sub: theta_adjustment
' Desc: Adjust 1 year theta to represent 1 business day theta
'=====================================
Public Sub theta_adjustment(ByRef greeks As clsGreeks, ByVal from_date As Date, ByVal to_date As Date)
    
    Dim days_between As Integer
    
On Error Resume Next

    days_between = get_business_days(from_date, to_date)

    If Err.number = 0 Then
    
On Error GoTo ErrorHandler
    
        greeks.theta = greeks.theta * days_between / 250
    
    Else
        
        greeks.theta = greeks.theta * (to_date - from_date) / 365
    
    End If
    
     'greeks.theta = greeks.theta / 365

    Exit Sub
    
ErrorHandler:

    raise_err "theta_adjustment", Err.description


End Sub



'Public Function cliquet_performance(local_floor() As Double, local_cap() As Double, fixing_value() As Double) As Double
'Public Function cliquet_performance(local_floor As Variant, local_cap As Variant, fixing_value As Variant) As Double
Public Function cliquet_performance(local_floor_in As Range, local_cap_in As Range, fixing_value_in As Range) As Double
    
    Dim local_floor() As Double
    Dim local_cap() As Double
    Dim fixing_value() As Double

    Dim rtn_value As Double
    Dim array_size As Long
    Dim floor_size As Long
    Dim cap_size As Long
    Dim l_floor As Double
    Dim l_cap As Double
    Dim inx As Long
    Dim prev_fixing As Double
    
    
On Error GoTo ErrorHandler

    local_floor = range_to_array(local_floor_in)
    local_cap = range_to_array(local_cap_in)
    fixing_value = range_to_array(fixing_value_in)
    
    array_size = UBound(fixing_value) - LBound(fixing_value) + 1
    floor_size = UBound(local_floor) - LBound(local_floor) + 1
    cap_size = UBound(local_cap) - LBound(local_cap) + 1
    
    If array_size <= 2 Or floor_size < 1 Or cap_size < 1 Then
    
        rtn_value = 0
        
    Else
    
        prev_fixing = fixing_value(LBound(fixing_value))
        l_floor = local_floor(LBound(local_floor))
        l_cap = local_floor(LBound(local_cap))
    
        For inx = LBound(fixing_value) + 1 To UBound(fixing_value)
        
            If fixing_value(inx) > 0 Then
            
                If UBound(local_floor) >= inx Then
                
                    l_floor = local_floor(inx)
                    
                End If
                
                If UBound(local_cap) >= inx Then
                
                    l_cap = local_cap(inx)
                    
                End If
                
                rtn_value = rtn_value + max(min(fixing_value(inx) / prev_fixing - 1, l_cap), l_floor)
                prev_fixing = fixing_value(inx)
            
            Else
            
                Exit For
                
            End If
        
        Next inx
    
    End If
    
    cliquet_performance = rtn_value

    Exit Function
    
ErrorHandler:

    raise_err "calculate_previous_performance"

End Function


Public Sub push_back_market(an_array() As clsMarket, an_obj As clsMarket)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsMarket
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_market"

End Sub

'Public Sub push_back_rate(an_array() As clsRateCurve, an_obj As clsRateCurve)
'
'
'    Dim initial_lbound As Integer
'    Dim initial_ubound As Integer
'
'    'Check if the array is initialized
'On Error Resume Next
'    Dim temp_inx As Integer
'    temp_inx = UBound(an_array)
'    If (Err.number = 9) Or temp_inx < 0 Then
'        initial_lbound = 1
'        initial_ubound = 0
'    Else
'        initial_lbound = LBound(an_array)
'        initial_ubound = UBound(an_array)
'    End If
'
'On Error GoTo ErrorHandler
'
'    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsRateCurve
'
'    Set an_array(UBound(an_array)) = an_obj
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "push_back_rate"
'
'End Sub

Public Function get_array_size_clsgreeks(an_array() As clsGreeks) As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next
    
    rtn_value = UBound(an_array) - LBound(an_array) + 1
    If Err.number = 9 Or rtn_value < 0 Then
        rtn_value = 0
    End If
    
    
    get_array_size_clsgreeks = rtn_value
    

End Function


Public Function get_array_size_clsJob(an_array() As clsJob) As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next
    
    rtn_value = UBound(an_array) - LBound(an_array) + 1
    If Err.number = 9 Or rtn_value < 0 Then
        rtn_value = 0
    End If
    
    
    get_array_size_clsJob = rtn_value
    

End Function
Public Function get_array_size_clsAcDealTicket(an_array() As clsACDealTicket) As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next
    
    rtn_value = UBound(an_array) - LBound(an_array) + 1
    If Err.number = 9 Or rtn_value < 0 Then
        rtn_value = 0
    End If
    
    
    get_array_size_clsAcDealTicket = rtn_value
    

End Function

Public Function get_array_size_clsCliquetDealTicket(an_array() As clsCliquetDealTicket) As Long
    
    Dim rtn_value As Long
    
On Error Resume Next
    
    rtn_value = UBound(an_array) - LBound(an_array) + 1
    If Err.number = 9 Or rtn_value < 0 Then
        rtn_value = 0
    End If
    
    
    get_array_size_clsCliquetDealTicket = rtn_value
    

End Function

Public Function get_array_size_clsVanillaOption(an_array() As clsVanillaOption) As Long
    
    Dim rtn_value As Long
    
On Error Resume Next
    
    rtn_value = UBound(an_array) - LBound(an_array) + 1
    If Err.number = 9 Or rtn_value < 0 Then
        rtn_value = 0
    End If
    
    
    get_array_size_clsVanillaOption = rtn_value
    

End Function

Public Function get_array_size_clsQuote(the_array() As clsQuote) As Integer
    
    Dim rtn_value As Integer
    
On Error Resume Next

    rtn_value = UBound(the_array)
    
    If Err.number = 9 Then
        rtn_value = 0
    End If
    
    
     get_array_size_clsQuote = rtn_value

End Function

'--------------------------
' mdl_common_ext_2
'--------------------------
Public Sub push_back_clsSabrParameter(an_array() As clsSabrParameter, an_obj As clsSabrParameter)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsSabrParameter
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_clsSabrParameter"

End Sub
Public Sub push_back_clsAutocallSchedule(an_array() As clsAutocallSchedule, an_obj As clsAutocallSchedule)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsAutocallSchedule
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_clsAutocallSchedule"

End Sub

Public Sub push_back_clsjob(an_array() As clsJob, an_obj As clsJob)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsJob
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_clsjob"

End Sub
Public Sub push_back_clsquote(an_array() As clsQuote, an_obj As clsQuote)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsQuote
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_clsquote"

End Sub
Public Sub push_back_greek(an_array() As clsGreeks, an_obj As clsGreeks)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = UBound(an_array)
    If (Err.number = 9) Or temp_inx < 0 Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsGreeks
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_greek"

End Sub

Public Sub push_back_vanilla(an_array() As clsVanillaOption, an_obj As clsVanillaOption)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = LBound(an_array)
    If (Err.number = 9) Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsVanillaOption
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back"

End Sub



Public Sub push_back_clsSwapSchedule(an_array() As clsSwapSchedule, an_obj As clsSwapSchedule)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = LBound(an_array)
    If (Err.number = 9) Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsSwapSchedule
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_clsSwapSchedule"

End Sub
Public Sub push_back_clsAcDealTicket(an_array() As clsACDealTicket, an_obj As clsACDealTicket)


    Dim initial_lbound As Integer
    Dim initial_ubound As Integer
    
    'Check if the array is initialized
On Error Resume Next
    Dim temp_inx As Integer
    temp_inx = LBound(an_array)
    If (Err.number = 9) Then
        initial_lbound = 1
        initial_ubound = 0
    Else
        initial_lbound = LBound(an_array)
        initial_ubound = UBound(an_array)
    End If

On Error GoTo ErrorHandler

    ReDim Preserve an_array(initial_lbound To initial_ubound + 1) As clsACDealTicket
    
    Set an_array(UBound(an_array)) = an_obj
    
    Exit Sub
    
ErrorHandler:

    raise_err "push_back_clsAcDealTicket"

End Sub

Public Function find_com_addIn() As Object

    Dim cai As COMAddIn
    Dim obj As Object
    
    For Each cai In Application.COMAddIns
        
        If InStr(cai.description, "SP Legacy System (COM Add-in Helper)") Then
            Set obj = cai.Object
            Exit For

        End If
    Next
    
    Set find_com_addIn = obj

End Function

Public Sub com_test()
    
    Dim com_obj As Object
    Dim rtn_value As Boolean
    
    Dim the_value As Double
    Dim asset_code As String
    Dim c_price(2) As Double
    Dim i_price(2) As Double
    Dim yyyymmdd As String
    
    Set com_obj = find_com_addIn
    
    asset_code = "IO122713438D"
    yyyymmdd = "20130903"
    c_price(0) = 100
    c_price(1) = 100
    i_price(0) = 100
    i_price(1) = 100
    
    
    rtn_value = com_obj.getValue(the_value, asset_code, c_price, i_price, yyyymmdd)

End Sub
Sub TestDnaComAddIn()
    Dim cai As COMAddIn
    Dim obj As Object
    For Each cai In Application.COMAddIns
        ' Could check cai.Connect to see if it is loaded.
        Debug.Print cai.description, cai.GUID
        If InStr(cai.description, "MyTitle (COM Add-in Helper)") Then
            Set obj = cai.Object
            If obj Is Nothing Then
              Debug.Print "ObjNothing"
            Else
              Debug.Print obj.SayHello(), obj.ActiveCell3
            End If
        End If
    Next
End Sub