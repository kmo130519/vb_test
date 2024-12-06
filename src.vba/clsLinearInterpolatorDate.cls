Option Explicit





Private abscissa_() As Date
Private data_() As Double

Public Function copy_obj() As clsLinearInterpolatorDate



    Dim obj As clsLinearInterpolatorDate
    Dim new_dates() As Long
    Dim new_dcfs() As Double
    Dim inx As Long
    
On Error GoTo ErrorHandler

    Set obj = New clsLinearInterpolatorDate
    
    obj.initialize abscissa_, data_
    
    Set copy_obj = obj
    
    
    Exit Function
    
ErrorHandler:

    raise_err "clsLinearInterpolator.copy_obj", Err.description
    
End Function

Public Sub initialize(abscissa() As Date, data() As Double)

    Dim inx As Integer
    
    Dim no_of_abscissa As Integer
    Dim no_of_data As Integer
        
On Error GoTo ErrorHandler

    no_of_abscissa = get_array_size_date(abscissa)
    no_of_data = get_array_size_double(data)
    
    If no_of_abscissa = no_of_data Then
    
        ReDim abscissa_(LBound(abscissa) To UBound(abscissa)) As Date
        ReDim data_(LBound(data) To UBound(data)) As Double
        
        For inx = LBound(abscissa) To UBound(abscissa)
            
            abscissa_(inx) = abscissa(inx)
        
        Next inx
        
       
        For inx = (LBound(data)) To UBound(data)
            
            data_(inx) = data(inx)
        
        Next inx
    
    Else
    
        raise_err "clsLinearINterpolator.initialize", "Array size mismatch.."
        
    End If
    
    Exit Sub
    
ErrorHandler:


    raise_err "clsLinearInterpolator.initialize", Err.description


End Sub

Public Function interpolated_value(x_value As Date) As Double

    Dim inx As Integer
    Dim rtn_value As Double
    
On Error GoTo ErrorHandler

    inx = find_location_date(abscissa_, x_value)
    
    If inx < LBound(abscissa_) Then
        rtn_value = data_(LBound(abscissa_))
    ElseIf inx >= UBound(abscissa_) Then
        rtn_value = data_(UBound(abscissa_))
    Else
        rtn_value = ((x_value - abscissa_(inx)) * data_(inx + 1) + (abscissa_(inx + 1) - x_value) * data_(inx)) / (abscissa_(inx + 1) - abscissa_(inx))
    End If

    interpolated_value = rtn_value
    
    Exit Function
    
ErrorHandler:
    
    raise_err "clsLinearInterpolator.interpolated_value", Err.description

End Function