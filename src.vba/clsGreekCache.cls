Option Explicit

Public asset_code As String

Public initialized As Boolean

Private abscissa_() As Double ' abscissa of precalculated greeks
Private greeks_() As clsGreeks ' This is 2 dimension array.

Public Function lower_bound_of_abscissa() As Integer
    
    lower_bound_of_abscissa = LBound(abscissa_)


End Function

Public Function upper_bound_of_abscissa() As Integer
    
    upper_bound_of_abscissa = UBound(abscissa_)


End Function

Public Function abscissa(inx As Integer) As Double

    abscissa = abscissa_(inx)

End Function

Public Function greeks(inx As Integer) As clsGreeks

    Set greeks = greeks_(inx)

End Function

Public Sub initialize(abscissa() As Double, greeks() As clsGreeks, Optional deal_in As clsCliquetDealTicket)

On Error GoTo ErrorHandler

    initialized = False
    
'    If Not IsNull(deal_in) And Not deal_in Is Nothing Then
'        set_deal deal_in
'    End If
        
    set_abscissa abscissa
    set_greeks greeks

    initialized = True
    
    Exit Sub
    
ErrorHandler:

    initialized = False

    raise_err "clsGreekCache.initialize"

End Sub

Public Sub null_greek(Optional deal_in As clsCliquetDealTicket)
    
    Dim null_abs(1 To 2) As Double
    Dim greeks(1 To 2) As clsGreeks
    
    null_abs(1) = 0
    null_abs(2) = 10000
    
    Set greeks(1) = New clsGreeks
    Set greeks(2) = New clsGreeks
    
    initialize null_abs, greeks, deal_in


End Sub

'Public Sub set_deal(deal_in As clsCliquetDealTicket)
'
'On Error GoTo ErrorHandler
'
'    Set deal = deal_in.copy_obj
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err " clsGreekCache.set_deal"
'
'
'End Sub

Public Sub set_abscissa(abscissa() As Double)

    Dim inx As Integer
    
On Error GoTo ErrorHandler
    
    ReDim abscissa_(LBound(abscissa) To UBound(abscissa)) As Double
        
    For inx = LBound(abscissa) To UBound(abscissa)
    
        abscissa_(inx) = abscissa(inx)
    
    Next inx
    
    Exit Sub
    
        
ErrorHandler:

    If Err.number = 9 Then
    
        Exit Sub
        
    End If

    raise_err "clsGreekCache.set_abscissa"


End Sub


Public Sub set_greeks(greeks() As clsGreeks)

    Dim inx As Integer
    
On Error GoTo ErrorHandler
    
    ReDim greeks_(LBound(greeks) To UBound(greeks)) As clsGreeks
    
    For inx = LBound(greeks) To UBound(greeks)
    
        Set greeks_(inx) = greeks(inx).copy_obj
    
    Next inx
    
    Exit Sub
    
ErrorHandler:

    If Err.number = 9 Then
    
        Exit Sub
        
    End If

    raise_err "clsGreekCache.set_greeks"


End Sub

Public Function get_greeks(stock_price As Double) As clsGreeks

    Dim rtn_greeks As clsGreeks
    Dim inx As Integer
    
On Error GoTo ErrorHandler
    
    Set rtn_greeks = New clsGreeks
    
    inx = find_location(stock_price)
    
    rtn_greeks.value = linear_interpolation(abscissa_(inx), abscissa_(inx + 1), greeks_(inx).value, greeks_(inx + 1).value, stock_price)
    rtn_greeks.delta = linear_interpolation(abscissa_(inx), abscissa_(inx + 1), greeks_(inx).delta, greeks_(inx + 1).delta, stock_price)
    rtn_greeks.gamma = linear_interpolation(abscissa_(inx), abscissa_(inx + 1), greeks_(inx).gamma, greeks_(inx + 1).gamma, stock_price)
    
    rtn_greeks.vega = greeks_(LBound(greeks_)).vega
    rtn_greeks.theta = greeks_(LBound(greeks_)).theta
    rtn_greeks.skew_s = greeks_(LBound(greeks_)).skew_s
    
    Set get_greeks = rtn_greeks
    
    Exit Function
    
ErrorHandler:

    raise_err "clsGreekCache.get_greeks"
    

End Function
Private Function linear_interpolation(x1 As Double, x2 As Double, y1 As Double, y2 As Double, x) As Double

    linear_interpolation = (x - x1) / (x2 - x1) * y2 + (x2 - x) / (x2 - x1) * y1

End Function


Private Function find_location(stock_price As Double) As Integer

    Dim inx_up As Integer
    Dim inx_low As Integer
    Dim inx_mid As Integer
    
    Dim rtn_inx As Integer
    
On Error GoTo ErrorHandler
    
    inx_low = LBound(abscissa_) - 1
    inx_up = UBound(abscissa_) + 1
    
    Do While inx_up - inx_low > 1
    
        inx_mid = Int((inx_up + inx_low) / 2)
        
        If stock_price >= abscissa_(inx_mid) Then
        
            inx_low = inx_mid
        
        Else
            
            inx_up = inx_mid
            
        End If

    Loop
    
    
    If stock_price = abscissa_(LBound(abscissa_)) Then
    
        rtn_inx = LBound(abscissa_)
        
    ElseIf stock_price = abscissa_(UBound(abscissa_)) Then
        
        rtn_inx = UBound(abscissa_) - 1
        
    Else
    
        rtn_inx = inx_low
    
    End If
    
    
    find_location = rtn_inx
    
    Exit Function
    
    
ErrorHandler:

    raise_err "clsGreekCache.find_location"

End Function