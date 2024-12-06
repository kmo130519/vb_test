Option Explicit

Public no_of_greeks As Integer

Private greeks_() As clsGreeks

Public Sub initialize(no_of_greeks_in As Integer)

    ReDim greeks_(1 To no_of_greeks_in) As clsGreeks
    
    no_of_greeks = no_of_greeks_in

End Sub

Public Sub set_greek(inx As Integer, a_greek As clsGreeks)

    If inx > no_of_greeks Then
        raise_err "clsGreekSet.set_greeks", "out of bound"
    End If
    
    Set greeks_(inx) = a_greek
    
End Sub


Public Function greeks(inx As Integer) As clsGreeks
    
    Set greeks = greeks_(inx)
    
End Function