Option Explicit

'Each function returns price, delta, gamma, crossgamma, theta, dDeltadt, dGammadt,dgammadspot from node point-ykjeon

Public Function ReturnPrice( _
ByRef currentprice() As Double, _
ByRef initialprice() As Double, _
ByRef nodeprice() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long) As Double

Dim i As Long
Dim Point() As Double
Dim nodecls As New NodeControl

    ReDim Point(1 To UBound(currentprice(), 1) - LBound(currentprice(), 1) + 1) As Double
    
    For i = LBound(currentprice(), 1) To UBound(currentprice(), 1)
    
        Point(i - LBound(currentprice(), 1) + 1) = currentprice(i) / initialprice(i)
    
    Next i

    ReturnPrice = nodecls.ReturnValue(Point(), nodeprice(), px(), dx(), nmin(), nmax())
    
    Set nodecls = Nothing
    Erase Point

End Function

Public Function ReturnDelta( _
ByRef currentprice() As Double, _
ByRef initialprice() As Double, _
ByRef nodeprice() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis As Long, _
Optional ByVal Direction As Long = 0, _
Optional ByVal NodeStep As Long = 1) As Double

Dim i As Long
Dim Point() As Double
Dim nodecls As New NodeControl

    ReDim Point(1 To UBound(currentprice(), 1) - LBound(currentprice(), 1) + 1) As Double
    
    For i = LBound(currentprice(), 1) To UBound(currentprice(), 1)
    
        Point(i - LBound(currentprice(), 1) + 1) = currentprice(i) / initialprice(i)
    
    Next i

    ReturnDelta = nodecls.ReturnTangent(Point(), nodeprice(), px(), dx(), nmin(), nmax(), Axis, _
     Direction, NodeStep) / initialprice(Axis)
    
    Set nodecls = Nothing
    Erase Point
    
End Function

Public Function ReturnGamma( _
ByRef currentprice() As Double, _
ByRef initialprice() As Double, _
ByRef nodeprice() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis1 As Long, _
ByVal Axis2 As Long, _
Optional ByVal NodeStep As Long = 1) As Double

Dim i As Long
Dim Point() As Double
Dim nodecls As New NodeControl

    If Axis1 = Axis2 Then

        ReDim Point(1 To UBound(currentprice(), 1) - LBound(currentprice(), 1) + 1) As Double
        
        For i = LBound(currentprice(), 1) To UBound(currentprice(), 1)
        
            Point(i - LBound(currentprice(), 1) + 1) = currentprice(i) / initialprice(i)
        
        Next i
    
        ReturnGamma = nodecls.ReturnConvexity(Point(), nodeprice(), px(), dx(), nmin(), nmax(), Axis1, NodeStep) _
        / (initialprice(Axis1) * initialprice(Axis1))
        
    Else
    
        ReturnGamma = ReturnCrossGamma(currentprice(), initialprice(), nodeprice(), px(), dx(), nmin(), nmax(), Axis1, Axis2)
    
    End If
    
    Set nodecls = Nothing
    Erase Point

End Function

Public Function ReturnCrossGamma( _
ByRef currentprice() As Double, _
ByRef initialprice() As Double, _
ByRef nodeprice() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis1 As Long, _
ByVal Axis2 As Long) As Double

Dim i As Long
Dim Point() As Double
Dim nodecls As New NodeControl

    ReDim Point(1 To UBound(currentprice(), 1) - LBound(currentprice(), 1) + 1) As Double
    
    For i = LBound(currentprice(), 1) To UBound(currentprice(), 1)
    
        Point(i - LBound(currentprice(), 1) + 1) = currentprice(i) / initialprice(i)
    
    Next i

    ReturnCrossGamma = nodecls.ReturnMixedDeriv(Point(), nodeprice(), px(), dx(), nmin(), nmax(), Axis1, Axis2) _
    / (initialprice(Axis1) * initialprice(Axis2))
    
    Set nodecls = Nothing
    Erase Point

End Function

Public Function ReturnTheta( _
ByRef currentprice() As Double, _
ByRef initialprice() As Double, _
ByRef nodeprice() As Double, _
ByRef NextTimePrice() As Double, _
ByVal timestep As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long) As Double

Dim i As Long
Dim Point() As Double
Dim nodecls As New NodeControl

    ReDim Point(1 To UBound(currentprice(), 1) - LBound(currentprice(), 1) + 1) As Double
    
    For i = LBound(currentprice(), 1) To UBound(currentprice(), 1)
    
        Point(i - LBound(currentprice(), 1) + 1) = currentprice(i) / initialprice(i)
    
    Next i

    ReturnTheta = (nodecls.ReturnValue(Point(), NextTimePrice(), px(), dx(), nmin(), nmax()) - _
    nodecls.ReturnValue(Point(), nodeprice(), px(), dx(), nmin(), nmax())) / timestep
        
    Set nodecls = Nothing
    Erase Point

End Function

Public Function ReturnDdeltaDt( _
ByRef currentprice() As Double, _
ByRef initialprice() As Double, _
ByRef nodeprice() As Double, _
ByRef NextTimePrice() As Double, _
ByVal timestep As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis As Long, _
Optional ByVal Direction As Long = 0, _
Optional ByVal NodeStep As Long = 1) As Double

Dim i As Long
Dim Point() As Double
Dim nodecls As New NodeControl

    ReDim Point(1 To UBound(currentprice(), 1) - LBound(currentprice(), 1) + 1) As Double
    
    For i = LBound(currentprice(), 1) To UBound(currentprice(), 1)
    
        Point(i - LBound(currentprice(), 1) + 1) = currentprice(i) / initialprice(i)
    
    Next i

    ReturnDdeltaDt = (nodecls.ReturnTangent(Point(), NextTimePrice(), px(), dx(), nmin(), nmax(), Axis, _
     Direction, NodeStep) - nodecls.ReturnTangent(Point(), nodeprice(), px(), dx(), nmin(), nmax(), Axis, _
     Direction, NodeStep)) / (initialprice(Axis) * timestep)
    
    Set nodecls = Nothing
    Erase Point

End Function

Public Function ReturnDgammaDt( _
ByRef currentprice() As Double, _
ByRef initialprice() As Double, _
ByRef nodeprice() As Double, _
ByRef NextTimePrice() As Double, _
ByVal timestep As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis As Long, _
Optional ByVal NodeStep As Long = 1) As Double

Dim i As Long
Dim Point() As Double
Dim nodecls As New NodeControl

    ReDim Point(1 To UBound(currentprice(), 1) - LBound(currentprice(), 1) + 1) As Double
    
    For i = LBound(currentprice(), 1) To UBound(currentprice(), 1)
    
        Point(i - LBound(currentprice(), 1) + 1) = currentprice(i) / initialprice(i)
    
    Next i

    ReturnDgammaDt = (nodecls.ReturnConvexity(Point(), NextTimePrice(), px(), dx(), nmin(), nmax(), Axis, _
    NodeStep) - nodecls.ReturnConvexity(Point(), nodeprice(), px(), dx(), nmin(), nmax(), Axis, _
    NodeStep)) / (initialprice(Axis) * initialprice(Axis) * timestep)
    
    Set nodecls = Nothing
    Erase Point

End Function

Public Function ReturnDgammaDspot( _
ByRef currentprice() As Double, _
ByRef initialprice() As Double, _
ByRef nodeprice() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis As Long) As Double

Dim i As Long
Dim Point() As Double
Dim nodecls As New NodeControl

    ReDim Point(1 To UBound(currentprice(), 1) - LBound(currentprice(), 1) + 1) As Double
    
    For i = LBound(currentprice(), 1) To UBound(currentprice(), 1)
    
        Point(i - LBound(currentprice(), 1) + 1) = currentprice(i) / initialprice(i)
    
    Next i
    
    ReturnDgammaDspot = nodecls.ReturnThirdMoment(Point(), nodeprice(), px(), dx(), nmin(), nmax(), Axis)
    
    Set nodecls = Nothing
    Erase Point

End Function