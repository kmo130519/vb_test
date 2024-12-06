Option Explicit


Private xAxis_vol_bump_() As Double
Private yAxis_vol_bump_() As Double
Private zAxis_vol_bump_() As Double
Private snapshot_value_vol_bump_() As Double
'Private snapshot_value_vol_bump_() As Double
'Private snapshot_value_skew_bump_() As Double



Public Function get_xAxis_vol_bump() As Double()

    get_xAxis_vol_bump = xAxis_vol_bump_

End Function

Public Function get_yAxis_vol_bump() As Double()

    get_yAxis_vol_bump = yAxis_vol_bump_

End Function
Public Function get_zAxis_vol_bump() As Double()

    get_zAxis_vol_bump = zAxis_vol_bump_

End Function
Public Function get_snapshot_value_vol_bump() As Double()

    get_snapshot_value_vol_bump = snapshot_value_vol_bump_

End Function
'Public Function ge



Public Sub set_xAxis_vol_bump(xAxis() As Double, Optional log_value As Boolean = True, Optional as_percentage_of As Double = 1#)
    
    Dim inx As Integer
    
On Error Resume Next

    ReDim xAxis_vol_bump_(LBound(xAxis) To UBound(xAxis)) As Double
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        
        raise_err "set_xAxis", "xAxis not initialized.."
    
    End If
    
    If log_value Then
    
        For inx = LBound(xAxis) To UBound(xAxis)
            
            xAxis_vol_bump_(inx) = Exp(xAxis(inx)) / as_percentage_of
        
        Next inx
    
    Else
    
        For inx = LBound(xAxis) To UBound(xAxis)
        
            xAxis_vol_bump_(inx) = xAxis(inx) / as_percentage_of
            
        Next inx
        
    End If
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_xAxis"


End Sub


Public Sub set_yAxis_vol_bump(yAxis() As Double, Optional log_value As Boolean = True, Optional as_percentage_of As Double = 1#)
    
    Dim inx As Integer
    
On Error Resume Next

    ReDim yAxis_vol_bump_(LBound(yAxis) To UBound(yAxis)) As Double
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        
        raise_err "set_yAxis", "yAxis not initialized.."
    
    End If
    
    If log_value Then
    
        For inx = LBound(yAxis) To UBound(yAxis)
            
            yAxis_vol_bump_(inx) = Exp(yAxis(inx)) / as_percentage_of
        
        Next inx
    
    Else
        
        For inx = LBound(yAxis) To UBound(yAxis)
            
            yAxis_vol_bump_(inx) = (yAxis(inx)) / as_percentage_of
        
        Next inx
        
    End If
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_yAxis"


End Sub
Public Sub set_zAxis_vol_bump(zAxis() As Double, Optional log_value As Boolean = True, Optional as_percentage_of As Double = 1#)
    
    Dim inx As Integer
    
On Error Resume Next

    ReDim zAxis_vol_bump_(LBound(zAxis) To UBound(zAxis)) As Double
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        
        raise_err "set_zAxis", "yAxis not initialized.."
    
    End If
    
    If log_value Then
    
        For inx = LBound(zAxis) To UBound(zAxis)
            
            zAxis_vol_bump_(inx) = Exp(zAxis(inx)) / as_percentage_of
        
        Next inx
    
    Else
        
    
        For inx = LBound(zAxis) To UBound(zAxis)
            
            zAxis_vol_bump_(inx) = (zAxis(inx)) / as_percentage_of
        
        Next inx
        
    End If
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_zAxis"


End Sub

Public Sub set_snapshot_value_bump(snapshot_value() As Double, Optional notional As Double = 1)
    
    Dim inx As Long
    
On Error Resume Next

    ReDim snapshot_value_vol_bump_(LBound(snapshot_value) To UBound(snapshot_value)) As Double
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        
        raise_err "set_snapshot_value", "snapshot_value_ not initialized.."
    
    End If
    
   
   For inx = LBound(snapshot_value) To UBound(snapshot_value)
            
        snapshot_value_vol_bump_(inx) = snapshot_value(inx) * notional
        
   Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_snaptshot_value", Err.description


End Sub


Public Sub set_snapshot_value_vol_bump(snapshot_value() As Double, Optional notional As Double = 1)
    
    Dim inx As Long
    
On Error Resume Next

    ReDim snapshot_value_vol_bump_(LBound(snapshot_value) To UBound(snapshot_value)) As Double
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        
        raise_err "set_snapshot_value", "snapshot_value_ not initialized.."
    
    End If
    
   
   For inx = LBound(snapshot_value) To UBound(snapshot_value)
            
        snapshot_value_vol_bump_(inx) = snapshot_value(inx) * notional
        
   Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_snaptshot_value", Err.description


End Sub