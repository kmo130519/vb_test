Option Explicit

'------------------------------------------------------
'Code:
' S : underlying price direction.
' VOL: volatility direction
'------------------------------------------------------

Private spot_grid_() As Double
Private vol_grid_() As Double
Private greek_value_() As clsGreeks ' 2 Dimension array

Public Sub set_greek_value(inx As Integer, jnx As Integer, greek As clsGreeks)
    
    Set greek_value_(inx, jnx) = greek


End Sub

Public Function greek_value(inx As Integer, jnx As Integer) As clsGreeks

    

    Set greek_value = greek_value_(inx, jnx)


End Function

Public Function spot_grid(inx As Integer) As Double

On Error GoTo ErrorHandler

    If inx >= 1 And inx <= get_number_of_grid("S") Then

        spot_grid = spot_grid_(inx)
    
    Else
    
        raise_err "spot_grid", "out of bound array"
        
    End If
    
    Exit Function
    
ErrorHandler:

    raise_err "clsSpotVolGrid.spot_grid"
    

End Function

Public Function vol_grid(inx As Integer) As Double

On Error GoTo ErrorHandler


    If inx >= 1 And inx <= get_number_of_grid("VOL") Then
    
        vol_grid = vol_grid_(inx)
    
    Else
    
        raise_err "spot_grid", "out of bound array"
        
    End If
    
    Exit Function
    
ErrorHandler:

    raise_err "clsSpotVolGrid.vol_grid"
        

End Function

'---------------------------------
' Redim grid arrays
' if the number of grids the user set is even number,
' the number of grid will be the nearest odd number
'---------------------------------
Public Sub initialize(no_of_grid_s As Integer, no_of_grid_vol As Integer, interval_s As Double, interval_vol As Double)

    Dim no_of_grid_s_mod As Integer
    Dim no_of_grid_vol_mod As Integer
    
    Dim inx As Integer
    
On Error GoTo ErrorHandler

    no_of_grid_s_mod = Int((no_of_grid_s - 1) / 2) * 2 + 1 ' Make it to be odd number.
    no_of_grid_vol_mod = Int((no_of_grid_vol - 1) / 2) * 2 + 1 ' Make it to be odd number.

    ReDim spot_grid_(1 To no_of_grid_s_mod) As Double
    ReDim vol_grid_(1 To no_of_grid_vol_mod) As Double
    
    ReDim greek_value_(1 To no_of_grid_s_mod, 1 To no_of_grid_vol_mod) As clsGreeks
    
    For inx = 1 To no_of_grid_s_mod
        
        spot_grid_(inx) = interval_s * (inx - get_center_grid_index("S"))
    
    Next inx
    
    For inx = 1 To no_of_grid_vol_mod
        
        vol_grid_(inx) = interval_vol * (inx - get_center_grid_index("VOL"))
    
    Next inx
    
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsSpotVolGrid.initialize"

End Sub

'------------------------
' Returns the index of the center grid.
'------------------------
Public Function get_center_grid_index(Optional Direction As String = "S") As Integer

    Dim no_of_grid As Integer
    Dim center_grid_index As Integer
    
On Error GoTo ErrorHandler

    no_of_grid = get_number_of_grid(Direction)
    
    center_grid_index = Int(no_of_grid / 2) + 1
    
    
    get_center_grid_index = center_grid_index
    
    Exit Function
    
ErrorHandler:

    raise_err "clsSpotVolGrid.get_center_grid_index"
    

End Function

'------------------------
' Returns the number of grid according to the direction
'------------------------
Public Function get_number_of_grid(Optional Direction As String = "S") As Integer
    
    Dim no_of_grid As Integer
  
On Error GoTo ErrorHandler

    If Direction = "S" Then

On Error Resume Next
        no_of_grid = UBound(spot_grid_)
        
        If Err.number = 9 Then
            
            no_of_grid = 0
        
        End If
        
On Error GoTo ErrorHandler
        
    ElseIf Direction = "VOL" Then
    
        no_of_grid = UBound(vol_grid_)
        
On Error Resume Next
        
        If Err.number = 9 Then
            
            no_of_grid = 0
        
        End If
        
On Error GoTo ErrorHandler

    Else
    
        raise_err "get_number_of_grid", "Invalid direction of grid"
    
    End If
    
    
    get_number_of_grid = no_of_grid
        
    Exit Function
    
ErrorHandler:

    raise_err "clsSpotVolGrid.get_number_of_grid"



End Function