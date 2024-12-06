Option Explicit


Public grid_ As clsPillarGrid
Public grid_origin_ As clsPillarGrid

Private vol_surface_() As Double
Private vol_surface_origin_() As Double

Public Function copy_vol_surface() As Double()

    Dim rtn_array() As Double
    Dim inx As Integer
    Dim jnx As Integer

    ReDim rtn_array(1 To grid_.no_of_dates, 1 To grid_.no_of_strikes) As Double

    For inx = 1 To grid_.no_of_dates
        For jnx = 1 To grid_.no_of_strikes

            rtn_array(inx, jnx) = vol_surface_(inx, jnx)
    
        Next jnx
    Next inx

    copy_vol_surface = rtn_array

End Function

Public Function copy_obj() As clsVolSurface
    
    Dim rtn_obj As clsVolSurface
    
    Set rtn_obj = New clsVolSurface
    
    rtn_obj.set_vol_surface grid_.copy_obj(), copy_vol_surface()
    rtn_obj.set_vol_surface_origin grid_.copy_obj(), copy_vol_surface()
    
    
    Set copy_obj = rtn_obj


End Function

Public Function get_vol_surface(Optional base As Integer = 1) As Double()
    
    Dim rtn_array() As Double
    Dim inx As Integer
    Dim jnx As Integer
    
On Error GoTo ErrorHandler
    
    ReDim rtn_array(base To grid_.no_of_dates * grid_.no_of_strikes + base - 1)
    
    For inx = 1 To grid_.no_of_strikes
        For jnx = 1 To grid_.no_of_dates
        
            rtn_array(base + (inx - 1) * grid_.no_of_dates + jnx - 1) = vol_surface_(jnx, inx)
            
        Next jnx
    Next inx
    
    get_vol_surface = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsVolSurface.get_vol_surface", Err.description
    
End Function


Public Function vol_surface() As Double()

    vol_surface = vol_surface_

End Function


Public Sub set_grid(grid_in As clsPillarGrid)  ', Optional max_date As Date = 0)

    Set grid_ = grid_in
        
'    If max_date > grid_.get_last_date Then
'
'        Dim tmp_dates() As Date
'
'        tmp_dates = grid_.get_all_dates()
'
'        push_back_date tmp_dates, max_date
'
'        grid_.set_dates tmp_dates
'
'    End If

End Sub

'==================================================
' 2015/10/05
'==================================================

Private Sub add_upto_date_to_grid(ByRef grid_in As clsPillarGrid, ByVal up_to_tenor As Date)
    
    Dim inx As Integer
    Dim new_date_array() As Date
    
On Error GoTo ErrorHandler

    If grid_in.dates(1) > up_to_tenor Then
        push_back_date new_date_array, up_to_tenor
    End If
    
    For inx = 1 To grid_in.no_of_dates
    
        push_back_date new_date_array, grid_in.dates(inx)
        
        If grid_in.dates(inx) < up_to_tenor Then
            If inx < grid_in.no_of_dates Then
                If grid_in.dates(inx + 1) > up_to_tenor Then
                    push_back_date new_date_array, up_to_tenor
                End If
            End If
        End If
    
    Next inx
    
    grid_in.set_dates new_date_array
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsvolSurface.add_upto_date_to_grid", Err.description

End Sub

Private Function find_origin_index(old_grid As clsPillarGrid, new_grid_date As Date) As Integer

    Dim inx As Integer
    Dim rtn_inx As Integer
    
On Error GoTo ErrorHandler

    
   rtn_inx = 0
    
    For inx = old_grid.no_of_dates To 1 Step -1
    
        If old_grid.dates(inx) <= new_grid_date Then
            rtn_inx = inx
            Exit For
        End If
            
    Next inx

   
    find_origin_index = rtn_inx
    
    Exit Function
    
ErrorHandler:

    raise_err "clsVolSurface.find_origin_index", Err.description

End Function

Private Sub change_vol_data_as_grid_change(ByRef vol_surface_to() As Double, ByRef vol_surface_from() As Double, ByRef old_grid As clsPillarGrid, ByRef new_grid As clsPillarGrid)

    Dim inx As Integer
    Dim jnx As Integer
    Dim tmp_origin_index As Integer
    
    
    
On Error GoTo ErrorHandler

    ReDim vol_surface_to(1 To new_grid.no_of_dates, 1 To new_grid.no_of_strikes) As Double
    
    For inx = 1 To new_grid.no_of_dates
    
    ' find origin date index d_origin_inx <= d_inx, d_origin_inx + 1 > d_inx
    '
        tmp_origin_index = find_origin_index(old_grid, new_grid.dates(inx))
        
        If tmp_origin_index = 0 Then
            
            For jnx = 1 To new_grid.no_of_strikes
                vol_surface_to(inx, jnx) = vol_surface_from(1, jnx)
            Next jnx
        
        ElseIf tmp_origin_index = old_grid.no_of_dates Then
            
            For jnx = 1 To new_grid.no_of_strikes
                vol_surface_to(inx, jnx) = vol_surface_from(old_grid.no_of_dates, jnx)
            Next jnx
            
        Else
            
            For jnx = 1 To new_grid.no_of_strikes
            
                vol_surface_to(inx, jnx) = vol_surface_from(tmp_origin_index, jnx) * (old_grid.dates(tmp_origin_index + 1) - new_grid.dates(inx))
                vol_surface_to(inx, jnx) = vol_surface_to(inx, jnx) + vol_surface_from(tmp_origin_index + 1, jnx) * (new_grid.dates(inx) - old_grid.dates(tmp_origin_index))
                vol_surface_to(inx, jnx) = vol_surface_to(inx, jnx) / (old_grid.dates(tmp_origin_index + 1) - old_grid.dates(tmp_origin_index))
                
            Next jnx
            
        End If
        
    Next inx
    

    Exit Sub
    
ErrorHandler:
    
    raise_err "clsVolSurface.add_upto_date_to_surface", Err.description
    
End Sub

Private Sub backup_vol_surface_data()
    
    Dim inx As Integer
    Dim jnx As Integer
    Dim no_of_rows As Integer
    Dim no_of_cols As Integer
    
    no_of_rows = grid_.no_of_dates
    no_of_cols = grid_.no_of_strikes

On Error GoTo ErrorHandler
    

    ReDim vol_surface_origin_(1 To no_of_rows, 1 To no_of_cols) As Double
    
    
    For inx = 1 To no_of_rows
        For jnx = 1 To no_of_cols
                
            vol_surface_origin_(inx, jnx) = vol_surface_(inx, jnx)
        
        Next jnx
    Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsVolSurface.backup_vol_surface_data", Err.description

End Sub

Public Sub rewind_bump_vol_upto()

    Set grid_ = grid_origin_.copy_obj
    vol_surface_ = vol_surface_origin_
    
End Sub

Public Sub bump_vol_upto(ByVal bump_size As Double, ByVal up_to_tenor As Date)
    
    '---------------------------------------------------------
    ' back up original vol surface
    '---------------------------------------------------------
    
    backup_vol_surface_data ' to vol_surface_origin_
    
    Set grid_origin_ = grid_.copy_obj
    add_upto_date_to_grid grid_, up_to_tenor
        
    change_vol_data_as_grid_change vol_surface_, vol_surface_origin_, grid_origin_, grid_
    
    '----------------------
    ' Bump
    '----------------------
    
    Dim inx As Integer
    Dim jnx As Integer
    Dim no_of_rows As Integer
    Dim no_of_cols As Integer
    
On Error GoTo ErrorHandler

    no_of_rows = grid_.no_of_dates
    no_of_cols = grid_.no_of_strikes
    
    For inx = 1 To no_of_rows
    
        If grid_.dates(inx) <= up_to_tenor Then
            For jnx = 1 To no_of_cols
            
                vol_surface_(inx, jnx) = vol_surface_(inx, jnx) + bump_size
            
            Next jnx
        End If
    Next inx
    
    
        
    
    'grid_origin_.set_dates
    
On Error GoTo ErrorHandler


    Exit Sub
    
ErrorHandler:

    raise_err "clsVolSurface.bump_vol_upto", Err.description

End Sub




Public Sub bump_skew_upto(ByVal bump_size As Double, ByVal up_to_tenor As Date, ByVal s As Double)
    
    '---------------------------------------------------------
    ' back up original vol surface
    '---------------------------------------------------------
    
    backup_vol_surface_data ' to vol_surface_origin_
    
    Set grid_origin_ = grid_.copy_obj
    add_upto_date_to_grid grid_, up_to_tenor
        
    change_vol_data_as_grid_change vol_surface_, vol_surface_origin_, grid_origin_, grid_
    
    '----------------------
    ' Bump
    '----------------------
    
    Dim inx As Integer
    Dim jnx As Integer
    Dim no_of_rows As Integer
    Dim no_of_cols As Integer
    
On Error GoTo ErrorHandler

    no_of_rows = grid_.no_of_dates
    no_of_cols = grid_.no_of_strikes
    
    For inx = 1 To no_of_rows
    
        If grid_.dates(inx) <= up_to_tenor Then
            For jnx = 1 To no_of_cols
            
                vol_surface_(inx, jnx) = vol_surface_(inx, jnx) + 1 / Sqr(grid_.tau(inx)) * (s - grid_.strikes(jnx)) / s * bump_size 'vol_surface_(inx, jnx) + bump_size
            
            Next jnx
        End If
    Next inx
    

On Error GoTo ErrorHandler


    Exit Sub
    
ErrorHandler:

    raise_err "clsVolSurface.bump_vol_upto", Err.description

End Sub



Public Sub bump_conv_upto(ByVal bump_size As Double, ByVal up_to_tenor As Date, ByVal s As Double)
    
    '---------------------------------------------------------
    ' back up original vol surface
    '---------------------------------------------------------
    
    backup_vol_surface_data ' to vol_surface_origin_
    
    Set grid_origin_ = grid_.copy_obj
    add_upto_date_to_grid grid_, up_to_tenor
        
    change_vol_data_as_grid_change vol_surface_, vol_surface_origin_, grid_origin_, grid_
    
    '----------------------
    ' Bump
    '----------------------
    
    Dim inx As Integer
    Dim jnx As Integer
    Dim no_of_rows As Integer
    Dim no_of_cols As Integer
    
On Error GoTo ErrorHandler

    no_of_rows = grid_.no_of_dates
    no_of_cols = grid_.no_of_strikes
    
    For inx = 1 To no_of_rows
    
        If grid_.dates(inx) <= up_to_tenor Then
            For jnx = 1 To no_of_cols
            
                vol_surface_(inx, jnx) = vol_surface_(inx, jnx) + 1 / grid_.tau(inx) * Abs((s - grid_.strikes(jnx)) / s) * bump_size 'vol_surface_(inx, jnx) + bump_size
            
            Next jnx
        End If
    Next inx
    

On Error GoTo ErrorHandler


    Exit Sub
    
ErrorHandler:

    raise_err "clsVolSurface.bump_conv_upto", Err.description

End Sub


'----------------------------------------------------------------------------------
' Sub: bump_vol_surface
' DESC: Bump vol values and local vol values.
'       Copy the values to XXX_origin variables. XXX_origin variables are used to rewind the bumping.
'----------------------------------------------------------------------------------
Public Sub bump_vol(ByVal bump_size As Double) ', tenor_index As Integer, strike_index As Integer)

    Dim inx As Integer
    Dim jnx As Integer
    Dim no_of_rows As Integer
    Dim no_of_cols As Integer
    Dim bump_local_vol As Boolean
    
On Error GoTo ErrorHandler

    no_of_rows = grid_.no_of_dates
    no_of_cols = grid_.no_of_strikes
    
    ReDim vol_surface_origin_(1 To no_of_rows, 1 To no_of_cols) As Double
    
    For inx = 1 To no_of_rows
        For jnx = 1 To no_of_cols
                
            vol_surface_origin_(inx, jnx) = vol_surface_(inx, jnx)
            vol_surface_(inx, jnx) = vol_surface_(inx, jnx) + bump_size
        
        Next jnx
    Next inx
    
    
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsVolSurface.bump_local", Err.description

End Sub


'----------------------------------------------------------------------------------
' Sub: bump_skew
'----------------------------------------------------------------------------------
Public Sub bump_skew(ByVal bump_size As Double, ByVal s As Double)

    Dim inx As Integer
    Dim jnx As Integer
    Dim no_of_rows As Integer
    Dim no_of_cols As Integer
        
    
On Error GoTo ErrorHandler

    no_of_rows = grid_.no_of_dates
    no_of_cols = grid_.no_of_strikes
    
    ReDim vol_surface_origin_(1 To no_of_rows, 1 To no_of_cols) As Double
    
    
    For inx = 1 To no_of_rows
        For jnx = 1 To no_of_cols
                
            vol_surface_origin_(inx, jnx) = vol_surface_(inx, jnx)
            vol_surface_(inx, jnx) = vol_surface_(inx, jnx) + 1 / Sqr(grid_.tau(inx)) * (s - grid_.strikes(jnx)) / s * bump_size
        
        Next jnx
    Next inx
    
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsVolSurface.bump_skew", Err.description

End Sub


'----------------------------------------------------------------------------------
' Sub: bump_conv
'----------------------------------------------------------------------------------
Public Sub bump_conv(ByVal bump_size As Double, ByVal s As Double)

    Dim inx As Integer
    Dim jnx As Integer
    Dim no_of_rows As Integer
    Dim no_of_cols As Integer
        
    
On Error GoTo ErrorHandler

    no_of_rows = grid_.no_of_dates
    no_of_cols = grid_.no_of_strikes
    
    ReDim vol_surface_origin_(1 To no_of_rows, 1 To no_of_cols) As Double
    
    
    For inx = 1 To no_of_rows
        For jnx = 1 To no_of_cols
                
            vol_surface_origin_(inx, jnx) = vol_surface_(inx, jnx)
            vol_surface_(inx, jnx) = vol_surface_(inx, jnx) + 1 / grid_.tau(inx) * Abs((s - grid_.strikes(jnx))) / s * bump_size
        
        Next jnx
    Next inx
    
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsVolSurface.bump_conv", Err.description

End Sub



'Public Sub rewind_local_vol_bump(tenor_index As Integer, strike_index As Integer)
'
'
'On Error Resume Next
'
'
'
'    local_vol_surface_(tenor_index, strike_index) = local_vol_surface_origin_(tenor_index, strike_index)
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "clsSABRSurface.bump_vol_surface", Err.description
'
'End Sub

'----------------------------------------------------------------------------------
' Sub: rewind_vol_bump
' DESC: Reset vol values and local vol values to be the initial values.
'----------------------------------------------------------------------------------
Public Sub rewind_vol_bump()

    Dim inx As Integer
    Dim jnx As Integer
    Dim no_of_rows As Integer
    Dim no_of_cols As Integer
    Dim bump_local_vol As Boolean
  
    
On Error GoTo ErrorHandler

    no_of_rows = grid_.no_of_dates
    no_of_cols = grid_.no_of_strikes
    
        
    For inx = 1 To no_of_rows
        For jnx = 1 To no_of_cols
                
            vol_surface_(inx, jnx) = vol_surface_origin_(inx, jnx)
        
        Next jnx
    Next inx

    Exit Sub
    
ErrorHandler:

    raise_err "clsVolSurface.bump_vol_surface", Err.description

End Sub




'---------------------------------------------
' Sub: set_local_vol_surface
' Input: grid (vol surface layout)
'        vol_data( vol surface data )
'        local_vol_data( vol surface data )
'---------------------------------------------
Public Sub set_vol_surface(grid As clsPillarGrid, vol_data() As Double) ', local_vol_data() As Double)

On Error GoTo ErrorHandler
    
    vol_surface_ = vol_data
    

    Set grid_ = grid

    Exit Sub
    
ErrorHandler:

    raise_err "clsVolSurface.set_local_vol_surface", Err.description
    

End Sub

Public Sub set_vol_surface_origin(grid As clsPillarGrid, vol_data() As Double) ', local_vol_data() As Double)

On Error GoTo ErrorHandler
    
    vol_surface_ = vol_data
    

    Set grid_ = grid

    Exit Sub
    
ErrorHandler:

    raise_err "clsVolSurface.set_vol_surface_origin", Err.description
    

End Sub