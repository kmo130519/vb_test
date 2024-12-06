'===============================================================
' CLASS : clsSABRSurface
' DESC:
' Modified on
'  2013-10-10
'===============================================================

Option Explicit

'-----------------------------------
' KEY
'-----------------------------------
Public eval_date_ As Date
Public underlying_asset_ As String

Public real_date As Date

Public s_ As Double


Public local_vol_surface As clsVolSurface




Private Const h_ As Double = 0.001

Public Function copy_obj() As clsSABRSurface

    Dim rtn_obj As clsSABRSurface
    
    Set rtn_obj = New clsSABRSurface
    
    rtn_obj.eval_date_ = eval_date_
    rtn_obj.underlying_asset_ = underlying_asset_
    rtn_obj.real_date = real_date
    rtn_obj.s_ = s_
    
    Set rtn_obj.local_vol_surface = local_vol_surface.copy_obj

    Set copy_obj = rtn_obj
    
End Function

Public Sub set_local_vol_surface_(vol_surface_in As clsVolSurface)

    Set local_vol_surface = vol_surface_in

End Sub

Public Sub set_local_vol_surface(local_vol_grid As clsPillarGrid, local_vol_data() As Double)

On Error GoTo ErrorHandler

   
    If local_vol_surface Is Nothing Then
    
        Set local_vol_surface = New clsVolSurface
        
    End If
    
    local_vol_surface.set_vol_surface local_vol_grid, local_vol_data
    

    

    Exit Sub
    
ErrorHandler:

    raise_err "clsSABRSurface.set_local_vol_surface", Err.description
    

End Sub




'-------------------------------------------
' Getters
'-------------------------------------------
Public Function eval_date() As Date

    eval_date = eval_date_


End Function

Public Function underlying_asset() As String

    underlying_asset = underlying_asset_


End Function





'===============================
' 2015/10/05
' bump local vol up to time
'===============================

Public Sub bump_local_vol_upto(ByVal bump_size As Double, ByVal up_to_date As Date)

    local_vol_surface.bump_vol_upto bump_size, up_to_date
    

End Sub

Public Sub bump_local_vol_skew_upto(ByVal bump_size As Double, ByVal up_to_date As Date)

    local_vol_surface.bump_skew_upto bump_size * 2, up_to_date, s_
    

End Sub
Public Sub bump_local_vol_conv_upto(ByVal bump_size As Double, ByVal up_to_date As Date)

    local_vol_surface.bump_conv_upto bump_size * 2, up_to_date, s_
    

End Sub


Public Sub bump_vol_surface(ByVal bump_size As Double)


'    vol_surface.bump_vol bump_size
    local_vol_surface.bump_vol bump_size
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsSABRSurface.bump_vol_surface", Err.description

End Sub

Public Sub rewind_vol_bump()
    
'    vol_surface.rewind_vol_bump
    local_vol_surface.rewind_vol_bump

End Sub

Public Sub bump_skew(ByVal bump_size As Double)
    
'    vol_surface.bump_skew bump_size, s_
    local_vol_surface.bump_skew bump_size * 2, s_

End Sub

Public Sub bump_conv(ByVal bump_size As Double)
    
'    vol_surface.bump_conv bump_size, s_
    local_vol_surface.bump_conv bump_size * 2, s_

End Sub
'
'Public Function get_skew(maturity_date As Date, Optional s_in As Double = 0, Optional eval_date_in As Date = -1, Optional lower_strike As Double = 0.9, Optional upper_strike As Double = 1.1) As Double
'
'On Error GoTo ErrorHandler
'
'    Dim rtn_value As Double
'
'    Dim s As Double
'    Dim eval_date As Date
'
'    If eval_date_in > 0 Then
'        eval_date = eval_date_in
'    Else
'        eval_date = eval_date_
'    End If
'
'    If s_in > 0 Then
'        s = s_in
'    Else
'        s = s_
'    End If
'
'    rtn_value = (implied_vol(eval_date, s * lower_strike, maturity_date) - implied_vol(eval_date, s * upper_strike, maturity_date)) / (lower_strike - upper_strike)
'
'    get_skew = rtn_value
'
'
'    Exit Function
'
'ErrorHandler:
'
'    raise_err "clsSABRSurface.get_skew", Err.description & ":" & Me.underlying_asset & ":" & maturity_date
'
'End Function
'
'Public Function get_smile(maturity_date As Date, Optional s_in As Double = 0, Optional eval_date_in As Date = -1, Optional lower_strike As Double = 0.9, Optional mid_strike As Double = 1#, Optional upper_strike As Double = 1.1) As Double
'
'On Error GoTo ErrorHandler
'
'    Dim rtn_value As Double
'
'    Dim s As Double
'    Dim eval_date As Date
'
'    Dim h1 As Double
'    Dim h2 As Double
'
'    Dim vol_l As Double
'    Dim vol_m As Double
'    Dim vol_u As Double
'
'    If eval_date_in > 0 Then
'        eval_date = eval_date_in
'    Else
'        eval_date = eval_date_
'    End If
'
'    h1 = mid_strike - lower_strike
'    h2 = upper_strike - mid_strike
'
'
'    If s_in > 0 Then
'        s = s_in
'    Else
'        s = s_
'    End If
'
'    vol_l = implied_vol(eval_date, s * lower_strike, maturity_date)
'    vol_m = implied_vol(eval_date, s * mid_strike, maturity_date)
'    vol_u = implied_vol(eval_date, s * upper_strike, maturity_date)
'
'    rtn_value = 2 / h1 / (h1 + h2) * vol_l + (-2 / h1 / (h1 + h2) - 2 / h2 / (h1 + h2)) * vol_m + 2 / h2 / (h1 + h2) * vol_u
'    'rtn_value = (implied_vol(eval_date, s * lower_strike, maturity_date) + implied_vol(eval_date, s * upper_strike, maturity_date) - 2 * implied_vol(eval_date, s * mid_strike, maturity_date)) _
'              / (lower_strike - upper_strike)
'
'    get_smile = rtn_value
'
'
'    Exit Function
'
'ErrorHandler:
'
'    raise_err "clsSABRSurface.get_skew", Err.description & ":" & Me.underlying_asset & ":" & maturity_date
'
'End Function