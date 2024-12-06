Option Explicit

Public value As Double
Public delta As Double
Public implied_tree_delta As Double ' Refer to Derman, 1999
Public implied_tree_gamma As Double ' Refer to Derman, 1999
Public sticky_strike_delta As Double ' Refer to Derman, 1999
Public sticky_strike_gamma As Double ' Refer to Derman, 1999
Public sticky_moneyness_delta As Double ' Refer to Derman, 1999
Public sticky_moneyness_gamma As Double ' Refer to Derman, 1999
Public percent_dollar_gamma_implied_tree As Double ' Refer to Derman, 1999
Public vega As Double
Public w_vega As Double

Public implied_tree_ratio As Double


Public gamma As Double


Public ul_code As String

Public no_of_ul As Integer '2015/10/19


Private deltas_() As Double
Private implied_tree_deltas_() As Double ' Refer to Derman, 1999
Private implied_tree_gammas_() As Double ' Refer to Derman, 1999
Private sticky_strike_deltas_() As Double ' Refer to Derman, 1999
Private sticky_strike_gammas_() As Double ' Refer to Derman, 1999
'added 2018.09.07
Private sticky_moneyness_deltas_() As Double
Private sticky_moneyness_gammas_() As Double

Private vegas_() As Double
Private volgas_() As Double
Private vannas_() As Double
Private gammas_() As Double
Public cross_gamma12 As Double
Public cross_gamma13 As Double
Public cross_gamma23 As Double
Private corr_sens_() As Double


Private partial_skews_() As Double ' 2015-10-05
Private term_skews_() As Double ' 2015-10-05
Private partial_vegas_() As Double ' 2015-10-05
Private term_vegas_() As Double ' 2015-10-05
Private partial_convs_() As Double ' 2015-10-05
Private term_convs_() As Double ' 2015-10-05
Private term_dates_() As Date ' 2015-10-05
Public no_of_tenors As Integer

Public duration As Double

Private eff_duration_() As Double '2016-01-22

Private skew_s_() As Double
Private conv_s_() As Double

Private xAxis_() As Double
Private yAxis_() As Double
Private zAxis_() As Double
Private snapshot_value_() As Double

Private value_matrix_() As clsValueMatrix



Public theta As Double
Public volga As Double
Public vanna As Double
Public rho As Double ' for PL Currency
Private rho_ul_() As Double ' for drift term

Public skew_s As Double
Public conv_s As Double

Public asset_code As String

Public implied_vol As Double

Public ul_price As Double
Public ul_initial_price As Double



Private ul_prices_() As Double
Private ul_initial_prices_() As Double

Public eval_date As Date

Public term_date As Date

Public simulation_variance As Double ' Variance of the simulation paths

' Greek Code
' VALUE, DELTA, GAMMA, VEGA, THETA, VOLGA, VANNA, SKEW, CONV


Public Sub set_value_matrix(inx As Integer, a_value_matrix As clsValueMatrix)

   
    
On Error Resume Next

    Dim temp_inx As Integer
    
    temp_inx = UBound(value_matrix_)
    
    If (Err.number = 9) Or temp_inx < inx Then
    
        ReDim Preserve value_matrix_(1 To inx) As clsValueMatrix
    
    End If
    
On Error GoTo ErrorHandler
    
    Set value_matrix_(inx) = a_value_matrix
    
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_value_matrix"
    

End Sub

Public Function value_matrix(inx As Integer) As clsValueMatrix

    Set value_matrix = value_matrix_(inx)

End Function





Public Sub redim_arrays(no_of_ul_in As Integer, Optional no_of_term As Integer = 6)

    no_of_ul = no_of_ul_in
    
    If no_of_ul_in > 0 Then
        ReDim Preserve deltas_(1 To no_of_ul_in) As Double
        ReDim Preserve vegas_(1 To no_of_ul_in) As Double
        ReDim Preserve volgas_(1 To no_of_ul_in) As Double
        ReDim Preserve vannas_(1 To no_of_ul_in) As Double
        ReDim Preserve gammas_(1 To no_of_ul_in) As Double
        ReDim Preserve skew_s_(1 To no_of_ul_in) As Double
        ReDim Preserve conv_s_(1 To no_of_ul_in) As Double
        ReDim Preserve ul_prices_(1 To no_of_ul_in) As Double
        
        ReDim Preserve eff_duration_(1 To no_of_ul_in) As Double
        
        ReDim Preserve ul_initial_prices_(1 To no_of_ul_in) As Double
        
        ReDim Preserve sticky_strike_deltas_(1 To no_of_ul_in) As Double
        ReDim Preserve sticky_strike_gammas_(1 To no_of_ul_in) As Double
        ReDim Preserve implied_tree_deltas_(1 To no_of_ul_in) As Double
        
        ReDim Preserve corr_sens_(1 To no_of_ul_in, 1 To no_of_ul_in) As Double
        
        ReDim Preserve rho_ul_(1 To no_of_ul_in) As Double
        
        If no_of_term > 0 Then
            ReDim Preserve term_vegas_(1 To no_of_ul_in, 1 To no_of_term) As Double
            ReDim Preserve partial_vegas_(1 To no_of_ul_in, 1 To no_of_term) As Double
            ReDim Preserve term_skews_(1 To no_of_ul_in, 1 To no_of_term) As Double
            ReDim Preserve partial_skews_(1 To no_of_ul_in, 1 To no_of_term) As Double
            ReDim Preserve term_convs_(1 To no_of_ul_in, 1 To no_of_term) As Double
            ReDim Preserve partial_convs_(1 To no_of_ul_in, 1 To no_of_term) As Double
            ReDim Preserve term_dates_(1 To no_of_ul_in, 1 To no_of_term) As Date
        End If
        
        no_of_tenors = no_of_term
    End If


End Sub


Public Function copy_obj() As clsGreeks

    Dim temp_cls As clsGreeks
    Dim no_of_uls As Integer
    Dim inx As Integer
    Dim jnx As Integer
    
    Set temp_cls = New clsGreeks
    
    temp_cls.value = value
    temp_cls.delta = delta
    temp_cls.vega = vega
    temp_cls.gamma = gamma
    temp_cls.theta = theta
    temp_cls.volga = volga
    temp_cls.vanna = vanna
    
    temp_cls.skew_s = skew_s
    temp_cls.conv_s = conv_s
    
    temp_cls.cross_gamma = cross_gamma
    
    For inx = 1 To no_of_uls
        For jnx = inx + 1 To no_of_uls
            temp_cls.set_corr_sens inx, jnx, corr_sens_(inx, jnx)
        Next jnx
    Next inx
    
    temp_cls.implied_tree_delta = implied_tree_delta
    temp_cls.sticky_strike_delta = sticky_strike_delta
    
    no_of_uls = get_array_size_double(deltas_)
    temp_cls.redim_arrays no_of_uls
    
    For inx = 1 To no_of_uls
        temp_cls.set_deltas inx, deltas_(inx)
    Next inx
    
    no_of_uls = get_array_size_double(vegas_)
    For inx = 1 To no_of_uls
        temp_cls.set_vegas inx, vegas_(inx)
    Next inx
    
    no_of_uls = get_array_size_double(vannas_)
    For inx = 1 To no_of_uls
        temp_cls.set_vannas inx, vannas_(inx)
    Next inx
    
    no_of_uls = get_array_size_double(volgas_)
    For inx = 1 To no_of_uls
        temp_cls.set_volgas inx, volgas_(inx)
    Next inx
        
    no_of_uls = get_array_size_double(sticky_strike_deltas_)
    For inx = 1 To no_of_uls
        temp_cls.set_deltas inx, sticky_strike_deltas_(inx)
    Next inx
    
            
    no_of_uls = get_array_size_double(implied_tree_deltas_)
    For inx = 1 To no_of_uls
        temp_cls.set_deltas inx, implied_tree_deltas_(inx)
    Next inx
    
    
    no_of_uls = get_array_size_double(gammas_)
    For inx = 1 To no_of_uls
        temp_cls.set_gammas inx, gammas_(inx)
    Next inx
    
    no_of_uls = get_array_size_double(skew_s_)
    For inx = 1 To no_of_uls
        temp_cls.set_skews inx, skew_s_(inx)
    Next inx
    
    no_of_uls = get_array_size_double(conv_s_)
    For inx = 1 To no_of_uls
        temp_cls.set_convs inx, conv_s_(inx)
    Next inx
    
    no_of_uls = get_array_size_double(ul_prices_)
    For inx = 1 To no_of_uls
        temp_cls.set_ul_prices inx, ul_prices_(inx)
    Next inx
                
    no_of_uls = get_array_size_double(ul_initial_prices_)
    For inx = 1 To no_of_uls
        temp_cls.set_ul_initial_prices inx, ul_initial_prices_(inx)
    Next inx
                
    no_of_uls = get_array_size_double(rho_ul_)
    
    For inx = 1 To no_of_uls
        temp_cls.set_rho_ul inx, rho_ul_(inx)
    Next inx
            
    temp_cls.simulation_variance = simulation_variance
    
    
    
    Set copy_obj = temp_cls

End Function

Public Sub set_xAxis(xAxis() As Double, Optional log_value As Boolean = True, Optional as_percentage_of As Double = 1#)
    
    Dim inx As Integer
    
On Error Resume Next

    ReDim xAxis_(LBound(xAxis) To UBound(xAxis)) As Double
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        
        raise_err "set_xAxis", "xAxis not initialized.."
    
    End If
    
    If log_value Then
    
        For inx = LBound(xAxis) To UBound(xAxis)
            
            xAxis_(inx) = Exp(xAxis(inx)) / as_percentage_of
        
        Next inx
    
    Else
    
        For inx = LBound(xAxis) To UBound(xAxis)
        
            xAxis_(inx) = xAxis(inx) / as_percentage_of
            
        Next inx
        
    End If
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_xAxis"


End Sub


Public Sub set_yAxis(yAxis() As Double, Optional log_value As Boolean = True, Optional as_percentage_of As Double = 1#)
    
    Dim inx As Integer
    
On Error Resume Next

    ReDim yAxis_(LBound(yAxis) To UBound(yAxis)) As Double
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        
        raise_err "set_yAxis", "yAxis not initialized.."
    
    End If
    
    If log_value Then
    
        For inx = LBound(yAxis) To UBound(yAxis)
            
            yAxis_(inx) = Exp(yAxis(inx)) / as_percentage_of
        
        Next inx
    
    Else
        
        For inx = LBound(yAxis) To UBound(yAxis)
            
            yAxis_(inx) = (yAxis(inx)) / as_percentage_of
        
        Next inx
        
    End If
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_yAxis"


End Sub
Public Sub set_zAxis(zAxis() As Double, Optional log_value As Boolean = True, Optional as_percentage_of As Double = 1#)
    
    Dim inx As Integer
    
On Error Resume Next

    ReDim zAxis_(LBound(zAxis) To UBound(zAxis)) As Double
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        
        raise_err "set_zAxis", "yAxis not initialized.."
    
    End If
    
    If log_value Then
    
        For inx = LBound(zAxis) To UBound(zAxis)
            
            zAxis_(inx) = Exp(zAxis(inx)) / as_percentage_of
        
        Next inx
    
    Else
        
    
        For inx = LBound(zAxis) To UBound(zAxis)
            
            zAxis_(inx) = (zAxis(inx)) / as_percentage_of
        
        Next inx
        
    End If
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_zAxis"


End Sub

Public Sub set_snapshot_value(snapshot_value() As Double, Optional notional As Double = 1)
    
    Dim inx As Long
    
On Error Resume Next

    ReDim snapshot_value_(LBound(snapshot_value) To UBound(snapshot_value)) As Double
    
On Error GoTo ErrorHandler
    
    If Err.number = 9 Then
        
        raise_err "set_snapshot_value", "snapshot_value_ not initialized.."
    
    End If
    
   
   For inx = LBound(snapshot_value) To UBound(snapshot_value)
            
        snapshot_value_(inx) = snapshot_value(inx) * notional
        
   Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "set_snaptshot_value", Err.description


End Sub



Public Function xAxis(ByVal inx As Integer) As Double

    xAxis = xAxis_(inx)

End Function

Public Function yAxis(ByVal inx As Integer) As Double

    yAxis = yAxis_(inx)

End Function
Public Function snapshot_value(ByVal inx As Integer) As Double

    snapshot_value = snapshot_value_(inx)

End Function
'Public Function snapshot_value_bump(ByVal inx As Integer) As Double
'
'    snapshot_value_bump = snapshot_value_vol_bump_(inx)
'
'End Function

Public Function get_xAxis() As Double()

    get_xAxis = xAxis_

End Function

Public Function get_yAxis() As Double()

    get_yAxis = yAxis_

End Function
Public Function get_zAxis() As Double()

    get_zAxis = zAxis_

End Function
Public Function get_snapshot_value() As Double()

    get_snapshot_value = snapshot_value_

End Function
'Public Function get_snapshot_value_bump() As Double()
'
'    get_snapshot_value_bump = snapshot_value_vol_bump_
'
'End Function


Public Function ul_prices(ByVal inx As Integer) As Double

    ul_prices = ul_prices_(inx)
    
End Function
Public Function ul_initial_prices(ByVal inx As Integer) As Double

    ul_initial_prices = ul_initial_prices_(inx)
    
End Function

Public Function deltas(ByVal inx As Integer) As Double

On Error GoTo ErrorHandler

    deltas = deltas_(inx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        deltas = 0
    Else
        raise_err "clsGreeks.deltas", Err.description
    End If
    
End Function


Public Function eff_duration(ByVal inx As Integer) As Double

On Error GoTo ErrorHandler

    eff_duration = eff_duration_(inx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        eff_duration = 0
    Else
        raise_err "clsGreeks.deltas", Err.description
    End If
    
End Function

'2015-10-05
Public Function term_vega(ByVal inx As Integer, ByVal jnx As Integer) As Double

On Error GoTo ErrorHandler

    term_vega = term_vegas_(inx, jnx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        term_vega = 0
    Else
        raise_err "clsGreeks.term_vega", Err.description
    End If
    
End Function

'2015-10-05
Public Function partial_vega(ByVal inx As Integer, ByVal jnx As Integer) As Double

On Error GoTo ErrorHandler

    partial_vega = partial_vegas_(inx, jnx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        partial_vega = 0
    Else
        raise_err "clsGreeks.partial_vega", Err.description
    End If
    
End Function
Public Function term_skew(ByVal inx As Integer, ByVal jnx As Integer) As Double

On Error GoTo ErrorHandler

    term_skew = term_skews_(inx, jnx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        term_skew = 0
    Else
        raise_err "clsGreeks.term_skew", Err.description
    End If
    
End Function

'2015-10-05
Public Function partial_skew(ByVal inx As Integer, ByVal jnx As Integer) As Double

On Error GoTo ErrorHandler

    partial_skew = partial_skews_(inx, jnx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        partial_skew = 0
    Else
        raise_err "clsGreeks.partial_skew", Err.description
    End If
    
End Function

Public Function term_conv(ByVal inx As Integer, ByVal jnx As Integer) As Double

On Error GoTo ErrorHandler

    term_conv = term_convs_(inx, jnx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        term_conv = 0
    Else
        raise_err "clsGreeks.term_conv", Err.description
    End If
    
End Function

'2015-10-05
Public Function partial_conv(ByVal inx As Integer, ByVal jnx As Integer) As Double

On Error GoTo ErrorHandler

    partial_conv = partial_convs_(inx, jnx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        partial_conv = 0
    Else
        raise_err "clsGreeks.partial_conv", Err.description
    End If
    
End Function
'2015-10-05
Public Function get_term_date(ByVal inx As Integer, ByVal jnx As Integer) As Date

On Error GoTo ErrorHandler

    get_term_date = term_dates_(inx, jnx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        get_term_date = 0
    Else
        raise_err "clsGreeks.term_date", Err.description
    End If
    
End Function

Public Function corr_sens(ByVal inx As Integer, ByVal jnx As Integer) As Double

On Error GoTo ErrorHandler

    corr_sens = corr_sens_(inx, jnx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        corr_sens = 0
    Else
        raise_err "clsGreeks.corr_sens", Err.description
    End If
    
End Function
Public Function implied_tree_deltas(ByVal inx As Integer) As Double

On Error GoTo ErrorHandler

    implied_tree_deltas = implied_tree_deltas_(inx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        implied_tree_deltas = 0
    Else
        raise_err "clsGreeks.implied_tree_deltas", Err.description
    End If
    
End Function
Public Function implied_tree_gammas(ByVal inx As Integer) As Double

On Error GoTo ErrorHandler

    implied_tree_gammas = implied_tree_gammas_(inx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        implied_tree_gammas = 0
    Else
        raise_err "clsGreeks.implied_tree_gammas", Err.description
    End If
    
End Function

'added 2018.9.7
Public Function sticky_moneyness_deltas(ByVal inx As Integer) As Double

On Error GoTo ErrorHandler

    sticky_moneyness_deltas = sticky_moneyness_deltas_(inx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        sticky_moneyness_deltas = 0
    Else
        raise_err "clsGreeks.sticky_moneyness_deltas", Err.description
    End If
    
End Function

Public Function sticky_moneyness_gammas(ByVal inx As Integer) As Double

On Error GoTo ErrorHandler

    sticky_moneyness_gammas = sticky_moneyness_gammas_(inx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        sticky_moneyness_gammas = 0
    Else
        raise_err "clsGreeks.sticky_moneyness_gammas", Err.description
    End If
    
End Function


Public Function sticky_strike_deltas(ByVal inx As Integer) As Double

On Error GoTo ErrorHandler

    sticky_strike_deltas = sticky_strike_deltas_(inx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        sticky_strike_deltas = 0
    Else
        raise_err "clsGreeks.sticky_strike_deltas", Err.description
    End If
    
End Function

Public Function sticky_strike_gammas(ByVal inx As Integer) As Double

On Error GoTo ErrorHandler

    sticky_strike_gammas = sticky_strike_gammas_(inx)
    
    Exit Function

ErrorHandler:
    
    If Err.number = 9 Then
        sticky_strike_gammas = 0
    Else
        raise_err "clsGreeks.sticky_strike_gammas", Err.description
    End If
    
End Function

Public Function gammas(ByVal inx As Integer) As Double

On Error GoTo ErrorHandler

    gammas = gammas_(inx)
        
    Exit Function
    
ErrorHandler:
    
    If Err.number = 9 Then
        gammas = 0
    Else
        raise_err "clsGreeks.deltas", Err.description
    End If
    
End Function
Public Function rho_ul(Optional ByVal inx As Integer = 1) As Double

On Error GoTo ErrorHandler

    rho_ul = rho_ul_(inx)
        
    Exit Function
    
ErrorHandler:
    
    If Err.number = 9 Then
        rho_ul = 0
    Else
        raise_err "clsGreeks.rho_ul", Err.description
    End If
    
End Function
Public Function skew_s_s(ByVal inx As Integer) As Double


On Error GoTo ErrorHandler

    skew_s_s = skew_s_(inx)
    
    Exit Function
    
ErrorHandler:
    
    If Err.number = 9 Then
        skew_s_s = 0
    Else
        raise_err "clsGreeks.skew_s_s", Err.description
    End If
        
End Function
Public Function conv_s_s(ByVal inx As Integer) As Double


On Error GoTo ErrorHandler

    conv_s_s = conv_s_(inx)
    
    Exit Function
    
ErrorHandler:
    
    If Err.number = 9 Then
        conv_s_s = 0
    Else
        raise_err "clsGreeks.conv_s_s", Err.description
    End If
        
End Function
Public Function vegas(ByVal inx As Integer) As Double


On Error GoTo ErrorHandler

    vegas = vegas_(inx)
    
    Exit Function
    
ErrorHandler:
    
    If Err.number = 9 Then
        vegas = 0
    Else
        raise_err "clsGreeks.vegas", Err.description
    End If
            
    
    
End Function

Public Function vannas(ByVal inx As Integer) As Double


On Error GoTo ErrorHandler

    vannas = vannas_(inx)
    
    Exit Function
    
ErrorHandler:
    
    If Err.number = 9 Then
        vannas = 0
    Else
        raise_err "clsGreeks.vannas", Err.description
    End If
            
    
    
End Function

Public Function volgas(ByVal inx As Integer) As Double


On Error GoTo ErrorHandler

    volgas = volgas_(inx)
    
    Exit Function
    
ErrorHandler:
    
    If Err.number = 9 Then
        volgas = 0
    Else
        raise_err "clsGreeks.volgas", Err.description
    End If
            
    
    
End Function


'------------------------
' 2015-10-19
'------------------------

Public Sub to_term_vega(current_date As Date) '<------

    Dim inx As Integer
    Dim jnx As Integer
    Dim knx As Integer
    
    
    
On Error GoTo ErrorHandler
    
    If no_of_tenors > 0 Then 'get_array_size_date(term_dates_) > 0 Then
    
        If no_of_tenors > 1 Then
        
            If term_dates_(1, no_of_tenors) > term_dates_(1, no_of_tenors - 1) Then  '<=====
        
                For inx = 1 To no_of_ul
                    
                    '-----
                    'treat the last term
                    term_vegas_(inx, no_of_tenors) = partial_vegas_(inx, no_of_tenors) - partial_vegas_(inx, no_of_tenors - 1)
                    term_vegas_(inx, no_of_tenors) = term_vegas_(inx, no_of_tenors) * (term_dates_(inx, no_of_tenors) - current_date) / (term_dates_(inx, no_of_tenors) - term_dates_(inx, no_of_tenors - 1))
                    '<----
                
                    For jnx = no_of_tenors - 1 To 2 Step -1
                    
                        term_vegas_(inx, jnx) = partial_vega(inx, jnx) - partial_vegas_(inx, jnx - 1) ' FWD Vega
                        
                        For knx = jnx + 1 To no_of_tenors
                            term_vegas_(inx, jnx) = term_vegas_(inx, jnx) - term_vegas_(inx, knx) * (term_dates_(inx, jnx) - term_dates_(inx, jnx - 1)) / (term_dates_(inx, knx) - current_date) 'Extract Q_k / T_k
                        Next knx
                        
                        term_vegas_(inx, jnx) = term_vegas_(inx, jnx) * (term_dates_(inx, jnx) - current_date) / (term_dates_(inx, jnx) - term_dates_(inx, jnx - 1))
                    
                    Next jnx
                    
                    '-- treat the first term
                    term_vegas_(inx, 1) = partial_vegas_(inx, 1)
                    
                    For knx = 2 To no_of_tenors
                        term_vegas_(inx, 1) = term_vegas_(inx, 1) - term_vegas_(inx, knx) * (term_dates_(inx, 1) - current_date) / (term_dates_(inx, knx) - current_date) 'Extract Q_k / T_k
                    Next knx
                        
                    term_vegas_(inx, 1) = term_vegas_(inx, 1) * (term_dates_(inx, 1) - current_date) / (term_dates_(inx, 1) - current_date)
                    '<-----
                
                Next inx
            
            End If
            
        Else
        
            For inx = 1 To no_of_ul
                term_vegas_(inx, no_of_tenors) = partial_vegas_(inx, no_of_tenors)
            Next inx
                
        End If
            

    
    End If


    Exit Sub
    
ErrorHandler:

    raise_err "clsGreeks.to_term_vega", Err.description


End Sub


Public Sub to_term_skew(current_date As Date) '<------

    Dim inx As Integer
    Dim jnx As Integer
    Dim knx As Integer
    
    
    
On Error GoTo ErrorHandler
    
    If no_of_tenors > 0 Then 'get_array_size_date(term_dates_) > 0 Then
    
        If no_of_tenors > 1 Then
        
            If term_dates_(1, no_of_tenors) > term_dates_(1, no_of_tenors - 1) Then  '<=====
        
                For inx = 1 To no_of_ul
                    
                    '-----
                    'treat the last term
                    term_skews_(inx, no_of_tenors) = partial_skews_(inx, no_of_tenors) - partial_skews_(inx, no_of_tenors - 1)
                    term_skews_(inx, no_of_tenors) = term_skews_(inx, no_of_tenors) * (term_dates_(inx, no_of_tenors) - current_date) / (term_dates_(inx, no_of_tenors) - term_dates_(inx, no_of_tenors - 1))
                    '<----
                
                    For jnx = no_of_tenors - 1 To 2 Step -1
                    
                        term_skews_(inx, jnx) = partial_skew(inx, jnx) - partial_skews_(inx, jnx - 1) ' FWD skew
                        
                        For knx = jnx + 1 To no_of_tenors
                            term_skews_(inx, jnx) = term_skews_(inx, jnx) - term_skews_(inx, knx) * (term_dates_(inx, jnx) - term_dates_(inx, jnx - 1)) / (term_dates_(inx, knx) - current_date) 'Extract Q_k / T_k
                        Next knx
                        
                        term_skews_(inx, jnx) = term_skews_(inx, jnx) * (term_dates_(inx, jnx) - current_date) / (term_dates_(inx, jnx) - term_dates_(inx, jnx - 1))
                    
                    Next jnx
                    
                    '-- treat the first term
                    term_skews_(inx, 1) = partial_skews_(inx, 1)
                    
                    For knx = 2 To no_of_tenors
                        term_skews_(inx, 1) = term_skews_(inx, 1) - term_skews_(inx, knx) * (term_dates_(inx, 1) - current_date) / (term_dates_(inx, knx) - current_date) 'Extract Q_k / T_k
                    Next knx
                        
                    term_skews_(inx, 1) = term_skews_(inx, 1) * (term_dates_(inx, 1) - current_date) / (term_dates_(inx, 1) - current_date)
                    '<-----
                
                Next inx
            End If
        
        Else
    
            For inx = 1 To no_of_ul
                term_skews_(inx, no_of_tenors) = partial_skews_(inx, no_of_tenors)
            Next inx
            
        End If
        
    
    
    End If


    Exit Sub
    
ErrorHandler:

    raise_err "clsGreeks.to_term_skew", Err.description


End Sub

Public Sub to_term_conv(current_date As Date) '<------

    Dim inx As Integer
    Dim jnx As Integer
    Dim knx As Integer
    
    
    
On Error GoTo ErrorHandler
    
    If no_of_tenors > 0 Then 'get_array_size_date(term_dates_) > 0 Then
    
        If no_of_tenors > 1 Then
            
            If term_dates_(1, no_of_tenors) > term_dates_(1, no_of_tenors - 1) Then  '<=====
            
                For inx = 1 To no_of_ul
                    
                    '-----
                    'treat the last term
                    term_convs_(inx, no_of_tenors) = partial_convs_(inx, no_of_tenors) - partial_convs_(inx, no_of_tenors - 1)
                    term_convs_(inx, no_of_tenors) = term_convs_(inx, no_of_tenors) * (term_dates_(inx, no_of_tenors) - current_date) / (term_dates_(inx, no_of_tenors) - term_dates_(inx, no_of_tenors - 1))
                    '<----
                
                    For jnx = no_of_tenors - 1 To 2 Step -1
                    
                        term_convs_(inx, jnx) = partial_conv(inx, jnx) - partial_convs_(inx, jnx - 1) ' FWD conv
                        
                        For knx = jnx + 1 To no_of_tenors
                            term_convs_(inx, jnx) = term_convs_(inx, jnx) - term_convs_(inx, knx) * (term_dates_(inx, jnx) - term_dates_(inx, jnx - 1)) / (term_dates_(inx, knx) - current_date) 'Extract Q_k / T_k
                        Next knx
                        
                        term_convs_(inx, jnx) = term_convs_(inx, jnx) * (term_dates_(inx, jnx) - current_date) / (term_dates_(inx, jnx) - term_dates_(inx, jnx - 1))
                    
                    Next jnx
                    
                    '-- treat the first term
                    term_convs_(inx, 1) = partial_convs_(inx, 1)
                    
                    For knx = 2 To no_of_tenors
                        term_convs_(inx, 1) = term_convs_(inx, 1) - term_convs_(inx, knx) * (term_dates_(inx, 1) - current_date) / (term_dates_(inx, knx) - current_date) 'Extract Q_k / T_k
                    Next knx
                        
                    term_convs_(inx, 1) = term_convs_(inx, 1) * (term_dates_(inx, 1) - current_date) / (term_dates_(inx, 1) - current_date)
                    '<-----
                
                Next inx
        
            End If
            
        Else
        
            For inx = 1 To no_of_ul
                term_convs_(inx, no_of_tenors) = partial_convs_(inx, no_of_tenors)
            Next inx
            
        End If
        
    
    
    End If


    Exit Sub
    
ErrorHandler:

    raise_err "clsGreeks.to_term_conv", Err.description


End Sub



Public Sub set_ul_prices(ByVal inx As Integer, ByVal ul_price_value As Double)

    ul_prices_(inx) = ul_price_value

End Sub
Public Sub set_ul_initial_prices(ByVal inx As Integer, ByVal ul_price_value As Double)

    ul_initial_prices_(inx) = ul_price_value

End Sub
Public Sub set_implied_tree_deltas(ByVal inx As Integer, ByVal delta_value As Double)


    implied_tree_deltas_(inx) = delta_value

End Sub
Public Sub set_sticky_strike_deltas(ByVal inx As Integer, ByVal delta_value As Double)


    sticky_strike_deltas_(inx) = delta_value

End Sub

Public Sub set_sticky_strike_gammas(ByVal inx As Integer, ByVal gamma_value As Double)


    sticky_strike_gammas_(inx) = gamma_value

End Sub
Public Sub set_eff_durations(ByVal inx As Integer, ByVal eff_duration_value As Double)


    eff_duration_(inx) = eff_duration_value

End Sub
Public Sub set_deltas(ByVal inx As Integer, ByVal delta_value As Double)


    deltas_(inx) = delta_value

End Sub

Public Sub set_gammas(ByVal inx As Integer, ByVal delta_value As Double)


    gammas_(inx) = delta_value

End Sub

Public Sub set_rho_ul(ByVal inx As Integer, ByVal rho_value As Double)


    rho_ul_(inx) = rho_value

End Sub


Public Sub set_vegas(ByVal inx As Integer, ByVal delta_value As Double)


    vegas_(inx) = delta_value

End Sub
Public Sub set_vannas(ByVal inx As Integer, ByVal value As Double)


    vannas_(inx) = value

End Sub
Public Sub set_volgas(ByVal inx As Integer, ByVal value As Double)


    volgas_(inx) = value

End Sub

Public Sub set_skews(ByVal inx As Integer, ByVal skew_value As Double)


    skew_s_(inx) = skew_value

End Sub

Public Sub set_convs(ByVal inx As Integer, ByVal conv_value As Double)


    conv_s_(inx) = conv_value

End Sub

'--------------------------
'2015-10-05
Public Sub set_term_vegas(ByVal inx As Integer, ByVal jnx As Integer, ByVal vega_value As Double)

    term_vegas_(inx, jnx) = vega_value
    

End Sub
'--------------------------
'2015-10-05
Public Sub set_partial_vegas(ByVal inx As Integer, ByVal jnx As Integer, ByVal vega_value As Double)

    partial_vegas_(inx, jnx) = vega_value
    

End Sub


'--------------------------
'2015-10-19
Public Sub set_term_skews(ByVal inx As Integer, ByVal jnx As Integer, ByVal skew_value As Double)

    term_skews_(inx, jnx) = skew_value
    

End Sub

'--------------------------
'2015-10-19
Public Sub set_term_convs(ByVal inx As Integer, ByVal jnx As Integer, ByVal conv_value As Double)

    term_convs_(inx, jnx) = conv_value
    

End Sub

'--------------------------
'2015-10-19
Public Sub set_term_skews_per_ul(ByVal ul_inx As Integer, skew_values() As Double)

    Dim inx As Integer
    
On Error GoTo ErrorHandler

    If get_array_size_double(skew_values) < no_of_tenors Then
    
        ReDim Preserve skew_values(1 To no_of_tenors) As Double
        
    End If

    For inx = 1 To no_of_tenors
        term_skews_(ul_inx, inx) = skew_values(inx)
    Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsGreeks.set_term_skews_per_ul", Err.description
    

End Sub

'--------------------------
'2015-10-19
Public Sub set_term_convs_per_ul(ByVal ul_inx As Integer, conv_values() As Double)

    Dim inx As Integer
    
On Error GoTo ErrorHandler

    If get_array_size_double(conv_values) < no_of_tenors Then
    
        ReDim Preserve conv_values(1 To no_of_tenors) As Double
        
    End If

    For inx = 1 To no_of_tenors
        term_convs_(ul_inx, inx) = conv_values(inx)
    Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsGreeks.set_term_convs_per_ul", Err.description
    

End Sub
'--------------------------
'2015-10-19
Public Sub set_term_vegas_per_ul(ByVal ul_inx As Integer, vega_values() As Double)

    Dim inx As Integer
    
On Error GoTo ErrorHandler

    If get_array_size_double(vega_values) < no_of_tenors Then
    
        ReDim Preserve vega_values(1 To no_of_tenors) As Double
        
    End If

    For inx = 1 To no_of_tenors
        term_vegas_(ul_inx, inx) = vega_values(inx)
    Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsGreeks.set_term_vegas_per_ul", Err.description
    

End Sub
'--------------------------
'2015-10-19
Public Sub set_partial_skews(ByVal inx As Integer, ByVal jnx As Integer, ByVal skew_value As Double)

    partial_skews_(inx, jnx) = skew_value
    

End Sub

'--------------------------
'2015-10-19
Public Sub set_partial_convs(ByVal inx As Integer, ByVal jnx As Integer, ByVal conv_value As Double)

    partial_convs_(inx, jnx) = conv_value
    

End Sub
Public Sub set_term_dates(ByVal inx As Integer, ByVal jnx As Integer, ByVal term_date_in As Date)

    term_dates_(inx, jnx) = term_date_in
    

End Sub

'--------------------------
'2015-10-19
Public Sub set_term_dates_per_ul(ByVal ul_inx As Integer, term_date_in() As Date)

    Dim inx As Integer
    
On Error GoTo ErrorHandler

    For inx = 1 To no_of_tenors
        term_dates_(ul_inx, inx) = term_date_in(inx)
    Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsGreeks.set_term_vegas_per_ul", Err.description
    

End Sub


Public Sub set_corr_sens(ByVal inx As Integer, ByVal jnx As Integer, ByVal corr_sens_value As Double)

    corr_sens_(inx, jnx) = corr_sens_value
    corr_sens_(jnx, inx) = corr_sens_value

End Sub

Public Sub set_all_ul_initial_prices(ul_initial_prices() As Double)


    ul_initial_prices_ = ul_initial_prices

End Sub
Public Sub set_all_ul_prices(ul_prices() As Double)


    ul_prices_ = ul_prices

End Sub
Public Function get_all_gammas() As Double()


    get_all_gammas = gammas_

End Function
Public Function get_all_deltas() As Double()


    get_all_deltas = deltas_

End Function

Public Function get_all_eff_durations() As Double()


    get_all_eff_durations = eff_duration_

End Function
Public Function get_all_rho_ul() As Double()


    get_all_rho_ul = rho_ul_

End Function

Public Function get_all_vegas() As Double()


    get_all_vegas = vegas_

End Function

Public Function get_all_skews() As Double()


    get_all_skews = skew_s_

End Function

Public Function get_all_term_skews() As Double()


    get_all_term_skews = term_skews_

End Function


Public Function get_all_convs() As Double()


    get_all_convs = conv_s_

End Function

Public Function get_all_term_convs() As Double()


    get_all_term_convs = term_convs_

End Function

Public Function get_all_term_vegas() As Double()


    get_all_term_vegas = term_vegas_

End Function

Public Function get_all_term_dates() As Date()


    get_all_term_dates = term_dates_

End Function
Public Function get_all_vannas() As Double()


    get_all_vannas = vannas_

End Function
Public Sub set_all_eff_durations(duration_value() As Double, Optional input_base As Integer = 1)

    Dim inx As Integer
    Dim array_size As Integer
    
    
    array_size = get_array_size_double(duration_value)
    
    ReDim eff_duration_(1 To array_size) As Double
    
    For inx = 1 To array_size
    
        eff_duration_(inx) = duration_value(inx + input_base - 1)

    Next inx
    
 '   deltas_ = deltas_value

End Sub
Public Sub set_all_deltas(deltas_value() As Double, Optional input_base As Integer = 1, Optional notional As Double = 1)

    Dim inx As Integer
    Dim array_size As Integer
    
    
    array_size = get_array_size_double(deltas_value)
    
    ReDim deltas_(1 To array_size) As Double
    
    For inx = 1 To array_size
    
        deltas_(inx) = deltas_value(inx + input_base - 1) * notional

    Next inx
    
 '   deltas_ = deltas_value

End Sub


Public Sub set_all_term_vega(term_vega() As Double)


    term_vegas_ = term_vega

End Sub


Public Sub set_all_partial_vega(partial_vega() As Double)


    partial_vegas_ = partial_vega

End Sub


Public Sub set_all_term_skew(term_skew() As Double)


    term_skews_ = term_skew

End Sub


Public Sub set_all_partial_skew(partial_skew() As Double)


    partial_skews_ = partial_skew

End Sub


Public Sub set_all_term_conv(term_conv() As Double)


    term_convs_ = term_conv

End Sub


Public Sub set_all_partial_conv(partial_conv() As Double)


    partial_convs_ = partial_conv

End Sub



Public Sub set_all_term_date(term_date() As Date)


    term_dates_ = term_date

End Sub


Public Sub set_all_corr_sens(corr_sens_value() As Double)


    corr_sens_ = corr_sens_value

End Sub
'added 2018.09.07
Public Sub set_all_sticky_moneyness_deltas(deltas_value() As Double)


    sticky_moneyness_deltas_ = deltas_value

End Sub

Public Sub set_all_sticky_moneyness_gammas(gammas_value() As Double)


    sticky_moneyness_gammas_ = gammas_value

End Sub

Public Sub set_all_sticky_strike_deltas(deltas_value() As Double)


    sticky_strike_deltas_ = deltas_value

End Sub

Public Sub set_all_sticky_strike_gammas(gammas_value() As Double)


    sticky_strike_gammas_ = gammas_value

End Sub
Public Sub set_all_implied_tree_deltas(deltas_value() As Double)


    implied_tree_deltas_ = deltas_value

End Sub

Public Sub set_all_implied_tree_gammas(gammas_value() As Double)


    implied_tree_gammas_ = gammas_value

End Sub

Public Sub set_all_gammas(gammas_value() As Double, Optional input_base As Integer = 1, Optional notional As Double = 1)

    Dim inx As Integer
    Dim array_size As Integer
    
    array_size = get_array_size_double(gammas_value)
    
    ReDim gammas_(1 To array_size) As Double
    
    For inx = 1 To array_size
    
        gammas_(inx) = gammas_value(inx + input_base - 1) * notional

    Next inx
       
 '   gammas_ = gammas_value

End Sub
Public Sub set_all_rho_ul(rho_ul_value() As Double, Optional input_base As Integer = 1, Optional notional As Double = 1)

    Dim inx As Integer
    Dim array_size As Integer
    
    array_size = get_array_size_double(rho_ul_value)
    
    ReDim rho_ul_(1 To array_size) As Double
    
    For inx = 1 To array_size
    
        rho_ul_(inx) = rho_ul_value(inx + input_base - 1) * notional

    Next inx
       
 '   gammas_ = gammas_value

End Sub


Public Sub set_all_vegas(vegas_value() As Double)


    vegas_ = vegas_value

End Sub
Public Sub set_all_vannas(vanna_value() As Double)


    vannas_ = vanna_value

End Sub

Public Sub set_all_volgas(volga_value() As Double)


    volgas_ = volga_value

End Sub

Public Sub set_all_skews(skew_values() As Double)


    skew_s_ = skew_values

End Sub

Public Sub set_all_convs(conv_values() As Double)


    conv_s_ = conv_values

End Sub


Public Function get_term_greek_value(greek_code As String, ul_inx As Integer, term_inx As Integer) As Double

On Error GoTo ErrorHandler

    Select Case UCase(greek_code)
                
        Case "VEGA": get_term_greek_value = term_vegas_(ul_inx, term_inx)
    
        Case "SKEW": get_term_greek_value = term_skews_(ul_inx, term_inx)
        
        Case "CONV": get_term_greek_value = term_convs_(ul_inx, term_inx)
            
        
    End Select
    
    

    Exit Function

ErrorHandler:
    
    raise_err "clsGreeks.get_term_greek_value", "Invalid Greek Code"

End Function
Public Function get_greek_value(greek_code As String) As Double

On Error GoTo ErrorHandler

    Select Case UCase(greek_code)
        
        Case "VALUE": get_greek_value = value
        
        Case "DELTA": get_greek_value = delta
        
        Case "GAMMA": get_greek_value = gamma
        
        Case "VEGA": get_greek_value = vega
        
        Case "THETA": get_greek_value = theta
        
        Case "VOLGA": get_greek_value = volga
        
        Case "VANNA": get_greek_value = vanna
    
        Case "SKEW": get_greek_value = skew_s
        
        Case "CONV": get_greek_value = conv_s
        
        Case "IMPLIED_TREE_DELTA": get_greek_value = implied_tree_delta
        
        Case "STICKY_STRIKE_DELTA": get_greek_value = sticky_strike_delta
        
    End Select
    
    

    Exit Function

ErrorHandler:
    
    raise_err "clsGreeks.get_greek_value", "Invalid Greek Code"

End Function

Public Function get_greek_values(greek_code As String, Optional index As Integer = 1) As Double

On Error GoTo ErrorHandler

    Select Case UCase(greek_code)
        
        Case "VALUE": get_greek_values = value
        
        Case "DELTA": get_greek_values = deltas(index)
        
        Case "GAMMA": get_greek_values = gammas(index)
        
        Case "VEGA": get_greek_values = vegas(index)
        
        Case "THETA": get_greek_values = theta
        
        Case "VOLGA": get_greek_values = volga
        
        Case "VANNA": get_greek_values = vanna
    
        Case "SKEW": get_greek_values = skew_s_s(index)
        
        Case "CONV": get_greek_values = conv_s
        
        Case "IMPLIED_TREE_DELTA": get_greek_values = implied_tree_deltas_(index)
        
        Case "STICKY_STRIKE_DELTA": get_greek_values = sticky_strike_deltas_(index)
        
    End Select
    
    

    Exit Function

ErrorHandler:
    
    raise_err "clsGreeks.get_greek_value", "Invalid Greek Code"

End Function


Public Sub set_greek_value(greek_code As String, greek_value As Double)

On Error GoTo ErrorHandler

    Select Case UCase(greek_code)
        
        Case "VALUE": value = greek_value
        
        Case "DELTA":   delta = greek_value
        
        Case "GAMMA":   gamma = greek_value
        
        Case "VEGA":   vega = greek_value
        
        Case "THETA":   theta = greek_value
        
        Case "VOLGA":   volga = greek_value
        
        Case "VANNA":   vanna = greek_value
    
        Case "SKEW":   skew_s = greek_value
        
        Case "CONV":   conv_s = greek_value
        
        Case "UL_PRICE":   ul_price = greek_value '---2015-01-05
        
    End Select
    
    

    Exit Sub

ErrorHandler:
    
    raise_err "clsGreeks.set_greek_value", "Invalid Greek Code"


End Sub

Public Sub greek_value_adjustment(adj_factor As Double, percent_vega As Boolean)

    value = value * adj_factor
    delta = delta * adj_factor
    gamma = gamma * adj_factor
    theta = theta * adj_factor
    rho = rho * adj_factor
    
    volga = volga * adj_factor
    skew_s = skew_s * adj_factor
    conv_s = conv_s * adj_factor
    implied_tree_delta = implied_tree_delta * adj_factor
    sticky_strike_delta = sticky_strike_delta * adj_factor
    implied_tree_gamma = implied_tree_gamma * adj_factor
    sticky_strike_gamma = sticky_strike_gamma * adj_factor
    
    If percent_vega Then
        vega = vega / 100 * adj_factor
        vanna = vanna / 100 * adj_factor
    Else
        vega = vega * adj_factor
        vanna = vanna * adj_factor
    End If
    
    

End Sub

Public Sub set_all_greek_value(value_in As Double, delta_in As Double, gamma_in As Double, vega_in As Double, theta_in As Double _
                            , rho_in As Double, vanna_in As Double, volga_in As Double, skew_in As Double, conv_in As Double)


On Error GoTo ErrorHandler

    value = value_in
    delta = delta_in
    gamma = gamma_in
    theta = theta_in
    rho = rho_in
    vanna = vanna_in
    volga = volga_in
    skew_s = skew_in
    
    vega = vega_in
    
    conv_s = conv_in
    
    

    Exit Sub

ErrorHandler:
    
    raise_err "clsGreeks.set_all_greek_value"



End Sub