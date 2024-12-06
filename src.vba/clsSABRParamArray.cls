Option Explicit

Public no_of_dates As Integer

Public eval_date_ As Date

Private sabr_param_() As clsSabrParameter

Public Function bump_fwd(ByVal bump_size As Double) As Double()
    
    Dim inx As Integer
    Dim rtn_array() As Double
    
On Error GoTo ErrorHandler

    For inx = 1 To no_of_dates
    
        sabr_param_(inx).forward = sabr_param_(inx).forward + bump_size
        push_back_double rtn_array, sabr_param_(inx).forward
    
    Next inx
    
    


    bump_fwd = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsSABRParamArray.bump_fwd", Err.description
        


End Function


Public Sub set_sabr_param(no_of_array As Integer, sabr_param() As clsSabrParameter)

    Dim inx As Integer
    
    ReDim sabr_param_(1 To no_of_array)
    
    For inx = 1 To no_of_array
    
        Set sabr_param_(inx) = sabr_param(inx).copy_obj
    
    Next inx
    


End Sub

Public Function copy_obj(Optional min_date As Date = 0) As clsSABRParamArray

    Dim rtn_obj As clsSABRParamArray
    Dim inx As Integer
    Dim tmp_sabr_parameter() As clsSabrParameter
    Dim dummy_obj As clsSabrParameter
    
    Set rtn_obj = New clsSABRParamArray
    
    rtn_obj.no_of_dates = 0 'no_of_dates
    rtn_obj.eval_date_ = eval_date_
    
    For inx = 1 To no_of_dates
        
        If sabr_param_(inx).Maturity > min_date Then
            Set dummy_obj = sabr_param_(inx).copy_obj()
            push_back_clsSabrParameter tmp_sabr_parameter, dummy_obj
            rtn_obj.no_of_dates = rtn_obj.no_of_dates + 1
        End If
    
    Next inx
    

    rtn_obj.set_sabr_param rtn_obj.no_of_dates, tmp_sabr_parameter

    Set copy_obj = rtn_obj

End Function

Public Sub initialize(no_of_dates_in As Integer _
                    , Maturity() As Date _
                    , fwd() As Double _
                    , alpha() As Double _
                    , beta As Double _
                    , nu() As Double _
                    , rho() As Double _
                    , atm() As Double _
                    , eval_date As Date _
                    )

    Dim inx As Integer
    
On Error GoTo ErrorHandler
    eval_date_ = eval_date

    no_of_dates = no_of_dates_in
    
    If no_of_dates_in > 0 Then
        ReDim sabr_param_(1 To no_of_dates) As clsSabrParameter
        
        For inx = 1 To no_of_dates
        
            Set sabr_param_(inx) = New clsSabrParameter
            sabr_param_(inx).initialize fwd(inx), alpha(inx), beta, nu(inx), rho(inx), (Maturity(inx) - eval_date_) / 365, atm(inx), Maturity(inx)
        
        Next inx
    End If
    
    Exit Sub
    
ErrorHandler:

    raise_err "clsSABRParamArray.initialize", Err.description

End Sub

Public Function sabr_param(inx As Integer) As clsSabrParameter

          
    Set sabr_param = sabr_param_(inx)

End Function



Public Function maturity_dates_array() As Date()
    
    Dim rtn_array() As Date
    Dim inx As Integer
    
    
    For inx = 1 To no_of_dates
        
        push_back_date rtn_array, sabr_param_(inx).Maturity
    
    Next inx
    
    maturity_dates_array = rtn_array
    

End Function
Public Function atm_vol_array() As Double()
    
    Dim rtn_array() As Double
    Dim inx As Integer
    
    
    For inx = 1 To no_of_dates
        
        push_back_double rtn_array, sabr_param_(inx).vol_atm
    
    Next inx
    
    atm_vol_array = rtn_array
    

End Function



Public Function nu_array() As Double()
    
    Dim rtn_array() As Double
    Dim inx As Integer
    
    
    For inx = 1 To no_of_dates
        
        push_back_double rtn_array, sabr_param_(inx).nu
    
    Next inx
    
    nu_array = rtn_array
    

End Function

Public Function rho_array() As Double()
    
    Dim rtn_array() As Double
    Dim inx As Integer
    
    
    For inx = 1 To no_of_dates
        
        push_back_double rtn_array, sabr_param_(inx).rho
    
    Next inx
    
    rho_array = rtn_array
    

End Function