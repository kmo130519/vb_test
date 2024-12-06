Option Explicit

Public v_long As Double 'long term volatility
Public v_initial As Double 'initial volatility
Public lamda As Double ' mean reversion
Public eta As Double  ' vol of vol
Public rho As Double  ' correlation

Private v_long_origin As Double 'long term volatility
Private v_initial_origin As Double 'initial volatility
Private lamda_origin As Double ' mean reversion
Private eta_origin As Double  ' vol of vol
Private rho_origin As Double  ' correlation

Private bump_parameters(1 To 4) As clsBumpParam ' 1. parallel shift up. 2. shift down. 3. skew up. 4. skew down
'Private bump_parameters_down As clsBumpParam 'parallel shift down
'Private bump_parameters_skew_up As clsBumpParam 'skew up
'Private bump_parameters_skew_down As clsBumpParam 'skew down



Public Function copy_obj() As clsHestonParameter

    Dim rtn_obj As clsHestonParameter
    Dim inx As Integer
    
    Set rtn_obj = New clsHestonParameter
    
    rtn_obj.initialize v_long, v_initial, lamda, eta, rho
    For inx = 1 To 4
        
        rtn_obj.set_bump_param inx, bump_parameters(inx).v_long, bump_parameters(inx).v_initial, bump_parameters(inx).lamda, bump_parameters(inx).eta, bump_parameters(inx).rho
    
    Next inx
    
    
    Set copy_obj = rtn_obj

End Function

'--------------------------------
' Sub: set_bum_param
' Desc: Fill the bump_parameters.
'       Bump parameters will be used to calculate senstivities related to volatility surface.
'--------------------------------
Public Sub set_bump_param(inx As Integer, v_long_in As Double, v_init_in As Double, lamda_in As Double, eta_in As Double, rho_in As Double)


On Error GoTo ErrorHandler

    Set bump_parameters(inx) = New clsBumpParam
    
    bump_parameters(inx).v_long = v_long_in
    bump_parameters(inx).v_initial = v_init_in
    bump_parameters(inx).lamda = lamda_in
    bump_parameters(inx).eta = eta_in
    bump_parameters(inx).rho = rho_in


    Exit Sub
    
ErrorHandler:

    raise_err "clsHestonParameter.set_bump_param"

End Sub

Public Sub initialize(v_long_in As Double, v_init_in As Double, lamda_in As Double, eta_in As Double, rho_in As Double)


On Error GoTo ErrorHandler

    v_long = v_long_in
    v_initial = v_init_in
    lamda = lamda_in
    eta = eta_in
    rho = rho_in
    
    v_long_origin = v_long
    v_initial_origin = v_initial
    lamda_origin = lamda
    eta_origin = eta
    rho_origin = rho
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsHestonParameter.initialize :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
                
    
    
End Sub



Public Sub parallel_shift(vol_difference As Double)


On Error GoTo ErrorHandler


    v_long = (Sqr(v_long) + vol_difference) ^ 2
    v_initial = (Sqr(v_initial) + vol_difference) ^ 2
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsHestonParameter.parallel_shift :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
                
    

End Sub

'--------------------------------
' Sub: rewind
' Desc: Set the parameters to be the original parameters.
'--------------------------------
Public Sub rewind()

    v_long = v_long_origin
    v_initial = v_initial_origin
    lamda = lamda_origin
    eta = eta_origin
    rho = rho_origin

End Sub


Public Sub bump_param(inx As Integer)

On Error GoTo ErrorHandler


    v_long = bump_parameters(inx).v_long
    v_initial = bump_parameters(inx).v_initial
    lamda = bump_parameters(inx).lamda
    eta = bump_parameters(inx).eta
    rho = bump_parameters(inx).rho
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsHestonParameter.bump_param :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
                


End Sub