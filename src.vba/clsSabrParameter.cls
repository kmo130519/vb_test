Option Explicit


'Public eval_date As Date
'Public ul_code As String

Public forward As Double
Public alpha As Double
Public beta As Double
Public nu As Double
Public rho As Double

Public vol_atm As Double

Public Maturity As Date

Public tau As Double

Public Sub initialize(forward_in As Double _
                    , alpha_in As Double _
                    , beta_in As Double _
                    , nu_in As Double _
                    , rho_in As Double _
                    , tau_in As Double _
                    , atm_in As Double _
                    , maturity_in As Date)

    forward = forward_in
    alpha = alpha_in
    beta = beta_in
    nu = nu_in
    rho = rho_in
    tau = tau_in
    vol_atm = atm_in
    
    Maturity = maturity_in

End Sub


Public Function copy_obj() As clsSabrParameter
    
    Dim rtn_obj As clsSabrParameter
    
    
    Set rtn_obj = New clsSabrParameter
    
    rtn_obj.forward = forward
    rtn_obj.alpha = alpha
    rtn_obj.beta = beta
    rtn_obj.nu = nu
    rtn_obj.rho = rho
    
    rtn_obj.vol_atm = vol_atm
    
    rtn_obj.tau = tau
    rtn_obj.Maturity = Maturity
    
    Set copy_obj = rtn_obj

End Function