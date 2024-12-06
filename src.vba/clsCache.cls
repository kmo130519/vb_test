Option Explicit

'############################################################
' Class: clsCache
' This class contains information used for the cliquet simulation.
' The information is calculated beforehand to save cost
'############################################################

Private A1_() As Double
Private A2_() As Double
Private A3_() As Double
Private K_0_() As Double
Private K_1_() As Double
Private K_2_() As Double
Private K_3_() As Double
Private K_4_() As Double
Private A_() As Double

Private fwd_rate_() As Double


Public Sub initialize(heston_param As clsHestonParameter, time_step As clsTimeStep, rate_curve As clsRateCurve, Optional gamma_1 As Double = 0.5)

    Dim inx As Integer
    
    Dim gamma_2 As Double
    
On Error GoTo ErrorHandler
    
    ReDim A1_(0 To time_step.no_of_steps) As Double
    ReDim A2_(0 To time_step.no_of_steps) As Double
    ReDim A3_(0 To time_step.no_of_steps) As Double
    ReDim A_(0 To time_step.no_of_steps) As Double
        
    ReDim K_0_(0 To time_step.no_of_steps) As Double
    ReDim K_1_(0 To time_step.no_of_steps) As Double
    ReDim K_2_(0 To time_step.no_of_steps) As Double
    ReDim K_3_(0 To time_step.no_of_steps) As Double
    ReDim K_4_(0 To time_step.no_of_steps) As Double
    ReDim K_Vega_(0 To time_step.no_of_steps) As Double
    ReDim fwd_rate_(0 To time_step.no_of_steps - 1) As Double
    
    Dim rho As Double
    Dim lamda As Double
    Dim eta As Double
    Dim v_l As Double
    Dim dt As Double
    
    rho = heston_param.rho
    lamda = heston_param.lamda
    v_l = heston_param.v_long
    eta = heston_param.eta
    
    gamma_2 = 1 - gamma_1
    
    If time_step.no_of_steps < 0 Then
    
        Err.Raise vbObjectError + 1001, "clsCache.initialize", "The number of time steps is less than zero"
        
    End If
           
    'rate_curve.make_obj
    
    For inx = 0 To time_step.no_of_steps
    
        dt = time_step.dt(inx)
    
        A1_(inx) = Exp(-lamda * dt)
        A2_(inx) = eta ^ 2 * A1_(inx) / lamda * (1 - A1_(inx))
        A3_(inx) = v_l * eta ^ 2 / 2 / lamda * (1 - A1_(inx)) ^ 2
        
        K_0_(inx) = -rho * lamda * v_l / eta * dt
        K_1_(inx) = gamma_1 * dt * (lamda * rho / eta - 0.5) - rho / eta
        K_2_(inx) = gamma_2 * dt * (lamda * rho / eta - 0.5) + rho / eta
        K_3_(inx) = gamma_1 * dt * (1 - rho ^ 2)
        K_4_(inx) = gamma_2 * dt * (1 - rho ^ 2)
            
        
        A_(inx) = K_2_(inx) + 0.5 * K_4_(inx)
        
        If inx <= time_step.no_of_steps - 1 Then
        
            fwd_rate_(inx) = rate_curve.get_fwd_rate(time_step.get_date(inx), time_step.get_date(inx + 1))
            
        End If
        
    Next inx
        
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsCache.initialize :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Sub


End Sub

Public Function A1(inx As Integer) As Double

On Error GoTo ErrorHandler:

    A1 = A1_(inx)
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCache.A1 :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function

    
End Function

Public Function A2(inx As Integer) As Double

On Error GoTo ErrorHandler:

    A2 = A2_(inx)
    
    Exit Function
        
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCache.A2 :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function

    
End Function

Public Function A3(inx As Integer) As Double

On Error GoTo ErrorHandler:

    A3 = A3_(inx)
    
    Exit Function
        
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCache.A3 :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function

    
End Function


Public Function a(inx As Integer) As Double

On Error GoTo ErrorHandler:

    a = A_(inx)
    
    Exit Function
        
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCache.A :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function

    
End Function

Public Function K_0(inx As Integer) As Double

On Error GoTo ErrorHandler:

    K_0 = K_0_(inx)
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCache.K_0 :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function

End Function

Public Function K_1(inx As Integer) As Double

On Error GoTo ErrorHandler:

    K_1 = K_1_(inx)
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCache.K_1 :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function


End Function

Public Function K_2(inx As Integer) As Double

On Error GoTo ErrorHandler:

    K_2 = K_2_(inx)
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCache.K_2 :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function


End Function

Public Function K_3(inx As Integer) As Double

On Error GoTo ErrorHandler:

    K_3 = K_3_(inx)
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCache.K_3 :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function


End Function

Public Function K_4(inx As Integer) As Double

On Error GoTo ErrorHandler:

    K_4 = K_4_(inx)
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCache.K_4 :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function

End Function

Public Function forward_rate(inx As Integer) As Double

On Error GoTo ErrorHandler:

    forward_rate = fwd_rate_(inx)
    
    Exit Function
    
ErrorHandler:
    
    Err.Raise vbObjectError + 1003, "clsCache.forward_rate :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function
        
End Function