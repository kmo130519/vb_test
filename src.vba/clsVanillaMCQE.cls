'----------------------------------------------------------------------------------------------
' CLASS: clsCliquetPathQE
' Author: Lee, Seok.
' Amended on: 2011.08.01
' DESC:
' This class performs monte carlo simulation of cliquet option using heston model.
' Quadratic-Exponential Scheme is implemented in this class.
' Refer to 'Efficient simulation of the Heston Stochastic Volatility Model, Leif Andersen, BOA'
' Initialize procedure initializes all the members of this class.
' ----------------------------------------------------------------------------------------------

Option Explicit


Public price As Double 'Option price.

Private deal_ticket_ As clsCliquetDealTicket 'Deal information.
Private s_ As Double 'Initial stock price
Public rate_curve_ As clsRateCurve 'IR Curve
Private div_schedule_ As clsDividendSchedule 'Discrete dividend schedule
Private heston_param_ As clsHestonParameter 'Parameters related to heston model

Private call_put_ As String

Private strike_ As Double

Private time_step_ As clsTimeStep  'time steps. This member has information on dt and fixing schedule.
Private cache_ As clsCache 'Some of the calcuations are performed in advance to save calculation cost.
Private random_numbers_() As Double 'The array of random numbers to be used during the simulation.
Private path_seq_ As Long 'This is the index of paths.
Private time_seq_ As Integer 'This is the index of time step.

'Private previous_performance As Double

Private Const psi_c_ As Double = 1.5  'Defined in the paper. Arbitrary.
Private Const gamma_1_ As Double = 0.5 'Defined in the paper. Arbitrary.
Private Const epsilon_ As Double = 0.001 'Used to smooth delta to calculate pathwise gamma.
Private Const tiny_ As Double = 0.000001

Private distribution_helper_ As Distributions 'Helper class. Written by Jeon, YK.
Private func As New Functions 'Helper class. Written by Jeon, YK.

Private random_generator_ As clsRandomGenerator 'Random number generator. Using Quantlib.

Private greeks_ As clsGreeks

' S(k-1). This is used to calculated pathwise delta
Private s_tk_1_ As Double
'Private s_tk_ As Double

Private rnd_mode_ As String
Private bump_greek_ As Boolean

Private Const bump_s_ As Double = 1#


'========================
Private temp_seq As Integer
'========================




'--------------------------------------------------------------------------------
' Procedure: initialize
' INPUT: deal_ticket (Deal information), S, rate_curve, div_schedule, heston_param(market information)
'        rnd_mode ( preference for the random number generation )
' Output: random_numbers, time_step_, cache_, implicitly.
' DESC: Creates objects and initializes the members used in this class.
'--------------------------------------------------------------------------------
Public Sub initialize(deal_ticket As clsCliquetDealTicket, call_put As String, strike As Double _
                      , s As Double, rate_curve As clsRateCurve, div_schedule As clsDividendSchedule, heston_param As clsHestonParameter _
                      , Optional rnd_mode As String = "SOBOL")
                      
    temp_seq = 1
    
'    Dim temp_fixing_schedule As clsCliquetFixing

On Error GoTo ErrorHandler

    Set deal_ticket_ = deal_ticket  'Temporary
    s_ = s
    Set rate_curve_ = rate_curve
    Set div_schedule_ = div_schedule
    Set heston_param_ = heston_param
    rnd_mode_ = rnd_mode
    
    price = 0
    
    Set time_step_ = New clsTimeStep
    Set cache_ = New clsCache
    Set distribution_helper_ = New Distributions
    Set func = New Functions
    Set greeks_ = New clsGreeks
    Set random_generator_ = New clsRandomGenerator
    
    call_put_ = call_put
    strike_ = strike
    
    reset_data True
    
    initialize_greek
        
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsCliquetMCQE.initialization :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
        
End Sub

'---------------------------------------------------
' initialize_greek
' Find the first fixing value before current day
' to calcuate delta using pahtwise differentiation
'----------------------------------------------------
Private Sub initialize_greek()

    Dim inx As Integer
    
On Error GoTo ErrorHandler
    
    inx = deal_ticket_.get_max_fixing_index()
    
    
    'If the initial stock price was set...
    If deal_ticket_.current_date >= deal_ticket_.get_fixing_date(0).fixing_date And deal_ticket_.get_fixing_date(0).fixing_value > 0 Then
        
        ' Find the last fixing value...
        Do While deal_ticket_.get_fixing_date(inx).fixing_date > deal_ticket_.current_date
        
            inx = inx - 1
            
        Loop
        
        ' If curren date is a fixing date but not fixed yet...
        If deal_ticket_.current_date = deal_ticket_.get_fixing_date(inx).fixing_date And deal_ticket_.get_fixing_date(inx).fixing_value <= 0 Then
            
            inx = inx - 1
            
        End If
        
        s_tk_1_ = deal_ticket_.get_fixing_date(inx).fixing_value
        
    Else
        
        s_tk_1_ = -1
        
    End If
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsCliquetMCQE.initialize_greek :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub

End Sub



'------------------------------------------------------------------
' reset
' Reset data to run the simulation again with different parameters.
' Set date_changed = True to regenerate time steps.
'-------------------------------------------------------------------
Public Sub reset_data(Optional date_changed As Boolean = False)


On Error GoTo ErrorHandler
    
    ' If the fixing value is not specified for the past fixing date, set the current market price to be the fixing value
    deal_ticket_.initialize s_
    
    If date_changed Then
    
        div_schedule_.initialize_div_inx deal_ticket_.current_date
        
        time_step_.initialize deal_ticket_, div_schedule_
        
        '---------------------------------------------------------------------
        ' If it's maturity date, do not initialize object other than deal_ticket
        '---------------------------------------------------------------------
        If deal_ticket_.current_date < deal_ticket_.maturity_date Then
        
    
'            cache_.initialize heston_param_, time_step_, rate_curve_, gamma_1_
            
            '******************************************************
            ' Subject to improve.
            ' Generate functions are not included in this class.
            ' Especially, 'SOBOL' sequence generator is included 'xll' provided by 3rd party.
            '******************************************************
            If rnd_mode_ = "BASIC" Then
                    
                random_numbers_ = random_generator_.uniform_numbers_2d(deal_ticket_.no_of_trials, time_step_.no_of_steps)
                
            ElseIf rnd_mode_ = "SOBOL" Then
            
                random_numbers_ = random_generator_.sobol_numbers_2d(deal_ticket_.no_of_trials, time_step_.no_of_steps)
                
            Else
            
                Exit Sub
                
            End If
            
            
        End If
        
    End If
    
    '---------------------------------------------------------------------
    ' If it's maturity date, do not initialize object other than deal_ticket
    '---------------------------------------------------------------------
    If deal_ticket_.current_date < deal_ticket_.maturity_date Then
    
        cache_.initialize heston_param_, time_step_, rate_curve_, gamma_1_
        
    End If
    
    path_seq_ = 1
    time_seq_ = 0

    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsCliquetMCQE.reset_data :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub

End Sub

'-----------------------------------------------------------------------------------------------
' RUN
' run the simulation and get the price and greeks.
' Should be initialized for this procedure run properly.
'------------------------------------------------------------------------------------------------

Public Sub run()

    Dim prices() As Double
    Dim deltas() As Double
    Dim gammas() As Double
    Dim vegas() As Double
    Dim s_origin As Double
    
    Dim price_temp As Double
'    Dim price_down As Double
    
On Error GoTo ErrorHandler

    s_origin = s_
        
    do_simulation prices, deltas, gammas, vegas
    
    greeks_.value = func.ReturnAVG(prices) * rate_curve_.get_discount_factor(func.max_date(deal_ticket_.value_date, deal_ticket_.current_date), deal_ticket_.maturity_date)
    greeks_.delta = func.ReturnAVG(deltas) * rate_curve_.get_discount_factor(func.max_date(deal_ticket_.value_date, deal_ticket_.current_date), deal_ticket_.maturity_date)
    greeks_.gamma = func.ReturnAVG(gammas) * rate_curve_.get_discount_factor(func.max_date(deal_ticket_.value_date, deal_ticket_.current_date), deal_ticket_.maturity_date)
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCliquetMCQE.run :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Sub
         
End Sub

Public Sub do_simulation(prices() As Double, deltas() As Double, gammas() As Double, vegas() As Double)

    Dim inx As Long
    
On Error GoTo ErrorHandler
    
    ReDim prices(1 To deal_ticket_.no_of_trials) As Double
    ReDim deltas(1 To deal_ticket_.no_of_trials) As Double
    ReDim gammas(1 To deal_ticket_.no_of_trials) As Double
    ReDim vegas(1 To deal_ticket_.no_of_trials) As Double
    
    For inx = 1 To deal_ticket_.no_of_trials
    
        run_a_trial prices(inx), deltas(inx), gammas(inx), vegas(inx)
        
    Next inx
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCliquetMCQE.do_simulation :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Sub
    

End Sub


' Getter
Public Function greeks() As clsGreeks

    Set greeks = greeks_

End Function

'------------------------------------------------------------------------
' Procedure: set_random_numbers
' DESC: Explicitly set the random number arrays. This sub routine can be used
' when the random numbers should be set manually.
'------------------------------------------------------------------------
Public Sub set_random_numbers(random_numbers() As Double)

    random_numbers_ = random_numbers

End Sub

' Setter
Public Sub set_s(spot As Double)

    s_ = spot

End Sub


' Setter
Public Sub set_date(c_date As Date)

On Error GoTo ErrorHandler

    deal_ticket_.current_date = c_date
    
    
Exit Sub

ErrorHandler:

    Err.Raise vbObjectError + 1001, "clsCliquetMCQE.set_date :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Sub
    
    
End Sub


'--------------------------------------------------------
' Sub: run_a_trial
' INPUT: void
' OUTPUT: Price calucated from a path, Pathwise delta.
'--------------------------------------------------------

Public Sub run_a_trial(ByRef prices As Double, ByRef deltas As Double, ByRef gammas As Double, ByRef vegas As Double)

    Dim return_value As Double
    
    Dim inx As Long
    Dim mirror_s As Double
    Dim mirror_v As Double
    Dim performance As Double
    Dim current_date As Date
    Dim previous_fixing As Double
    Dim s_tk_1 As Double
    
    Dim ds_dvol As Double
    Dim dp_dvol As Double
    
    Dim s_tk_checked As Boolean
    Dim pathwise_delta As Double
    Dim pathwise_gamma As Double
    
Dim temp_sheet As Worksheet

        
On Error GoTo ErrorHandler

    Set temp_sheet = Sheets("temp")

    s_tk_checked = False
    s_tk_1 = s_tk_1_
    
    pathwise_delta = 0
    pathwise_gamma = 0
    '------------------------------------------------------
    ' Initialize simulation parameters
    '------------------------------------------------------
    mirror_s = s_
    mirror_v = heston_param_.v_initial
    
    '--------------------------------------------------------
    ' Set the fixing index to be the first one.
    '--------------------------------------------------------
    deal_ticket_.rewind_fixing_inx
    div_schedule_.rewind_fixing_inx
    
    time_seq_ = 0
    
    
    'previous_fixing = s_
    
    
    Do While time_seq_ < time_step_.no_of_steps
    

    
        '--------------------------------------------------------
        ' Check if it is fixing date and calcuated performance
        '--------------------------------------------------------
        check_dividend mirror_s, time_seq_
'        check_fixing performance, previous_fixing, mirror_s, time_seq_, s_tk_checked, s_tk_1, ds_dvol, pathwise_delta, pathwise_gamma, dp_dvol
                
        '--------------------------------------------------------
        ' Move to next time step
        '--------------------------------------------------------
        progress_a_step mirror_s, mirror_v, ds_dvol
        
        
    Loop
    
    
    'calc_performance
    
    '--------------------------------------------------------------------
    ' check if the last time step is fixing date and update performance
    '--------------------------------------------------------------------
   ' check_fixing performance, previous_fixing, mirror_s, time_seq_, s_tk_checked, s_tk_1, ds_dvol, pathwise_delta, pathwise_gamma, dp_dvol
    '--------------------------------------------------------------------
    
    
  '  performance = performance + deal_ticket_.previous_performance
        
    '--------------------------------------------------------------------
    ' Apply global condition like 'Global Cap', 'Global Floor'
    '--------------------------------------------------------------------
 '   apply_global_condition performance, pathwise_delta, pathwise_gamma, dp_dvol
 
    If UCase(call_put_) = "CALL" Then
        performance = func.max((mirror_s - strike_), 0)
    Else
        performance = func.max((strike_ - mirror_s), 0)
    End If
        
        
    temp_sheet.Cells(path_seq_ + 3, 2).value = performance
    
    path_seq_ = path_seq_ + 1
    
    prices = performance
    deltas = pathwise_delta
    gammas = pathwise_gamma
    vegas = dp_dvol
    
    
    
    Exit Sub
    
ErrorHandler:


    Err.Raise vbObjectError + 1003, "clsCliquetMCQE.run_a_trial :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Sub
    
    

End Sub

'-----------------------------------------------
' Sub: check_fixing
' Desc: Check if it is fixing date and accumulate performance
'       If it is the first fixing date, performance should not be calculated
'       If it it the first fixing since s_t(k-1) is set,calculate pathwise delta
'-----------------------------------------------

Private Sub check_fixing(ByRef performance As Double, ByRef previous_fixing As Double, mirror_s As Double, inx As Integer _
                       , s_tk_checked As Boolean, ByRef s_tk_1 As Double, ByRef ds_dvol As Double _
                       , ByRef pathwise_delta As Double, ByRef pathwise_gamma As Double, ByRef dp_dvol As Double)

    Dim cur_fixing_schedule As clsCliquetFixing
    
    Dim temp As Double
    
On Error GoTo ErrorHandler
    
    Set cur_fixing_schedule = deal_ticket_.get_fixing_date()
    
    If time_step_.is_fixing_date(inx) Then
    
        If Not time_step_.is_first_fixing_date(inx) Then
        
            ' Min( Max ( S(t) / S(t-1) -1 , local floor ), local cap )
            performance = performance + func.min(func.max(mirror_s / previous_fixing - 1, cur_fixing_schedule.local_floor), cur_fixing_schedule.local_cap)
        
            dp_dvol = dp_dvol + ds_dvol * (func.indicator(mirror_s / s_tk_1 > cur_fixing_schedule.local_floor + 1) _
                                    - func.indicator(mirror_s / s_tk_1 > cur_fixing_schedule.local_cap + 1) _
                                    ) / s_tk_1
                                    
            ds_dvol = 0
            
            ' to calculated pathwise delta. In the case of cliquet, a series of forward starting options, only the first performance affects to delta and gamma.
            If Not s_tk_checked Then
            
'                s_tk = mirror_s
                s_tk_checked = True
                's_tk_ = mirror_s
                
                temp = mirror_s / s_tk_1 / s_ * func.indicator(s_tk_1 > 0)
                
                pathwise_delta = (func.indicator(mirror_s / s_tk_1 > cur_fixing_schedule.local_floor + 1) _
                                    - func.indicator(mirror_s / s_tk_1 > cur_fixing_schedule.local_cap + 1) _
                                    ) * temp
                
                pathwise_gamma = temp ^ 2 * (distribution_helper_.NormD((mirror_s / s_tk_1 - cur_fixing_schedule.local_floor - 1) / epsilon_) _
                                           - distribution_helper_.NormD((mirror_s / s_tk_1 - cur_fixing_schedule.local_cap - 1) / epsilon_)) / epsilon_
                                           
            
            End If
            
            s_tk_1 = mirror_s
        
        End If
        
        previous_fixing = mirror_s
        
        'indicate next fixing index
        deal_ticket_.move_next
        
    End If
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCliquetMCQE.check_fixing :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Sub
        

End Sub

Private Sub check_dividend(ByRef mirror_s As Double, inx As Integer)

    
On Error GoTo ErrorHandler

    If time_step_.is_dividend_date(inx) Then
    
        mirror_s = mirror_s - div_schedule_.get_dividend().dividend * mirror_s / s_
        div_schedule_.move_next
        
    End If

    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCliquetMCQE.check_dividend :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Sub
        
End Sub

Private Sub apply_global_condition(performance As Double, pathwise_delta As Double, pathwise_gamma As Double, pathwise_vega As Double)

    Dim d2v_dp2 As Double
    Dim dv_dp As Double
    'Dim d2p_ds2 As Double

    ' Min( Max( Accumulated Performance, Global Floor ), Global Cap )
    
'    dv_dp = (func.indicator(performance > deal_ticket_.global_floor) - func.indicator(performance > deal_ticket_.global_cap))
    
    dv_dp = distribution_helper_.NormCDF((performance - deal_ticket_.global_floor) / epsilon_) - distribution_helper_.NormCDF((performance - deal_ticket_.global_cap) / epsilon_)
    
    d2v_dp2 = distribution_helper_.NormD((performance - deal_ticket_.global_floor) / epsilon_) / epsilon_ - distribution_helper_.NormD((performance - deal_ticket_.global_cap) / epsilon_) / epsilon_
    
    performance = func.min(func.max(performance, deal_ticket_.global_floor), deal_ticket_.global_cap)
    
    pathwise_gamma = pathwise_delta ^ 2 * d2v_dp2 + dv_dp * pathwise_gamma
    
    ' "func.indicator" is the indicator function which returns 1 if the condition is true otherwise 0
    pathwise_delta = dv_dp * pathwise_delta
    
    pathwise_vega = dv_dp * pathwise_vega / 100 ' / 100 * 2 * heston_param_.v_initial '% vega

End Sub



'-------------------------------------------------------
' Sub: progress_a_step
' Desc: Move to next step. QE Scheme for variance.
'-------------------------------------------------------
Private Sub progress_a_step(ByRef s As Double, ByRef v As Double, ByRef ds_dvol As Double)

    Dim v_dt As Double
    Dim s_dt As Double
    Dim large_psi As Boolean
    Dim k_star As Double
    
On Error GoTo ErrorHandler
    
    
    
    v_dt = next_v(v, k_star)
    s_dt = next_s(ds_dvol, s, v, v_dt, k_star)
    
    '---------------------------------------
    ' subject to improve. 2011.07.15
    '----------------------------------------
    If v < 0 Then
        v = -v
    End If
    
    v = v_dt
    s = s_dt
    
    time_seq_ = time_seq_ + 1
    
    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCliquetMCQE.progress_a_step :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Sub
        

End Sub



'---------------------------------------
' Function: next_v
' input : v
' output : next_v, k_star
' Desc: QE Scheme for heston model. Refer to Andersen
'-------------------------------------------
Private Function next_v(v As Double, ByRef k_star As Double) As Double

    Dim m As Double
    Dim s_2 As Double
    Dim Psi As Double
    Dim a As Double
    Dim b_2 As Double

    Dim beta As Double
    Dim p As Double

    Dim A_ As Double

On Error GoTo ErrorHandler
    
    m = heston_param_.v_long + (v - heston_param_.v_long) * cache_.A1(time_seq_)
    s_2 = v * cache_.A2(time_seq_) + cache_.A3(time_seq_)
    
    Psi = s_2 / m ^ 2
    
    A_ = cache_.K_2(time_seq_) + 0.5 * cache_.K_4(time_seq_)
    
    If Psi <= psi_c_ Then
    
        b_2 = 2 / Psi - 1 + Sqr(2 / Psi) * Sqr(2 / Psi - 1)
        a = m / (1 + b_2)
        
        next_v = a * (Sqr(b_2) + distribution_helper_.NormInv(random_numbers_(path_seq_, time_seq_ + 1, 2))) ^ 2

        ' Martingale correction
        k_star = -A_ * b_2 * a / (1 - 2 * A_ * a) + 0.5 * Log(1 - 2 * A_ * a) - (cache_.K_1(time_seq_) + 0.5 * cache_.K_3(time_seq_)) * v
'        k_star = cache_.K_0(time_seq_)

        
    Else
        
        beta = 2 / m / (Psi + 1)
        p = (Psi - 1) / (Psi + 1)
        
        If random_numbers_(path_seq_, time_seq_ + 1, 2) <= p Then
        
            'next_v = 0
            next_v = tiny_
            
        Else
            next_v = Log((1 - p) / (1 - random_numbers_(path_seq_, time_seq_ + 1, 2))) / beta
        
        End If
        
        ' Martingale correction
       k_star = -Log(p + beta * (1 - p) / (beta - A_)) - (cache_.K_1(time_seq_) + 0.5 * cache_.K_3(time_seq_)) * v

'       k_star = cache_.K_0(time_seq_)


        
    End If
    
    
    Exit Function
    
        
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCliquetMCQE.next_v :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function

    
End Function


'---------------------------------------
' Function: next_s
' input : s, v, nextv, k_start
' output : next_s
' Desc: Refer to Andersen
'-------------------------------------------

Private Function next_s(ByRef ds_dvol As Double, s As Double, v As Double, next_v As Double, k_star As Double) As Double

On Error GoTo ErrorHandler
'Dim temp_sheet As Worksheet
'Set temp_sheet = Application.Sheets("Vega Test")

    Dim K_all As Double
    
    K_all = Exp(cache_.forward_rate(time_seq_) * time_step_.dt(time_seq_) + k_star + cache_.K_1(time_seq_) * v + cache_.K_2(time_seq_) * next_v _
          + Sqr(cache_.K_3(time_seq_) * v + cache_.K_4(time_seq_) * next_v) * distribution_helper_.NormInv(random_numbers_(path_seq_, time_seq_ + 1, 1)))
          
    
'    ds_dvol = K_all * ds_dvol _
          + s * K_all * (cache_.K_1(time_seq_) + cache_.K_2(time_seq_) _
                       + 0.5 / Sqr(cache_.K_3(time_seq_) * v + cache_.K_4(time_seq_) * next_v) * (cache_.K_3(time_seq_) + cache_.K_4(time_seq_)) * distribution_helper_.NormInv(random_numbers_(path_seq_, time_seq_ + 1, 1)) _
                        )
                        
'    ds_dvol = (K_all * ds_dvol _
          + s * K_all * (cache_.K_1(time_seq_) + cache_.K_2(time_seq_) _
                       + 0.5 / Sqr(cache_.K_3(time_seq_) * v + cache_.K_4(time_seq_) * next_v) * (cache_.K_3(time_seq_) + cache_.K_4(time_seq_)) * distribution_helper_.NormInv(random_numbers_(path_seq_, time_seq_ + 1, 1)) _
                        )) * 2 * Sqr(v)
                        
                        
    ds_dvol = K_all * ds_dvol _
          + s * K_all * (2 * cache_.K_1(time_seq_) * Sqr(v) + 2 * cache_.K_2(time_seq_) * Sqr(next_v) _
                       + (cache_.K_3(time_seq_) * Sqr(v) + cache_.K_4(time_seq_) * Sqr(next_v)) / Sqr(cache_.K_3(time_seq_) * v + cache_.K_4(time_seq_) * next_v) * distribution_helper_.NormInv(random_numbers_(path_seq_, time_seq_ + 1, 1)) _
                        )
                        
    '-------------------------------------
    ' subject to improve
    '-------------------------------------
    next_s = s * K_all

'##############################################
'temp
'temp_sheet.Cells(path_seq_ + 2, time_seq_ + 2).value = ds_dvol
    
    Exit Function
        
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsCliquetMCQE.next_s :" & Chr(13) & Err.source, Err.description '1003 <-- Run error
    Exit Function

End Function

Public Sub set_fixing_value(inx As Integer, fixing_value As Double)

    deal_ticket_.set_fixing_value inx, fixing_value
    
End Sub

Public Sub set_current_date(cur_date As Date)

    deal_ticket_.current_date = cur_date

End Sub