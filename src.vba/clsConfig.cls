Option Explicit

Public ip_address_ As String

Public next_date_ As Date
Public last_date_ As Date
Public date_before_yesterday_ As Date
Public two_days_before_yesterday_ As Date
Public current_date_ As Date

Public date_week_ago_ As Date
Public date_month_ago_ As Date
Public date_3_month_ago_ As Date
Public date_6_month_ago_ As Date

Public time_step_closing_ As Integer
Public no_of_trials_closing_ As Long
Public required_tolerance_closing_ As Double

Public time_step_ As Integer
Public no_of_trials_ As Long
Public required_tolerance_ As Double

Public grid_interval_ As Double

Public min_s_ As Double
Public max_s_ As Double

Public market_refresh_interval_ As Long
Public auto_calculation_ As Boolean
Public intra_day_greek_ As String
Public neglect_barrier_smoothing_ As Boolean

Public x_grid_ As Integer
Public v_grid_ As Integer
Public time_step_per_day As Double
Public fdm_scheme_ As Integer

Public vega_reference_maturity As Double
Public empirical_vega_weighting As Double

Public file_path_ As String

Public sparse_grid_level As Integer
Public sparse_grid_min As Double
Public sparse_grid_max As Double

Public dll_version_ As String

Public no_of_strike_grid As Integer
Public width_of_strike As Double

Public max_retrial_count_ As Integer
Public snapshot_file_extension As String
Public position_file_extension As String
Public position_summary_file_extension As String
Public realtime_file_name As String

Public adjust_strike_shift_percent As Double

Public batch_size As Integer

Private term_vega_tenor_() As Date



Public Sub set_term_vega_tenor(term_vega_tenor() As Date)

    Dim inx As Integer
    
    ReDim term_vega_tenor_(1 To get_array_size_date(term_vega_tenor))
    
    
    For inx = 1 To get_array_size_date(term_vega_tenor)
        term_vega_tenor_(inx) = term_vega_tenor(inx)
    Next inx

End Sub

Public Function term_vega_tenor(inx As Integer) As Date

    term_vega_tenor = term_vega_tenor_(inx)
    
End Function

Public Function term_vega_tenor_array() As Date()

    term_vega_tenor_array = term_vega_tenor_
    
End Function





Public Function copy_obj() As clsConfig

    Dim rtn_obj As clsConfig
    
    Set rtn_obj = New clsConfig

    rtn_obj.ip_address_ = ip_address_
    
    rtn_obj.next_date_ = next_date_
    rtn_obj.last_date_ = last_date_
    rtn_obj.date_before_yesterday_ = date_before_yesterday_
    rtn_obj.two_days_before_yesterday_ = two_days_before_yesterday_
    rtn_obj.current_date_ = current_date_
    
    rtn_obj.date_week_ago_ = date_week_ago_
    rtn_obj.date_month_ago_ = date_month_ago_
    rtn_obj.date_3_month_ago_ = date_3_month_ago_
    rtn_obj.date_6_month_ago_ = date_6_month_ago_
    
    rtn_obj.time_step_closing_ = time_step_closing_
    rtn_obj.no_of_trials_closing_ = no_of_trials_closing_
    rtn_obj.required_tolerance_closing_ = required_tolerance_closing_
    
    rtn_obj.time_step_ = time_step_
    rtn_obj.no_of_trials_ = no_of_trials_
    rtn_obj.required_tolerance_ = required_tolerance_
    
    rtn_obj.grid_interval_ = grid_interval_
    
    rtn_obj.min_s_ = min_s_
    rtn_obj.max_s_ = max_s_
    
    rtn_obj.market_refresh_interval_ = market_refresh_interval_
    rtn_obj.auto_calculation_ = auto_calculation_
    rtn_obj.intra_day_greek_ = intra_day_greek_
    rtn_obj.neglect_barrier_smoothing_ = neglect_barrier_smoothing_
    
    rtn_obj.x_grid_ = x_grid_
    rtn_obj.v_grid_ = v_grid_
    rtn_obj.time_step_per_day = time_step_per_day
    rtn_obj.fdm_scheme_ = fdm_scheme_
    
    rtn_obj.vega_reference_maturity = vega_reference_maturity
    rtn_obj.empirical_vega_weighting = empirical_vega_weighting
    
    rtn_obj.file_path_ = file_path_
    
    rtn_obj.sparse_grid_level = sparse_grid_level
    rtn_obj.sparse_grid_min = sparse_grid_min
    rtn_obj.sparse_grid_max = sparse_grid_max
    
    rtn_obj.dll_version_ = dll_version_
    
    rtn_obj.no_of_strike_grid = no_of_strike_grid
    rtn_obj.width_of_strike = width_of_strike
    
    rtn_obj.max_retrial_count_ = max_retrial_count_
    rtn_obj.snapshot_file_extension = snapshot_file_extension
    rtn_obj.position_file_extension = position_file_extension
    rtn_obj.position_summary_file_extension = position_summary_file_extension
    rtn_obj.realtime_file_name = realtime_file_name
    
    rtn_obj.adjust_strike_shift_percent = adjust_strike_shift_percent
    
    
    rtn_obj.set_term_vega_tenor term_vega_tenor_
    
    rtn_obj.batch_size = batch_size

    Set copy_obj = rtn_obj

End Function