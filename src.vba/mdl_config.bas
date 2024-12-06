Option Explicit

Public realtime_running_3d As Boolean


Public Sub display_time_checker()
    
    Dim inx As Integer
    Dim array_size As Integer
    
    array_size = time_checker.get_array_count
    
    shtConfig.Range("time_check_start").Range("A2:C1000").ClearContents
    
    For inx = 1 To array_size
    
        shtConfig.Range("time_check_start").Cells(inx + 1, 1).value = time_checker.stop_point_name()(inx)
        shtConfig.Range("time_check_start").Cells(inx + 1, 2).value = time_checker.tick_counter()(inx)
        
        If inx > 1 Then
            shtConfig.Range("time_check_start").Cells(inx + 1, 3).value = time_checker.tick_counter()(inx) - time_checker.tick_counter()(inx - 1)
        End If
        
    Next inx
        
    time_checker.initailize

End Sub


Public Sub real_time_on()

'    shtMarket.Cells.Replace "SP_RTD(", "=SP_RTD(", xlPart, xlByRows, False, False, False, False
'    shtMarketForeign.Cells.Replace "SP_RTD(", "=SP_RTD(", xlPart, xlByRows, False, False, False, False
End Sub

Public Sub real_time_off()

'    shtMarket.Cells.Replace "=SP_RTD(", "SP_RTD(", xlPart, xlByRows, False, False, False, False
'    shtMarketForeign.Cells.Replace "=SP_RTD(", "SP_RTD(", xlPart, xlByRows, False, False, False, False
    
End Sub


Public Sub cmd_init()


On Error GoTo ErrorHandler

time_checker.add_tick_counter "START"

    shtIndexPosition.Range("tglRealTime").value = False
    shtIndexPosition.Range("tglTimer").value = False
    shtIndexPosition.Range("tglEndofDay").value = False
    shtIndexPosition.Range("tglExcludeIntraday").value = False
    
    shtConfig.Range("tglRetrieveDb").value = True
    shtConfig.Range("tglNeglectCurrentDateVol").value = True
        
    
    real_time_off
    
time_checker.add_tick_counter "Button Toggle Completed"

    initialize_global_variables
    
time_checker.add_tick_counter "Init g variable Completed"
    '------------------------------------------------------------------------------------
    ' Reset current date to today and retrieve next date and last date from "DATE_INF"
    '------------------------------------------------------------------------------------
    reset_closing_date
    reset_config
    
time_checker.add_tick_counter "Date setting Completed"

    load_market True, True ', config__.current_date_
    
time_checker.add_tick_counter "Load market Completed"

    display_config_date
    


    initialized__ = True
    
    Application.Calculation = xlCalculationManual
        
    initialize_holiday_list
    
time_checker.add_tick_counter "Initialize_holiday_list Completed"
    
    reset_search_condition
    
    cmd_retrieve_deal_list
    
time_checker.add_tick_counter "cmd_retrieve_deal_list Completed"
    
    cmd_retrieve_ac_deal_list
    
time_checker.add_tick_counter "cmd_retrieve_ac_deal_list Completed"

    cmd_retrieve_vanilla
    
time_checker.add_tick_counter "cmd_retrieve_vanilla Completed"
    
    cmd_retrieve_futures
    
time_checker.add_tick_counter "cmd_retrieve_futures Completed"
    
    Application.Calculation = xlCalculationAutomatic
    
time_checker.add_tick_counter "Cell Calculation Completed"

    
    cmd_TimerOff
    shtACList.Range("tgl_3d").value = False
    realtime_running_3d = False
    
time_checker.add_tick_counter "cmd_TimerOff Completed"

display_time_checker

    Exit Sub
    
ErrorHandler:
    initialized__ = False
    raise_err "cmd_init"
    
End Sub

Public Function read_term_vega_tenor() As Date()

    Dim rtn_array() As Date
    Dim inx As Integer
    
    inx = 1
    
    While shtConfig.Range("term_vega_tenor").Cells(inx, 1) <> ""
        
        push_back_date rtn_array, shtConfig.Range("date_config").Cells(1, 1).value + shtConfig.Range("term_vega_tenor").Cells(inx, 1)
        inx = inx + 1
        
    Wend
    
    
    read_term_vega_tenor = rtn_array

End Function

Public Sub cmd_reset_config()

On Error GoTo ErrorHandler


    If shtConfig.Range("tglClearGloablVar") Then

        initialize_global_variables
    
    End If
    
    reset_config
    
    display_config_date
    
    initialize_holiday_list
    
    cmd_TimerOff
    
    initialized__ = True
    
    Exit Sub
    
ErrorHandler:

    initialized__ = False
    raise_err "cmd_reset_config"

End Sub

Public Sub cmd_morning_closing()

On Error GoTo ErrorHandler
                                
    run_morning_closing
    
    MsgBox "Closing completed." & Chr(13) & output_string
    
    Exit Sub
    
ErrorHandler:

    raise_err "cmd_morning_closing", Err.description & output_string
    

End Sub
'
'Public Sub cmd_closing()
'
'On Error GoTo ErrorHandler
'
'    run_closing
'
'    MsgBox "Closing completed." & Chr(13) & output_string
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "cmd_closing", Err.description & output_string
'
'
'End Sub

Public Sub cmd_market_closing()

On Error GoTo ErrorHandler
                                
    closing_greek_ac_market
    
    MsgBox "Closing completed." & Chr(13) & output_string
    
    Exit Sub
    
ErrorHandler:

    raise_err "cmd_market_closing", Err.description & output_string
    

End Sub


Public Sub reset_closing_date(Optional closing_date As Date = -1)
    

On Error GoTo ErrorHandler
                                
                                
    If closing_date <= 0 Then
        closing_date = Date
    End If
    
    shtConfig.Range("date_config").Cells(1, 1).value = closing_date

    Exit Sub
    
ErrorHandler:

    raise_err "reset_closing_date"


End Sub

Public Sub reset_search_condition(Optional LIVE_YN As String = "Y", Optional CONFIRM_YN As String = "Y")

On Error GoTo ErrorHandler
    
    shtCliquetList.Range("search_condition").Cells(2, 1).value = LIVE_YN
    shtCliquetList.Range("search_condition").Cells(2, 2).value = CONFIRM_YN
    
    shtACList.Range("search_condition").Cells(2, 1).value = LIVE_YN
    shtACList.Range("search_condition").Cells(2, 2).value = CONFIRM_YN

    Exit Sub
    
ErrorHandler:

    raise_err "reset_search_condition"

End Sub


Public Sub display_config_date()


On Error GoTo ErrorHandler
                                      
    shtConfig.Range("date_config").Cells(1, 1).value = config__.current_date_
    shtConfig.Range("date_config").Cells(2, 1).value = config__.next_date_
    shtConfig.Range("date_config").Cells(3, 1).value = config__.last_date_
    shtConfig.Range("date_config").Cells(4, 1).value = config__.date_before_yesterday_
    shtConfig.Range("date_config").Cells(5, 1).value = config__.date_week_ago_
    shtConfig.Range("date_config").Cells(6, 1).value = config__.date_month_ago_
    shtConfig.Range("date_config").Cells(7, 1).value = config__.date_3_month_ago_
    shtConfig.Range("date_config").Cells(8, 1).value = config__.date_6_month_ago_

    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 10000, "refresh_config : " & Chr(13) & Err.source, Err.description
    

End Sub