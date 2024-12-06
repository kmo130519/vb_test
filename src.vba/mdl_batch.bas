Sub ScheduledMacro()

    Dim time_value As Date
    time_value = TimeValue(shtELS.Range("Scheduled_Time").Text)
    
    Application.OnTime time_value, "Batch"
    Application.StatusBar = "The batch procedure is supposed to run at " & CStr(time_value)
    
End Sub

Sub Batch()

    shtMarket.btnReadMarketData_Click
    shtLocalVol.btnReadLocalVol_Click
    shtELS.btnCalculatePriceELS_Click
    Application.StatusBar = False
    
End Sub