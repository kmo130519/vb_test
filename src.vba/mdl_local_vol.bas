Option Explicit

Private Const no_of_indices As Integer = 5
Private Const index_names_1 As String = "KOSPI200"
Private Const index_names_2 As String = "SPX"
Private Const index_names_5 As String = "HSCEI"
Private Const index_names_3 As String = "SX5E"
Private Const index_names_4 As String = "NKY"

Public Function get_vol_mod_alpha(x As Double, beta As Double, fwd As Double, tau As Double, alpha As Double, rho As Double, nu As Double)

   ' Dim alpha As Double
    Dim vol_mod As Double
    Dim z As Double
    Dim chi As Double
    
    'alpha = get_alpha(beta, fwd, tau, vol_atm, rho, nu)
    
    If fwd <> x Then
    
        vol_mod = alpha * (1 + ((1 - beta) ^ 2 / 24 * alpha ^ 2 / (fwd * x) ^ (1 - beta) + 1 / 4 * rho * beta * nu * alpha / (fwd * x) ^ ((1 - beta) / 2) + (2 - 3 * rho ^ 2) / 24 * nu ^ 2) * tau)
        
        vol_mod = vol_mod / ((fwd * x) ^ ((1 - beta) / 2) * (1 + (1 - beta) ^ 2 / 24 * Log(fwd / x) ^ 2 + (1 - beta) ^ 4 / 1920 * Log(fwd / x) ^ 4))
        
        z = nu / alpha * (fwd * x) ^ ((1 - beta) / 2) * Log(fwd / x)
        
        chi = Log((Sqr(1 - 2 * rho * z + z ^ 2) + z - rho) / (1 - rho))
        
        vol_mod = vol_mod * z / chi
        
        
    
    Else
        
        vol_mod = alpha / fwd ^ (1 - beta)
        
        vol_mod = vol_mod * (1 + ((1 - beta) ^ 2 / 24 * alpha ^ 2 / (fwd ^ (2 - 2 * beta)) + 1 / 4 * rho * beta * alpha * nu / fwd ^ (1 - beta) + (2 - 3 * rho ^ 2) / 24 * nu ^ 2) * tau)
        
      
    End If
        
    get_vol_mod_alpha = vol_mod

End Function

Public Sub cmd_retrieve_local_vol_surface(Optional market_date As Date = -1, Optional prev_date_in As Date = -1)


    Dim eval_date As Date
    Dim prev_date As Date
    
    Dim sabr_surfaces() As clsSABRSurface
    Dim iv_surfaces()  As clsImpliedVolSurface
    Dim tmp_implied_vol_surface As clsImpliedVolSurface
    Dim index_list(1 To no_of_indices) As String
    Dim inx As Integer
    
    'Dim eff_vol_date As Date
    
    
    
    
    Dim data_count As Integer
    
On Error GoTo ErrorHandler

'----------------------------
' Check market date
'----------------------------
    If market_date > 0 Then
    
        shtMarket.Range("market_date").Cells(1, 1).value = market_date
        
    End If
    
    If prev_date > 0 Then
    
        shtMarket.Range("market_date").Cells(1, 2).value = prev_date
        
    End If
        
    eval_date = shtMarket.Range("market_date").Cells(1, 1).value
    prev_date = shtMarket.Range("market_date").Cells(1, 2).value


'----------------------------
' Set index names
'----------------------------
    index_list(1) = index_names_1
    index_list(2) = index_names_2
    index_list(3) = index_names_3
    index_list(4) = index_names_4
    index_list(5) = index_names_5
    
    
'----------------------------
' Retrieve Local vol Surfaces
'----------------------------
    data_count = retrieve_local_vol_surface(sabr_surfaces, index_list, eval_date, shtConfig.Range("tglNeglectCurrentDateVol").value)
    
'    If data_count <= 0 Then
'        data_count = retrieve_local_vol_surface(sabr_surfaces, index_list, prev_date)
'        If data_count <= 0 Then
'            MsgBox "No Surface"
'        End If
'    End If
            

'----------------------------
' Retrieve vol Surfaces
'----------------------------


    ReDim iv_surfaces(1 To no_of_indices) As clsImpliedVolSurface

    For inx = 1 To no_of_indices
    
        Set tmp_implied_vol_surface = New clsImpliedVolSurface
        tmp_implied_vol_surface.Init_Extract_Vol get_max_date(index_list(inx), eval_date, "UL_VOL_SURFACE"), index_list(inx)
        
'        If Not tmp_implied_vol_surface.initialized Then
'
'            tmp_implied_vol_surface.Init_Extract_Vol prev_date, index_list(inx)
'
'            If Not tmp_implied_vol_surface.initialized Then
'                MsgBox "No Surface found for " & index_list(inx)
'            End If
'
'        End If
        
        If tmp_implied_vol_surface.initialized Then
        
            Set iv_surfaces(inx) = tmp_implied_vol_surface
        
        End If
        

    Next inx
    
    shtLocalVol.lstIndices.clear
    
    shtLocalVol.lstIndices.List = index_list
    
    shtLocalVol.lstIndices.Height = 80
    shtLocalVol.lstIndices.width = 80
    
    For inx = 1 To no_of_indices
        
        clear_data index_list(inx)
        
        Dim tmp_grid As clsPillarGrid
        
        Set tmp_grid = New clsPillarGrid
        tmp_grid.set_dates iv_surfaces(inx).dates()
        tmp_grid.set_strikes iv_surfaces(inx).strikes()
        
        display_grid index_list(inx), tmp_grid, index_list(inx) & "_Vol_Surface"
        display_iv_surface index_list(inx), iv_surfaces(inx), tmp_grid, eval_date

        display_grid index_list(inx), sabr_surfaces(inx).local_vol_surface.grid_, index_list(inx) & "_local_Vol_Surface"
        display_local_vol_surface index_list(inx), sabr_surfaces(inx)
        
        'shtLocalVol.lstIndices.AddItem index_list(inx), inx - 1
        shtLocalVol.lstIndices.Selected(inx - 1) = True
        
        shtLocalVol.Range(index_list(inx) & "_local_Vol_Surface").Cells(-1, 2).value = sabr_surfaces(inx).real_date
    Next inx
    
    
    
    Exit Sub
    
ErrorHandler:

    raise_err " cmd_retrieve_local_vol_surface", Err.description

End Sub
Public Sub cmd_save_local_vol_surface()

    Dim inx As Integer
    Dim index_list(1 To no_of_indices) As String
    Dim no_of_indices_to_save As Integer
    
On Error GoTo ErrorHandler

    
    Dim markets(1 To no_of_indices) As clsMarket
    Dim sabr_surfaces() As clsSABRSurface
    Dim spot() As Double
    
    no_of_indices_to_save = 0
    
    '-----------------------------------------
    ' READ Surfaces >>>
    '-----------------------------------------
    index_list(1) = index_names_1
    index_list(2) = index_names_2
    index_list(3) = index_names_3
    index_list(4) = index_names_4
    index_list(5) = index_names_5
    
    Set markets(1) = read_kospi_index_market(index_list(1), True)

    For inx = 2 To no_of_indices
        Set markets(inx) = read_index_market(index_list(inx), True)
    Next inx
    '-----------------------------------------
    ' <<< READ Surfaces
    '-----------------------------------------
    
    For inx = 1 To no_of_indices
    
        If shtLocalVol.lstIndices.Selected(inx - 1) Then
        
            push_back_clsSABRSurface sabr_surfaces, markets(inx).sabr_surface_
            push_back_double spot, markets(inx).s_
            no_of_indices_to_save = no_of_indices_to_save + 1
            
            If Left(index_list(inx), 5) = "KOSPI" Then
                cmd_save_sabr_ma_kr
            Else
                save_a_sabr_ma index_list(inx)
            End If
         
'            Set sabr_surfaces(inx) = markets(inx).sabr_surface_
'            spot(inx) = markets(inx).s_
        
        End If
        
    Next inx
    
    save_sabr_surfaces sabr_surfaces, spot, no_of_indices_to_save 'no_of_indices
    save_sabr_param_loc_grid sabr_surfaces, no_of_indices_to_save 'no_of_indices

    Exit Sub
    
ErrorHandler:

    raise_err "cmd_save_local_vol_survace", Err.description

End Sub

Private Function read_local_vol_dates(Optional seq As Integer = 1) As Date()
    
    Dim rtn_array() As Date
    Dim inx As Integer
    
    inx = 0
    
    Do While shtLocalVol.Range("LocalVol_Tenor").Cells(1 + seq, inx + 1).value <> ""
        
        push_back_date rtn_array, shtLocalVol.Range("LocalVol_Tenor").Cells(1 + seq, inx + 1).value
        inx = inx + 1
    
    Loop
    
    read_local_vol_dates = rtn_array

End Function

Public Sub cmd_draw_surface()

    Dim data_count As Integer
    Dim eval_date As Date
    Dim prev_date As Date
    Dim prev_sabr_surfaces() As clsSABRSurface
    Dim weight As Double
    Dim local_vol_grid(1 To no_of_indices) As clsPillarGrid
    
On Error GoTo ErrorHandler

    Dim index_list(1 To no_of_indices) As String
    Dim tmp_market(1 To no_of_indices) As clsMarket
    Dim tmp_market2 As clsMarket
    Dim inx As Integer
    
    Dim local_vol_dates() As Date
    
gCounter = 1
    
    eval_date = shtMarket.Range("market_date").Cells(1, 1).value
    local_vol_dates = read_local_vol_dates(1)
    
    index_list(1) = index_names_1
    index_list(2) = index_names_2
    index_list(3) = index_names_3
    index_list(4) = index_names_4
    index_list(5) = index_names_5
    
'------------------------------
    Set tmp_market(1) = read_kospi_index_market(index_list(1), False)

    Set local_vol_grid(1) = New clsPillarGrid
    local_vol_grid(1).initialize tmp_market(1).s_, local_vol_dates, eval_date

    If shtLocalVol.chkSmoothing Then
        tmp_market(1).sabr_surface_.set_prev_surface prev_market_set__.market_by_ul(index_list(1)).sabr_surface_
    End If

    tmp_market(1).sabr_surface_.calc_local_vol_surface local_vol_grid(1), shtLocalVol.chkSmoothing, shtLocalVol.Range("LocalVol_Tenor").Cells(2, -1).value 'local_vol_dates, vol_strikes
'------------------------------
    
    For inx = 2 To no_of_indices
        Set tmp_market(inx) = read_index_market(index_list(inx), False)
        local_vol_dates = read_local_vol_dates(inx)
         
        Set local_vol_grid(inx) = New clsPillarGrid
        local_vol_grid(inx).initialize tmp_market(inx).s_, local_vol_dates, eval_date
        
        If shtLocalVol.chkSmoothing Then
            tmp_market(inx).sabr_surface_.set_prev_surface prev_market_set__.market_by_ul(index_list(inx)).sabr_surface_
        End If
        
        tmp_market(inx).sabr_surface_.calc_local_vol_surface local_vol_grid(inx), shtLocalVol.chkSmoothing, shtLocalVol.Range("LocalVol_Tenor").Cells(inx + 1, -1).value  'local_vol_dates, vol_strikes

    Next inx
    
    
    For inx = 1 To no_of_indices
    
        If shtLocalVol.lstIndices.Selected(inx - 1) Then
            clear_data index_list(inx)
            display_grid index_list(inx), tmp_market(inx).sabr_surface_.vol_surface.grid_, index_list(inx) & "_Vol_Surface"  ' .grid_, index_list(inx) & "_Vol_Surface"
            display_vol_surface index_list(inx), tmp_market(inx).sabr_surface_
                                
            If shtLocalVol.chkShift Then
                    
                 tmp_market(inx).sabr_surface_.shift_surface shtLocalVol.Range(index_list(inx) & "_local_Vol_Surface").Cells(0, 12).value
                
            End If
            
            display_grid index_list(inx), tmp_market(inx).sabr_surface_.local_vol_surface.grid_, index_list(inx) & "_local_Vol_Surface"
            display_local_vol_surface index_list(inx), tmp_market(inx).sabr_surface_
            
        End If
        
    Next inx
        
    Exit Sub
    
ErrorHandler:

    raise_err "cmd_draw_surface", Err.description

End Sub
'
'Private Sub smoothe_local_vol(today_surface As clsSABRSurface, prev_sabr_surface As clsSABRSurface, eval_date As Date, prev_weight As Double)
'
'    Dim inx As Integer
'    Dim jnx As Integer
'
'    Dim tmp_vol As Double
'
'On Error GoTo ErrorHandler
'
'    For inx = 1 To today_surface.local_vol_grid_.no_of_dates
'        For jnx = 1 To today_surface.local_vol_grid_.no_of_strikes
'            tmp_vol = prev_sabr_surface.interpolated_local_vol(eval_date, today_surface.local_vol_grid_.strikes(jnx), today_surface.local_vol_grid_.dates(inx))
'            today_surface.set_local_vol tmp_vol * prev_weight + today_surface.local_vol_surface()(inx, jnx) * (1 - prev_weight), inx, jnx
'        Next jnx
'    Next inx
'
'    Exit Sub
'
'ErrorHandler:
'
'    'raise_err "smoothe_local_vol", Err.description
'    MsgBox "No data to smooothe.. Continue the process anyway.. "
'
'End Sub

Private Sub clear_data(index_name As String, Optional ByVal local_vol_only As Boolean = False)

    Dim row_count As Integer
    Dim local_vol_row_count As Integer
    
    
    If Not local_vol_only Then
        local_vol_row_count = shtLocalVol.Range(index_name & "_Local_Vol_Surface").Cells(0, 1).value
        shtLocalVol.Range(index_name & "_Local_Vol_Surface").Cells(2, -6).Range("A1:BZ" & local_vol_row_count + 1).ClearContents
        shtLocalVol.Range(index_name & "_Local_Vol_Surface").Range("B1:BZ1").ClearContents
    End If
    
    
    row_count = shtLocalVol.Range(index_name & "_Vol_Surface").Cells(0, 1).value
    shtLocalVol.Range(index_name & "_Vol_Surface").Cells(2, -6).Range("A1:BZ" & row_count + 1).ClearContents
    shtLocalVol.Range(index_name & "_Vol_Surface").Range("B1:BZ1").ClearContents

End Sub


Public Sub display_local_vol_surface(ByVal index_name As String, sabr_surface As clsSABRSurface)

    Dim inx As Integer
    Dim jnx As Integer
    Dim the_range As Range
    
On Error GoTo ErrorHandler

    Set the_range = shtLocalVol.Range(index_name & "_local_Vol_Surface")
    
    
    For jnx = 1 To UBound(sabr_surface.local_vol_surface.grid_.get_all_dates()) 'sabr_surface.local_vol_grid_.get_all_dates())
        

        For inx = 1 To UBound(sabr_surface.local_vol_surface.grid_.get_all_strikes()) '.local_vol_grid_.get_all_strikes())

            the_range.Cells(1 + jnx, 1 + inx).value = sabr_surface.local_vol_surface.vol_surface()(jnx, inx)
        
        Next inx
    Next jnx
    
    
    For jnx = 1 To sabr_surface.sabr_parameters_loc_.no_of_dates
        the_range.Cells(1 + jnx, -6).value = sabr_surface.sabr_parameters_loc_.sabr_param(jnx).forward
        the_range.Cells(1 + jnx, -5).value = sabr_surface.sabr_parameters_loc_.sabr_param(jnx).alpha
        the_range.Cells(1 + jnx, -4).value = sabr_surface.sabr_parameters_loc_.sabr_param(jnx).beta
        the_range.Cells(1 + jnx, -3).value = sabr_surface.sabr_parameters_loc_.sabr_param(jnx).nu
        the_range.Cells(1 + jnx, -2).value = sabr_surface.sabr_parameters_loc_.sabr_param(jnx).rho
        the_range.Cells(1 + jnx, -1).value = sabr_surface.sabr_parameters_loc_.sabr_param(jnx).vol_atm
        the_range.Cells(1 + jnx, 0).value = sabr_surface.sabr_parameters_loc_.sabr_param(jnx).tau
    Next jnx
    
    Exit Sub
    
ErrorHandler:

    raise_err "display_local_vol_surface", Err.description

End Sub

Private Sub display_grid(ByVal index_name As String, grid As clsPillarGrid, range_name As String)

    Dim the_range As Range
    Dim inx As Integer
    Dim jnx As Integer
    
    Set the_range = shtLocalVol.Range(range_name)
    
    the_range.Cells(-1, 1).value = index_name & " Index"
    
    For inx = 1 To UBound(grid.get_all_strikes())
    
        the_range.Cells(1, 1 + inx).value = grid.strikes(inx)
    
    Next inx
    
    For inx = 1 To UBound(grid.get_all_dates())
    
        the_range.Cells(1 + inx, 1).value = grid.dates(inx)
    
    Next inx

    
End Sub
'
'Public Sub display_iv_surface(ByVal index_name As String, iv_surface As clsImpliedVolSurface, grid As clsPillarGrid, eval_date As Date)
'
'    Dim inx As Integer
'    Dim jnx As Integer
'    Dim the_range As Range
'
'On Error GoTo ErrorHandler
'
'    Set the_range = shtLocalVol.Range(index_name & "_Vol_Surface")
'
'    For inx = 1 To UBound(grid.get_all_strikes())
'
'        For jnx = 1 To UBound(grid.get_all_dates())
'
'            the_range.Cells(1 + jnx, 1 + inx).value = iv_surface.Extract_Index_Vol(eval_date, grid.dates(jnx), grid.strikes(inx))
'
'        Next jnx
'
'    Next inx
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "display_iv_surface", Err.description
'
'End Sub

Public Sub display_vol_surface(ByVal index_name As String, sabr_surface As clsSABRSurface)

    Dim inx As Integer
    Dim jnx As Integer
    Dim the_range As Range
    
On Error GoTo ErrorHandler

    Set the_range = shtLocalVol.Range(index_name & "_Vol_Surface")
    
    the_range.Cells(-1, 1).value = index_name & " Index"
    
    For inx = 1 To UBound(sabr_surface.vol_surface.grid_.get_all_strikes())  '.grid().get_all_strikes())
    
        the_range.Cells(1, 1 + inx).value = sabr_surface.vol_surface.grid_.get_all_strikes()(inx)
    
    Next inx
    
    For inx = 1 To UBound(sabr_surface.vol_surface.grid_.get_all_dates()) ' sabr_surface.grid().get_all_dates())
    
        the_range.Cells(1 + inx, 1).value = sabr_surface.vol_surface.grid_.get_all_dates()(inx)
    
    Next inx
    
    For inx = 1 To UBound(sabr_surface.vol_surface.grid_.get_all_strikes())
    
        For jnx = 1 To UBound(sabr_surface.vol_surface.grid_.get_all_dates())
    
         the_range.Cells(1 + jnx, 1 + inx).value = sabr_surface.vol_surface.vol_surface()(jnx, inx)
    
        Next jnx
    
    Next inx
    
    Exit Sub
    
ErrorHandler:

    raise_err "display_vol_surface", Err.description

End Sub