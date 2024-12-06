Option Explicit

Public Sub btnReadLocalVol_Click()

    '전역변수 설정
    Call SET_GLOBAL

    '위기상황분석 옵션
    Dim scenario_id As String
    If SCENARIO_ENABLE = False Then
        scenario_id = "0"
    Else
        scenario_id = shtMarket.Range("SCENARIO_ID").value
    End If
    
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    Dim adoCon As New adoDB.Connection
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
    Dim source As String
    source = shtLocalVol.Range("market_source").value
    
    Dim vol_type As String
    vol_type = shtLocalVol.Range("vol_type").value
    
    Dim market_date_str As String
    market_date_str = date2str(shtLocalVol.Range("market_date").value)
    
    Dim market_prevdate_str As String
    market_prevdate_str = date2str(shtLocalVol.Range("market_prevdate").value)
    
    Dim the_range As Range
    Dim ua_code As String
    
    With Application
        .ScreenUpdating = False
        .StatusBar = False
        .DisplayStatusBar = True
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Dim i As Integer
    
    For i = 1 To NUM_UA

        ua_code = get_ua_code(i)
        
        If get_ua_idx(ua_code) <> ua.KRD020021147 Then 'KOSPI200 레버리지는 KOSPI200 vol. surface 적용하므로 skip
        
            Set the_range = shtLocalVol.Range(ua_code & "_local_vol_surface")
            the_range.Range(Cells(1, 1), Cells(51, 104)).ClearContents
            the_range.Offset(0, 105).Range(Cells(1, 1), Cells(51, 104)).ClearContents
        
            If is_active_ua(ua_code) Then
                
                Application.StatusBar = "Importing volatility data... " & ua_code
                
                'If Left(ua_code, 3) = "KR7" Then 'single stock은 flat vol. 적용
                If is_flatvol_ua(ua_code) = True Then
                    'flat_vol
                    Call fill_flat_vol(ua_code, the_range, market_date_str, source, scenario_id, adoCon)
                    Call fill_flat_vol(ua_code, the_range.Offset(0, 105), market_prevdate_str, source, "0", adoCon)
                Else
                    'vol_surface
                    Call fill_local_vol(ua_code, vol_type, the_range, market_date_str, source, scenario_id, adoCon)
                    Call fill_local_vol(ua_code, vol_type, the_range.Offset(0, 105), market_prevdate_str, source, "0", adoCon)
                End If
                
                With the_range.Cells(-1, 1).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                
            Else
            
                the_range.Cells(1, 1) = "미사용"
                With the_range.Cells(-1, 1).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.499984740745262
                    .PatternTintAndShade = 0
                End With
            
            End If
        
        End If
        
    Next i
   
    Call disconnectDB(adoCon)
    
    With Application
        .ScreenUpdating = True
        .StatusBar = False
        .DisplayStatusBar = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    
    Set the_range = Nothing
    
    Set oRS = Nothing
    Set oCmd = Nothing
    Set adoCon = Nothing

End Sub

Private Sub fill_local_vol(ul_code As String, vol_type As String, the_range As Range, tdate As String, source As String, scenario_id As String, oDB As adoDB.Connection)

    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(2) As String
    ReDim bind_value(2) As Variant
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":ul_code"
    bind_value(1) = tdate
    bind_value(2) = ul_code

    If SCENARIO_ENABLE = True Then
        ReDim Preserve bind_variable(3) As String
        ReDim Preserve bind_value(3) As Variant
        bind_variable(3) = ":scenarioid"
        bind_value(3) = scenario_id
    End If
    
    Dim size_t As Integer
    Dim size_k As Integer
    
    Call get_surface_size(size_t, size_k, ul_code, vol_type, tdate, source, oDB)
    
    Select Case vol_type
    Case "Implied"
        If source = "FRONT" Then
            sql = getSQL(SQL_PATH_IV_SURFACE_FRONT, bind_variable, bind_value)
        Else
            sql = getSQL(SQL_PATH_IV_SURFACE, bind_variable, bind_value)
        End If
    Case "Local"
        If source = "FRONT" Then
            If SCENARIO_ENABLE Then
                sql = getSQL(SQL_PATH_LV_SURFACE_ST, bind_variable, bind_value)
            Else
                sql = getSQL(SQL_PATH_LV_SURFACE_FRONT, bind_variable, bind_value)
            End If
        Else
            If SCENARIO_ENABLE Then
                sql = getSQL(SQL_PATH_LV_SURFACE_ST, bind_variable, bind_value)
            Else
                sql = getSQL(SQL_PATH_LV_SURFACE, bind_variable, bind_value)
            End If
        End If
    End Select

    With oCmd
        .ActiveConnection = oDB
        .CommandType = adCmdText
        .CommandText = sql

        oRS.Open .Execute
    End With

    Dim i, j As Integer
    i = 2
    j = 2
    
    the_range.Cells(1, 1) = str2date(tdate)
    
    Do Until oRS.EOF
    
        the_range.Cells(1, j) = oRS(2) 'strike
        the_range.Cells(i, 1) = str2date(oRS(3)) 'maturity
        the_range.Cells(i, j) = oRS(4) 'vol

        i = i + 1
        If (i = size_t + 2) Then
            i = 2
            j = j + 1
        End If
        oRS.MoveNext
    Loop

    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing

End Sub

Private Sub fill_flat_vol(ul_code As String, the_range As Range, tdate As String, source As String, scenario_id As String, oDB As adoDB.Connection)

    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(2) As String
    ReDim bind_value(2) As Variant
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":code"
    bind_value(1) = tdate
    bind_value(2) = ul_code

    If SCENARIO_ENABLE = True Then
        ReDim Preserve bind_variable(3) As String
        ReDim Preserve bind_value(3) As Variant
        bind_variable(3) = ":scenarioid"
        bind_value(3) = scenario_id
    End If
    
    sql = getSQL(SQL_PATH_FLAT_VOL, bind_variable, bind_value)

    With oCmd
        .ActiveConnection = oDB
        .CommandType = adCmdText
        .CommandText = sql

        oRS.Open .Execute
    End With

    Dim num_dates As Integer
    Dim num_strikes As Integer
    num_dates = 36 '3y
    num_strikes = 101 'every 0.2%
    
    Dim s0 As Double
    s0 = get_spot_price(ul_code, tdate, scenario_id, oDB)
    
    Dim i, j As Integer
    
    the_range.Cells(1, 1) = str2date(tdate)
    
    Do Until oRS.EOF

        For i = 1 To num_dates
            the_range.Cells(i + 1, 1) = DateAdd("m", i - 1, str2date(tdate) + 1)

            For j = 1 To num_strikes
                If i = 1 Then
                    the_range.Cells(1, j + 1) = s0 * (1 + 0.002 * (j - 51))
                End If
                the_range.Cells(i + 1, j + 1) = oRS(0)
            Next j
        Next i

        oRS.MoveNext
    Loop

    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
End Sub