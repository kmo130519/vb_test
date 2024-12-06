Option Explicit

Public Function str2date(tdate As String) As Date
'String YYYYMMDD -> Date YYYY-MM-DD

    str2date = CDate(Left(tdate, 4) + "-" + Mid(tdate, 5, 2) + "-" + Right(tdate, 2))

End Function

Public Function date2str(tdate As Date) As String
'Date YYYY-MM-DD -> String YYYYMMDD

    Dim str_tdate As String
    str_tdate = CStr(tdate)

    date2str = Left(str_tdate, 4) & Mid(str_tdate, 6, 2) & Right(str_tdate, 2)

End Function

Public Function ul_ofs(ul_code As String) As Integer

    Dim ofs As Integer
    
    Select Case ul_code
    Case "KOSPI200": ofs = 1
    Case "HSCEI": ofs = 2
    Case "HSI": ofs = 3
    Case "NKY": ofs = 4
    Case "SX5E": ofs = 5
    Case "SPX": ofs = 6
    Case "KR7005380001": ofs = 7 '현대차
    Case "KR7005930003": ofs = 8 '삼성전자
    Case "KR7000030007": ofs = 9 '우리은행
    Case "KR7028260008": ofs = 10 '삼성물산
    Case "KR7105560007": ofs = 11 'KB금융
    Case "KR7035420009": ofs = 12 'NAVER
    Case "KR7018260000": ofs = 13 '삼성SDS
    Case "KR7005490008": ofs = 14 'POSCO
    Case "KR7034220004": ofs = 15 'LG디스플레이
    End Select

    ul_ofs = ofs

End Function

Public Function get_ua_idx(ua_code As String) As Integer

    Dim rtn As Integer
    
    Select Case ua_code
    Case "HSCEI": rtn = ua.HSCEI
    Case "HSI": rtn = ua.HSI
    Case "SX5E": rtn = ua.SX5E
    Case "SPX": rtn = ua.SPX
    Case "NKY": rtn = ua.NKY
    Case "KOSPI200": rtn = ua.KOSPI200
    Case "KRD020021147": rtn = ua.KRD020021147
    Case "KR7005380001": rtn = ua.KR7005380001
    Case "KR7005930003": rtn = ua.KR7005930003
    Case "KR7000030007": rtn = ua.KR7000030007
    Case "KR7028260008": rtn = ua.KR7028260008
    Case "KR7105560007": rtn = ua.KR7105560007
    Case "KR7035420009": rtn = ua.KR7035420009
    Case "KR7018260000": rtn = ua.KR7018260000
    Case "KR7005490008": rtn = ua.KR7005490008
    Case "KR7034220004": rtn = ua.KR7034220004
    Case "TSLA", "SPTESLA", "US88160R1014": rtn = ua.TSLA '2024.04.29 SP code, ISIN code 추가
    Case "NVDA", "SPNVDA", "US67066G1040": rtn = ua.NVDA '2024.06.19 SP code, ISIN code 추가
    End Select
    
    get_ua_idx = rtn
    
End Function

Public Function get_ua_code(idx As Integer, Optional code_type As Integer = UA_CODE_TYPE.BLP) As String

    Dim rtn As String
    
    Select Case idx
    Case ua.HSCEI: rtn = "HSCEI"
    Case ua.HSI: rtn = "HSI"
    Case ua.SX5E: rtn = "SX5E"
    Case ua.SPX: rtn = "SPX"
    Case ua.NKY: rtn = "NKY"
    Case ua.KOSPI200: rtn = "KOSPI200"
    Case ua.KRD020021147: rtn = "KRD020021147"
    Case ua.KR7005380001: rtn = "KR7005380001"
    Case ua.KR7005930003: rtn = "KR7005930003"
    Case ua.KR7000030007: rtn = "KR7000030007"
    Case ua.KR7028260008: rtn = "KR7028260008"
    Case ua.KR7105560007: rtn = "KR7105560007"
    Case ua.KR7035420009: rtn = "KR7035420009"
    Case ua.KR7018260000: rtn = "KR7018260000"
    Case ua.KR7005490008: rtn = "KR7005490008"
    Case ua.KR7034220004: rtn = "KR7034220004"
    Case ua.TSLA:
        Select Case code_type
        Case UA_CODE_TYPE.BLP: rtn = "TSLA"
        Case UA_CODE_TYPE.SP: rtn = "SPTESLA"
        Case UA_CODE_TYPE.ISIN: rtn = "US88160R1014"
        End Select
    Case ua.NVDA:
        Select Case code_type
        Case UA_CODE_TYPE.BLP: rtn = "NVDA"
        Case UA_CODE_TYPE.SP: rtn = "SPNVDA"
        Case UA_CODE_TYPE.ISIN: rtn = "US67066G1040"
        End Select
    End Select
    
    get_ua_code = rtn
    
End Function

Public Function get_ua_name(ua_code As String) As String

    Dim rtn As String
    
    Select Case get_ua_idx(ua_code)
    Case ua.HSCEI: rtn = "HSCEI"
    Case ua.HSI: rtn = "HSI"
    Case ua.SX5E: rtn = "EuroStoxx50"
    Case ua.SPX: rtn = "S&P500"
    Case ua.NKY: rtn = "NIKKEI225"
    Case ua.KOSPI200: rtn = "KOSPI200"
    Case ua.KRD020021147: rtn = "KOSPI200레버리지"
    Case ua.KR7005380001: rtn = "현대차"
    Case ua.KR7005930003: rtn = "삼성전자"
    Case ua.KR7000030007: rtn = "우리은행"
    Case ua.KR7028260008: rtn = "삼성물산"
    Case ua.KR7105560007: rtn = "KB금융"
    Case ua.KR7035420009: rtn = "NAVER"
    Case ua.KR7018260000: rtn = "삼성SDS"
    Case ua.KR7005490008: rtn = "POSCO"
    Case ua.KR7034220004: rtn = "LG디스플레이"
    Case ua.TSLA: rtn = "TESLA"
    Case ua.NVDA: rtn = "NVIDIA"
    End Select
    
    get_ua_name = rtn
    
End Function

Public Function get_ua_currency(ua_code As String) As String

    Dim rtn As String
    
    Select Case get_ua_idx(ua_code)
    Case ua.HSCEI, ua.HSI: rtn = "HKD"
    Case ua.SX5E: rtn = "EUR"
    Case ua.SPX, ua.TSLA, ua.NVDA: rtn = "USD"
    Case ua.NKY: rtn = "JPY"
    Case Else: rtn = "KRW"
    End Select
    
    get_ua_currency = rtn
    
End Function

Public Function get_ccy_code(ccy_idx As Integer) As String

    Dim rtn As String
    
    Select Case ccy_idx
    Case ccy.EUR: rtn = "EUR"
    Case ccy.HKD: rtn = "HKD"
    Case ccy.JPY: rtn = "JPY"
    Case ccy.KRW: rtn = "KRW"
    Case ccy.USD: rtn = "USD"
    End Select
    
    get_ccy_code = rtn
    
End Function

Public Function is_active_ua(ua_code As String) As Boolean

    Dim rtn As Boolean
    
    Select Case get_ua_idx(ua_code)
    Case ua.HSCEI, ua.HSI, ua.KOSPI200, ua.NKY, ua.SX5E, ua.SPX, ua.TSLA, ua.NVDA: rtn = True
    'Case ua.HSCEI, ua.HSI, ua.KOSPI200, ua.NKY, ua.SX5E, ua.SPX: rtn = True
    Case Else: rtn = False
    End Select
        
    is_active_ua = rtn
    
End Function

Public Function is_active_ccy(ccy_code As String) As Boolean

    Dim rtn As Boolean
    
    Select Case get_ccy_idx(ccy_code)
    Case ccy.KRW, ccy.USD, ccy.EUR, ccy.JPY, ccy.HKD: rtn = True
    Case Else: rtn = False
    End Select
    
    is_active_ccy = rtn
    
End Function

Public Function get_ccy_idx(ccy_code As String) As Integer

    Dim rtn As Integer
    
    Select Case ccy_code
    Case "KRW": rtn = ccy.KRW
    Case "USD": rtn = ccy.USD
    Case "EUR": rtn = ccy.EUR
    Case "JPY": rtn = ccy.JPY
    Case "HKD": rtn = ccy.HKD
    End Select
    
    get_ccy_idx = rtn
    
End Function

Public Function get_rf_num_term(ccy_code As String) As Integer

    Dim rtn As Integer
    
    Select Case get_ccy_idx(ccy_code)
    Case ccy.HKD: rtn = RF_NUM_TERM.HKD
    Case ccy.EUR: rtn = RF_NUM_TERM.EUR
    Case ccy.USD: rtn = RF_NUM_TERM.USD
    Case ccy.JPY: rtn = RF_NUM_TERM.JPY
    Case ccy.KRW: rtn = RF_NUM_TERM.KRW
    End Select
    
    get_rf_num_term = rtn
    
End Function

Public Function get_dcf_ccy(idx As Integer) As String

    Dim rtn As String
    
    Select Case idx
    Case DCF.KRW: rtn = "KRW"
    Case DCF.USD: rtn = "USD"
    End Select
    
    get_dcf_ccy = rtn
    
End Function

Public Function get_dcf_idx(ccy_code As String) As Integer

    Dim rtn As String
    
    Select Case get_ccy_idx(ccy_code)
    Case ccy.KRW: rtn = DCF.KRW
    Case ccy.USD: rtn = DCF.USD
    End Select
    
    get_dcf_idx = rtn
    
End Function

Public Function get_spot_price(ul_code As String, tdate As String, Optional scenario_id As String, Optional adoCon As adoDB.Connection) As Double

    If adoCon Is Nothing Then
        Set adoCon = New adoDB.Connection
        Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    End If

    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
   
    Dim bind_variable() As String
    ReDim bind_variable(2)
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":code"

    Dim bind_value() As Variant
    ReDim bind_value(2)
    bind_value(1) = tdate
    bind_value(2) = ul_code
    
    If SCENARIO_ENABLE = True Then
        ReDim Preserve bind_variable(3) As String
        ReDim Preserve bind_value(3) As Variant
        bind_variable(3) = ":scenarioid"
        bind_value(3) = scenario_id
        sql = getSQL(SQL_PATH_UA_ENDPRICE_ST, bind_variable, bind_value)
    Else
        sql = getSQL(SQL_PATH_UA_ENDPRICE, bind_variable, bind_value)
    End If
        
    With oCmd
        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql
        
        oRS.Open .Execute
    End With

    Dim spot_price As Double
    
    Dim i As Integer
    Do Until oRS.EOF
        spot_price = oRS(0)
        i = i + 1
        oRS.MoveNext
    Loop
    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
    get_spot_price = spot_price
    
End Function


Public Sub get_div_schedule(ByVal ul_code As String, tdate As String, source As String, scenario_id As String, div_schedule As Variant, oDB As adoDB.Connection)

    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(2) As String
    ReDim bind_value(2) As Variant
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":ul_code"
    
    If is_eval_shift_ua(ul_code) = True Then
        bind_value(1) = date2str(str2date(tdate) - 1)
    Else
        bind_value(1) = tdate
    End If
    
    If Left(ul_code, 3) = "KR7" Then
        bind_value(2) = Mid(ul_code, 4, 6)
    Else
        bind_value(2) = ul_code
    End If
    
    If SCENARIO_ENABLE = True Then
        ReDim Preserve bind_variable(3) As String
        ReDim Preserve bind_value(3) As Variant
        bind_variable(3) = ":scenarioid"
        bind_value(3) = scenario_id
    End If
            
    Select Case source
    Case "FRONT"
        
        If SCENARIO_ENABLE = True Then
            '시나리오 적용 완료까지 당분간 RM 값 사용
            sql = getSQL(SQL_PATH_DIV_SCHEDULE_ST, bind_variable, bind_value)
        Else
            sql = getSQL(SQL_PATH_DIV_SCHEDULE_FRONT, bind_variable, bind_value)
        End If
              
    Case "RM"
        If SCENARIO_ENABLE = True Then
            sql = getSQL(SQL_PATH_DIV_SCHEDULE_ST, bind_variable, bind_value)
        Else
            'sql = getSQL(SQL_PATH_DIV_SCHEDULE, bind_variable, bind_value)
            '시장 포워드 로직 개발까지 당분간 FRONT 값 사용 2019.3.27
            sql = getSQL(SQL_PATH_DIV_SCHEDULE_FRONT, bind_variable, bind_value)
        End If

    End Select
    
    With oCmd
        .ActiveConnection = oDB
        .CommandType = adCmdText
        .CommandText = sql
    
        oRS.Open .Execute
    End With
    
    Dim i As Integer
    Dim ofs As Integer
    i = 0
    ofs = 0
    
    '에러 방지를 위해 첫 배당락일과 평가일 사이에 0을 넣는다.
    If i = 0 Then
        div_schedule(1 + i, 1) = str2date(tdate)
        div_schedule(1 + i, 2) = 0
        ofs = 1
    End If
    
    Do Until oRS.EOF
        
        If str2date(oRS(0)) > str2date(tdate) And str2date(oRS(0)) <= str2date(tdate) + 365 * 3 Then '2024.07.22 금리 테너 범위를 넘어서는 배당 스케줄은 에러 발생되므로 배당 스케줄을 3년으로 제한.
            div_schedule(1 + ofs + i, 1) = str2date(oRS(0))
            div_schedule(1 + ofs + i, 2) = oRS(1)
            i = i + 1
        End If
    
        oRS.MoveNext
    
    Loop

    '조회된 배당 스케줄이 없을 경우, 에러 방지를 위해 추가로 dummy schedule에 0을 넣는다. 2024.07.19
    If i = 0 Then
        div_schedule(1 + ofs + i, 1) = str2date(tdate) + 1
        div_schedule(1 + ofs + i, 2) = 0
    End If
    
    oRS.Close
        
    Set oRS = Nothing
    Set oCmd = Nothing
    
End Sub

Public Function get_div_yield(ul_code As String, tdate As String, source As String, scenario_id As String, oDB As adoDB.Connection) As Double

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
    
    Dim div_yield As Double
    
    Select Case source
    Case "FRONT"
    
        sql = getSQL(SQL_PATH_DIV_YIELD, bind_variable, bind_value)
               
        With oCmd
            .ActiveConnection = oDB
            .CommandType = adCmdText
            .CommandText = sql

            oRS.Open .Execute
        End With

        Do Until oRS.EOF
            div_yield = oRS(0)
            oRS.MoveNext
        Loop
        oRS.Close
    
    Case "RM"
        div_yield = 0
    End Select
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
    get_div_yield = div_yield
        
End Function

Public Sub get_rf_curve(ByVal ccy_str As String, num_limit As Integer, tdate_str As String, source As String, scenario_id As String, rf_curve As Variant, oDB As adoDB.Connection)

    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    'USD Libor 산출중단 반영
    If get_ccy_idx(ccy_str) = ccy.USD And tdate_str >= "20230701" Then
        ccy_str = "USD_SOFR"
    End If
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(2) As String
    ReDim bind_value(2) As Variant
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":ccy"
    bind_value(1) = tdate_str
    bind_value(2) = ccy_str
    
    '금리시나리오는 DV01 계산 중 반영되므로 시나리오를 적용하지 않음
'    If SCENARIO_ENABLE = True Then
'        ReDim Preserve bind_variable(3) As String
'        ReDim Preserve bind_value(3) As Variant
'        bind_variable(3) = ":scenarioid"
'        bind_value(3) = scenario_id
'
'        sql = getSQL(SQL_PATH_RF_CURVE_ST, bind_variable, bind_value)
'    Else
        sql = getSQL(SQL_PATH_RF_CURVE, bind_variable, bind_value)
'    End If

    With oCmd
        .ActiveConnection = oDB
        .CommandType = adCmdText
        .CommandText = sql

        oRS.Open .Execute
    End With
    
    Dim i As Integer
    i = 1

    Do Until (oRS.EOF Or i = num_limit + 1)
        
        rf_curve(i, 1) = oRS(0)
        rf_curve(i, 2) = oRS(1)
        rf_curve(i, 3) = oRS(2)
        If i > 1 Then
            rf_curve(i, 4) = -Log(oRS(2)) / oRS(0) * 365
        End If
        
        i = i + 1
        oRS.MoveNext
    Loop
    
    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
End Sub

Public Function get_fx_vol(ua_ccy As String, base_ccy As String, tdate As String, source As String, scenario_id As String, oDB As adoDB.Connection) As Double

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
    bind_value(2) = ua_ccy & base_ccy

    If SCENARIO_ENABLE = True Then
        ReDim Preserve bind_variable(3) As String
        ReDim Preserve bind_value(3) As Variant
        bind_variable(3) = ":scenarioid"
        bind_value(3) = scenario_id
    End If
    
    Dim fx_vol As Double
    
    If get_ccy_idx(ua_ccy) = get_ccy_idx(base_ccy) Then
        fx_vol = 0
    Else
        Select Case source
        Case "FRONT"
            sql = getSQL(SQL_PATH_FX_VOL_FRONT, bind_variable, bind_value)
        Case "RM"
            If SCENARIO_ENABLE = True Then
                sql = getSQL(SQL_PATH_FX_VOL_ST, bind_variable, bind_value)
            Else
                sql = getSQL(SQL_PATH_FX_VOL, bind_variable, bind_value)
            End If
        End Select
        
        With oCmd
            .ActiveConnection = oDB
            .CommandType = adCmdText
            .CommandText = sql
    
            oRS.Open .Execute
        End With
            
        Do Until oRS.EOF
            fx_vol = oRS(1)
            oRS.MoveNext
        Loop
        
        oRS.Close
    End If
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
    get_fx_vol = fx_vol
            
End Function


Public Function get_corr(pair_code() As String, isLocMinCorr As Boolean, tdate_str As String, source As String, scenario_id As String, oDB As adoDB.Connection) As Double
    
    'local correlation 추가 2019.3.27
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(3) As String
    ReDim bind_value(3) As Variant
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":code1"
    bind_variable(3) = ":code2"
    bind_value(1) = tdate_str
    bind_value(2) = pair_code(0)
    bind_value(3) = pair_code(1)
    If SCENARIO_ENABLE = True Or (source = "RM" And isLocMinCorr = False) Then
        ReDim Preserve bind_variable(4) As String
        ReDim Preserve bind_value(4) As Variant
        bind_variable(4) = ":scenarioid"
        bind_value(4) = scenario_id
    End If
    
    Select Case source
    Case "FRONT"
        If isLocMinCorr Then
            sql = getSQL(SQL_PATH_MIN_CORR_FRONT, bind_variable, bind_value)
        Else
            sql = getSQL(SQL_PATH_CORR_FRONT, bind_variable, bind_value)
        End If
    Case "RM"
        If isLocMinCorr Then
            sql = getSQL(SQL_PATH_MIN_CORR_FRONT, bind_variable, bind_value)
        Else
            sql = getSQL(SQL_PATH_CORR, bind_variable, bind_value)
        End If
    End Select
                
    With oCmd
        .ActiveConnection = oDB
        .CommandType = adCmdText
        .CommandText = sql
        
        oRS.Open .Execute
        
    End With
    
    Dim corr As Double
    
    Do Until oRS.EOF
        corr = oRS(0)

        oRS.MoveNext
    Loop
    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
    get_corr = corr
    
End Function

'local correlation 추가 2019.3.27
Public Function get_lambda_neutral(ul_code As String, tdate As String, source As String, scenario_id As String, oDB As adoDB.Connection) As Double
    
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    Dim lambda_neutral As Double
    
'source, scenario_id 아직 적용 안 함
'    Select Case source
'    Case "FRONT"
'    Case "RM"
'    End Select
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(2) As String
    ReDim bind_value(2) As Variant
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":ul_code"
    bind_value(1) = tdate
    bind_value(2) = ul_code
    
    sql = getSQL(SQL_PATH_LC_LAMBDA_NEUTRAL, bind_variable, bind_value)
    
    With oCmd
       .ActiveConnection = oDB
       .CommandType = adCmdText
       .CommandText = sql
    
       oRS.Open .Execute
    End With

    Do Until oRS.EOF
        lambda_neutral = oRS(0)
        oRS.MoveNext
    Loop
    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
    get_lambda_neutral = lambda_neutral
    
End Function

Public Sub get_dcf_curve(rate_id As String, ccy_str As String, tdate_str As String, source As String, scenario_id As String, dcf_date As Date, dcf_curve_name As String, dcf_curve As Variant, oDB As adoDB.Connection)

    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(2) As String
    ReDim bind_value(2) As Variant
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":rateid"
    bind_value(1) = tdate_str
    bind_value(2) = rate_id

    If SCENARIO_ENABLE = True Then
        ReDim Preserve bind_variable(3) As String
        ReDim Preserve bind_value(3) As Variant
        bind_variable(3) = ":scenarioid"
        bind_value(3) = scenario_id
    End If
    
    'source, 시나리오 적용 안됨
    If SCENARIO_ENABLE = True Then
        Select Case get_ccy_idx(ccy_str)
        Case ccy.KRW
            'sql = getSQL(SQL_PATH_DCF_CURVE_KRW_ST, bind_variable, bind_value)
            sql = getSQL(SQL_PATH_DCF_CURVE_KRW, bind_variable, bind_value)
        Case ccy.USD
            'sql = getSQL(SQL_PATH_DCF_CURVE_USD_ST, bind_variable, bind_value)
            sql = getSQL(SQL_PATH_DCF_CURVE_USD, bind_variable, bind_value)
        End Select
    Else
        Select Case get_ccy_idx(ccy_str)
        Case ccy.KRW
            sql = getSQL(SQL_PATH_DCF_CURVE_KRW, bind_variable, bind_value)
        Case ccy.USD
            sql = getSQL(SQL_PATH_DCF_CURVE_USD, bind_variable, bind_value)
        End Select
    End If

    With oCmd
        .ActiveConnection = oDB
        .CommandType = adCmdText
        .CommandText = sql

        oRS.Open .Execute
    End With
    
    dcf_date = oRS(3)
    dcf_curve_name = oRS(4)
    
    Dim i As Integer
    i = 1
    dcf_curve(i, 1) = 0
    dcf_curve(i, 2) = str2date(tdate_str)
    dcf_curve(i, 3) = 1
    dcf_curve(i, 4) = 0
        
    i = 2
    Do Until oRS.EOF
        dcf_curve(i, 1) = oRS(0)
        dcf_curve(i, 2) = oRS(1)
        dcf_curve(i, 3) = Exp(-oRS(2) / 100 * oRS(0) / 365)
        dcf_curve(i, 4) = oRS(2) / 100
        
        i = i + 1
        oRS.MoveNext
    Loop
   
    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
End Sub

Public Sub get_surface_size(size_t As Integer, size_k As Integer, ul_code As String, vol_type As String, tdate As String, source As String, oDB As adoDB.Connection)
    
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
    
    If source = "FRONT" Then
        Select Case vol_type
        Case "Implied"
            sql = getSQL(SQL_PATH_IV_SURFACE_SIZE_FRONT, bind_variable, bind_value)
        Case "Local"
            sql = getSQL(SQL_PATH_LV_SURFACE_SIZE_FRONT, bind_variable, bind_value)
        End Select
    Else
        ReDim Preserve bind_variable(3) As String
        ReDim Preserve bind_value(3) As Variant
        bind_variable(3) = ":source"
        bind_value(3) = source
        Select Case vol_type
        Case "Implied"
            sql = getSQL(SQL_PATH_IV_SURFACE_SIZE, bind_variable, bind_value)
        Case "Local"
            sql = getSQL(SQL_PATH_LV_SURFACE_SIZE, bind_variable, bind_value)
        End Select
    End If
        
    With oCmd
        .ActiveConnection = oDB
        .CommandType = adCmdText
        .CommandText = sql
       
        oRS.Open .Execute
    End With
    
    Do Until oRS.EOF
        size_k = oRS(1)
        size_t = oRS(2)
        
        oRS.MoveNext
    Loop
    
    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
End Sub

Public Function get_vol_on_surface(targetDate As String, ua_code As String, ua_spot As Double, tau As Double, k As Double, voltype As String, source As String, Optional adoCon As adoDB.Connection) As Double

    If IsMissing(adoCon) = True Then
        Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    End If
    
    Dim adoRst As New adoDB.Recordset
    
    Dim sqlSelect As String
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(4) As String
    ReDim bind_value(4) As Variant
    bind_variable(1) = ":tdate"
    bind_variable(2) = ":ul_code"
    bind_variable(3) = ":tau"
    bind_variable(4) = ":k"
    bind_value(1) = targetDate
    bind_value(2) = ua_code
    bind_value(3) = tau
    bind_value(4) = k
    
    Select Case voltype
    Case "Implied"
        If source = "FRONT" Then
            sqlSelect = getSQL(SQL_PATH_IV_SURFACE_4PTS_FRONT, bind_variable, bind_value)
        Else
            sqlSelect = getSQL(SQL_PATH_IV_SURFACE_4PTS, bind_variable, bind_value)
        End If
    Case "Local"
        If source = "FRONT" Then
            sqlSelect = getSQL(SQL_PATH_LV_SURFACE_4PTS_FRONT, bind_variable, bind_value)
        Else
            sqlSelect = getSQL(SQL_PATH_LV_SURFACE_4PTS, bind_variable, bind_value)
        End If
    End Select
    
    Call adoRst.Open(sqlSelect, adoCon, adOpenStatic)
    
    Dim strike(3) As Double
    Dim ttm(3) As Double
    Dim vsqare(3) As Double
    Dim i As Integer
    
    Do While Not adoRst.EOF
        
        strike(i) = adoRst.Fields("STRIKE")
        ttm(i) = DateDiff("d", str2date(targetDate), str2date(adoRst.Fields("MATURITY_DATE"))) / 365
        vsqare(i) = adoRst.Fields("VOLATILITY") ^ 2
        
        adoRst.MoveNext
        i = i + 1
    Loop
        
    Dim p(1) As Double
    get_vol_on_surface = 0
    
    If Not adoRst.BOF Then
        If source = "FRONT" Then
            p(0) = k * ua_spot
        Else
            p(0) = k
        End If
        p(1) = tau
        
        '만기방향으로만 제곱해서 선형 보간
        get_vol_on_surface = Sqr(bilin(strike, ttm, vsqare, p))
    End If
    
    adoRst.Close
    
    If IsMissing(adoCon) = True Then
        Call disconnectDB(adoCon)
    End If
    
End Function

Public Function get_fx(against_ccy As String, base_ccy As String, tdate As Date, Optional adoCon As adoDB.Connection, Optional sceanrio_id As String = "0") As Double

    If IsMissing(adoCon) = True Then
        Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    End If
    
    Dim fxcode As String
    fxcode = "FX" & against_ccy & base_ccy
    
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    If sceanrio_id = "0" Then
        sql = "select get_fxrate('" & date2str(tdate) & "','" & fxcode & "') from dual"
    Else
        sql = "select endprice from RCS.PML_FX_DATA_ST where tdate='" & date2str(tdate) & "' and scenarioid ='" & sceanrio_id & "' and code='" & fxcode & "'"
    End If
    
    With oCmd
    
        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql
        
        oRS.Open .Execute
    
    End With
    
    Dim rtn As Double
    Do Until oRS.EOF
        rtn = oRS(0)
        oRS.MoveNext
    Loop
    
    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
    If IsMissing(adoCon) = True Then
        Call disconnectDB(adoCon)
    End If
    
    get_fx = rtn
            
End Function

Public Function get_spot(ua_code As String, tdate As Date, isDelayed As Boolean, Optional adoCon As adoDB.Connection) As Double

    If IsMissing(adoCon) = True Then
        Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    End If
    
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    Dim sql As String
    
    If isDelayed = True Then
        sql = " select code, endprice from ras.if_stock_data where code= '" + ua_code + "' and tdate='" + date2str(tdate) + "' union select indexid, endprice from ras.if_index_data where indexid='" + ua_code + "' and tdate='" + date2str(tdate) + "'"
    Else
        sql = " select code, endprice from ras.if_stock_data where code= '" + ua_code + "' and tdate='" + date2str(tdate) + "' union select code, price from ras.bl_quanto_base where code='" + ua_code + "' and tdate='" + date2str(tdate) + "'"
    End If
    
    With oCmd
        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql
       
        oRS.Open .Execute
    End With

    Dim rtn As Double
    rtn = 0
    
    Do Until oRS.EOF
        rtn = oRS(1)
        oRS.MoveNext
    Loop
    
    oRS.Close
    
    Set oRS = Nothing
    Set oCmd = Nothing
    
    If IsMissing(adoCon) = True Then
        Call disconnectDB(adoCon)
    End If
    
    get_spot = rtn

End Function

Public Function is_eval_shift_ua(ua_code As String) As Boolean

    Dim rtn As Boolean
    rtn = False
    
    Dim enum_ua As Variant
    
    For Each enum_ua In eval_shift_ua
        If enum_ua = get_ua_idx(ua_code) Then
            rtn = True
            Exit For
        End If
    Next
    
    is_eval_shift_ua = rtn

End Function

Public Function is_flatvol_ua(ua_code As String) As Boolean

    Dim rtn As Boolean
    rtn = False
    
    Dim enum_ua As Variant
    
    For Each enum_ua In FLAT_VOL_UA
        If enum_ua = get_ua_idx(ua_code) Then
            rtn = True
            Exit For
        End If
    Next
    
    is_flatvol_ua = rtn

End Function

Public Function chk_vol_index(ul_code As String, ByRef base_ul_code As String) As Boolean

    Dim rtn As Boolean
    
    base_ul_code = ul_code
    
    Dim adoCon As New adoDB.Connection
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(1) As String
    ReDim bind_value(1) As Variant
    bind_variable(1) = ":code"
    bind_value(1) = ul_code
    
    Dim sql As String
    sql = getSQL(SQL_PATH_CHK_VOL_INDEX, bind_variable, bind_value)

    With oCmd

        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql

        oRS.Open .Execute

    End With
    
    Do Until oRS.EOF
        
        If IsNull(oRS(0)) = False Then
            base_ul_code = oRS(0)
        End If
        oRS.MoveNext

    Loop
    
    oRS.Close
    
    Call disconnectDB(adoCon)
    
    If base_ul_code = ul_code Then
        rtn = False
    Else
        rtn = True
    End If
    
    chk_vol_index = rtn
        
End Function