Option ExplicitnnPublic Function str2date(tdate As String) As Daten'String YYYYMMDD -> Date YYYY-MM-DDnn    str2date = CDate(Left(tdate, 4) + "-" + Mid(tdate, 5, 2) + "-" + Right(tdate, 2))nnEnd FunctionnnPublic Function date2str(tdate As Date) As Stringn'Date YYYY-MM-DD -> String YYYYMMDDnn    Dim str_tdate As Stringn    str_tdate = CStr(tdate)nn    date2str = Left(str_tdate, 4) & Mid(str_tdate, 6, 2) & Right(str_tdate, 2)nnEnd FunctionnnPublic Function ul_ofs(ul_code As String) As Integernn    Dim ofs As Integern    n    Select Case ul_coden    Case "KOSPI200": ofs = 1n    Case "HSCEI": ofs = 2n    Case "HSI": ofs = 3n    Case "NKY": ofs = 4n    Case "SX5E": ofs = 5n    Case "SPX": ofs = 6n    Case "KR7005380001": ofs = 7 '현대차n    Case "KR7005930003": ofs = 8 '삼성전자n    Case "KR7000030007": ofs = 9 '우리은행n    Case "KR7028260008": ofs = 10 '삼성물산n    Case "KR7105560007": ofs = 11 'KB금융n    Case "KR7035420009": ofs = 12 'NAVERn    Case "KR7018260000": ofs = 13 '삼성SDSn    Case "KR7005490008": ofs = 14 'POSCOn    Case "KR7034220004": ofs = 15 'LG디스플레이n    End Selectnn    ul_ofs = ofsnnEnd FunctionnnPublic Function get_ua_idx(ua_code As String) As Integernn    Dim rtn As Integern    n    Select Case ua_coden    Case "HSCEI": rtn = ua.HSCEIn    Case "HSI": rtn = ua.HSIn    Case "SX5E": rtn = ua.SX5En    Case "SPX": rtn = ua.SPXn    Case "NKY": rtn = ua.NKYn    Case "KOSPI200": rtn = ua.KOSPI200n    Case "KRD020021147": rtn = ua.KRD020021147n    Case "KR7005380001": rtn = ua.KR7005380001n    Case "KR7005930003": rtn = ua.KR7005930003n    Case "KR7000030007": rtn = ua.KR7000030007n    Case "KR7028260008": rtn = ua.KR7028260008n    Case "KR7105560007": rtn = ua.KR7105560007n    Case "KR7035420009": rtn = ua.KR7035420009n    Case "KR7018260000": rtn = ua.KR7018260000n    Case "KR7005490008": rtn = ua.KR7005490008n    Case "KR7034220004": rtn = ua.KR7034220004n    Case "TSLA", "SPTESLA", "US88160R1014": rtn = ua.TSLA '2024.04.29 SP code, ISIN code 추가n    Case "NVDA", "SPNVDA", "US67066G1040": rtn = ua.NVDA '2024.06.19 SP code, ISIN code 추가n    End Selectn    n    get_ua_idx = rtnn    nEnd FunctionnnPublic Function get_ua_code(idx As Integer, Optional code_type As Integer = UA_CODE_TYPE.BLP) As Stringnn    Dim rtn As Stringn    n    Select Case idxn    Case ua.HSCEI: rtn = "HSCEI"n    Case ua.HSI: rtn = "HSI"n    Case ua.SX5E: rtn = "SX5E"n    Case ua.SPX: rtn = "SPX"n    Case ua.NKY: rtn = "NKY"n    Case ua.KOSPI200: rtn = "KOSPI200"n    Case ua.KRD020021147: rtn = "KRD020021147"n    Case ua.KR7005380001: rtn = "KR7005380001"n    Case ua.KR7005930003: rtn = "KR7005930003"n    Case ua.KR7000030007: rtn = "KR7000030007"n    Case ua.KR7028260008: rtn = "KR7028260008"n    Case ua.KR7105560007: rtn = "KR7105560007"n    Case ua.KR7035420009: rtn = "KR7035420009"n    Case ua.KR7018260000: rtn = "KR7018260000"n    Case ua.KR7005490008: rtn = "KR7005490008"n    Case ua.KR7034220004: rtn = "KR7034220004"n    Case ua.TSLA:n        Select Case code_typen        Case UA_CODE_TYPE.BLP: rtn = "TSLA"n        Case UA_CODE_TYPE.SP: rtn = "SPTESLA"n        Case UA_CODE_TYPE.ISIN: rtn = "US88160R1014"n        End Selectn    Case ua.NVDA:n        Select Case code_typen        Case UA_CODE_TYPE.BLP: rtn = "NVDA"n        Case UA_CODE_TYPE.SP: rtn = "SPNVDA"n        Case UA_CODE_TYPE.ISIN: rtn = "US67066G1040"n        End Selectn    End Selectn    n    get_ua_code = rtnn    nEnd FunctionnnPublic Function get_ua_name(ua_code As String) As Stringnn    Dim rtn As Stringn    n    Select Case get_ua_idx(ua_code)n    Case ua.HSCEI: rtn = "HSCEI"n    Case ua.HSI: rtn = "HSI"n    Case ua.SX5E: rtn = "EuroStoxx50"n    Case ua.SPX: rtn = "S&P500"n    Case ua.NKY: rtn = "NIKKEI225"n    Case ua.KOSPI200: rtn = "KOSPI200"n    Case ua.KRD020021147: rtn = "KOSPI200레버리지"n    Case ua.KR7005380001: rtn = "현대차"n    Case ua.KR7005930003: rtn = "삼성전자"n    Case ua.KR7000030007: rtn = "우리은행"n    Case ua.KR7028260008: rtn = "삼성물산"n    Case ua.KR7105560007: rtn = "KB금융"n    Case ua.KR7035420009: rtn = "NAVER"n    Case ua.KR7018260000: rtn = "삼성SDS"n    Case ua.KR7005490008: rtn = "POSCO"n    Case ua.KR7034220004: rtn = "LG디스플레이"n    Case ua.TSLA: rtn = "TESLA"n    Case ua.NVDA: rtn = "NVIDIA"n    End Selectn    n    get_ua_name = rtnn    nEnd FunctionnnPublic Function get_ua_currency(ua_code As String) As Stringnn    Dim rtn As Stringn    n    Select Case get_ua_idx(ua_code)n    Case ua.HSCEI, ua.HSI: rtn = "HKD"n    Case ua.SX5E: rtn = "EUR"n    Case ua.SPX, ua.TSLA, ua.NVDA: rtn = "USD"n    Case ua.NKY: rtn = "JPY"n    Case Else: rtn = "KRW"n    End Selectn    n    get_ua_currency = rtnn    nEnd FunctionnnPublic Function get_ccy_code(ccy_idx As Integer) As Stringnn    Dim rtn As Stringn    n    Select Case ccy_idxn    Case ccy.EUR: rtn = "EUR"n    Case ccy.HKD: rtn = "HKD"n    Case ccy.JPY: rtn = "JPY"n    Case ccy.KRW: rtn = "KRW"n    Case ccy.USD: rtn = "USD"n    End Selectn    n    get_ccy_code = rtnn    nEnd FunctionnnPublic Function is_active_ua(ua_code As String) As Booleannn    Dim rtn As Booleann    n    Select Case get_ua_idx(ua_code)n    Case ua.HSCEI, ua.HSI, ua.KOSPI200, ua.NKY, ua.SX5E, ua.SPX, ua.TSLA, ua.NVDA: rtn = Truen    'Case ua.HSCEI, ua.HSI, ua.KOSPI200, ua.NKY, ua.SX5E, ua.SPX: rtn = Truen    Case Else: rtn = Falsen    End Selectn        n    is_active_ua = rtnn    nEnd FunctionnnPublic Function is_active_ccy(ccy_code As String) As Booleannn    Dim rtn As Booleann    n    Select Case get_ccy_idx(ccy_code)n    Case ccy.KRW, ccy.USD, ccy.EUR, ccy.JPY, ccy.HKD: rtn = Truen    Case Else: rtn = Falsen    End Selectn    n    is_active_ccy = rtnn    nEnd FunctionnnPublic Function get_ccy_idx(ccy_code As String) As Integernn    Dim rtn As Integern    n    Select Case ccy_coden    Case "KRW": rtn = ccy.KRWn    Case "USD": rtn = ccy.USDn    Case "EUR": rtn = ccy.EURn    Case "JPY": rtn = ccy.JPYn    Case "HKD": rtn = ccy.HKDn    End Selectn    n    get_ccy_idx = rtnn    nEnd FunctionnnPublic Function get_rf_num_term(ccy_code As String) As Integernn    Dim rtn As Integern    n    Select Case get_ccy_idx(ccy_code)n    Case ccy.HKD: rtn = RF_NUM_TERM.HKDn    Case ccy.EUR: rtn = RF_NUM_TERM.EURn    Case ccy.USD: rtn = RF_NUM_TERM.USDn    Case ccy.JPY: rtn = RF_NUM_TERM.JPYn    Case ccy.KRW: rtn = RF_NUM_TERM.KRWn    End Selectn    n    get_rf_num_term = rtnn    nEnd FunctionnnPublic Function get_dcf_ccy(idx As Integer) As Stringnn    Dim rtn As Stringn    n    Select Case idxn    Case DCF.KRW: rtn = "KRW"n    Case DCF.USD: rtn = "USD"n    End Selectn    n    get_dcf_ccy = rtnn    nEnd FunctionnnPublic Function get_dcf_idx(ccy_code As String) As Integernn    Dim rtn As Stringn    n    Select Case get_ccy_idx(ccy_code)n    Case ccy.KRW: rtn = DCF.KRWn    Case ccy.USD: rtn = DCF.USDn    End Selectn    n    get_dcf_idx = rtnn    nEnd FunctionnnPublic Function get_spot_price(ul_code As String, tdate As String, Optional scenario_id As String, Optional adoCon As adoDB.Connection) As Doublenn    If adoCon Is Nothing Thenn        Set adoCon = New adoDB.Connectionn        Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)n    End Ifnn    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringn   n    Dim bind_variable() As Stringn    ReDim bind_variable(2)n    bind_variable(1) = ":tdate"n    bind_variable(2) = ":code"nn    Dim bind_value() As Variantn    ReDim bind_value(2)n    bind_value(1) = tdaten    bind_value(2) = ul_coden    n    If SCENARIO_ENABLE = True Thenn        ReDim Preserve bind_variable(3) As Stringn        ReDim Preserve bind_value(3) As Variantn        bind_variable(3) = ":scenarioid"n        bind_value(3) = scenario_idn        sql = getSQL(SQL_PATH_UA_ENDPRICE_ST, bind_variable, bind_value)n    Elsen        sql = getSQL(SQL_PATH_UA_ENDPRICE, bind_variable, bind_value)n    End Ifn        n    With oCmdn        .ActiveConnection = adoConn        .CommandType = adCmdTextn        .CommandText = sqln        n        oRS.Open .Executen    End Withnn    Dim spot_price As Doublen    n    Dim i As Integern    Do Until oRS.EOFn        spot_price = oRS(0)n        i = i + 1n        oRS.MoveNextn    Loopn    oRS.Closen    n    Set oRS = Nothingn    Set oCmd = Nothingn    n    get_spot_price = spot_pricen    nEnd FunctionnnnPublic Sub get_div_schedule(ByVal ul_code As String, tdate As String, source As String, scenario_id As String, div_schedule As Variant, oDB As adoDB.Connection)nn    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringn    n    Dim bind_variable() As Stringn    Dim bind_value() As Variantn    ReDim bind_variable(2) As Stringn    ReDim bind_value(2) As Variantn    bind_variable(1) = ":tdate"n    bind_variable(2) = ":ul_code"n    n    If is_eval_shift_ua(ul_code) = True Thenn        bind_value(1) = date2str(str2date(tdate) - 1)n    Elsen        bind_value(1) = tdaten    End Ifn    n    If Left(ul_code, 3) = "KR7" Thenn        bind_value(2) = Mid(ul_code, 4, 6)n    Elsen        bind_value(2) = ul_coden    End Ifn    n    If SCENARIO_ENABLE = True Thenn        ReDim Preserve bind_variable(3) As Stringn        ReDim Preserve bind_value(3) As Variantn        bind_variable(3) = ":scenarioid"n        bind_value(3) = scenario_idn    End Ifn            n    Select Case sourcen    Case "FRONT"n        n        If SCENARIO_ENABLE = True Thenn            '시나리오 적용 완료까지 당분간 RM 값 사용n            sql = getSQL(SQL_PATH_DIV_SCHEDULE_ST, bind_variable, bind_value)n        Elsen            sql = getSQL(SQL_PATH_DIV_SCHEDULE_FRONT, bind_variable, bind_value)n        End Ifn              n    Case "RM"n        If SCENARIO_ENABLE = True Thenn            sql = getSQL(SQL_PATH_DIV_SCHEDULE_ST, bind_variable, bind_value)n        Elsen            'sql = getSQL(SQL_PATH_DIV_SCHEDULE, bind_variable, bind_value)n            '시장 포워드 로직 개발까지 당분간 FRONT 값 사용 2019.3.27n            sql = getSQL(SQL_PATH_DIV_SCHEDULE_FRONT, bind_variable, bind_value)n        End Ifnn    End Selectn    n    With oCmdn        .ActiveConnection = oDBn        .CommandType = adCmdTextn        .CommandText = sqln    n        oRS.Open .Executen    End Withn    n    Dim i As Integern    Dim ofs As Integern    i = 0n    ofs = 0n    n    '에러 방지를 위해 첫 배당락일과 평가일 사이에 0을 넣는다.n    If i = 0 Thenn        div_schedule(1 + i, 1) = str2date(tdate)n        div_schedule(1 + i, 2) = 0n        ofs = 1n    End Ifn    n    Do Until oRS.EOFn        n        If str2date(oRS(0)) > str2date(tdate) And str2date(oRS(0)) <= str2date(tdate) + 365 * 3 Then '2024.07.22 금리 테너 범위를 넘어서는 배당 스케줄은 에러 발생되므로 배당 스케줄을 3년으로 제한.n            div_schedule(1 + ofs + i, 1) = str2date(oRS(0))n            div_schedule(1 + ofs + i, 2) = oRS(1)n            i = i + 1n        End Ifn    n        oRS.MoveNextn    n    Loopnn    '조회된 배당 스케줄이 없을 경우, 에러 방지를 위해 추가로 dummy schedule에 0을 넣는다. 2024.07.19n    If i = 0 Thenn        div_schedule(1 + ofs + i, 1) = str2date(tdate) + 1n        div_schedule(1 + ofs + i, 2) = 0n    End Ifn    n    oRS.Closen        n    Set oRS = Nothingn    Set oCmd = Nothingn    nEnd SubnnPublic Function get_div_yield(ul_code As String, tdate As String, source As String, scenario_id As String, oDB As adoDB.Connection) As Doublenn    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringn    n    Dim bind_variable() As Stringn    Dim bind_value() As Variantn    ReDim bind_variable(2) As Stringn    ReDim bind_value(2) As Variantn    bind_variable(1) = ":tdate"n    bind_variable(2) = ":ul_code"n    bind_value(1) = tdaten    bind_value(2) = ul_coden    n    Dim div_yield As Doublen    n    Select Case sourcen    Case "FRONT"n    n        sql = getSQL(SQL_PATH_DIV_YIELD, bind_variable, bind_value)n               n        With oCmdn            .ActiveConnection = oDBn            .CommandType = adCmdTextn            .CommandText = sqlnn            oRS.Open .Executen        End Withnn        Do Until oRS.EOFn            div_yield = oRS(0)n            oRS.MoveNextn        Loopn        oRS.Closen    n    Case "RM"n        div_yield = 0n    End Selectn    n    Set oRS = Nothingn    Set oCmd = Nothingn    n    get_div_yield = div_yieldn        nEnd FunctionnnPublic Sub get_rf_curve(ByVal ccy_str As String, num_limit As Integer, tdate_str As String, source As String, scenario_id As String, rf_curve As Variant, oDB As adoDB.Connection)nn    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringn    n    'USD Libor 산출중단 반영n    If get_ccy_idx(ccy_str) = ccy.USD And tdate_str >= "20230701" Thenn        ccy_str = "USD_SOFR"n    End Ifn    n    Dim bind_variable() As Stringn    Dim bind_value() As Variantn    ReDim bind_variable(2) As Stringn    ReDim bind_value(2) As Variantn    bind_variable(1) = ":tdate"n    bind_variable(2) = ":ccy"n    bind_value(1) = tdate_strn    bind_value(2) = ccy_strn    n    '금리시나리오는 DV01 계산 중 반영되므로 시나리오를 적용하지 않음n'    If SCENARIO_ENABLE = True Thenn'        ReDim Preserve bind_variable(3) As Stringn'        ReDim Preserve bind_value(3) As Variantn'        bind_variable(3) = ":scenarioid"n'        bind_value(3) = scenario_idn'n'        sql = getSQL(SQL_PATH_RF_CURVE_ST, bind_variable, bind_value)n'    Elsen        sql = getSQL(SQL_PATH_RF_CURVE, bind_variable, bind_value)n'    End Ifnn    With oCmdn        .ActiveConnection = oDBn        .CommandType = adCmdTextn        .CommandText = sqlnn        oRS.Open .Executen    End Withn    n    Dim i As Integern    i = 1nn    Do Until (oRS.EOF Or i = num_limit + 1)n        n        rf_curve(i, 1) = oRS(0)n        rf_curve(i, 2) = oRS(1)n        rf_curve(i, 3) = oRS(2)n        If i > 1 Thenn            rf_curve(i, 4) = -Log(oRS(2)) / oRS(0) * 365n        End Ifn        n        i = i + 1n        oRS.MoveNextn    Loopn    n    oRS.Closen    n    Set oRS = Nothingn    Set oCmd = Nothingn    nEnd SubnnPublic Function get_fx_vol(ua_ccy As String, base_ccy As String, tdate As String, source As String, scenario_id As String, oDB As adoDB.Connection) As Doublenn    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringnn    Dim bind_variable() As Stringn    Dim bind_value() As Variantn    ReDim bind_variable(2) As Stringn    ReDim bind_value(2) As Variantn    bind_variable(1) = ":tdate"n    bind_variable(2) = ":code"n    bind_value(1) = tdaten    bind_value(2) = ua_ccy & base_ccynn    If SCENARIO_ENABLE = True Thenn        ReDim Preserve bind_variable(3) As Stringn        ReDim Preserve bind_value(3) As Variantn        bind_variable(3) = ":scenarioid"n        bind_value(3) = scenario_idn    End Ifn    n    Dim fx_vol As Doublen    n    If get_ccy_idx(ua_ccy) = get_ccy_idx(base_ccy) Thenn        fx_vol = 0n    Elsen        Select Case sourcen        Case "FRONT"n            sql = getSQL(SQL_PATH_FX_VOL_FRONT, bind_variable, bind_value)n        Case "RM"n            If SCENARIO_ENABLE = True Thenn                sql = getSQL(SQL_PATH_FX_VOL_ST, bind_variable, bind_value)n            Elsen                sql = getSQL(SQL_PATH_FX_VOL, bind_variable, bind_value)n            End Ifn        End Selectn        n        With oCmdn            .ActiveConnection = oDBn            .CommandType = adCmdTextn            .CommandText = sqln    n            oRS.Open .Executen        End Withn            n        Do Until oRS.EOFn            fx_vol = oRS(1)n            oRS.MoveNextn        Loopn        n        oRS.Closen    End Ifn    n    Set oRS = Nothingn    Set oCmd = Nothingn    n    get_fx_vol = fx_voln            nEnd FunctionnnnPublic Function get_corr(pair_code() As String, isLocMinCorr As Boolean, tdate_str As String, source As String, scenario_id As String, oDB As adoDB.Connection) As Doublen    n    'local correlation 추가 2019.3.27n    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringn    n    Dim bind_variable() As Stringn    Dim bind_value() As Variantn    ReDim bind_variable(3) As Stringn    ReDim bind_value(3) As Variantn    bind_variable(1) = ":tdate"n    bind_variable(2) = ":code1"n    bind_variable(3) = ":code2"n    bind_value(1) = tdate_strn    bind_value(2) = pair_code(0)n    bind_value(3) = pair_code(1)n    If SCENARIO_ENABLE = True Or (source = "RM" And isLocMinCorr = False) Thenn        ReDim Preserve bind_variable(4) As Stringn        ReDim Preserve bind_value(4) As Variantn        bind_variable(4) = ":scenarioid"n        bind_value(4) = scenario_idn    End Ifn    n    Select Case sourcen    Case "FRONT"n        If isLocMinCorr Thenn            sql = getSQL(SQL_PATH_MIN_CORR_FRONT, bind_variable, bind_value)n        Elsen            sql = getSQL(SQL_PATH_CORR_FRONT, bind_variable, bind_value)n        End Ifn    Case "RM"n        If isLocMinCorr Thenn            sql = getSQL(SQL_PATH_MIN_CORR_FRONT, bind_variable, bind_value)n        Elsen            sql = getSQL(SQL_PATH_CORR, bind_variable, bind_value)n        End Ifn    End Selectn                n    With oCmdn        .ActiveConnection = oDBn        .CommandType = adCmdTextn        .CommandText = sqln        n        oRS.Open .Executen        n    End Withn    n    Dim corr As Doublen    n    Do Until oRS.EOFn        corr = oRS(0)nn        oRS.MoveNextn    Loopn    oRS.Closen    n    Set oRS = Nothingn    Set oCmd = Nothingn    n    get_corr = corrn    nEnd Functionnn'local correlation 추가 2019.3.27nPublic Function get_lambda_neutral(ul_code As String, tdate As String, source As String, scenario_id As String, oDB As adoDB.Connection) As Doublen    n    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringn    n    Dim lambda_neutral As Doublen    n'source, scenario_id 아직 적용 안 함n'    Select Case sourcen'    Case "FRONT"n'    Case "RM"n'    End Selectn    Dim bind_variable() As Stringn    Dim bind_value() As Variantn    ReDim bind_variable(2) As Stringn    ReDim bind_value(2) As Variantn    bind_variable(1) = ":tdate"n    bind_variable(2) = ":ul_code"n    bind_value(1) = tdaten    bind_value(2) = ul_coden    n    sql = getSQL(SQL_PATH_LC_LAMBDA_NEUTRAL, bind_variable, bind_value)n    n    With oCmdn       .ActiveConnection = oDBn       .CommandType = adCmdTextn       .CommandText = sqln    n       oRS.Open .Executen    End Withnn    Do Until oRS.EOFn        lambda_neutral = oRS(0)n        oRS.MoveNextn    Loopn    oRS.Closen    n    Set oRS = Nothingn    Set oCmd = Nothingn    n    get_lambda_neutral = lambda_neutraln    nEnd FunctionnnPublic Sub get_dcf_curve(rate_id As String, ccy_str As String, tdate_str As String, source As String, scenario_id As String, dcf_date As Date, dcf_curve_name As String, dcf_curve As Variant, oDB As adoDB.Connection)nn    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringn    n    Dim bind_variable() As Stringn    Dim bind_value() As Variantn    ReDim bind_variable(2) As Stringn    ReDim bind_value(2) As Variantn    bind_variable(1) = ":tdate"n    bind_variable(2) = ":rateid"n    bind_value(1) = tdate_strn    bind_value(2) = rate_idnn    If SCENARIO_ENABLE = True Thenn        ReDim Preserve bind_variable(3) As Stringn        ReDim Preserve bind_value(3) As Variantn        bind_variable(3) = ":scenarioid"n        bind_value(3) = scenario_idn    End Ifn    n    'source, 시나리오 적용 안됨n    If SCENARIO_ENABLE = True Thenn        Select Case get_ccy_idx(ccy_str)n        Case ccy.KRWn            'sql = getSQL(SQL_PATH_DCF_CURVE_KRW_ST, bind_variable, bind_value)n            sql = getSQL(SQL_PATH_DCF_CURVE_KRW, bind_variable, bind_value)n        Case ccy.USDn            'sql = getSQL(SQL_PATH_DCF_CURVE_USD_ST, bind_variable, bind_value)n            sql = getSQL(SQL_PATH_DCF_CURVE_USD, bind_variable, bind_value)n        End Selectn    Elsen        Select Case get_ccy_idx(ccy_str)n        Case ccy.KRWn            sql = getSQL(SQL_PATH_DCF_CURVE_KRW, bind_variable, bind_value)n        Case ccy.USDn            sql = getSQL(SQL_PATH_DCF_CURVE_USD, bind_variable, bind_value)n        End Selectn    End Ifnn    With oCmdn        .ActiveConnection = oDBn        .CommandType = adCmdTextn        .CommandText = sqlnn        oRS.Open .Executen    End Withn    n    dcf_date = oRS(3)n    dcf_curve_name = oRS(4)n    n    Dim i As Integern    i = 1n    dcf_curve(i, 1) = 0n    dcf_curve(i, 2) = str2date(tdate_str)n    dcf_curve(i, 3) = 1n    dcf_curve(i, 4) = 0n        n    i = 2n    Do Until oRS.EOFn        dcf_curve(i, 1) = oRS(0)n        dcf_curve(i, 2) = oRS(1)n        dcf_curve(i, 3) = Exp(-oRS(2) / 100 * oRS(0) / 365)n        dcf_curve(i, 4) = oRS(2) / 100n        n        i = i + 1n        oRS.MoveNextn    Loopn   n    oRS.Closen    n    Set oRS = Nothingn    Set oCmd = Nothingn    nEnd SubnnPublic Sub get_surface_size(size_t As Integer, size_k As Integer, ul_code As String, vol_type As String, tdate As String, source As String, oDB As adoDB.Connection)n    n    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringn    n    Dim bind_variable() As Stringn    Dim bind_value() As Variantn    ReDim bind_variable(2) As Stringn    ReDim bind_value(2) As Variantn    bind_variable(1) = ":tdate"n    bind_variable(2) = ":ul_code"n    bind_value(1) = tdaten    bind_value(2) = ul_coden    n    If source = "FRONT" Thenn        Select Case vol_typen        Case "Implied"n            sql = getSQL(SQL_PATH_IV_SURFACE_SIZE_FRONT, bind_variable, bind_value)n        Case "Local"n            sql = getSQL(SQL_PATH_LV_SURFACE_SIZE_FRONT, bind_variable, bind_value)n        End Selectn    Elsen        ReDim Preserve bind_variable(3) As Stringn        ReDim Preserve bind_value(3) As Variantn        bind_variable(3) = ":source"n        bind_value(3) = sourcen        Select Case vol_typen        Case "Implied"n            sql = getSQL(SQL_PATH_IV_SURFACE_SIZE, bind_variable, bind_value)n        Case "Local"n            sql = getSQL(SQL_PATH_LV_SURFACE_SIZE, bind_variable, bind_value)n        End Selectn    End Ifn        n    With oCmdn        .ActiveConnection = oDBn        .CommandType = adCmdTextn        .CommandText = sqln       n        oRS.Open .Executen    End Withn    n    Do Until oRS.EOFn        size_k = oRS(1)n        size_t = oRS(2)n        n        oRS.MoveNextn    Loopn    n    oRS.Closen    n    Set oRS = Nothingn    Set oCmd = Nothingn    nEnd SubnnPublic Function get_vol_on_surface(targetDate As String, ua_code As String, ua_spot As Double, tau As Double, k As Double, voltype As String, source As String, Optional adoCon As adoDB.Connection) As Doublenn    If IsMissing(adoCon) = True Thenn        Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)n    End Ifn    n    Dim adoRst As New adoDB.Recordsetn    n    Dim sqlSelect As Stringn    n    Dim bind_variable() As Stringn    Dim bind_value() As Variantn    ReDim bind_variable(4) As Stringn    ReDim bind_value(4) As Variantn    bind_variable(1) = ":tdate"n    bind_variable(2) = ":ul_code"n    bind_variable(3) = ":tau"n    bind_variable(4) = ":k"n    bind_value(1) = targetDaten    bind_value(2) = ua_coden    bind_value(3) = taun    bind_value(4) = kn    n    Select Case voltypen    Case "Implied"n        If source = "FRONT" Thenn            sqlSelect = getSQL(SQL_PATH_IV_SURFACE_4PTS_FRONT, bind_variable, bind_value)n        Elsen            sqlSelect = getSQL(SQL_PATH_IV_SURFACE_4PTS, bind_variable, bind_value)n        End Ifn    Case "Local"n        If source = "FRONT" Thenn            sqlSelect = getSQL(SQL_PATH_LV_SURFACE_4PTS_FRONT, bind_variable, bind_value)n        Elsen            sqlSelect = getSQL(SQL_PATH_LV_SURFACE_4PTS, bind_variable, bind_value)n        End Ifn    End Selectn    n    Call adoRst.Open(sqlSelect, adoCon, adOpenStatic)n    n    Dim strike(3) As Doublen    Dim ttm(3) As Doublen    Dim vsqare(3) As Doublen    Dim i As Integern    n    Do While Not adoRst.EOFn        n        strike(i) = adoRst.Fields("STRIKE")n        ttm(i) = DateDiff("d", str2date(targetDate), str2date(adoRst.Fields("MATURITY_DATE"))) / 365n        vsqare(i) = adoRst.Fields("VOLATILITY") ^ 2n        n        adoRst.MoveNextn        i = i + 1n    Loopn        n    Dim p(1) As Doublen    get_vol_on_surface = 0n    n    If Not adoRst.BOF Thenn        If source = "FRONT" Thenn            p(0) = k * ua_spotn        Elsen            p(0) = kn        End Ifn        p(1) = taun        n        '만기방향으로만 제곱해서 선형 보간n        get_vol_on_surface = Sqr(bilin(strike, ttm, vsqare, p))n    End Ifn    n    adoRst.Closen    n    If IsMissing(adoCon) = True Thenn        Call disconnectDB(adoCon)n    End Ifn    nEnd FunctionnnPublic Function get_fx(against_ccy As String, base_ccy As String, tdate As Date, Optional adoCon As adoDB.Connection, Optional sceanrio_id As String = "0") As Doublenn    If IsMissing(adoCon) = True Thenn        Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)n    End Ifn    n    Dim fxcode As Stringn    fxcode = "FX" & against_ccy & base_ccyn    n    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringn    n    If sceanrio_id = "0" Thenn        sql = "select get_fxrate('" & date2str(tdate) & "','" & fxcode & "') from dual"n    Elsen        sql = "select endprice from RCS.PML_FX_DATA_ST where tdate='" & date2str(tdate) & "' and scenarioid ='" & sceanrio_id & "' and code='" & fxcode & "'"n    End Ifn    n    With oCmdn    n        .ActiveConnection = adoConn        .CommandType = adCmdTextn        .CommandText = sqln        n        oRS.Open .Executen    n    End Withn    n    Dim rtn As Doublen    Do Until oRS.EOFn        rtn = oRS(0)n        oRS.MoveNextn    Loopn    n    oRS.Closen    n    Set oRS = Nothingn    Set oCmd = Nothingn    n    If IsMissing(adoCon) = True Thenn        Call disconnectDB(adoCon)n    End Ifn    n    get_fx = rtnn            nEnd FunctionnnPublic Function get_spot(ua_code As String, tdate As Date, isDelayed As Boolean, Optional adoCon As adoDB.Connection) As Doublenn    If IsMissing(adoCon) = True Thenn        Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)n    End Ifn    n    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    Dim sql As Stringn    n    If isDelayed = True Thenn        sql = " select code, endprice from ras.if_stock_data where code= '" + ua_code + "' and tdate='" + date2str(tdate) + "' union select indexid, endprice from ras.if_index_data where indexid='" + ua_code + "' and tdate='" + date2str(tdate) + "'"n    Elsen        sql = " select code, endprice from ras.if_stock_data where code= '" + ua_code + "' and tdate='" + date2str(tdate) + "' union select code, price from ras.bl_quanto_base where code='" + ua_code + "' and tdate='" + date2str(tdate) + "'"n    End Ifn    n    With oCmdn        .ActiveConnection = adoConn        .CommandType = adCmdTextn        .CommandText = sqln       n        oRS.Open .Executen    End Withnn    Dim rtn As Doublen    rtn = 0n    n    Do Until oRS.EOFn        rtn = oRS(1)n        oRS.MoveNextn    Loopn    n    oRS.Closen    n    Set oRS = Nothingn    Set oCmd = Nothingn    n    If IsMissing(adoCon) = True Thenn        Call disconnectDB(adoCon)n    End Ifn    n    get_spot = rtnnnEnd FunctionnnPublic Function is_eval_shift_ua(ua_code As String) As Booleannn    Dim rtn As Booleann    rtn = Falsen    n    Dim enum_ua As Variantn    n    For Each enum_ua In eval_shift_uan        If enum_ua = get_ua_idx(ua_code) Thenn            rtn = Truen            Exit Forn        End Ifn    Nextn    n    is_eval_shift_ua = rtnnnEnd FunctionnnPublic Function is_flatvol_ua(ua_code As String) As Booleannn    Dim rtn As Booleann    rtn = Falsen    n    Dim enum_ua As Variantn    n    For Each enum_ua In FLAT_VOL_UAn        If enum_ua = get_ua_idx(ua_code) Thenn            rtn = Truen            Exit Forn        End Ifn    Nextn    n    is_flatvol_ua = rtnnnEnd FunctionnnPublic Function chk_vol_index(ul_code As String, ByRef base_ul_code As String) As Booleannn    Dim rtn As Booleann    n    base_ul_code = ul_coden    n    Dim adoCon As New adoDB.Connectionn    Dim oCmd As New adoDB.Commandn    Dim oRS As New adoDB.Recordsetn    n    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)n    n    Dim bind_variable() As Stringn    Dim bind_value() As Variantn    ReDim bind_variable(1) As Stringn    ReDim bind_value(1) As Variantn    bind_variable(1) = ":code"n    bind_value(1) = ul_coden    n    Dim sql As Stringn    sql = getSQL(SQL_PATH_CHK_VOL_INDEX, bind_variable, bind_value)nn    With oCmdnn        .ActiveConnection = adoConn        .CommandType = adCmdTextn        .CommandText = sqlnn        oRS.Open .Executenn    End Withn    n    Do Until oRS.EOFn        n        If IsNull(oRS(0)) = False Thenn            base_ul_code = oRS(0)n        End Ifn        oRS.MoveNextnn    Loopn    n    oRS.Closen    n    Call disconnectDB(adoCon)n    n    If base_ul_code = ul_code Thenn        rtn = Falsen    Elsen        rtn = Truen    End Ifn    n    chk_vol_index = rtnn        nEnd Function