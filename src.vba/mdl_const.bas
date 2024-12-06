Option Explicit

#If Win64 Then
    Public Const PROVIDER As String = "OraOLEDB.Oracle"
#Else
    Public Const PROVIDER As String = "MSDAORA.1"
#End If

Public Const TNS_SERVICE_NAME As String = "RM01"
Public Const TNS_SERVICE_NAME_TEST As String = "RMSDEV"
Public Const USER_ID As String = "RMS"
Public Const PASSWORD As String = "RMS"

Public Const SQL_PATH_ELS As String = "C:\cpp_dll\els_list.sql"
Public Const SQL_PATH_IDS As String = "C:\cpp_dll\ids_list.sql" '2024.06.13
'Public Const SQL_PATH_ELS_PRICING As String = "C:\cpp_dll\els_pricing_list.sql"
Public Const SQL_PATH_ELS_PRICE_FR As String = "C:\cpp_dll\els_price_fr.sql"
Public Const SQL_PATH_UA_ENDPRICE As String = "C:\cpp_dll\ua_endprice.sql"
Public Const SQL_PATH_UA_ENDPRICE_ST As String = "C:\cpp_dll\ua_endprice_st.sql"
Public Const SQL_PATH_RF_CURVE As String = "C:\cpp_dll\rf_curve.sql"
Public Const SQL_PATH_RF_CURVE_ST As String = "C:\cpp_dll\rf_curve_st.sql"
Public Const SQL_PATH_LC_LAMBDA_NEUTRAL As String = "C:\cpp_dll\lc_lambda_neutral.sql"
Public Const SQL_PATH_IV_SURFACE As String = "C:\cpp_dll\iv_surface.sql"
Public Const SQL_PATH_LV_SURFACE As String = "C:\cpp_dll\lv_surface.sql"
Public Const SQL_PATH_LV_SURFACE_ST As String = "C:\cpp_dll\lv_surface_st.sql"
Public Const SQL_PATH_IV_SURFACE_FRONT As String = "C:\cpp_dll\iv_surface_front.sql"
Public Const SQL_PATH_LV_SURFACE_FRONT As String = "C:\cpp_dll\lv_surface_front.sql"
Public Const SQL_PATH_IV_SURFACE_4PTS As String = "C:\cpp_dll\iv_surface_4pts.sql"
Public Const SQL_PATH_LV_SURFACE_4PTS As String = "C:\cpp_dll\lv_surface_4pts.sql"
Public Const SQL_PATH_IV_SURFACE_4PTS_FRONT As String = "C:\cpp_dll\iv_surface_4pts_front.sql"
Public Const SQL_PATH_LV_SURFACE_4PTS_FRONT As String = "C:\cpp_dll\lv_surface_4pts_front.sql"
Public Const SQL_PATH_IV_SURFACE_SIZE As String = "C:\cpp_dll\iv_surface_size.sql"
Public Const SQL_PATH_LV_SURFACE_SIZE As String = "C:\cpp_dll\lv_surface_size.sql"
Public Const SQL_PATH_IV_SURFACE_SIZE_FRONT As String = "C:\cpp_dll\iv_surface_size_front.sql"
Public Const SQL_PATH_LV_SURFACE_SIZE_FRONT As String = "C:\cpp_dll\lv_surface_size_front.sql"
Public Const SQL_PATH_DIV_YIELD As String = "C:\cpp_dll\div_yield.sql"
Public Const SQL_PATH_DIV_SCHEDULE_FRONT As String = "C:\cpp_dll\div_schedule_front.sql"
Public Const SQL_PATH_DIV_SCHEDULE As String = "C:\cpp_dll\div_schedule.sql"
Public Const SQL_PATH_DIV_SCHEDULE_ST As String = "C:\cpp_dll\div_schedule_st.sql"
Public Const SQL_PATH_FX_VOL As String = "C:\cpp_dll\fx_vol.sql"
Public Const SQL_PATH_FX_VOL_ST As String = "C:\cpp_dll\fx_vol_st.sql"
Public Const SQL_PATH_FX_VOL_FRONT As String = "C:\cpp_dll\fx_vol_front.sql"
Public Const SQL_PATH_CORR As String = "C:\cpp_dll\corr.sql"
Public Const SQL_PATH_CORR_FRONT As String = "C:\cpp_dll\corr_front.sql"
Public Const SQL_PATH_MIN_CORR_FRONT As String = "C:\cpp_dll\min_corr_front.sql"
Public Const SQL_PATH_DCF_CURVE_KRW As String = "C:\cpp_dll\dcf_curve_krw.sql"
Public Const SQL_PATH_DCF_CURVE_USD As String = "C:\cpp_dll\dcf_curve_usd.sql"
Public Const SQL_PATH_DCF_CURVE_KRW_ST As String = "C:\cpp_dll\dcf_curve_krw_st.sql"
Public Const SQL_PATH_FLAT_VOL As String = "C:\cpp_dll\flat_vol.sql"
Public Const SQL_PATH_FUT_OPT As String = "C:\cpp_dll\fut_opt.sql"
Public Const SQL_PATH_AC_DEAL As String = "C:\cpp_dll\ac_deal.sql"
Public Const SQL_PATH_AC_DEAL_UL As String = "C:\cpp_dll\ac_deal_ul.sql"
Public Const SQL_PATH_AC_DEAL_DUMMY As String = "C:\cpp_dll\ac_deal_dummy.sql"
Public Const SQL_PATH_AC_SCHEDULE As String = "C:\cpp_dll\ac_schedule.sql"
Public Const SQL_PATH_FLOATING_LEG As String = "C:\cpp_dll\floating_leg.sql"
Public Const SQL_PATH_CHK_VOL_INDEX As String = "C:\cpp_dll\chk_vol_index.sql"
Public Const SQL_PATH_INS_VOLTARGET_GREEKS As String = "C:\cpp_dll\insert_voltarget_greeks.sql"
Public Const SQL_PATH_INS_SPOT_CASH_GREEKS As String = "C:\cpp_dll\insert_spot_cash_greeks.sql"

Public Enum INST_TYPE
    note = 1
    SWAP = 2
End Enum

Public Const NUM_UA As Integer = 18

Public Enum ua
    HSCEI = 1
    HSI = 2
    SX5E = 3
    SPX = 4
    NKY = 5
    KOSPI200 = 6
    KRD020021147 = 7 'KOSPI200레버리지
    KR7005380001 = 8 '현대차
    KR7005930003 = 9 '삼성전자
    KR7000030007 = 10 '우리은행(상장폐지)
    KR7028260008 = 11 '삼성물산
    KR7105560007 = 12 'KB금융
    KR7035420009 = 13 'NAVER
    KR7018260000 = 14 '삼성SDS
    KR7005490008 = 15 'POSCO
    KR7034220004 = 16 'LG디스플레이
    TSLA = 17 'TESLA
    NVDA = 18 'Nvidia
End Enum

Public eval_shift_ua(3) As Integer
Public FLAT_VOL_UA(10) As Integer

Public Const MAX_UA_PCT_PRICE = 1000

Public Const NUM_CCY As Integer = 5

Public Enum ccy
    KRW = 1
    USD = 2
    EUR = 3
    JPY = 4
    HKD = 5
End Enum

Public Const NUM_DCF As Integer = 2

Public Enum DCF
    KRW = ccy.KRW
    USD = ccy.USD
End Enum

'할인금리커브 ID
Public Const DCF_KRW_RATEID As String = "E300"
Public Const DCF_USD_RATEID As String = "BBB+"

'무위험금리커브 테너 수
Public Enum RF_NUM_TERM
    KRW = 12
    JPY = 11
    HKD = 14
    EUR = 13
    USD = 15
End Enum

Public Const NUM_DIV_DATES = 1300

Public SCENARIO_ENABLE As Boolean
Public POPUP_WARNING_ENABLE As Boolean
Public CORR_SKEW_ENABLE As Boolean
Public PAYOFF_SMOOTHING_ENABLE As Boolean
Public KI_SHIFT_ENABLE As Boolean
Public GREEKS_ENABLE As Boolean

Public Enum DATA_FROM
    bsys = 1
    FRONT = 2
    RISK = 3
End Enum

Public Enum UA_CODE_TYPE
    SP = 1
    SP_EVENT = 2
    BLP = 3
    ISIN = 4
End Enum
    
Public Sub SET_GLOBAL()

    SCENARIO_ENABLE = Range("SCENARIO_ENABLE").value
    
    If SCENARIO_ENABLE = False Then
        POPUP_WARNING_ENABLE = False
    Else
        POPUP_WARNING_ENABLE = Range("POPUP_WARNING_ENABLE").value
    End If
    
    CORR_SKEW_ENABLE = Range("CORR_SKEW_ENABLE").value
    PAYOFF_SMOOTHING_ENABLE = Range("PAYOFF_SMOOTHING_ENABLE").value
    KI_SHIFT_ENABLE = Range("KI_SHIFT_ENABLE").value
    GREEKS_ENABLE = Range("GREEKS_ENABLE").value
    
    SET_EVAL_SHIFT_UA
    SET_FLAT_VOL_UA
    
End Sub

Public Sub SET_EVAL_SHIFT_UA()
    eval_shift_ua(0) = ua.SPX
    eval_shift_ua(1) = ua.SX5E
    eval_shift_ua(2) = ua.TSLA
    eval_shift_ua(3) = ua.NVDA
End Sub

Public Sub SET_FLAT_VOL_UA()
    FLAT_VOL_UA(0) = ua.KR7005380001 '현대차
    FLAT_VOL_UA(1) = ua.KR7005930003 '삼성전자
    FLAT_VOL_UA(2) = ua.KR7000030007 '우리은행(상장폐지)
    FLAT_VOL_UA(3) = ua.KR7028260008 '삼성물산
    FLAT_VOL_UA(4) = ua.KR7105560007 'KB금융
    FLAT_VOL_UA(5) = ua.KR7035420009 'NAVER
    FLAT_VOL_UA(6) = ua.KR7018260000 '삼성SDS
    FLAT_VOL_UA(7) = ua.KR7005490008 'POSCO
    FLAT_VOL_UA(8) = ua.KR7034220004 'LG디스플레이
    FLAT_VOL_UA(9) = ua.TSLA 'TESLA
    FLAT_VOL_UA(10) = ua.NVDA 'Nvidia
End Sub