'수정사항
'
'----clsACDealTicket----
'Public performance_type As Long -> Public performance_type() As Long
'Public strike_shift As Double -> Public strike_shift() As Double
'[add] Public early_exit_strike_shift() As Double
'[add] Public early_exit_touched_flag As Boolean

'----sy_shared_ptr_FdAutocallableEngine1D::initializeEngine
'[delete] double strikeShift

'----sy_shared_ptr_FdAutocallableEngine2D::initializeEngine
'[delete] double strikeShift

'----shared_ptr_AutocallablePayoffND::initializePayoff
'performanceType: 만기 베리어의 performace type

'Private Sub Workbook_Open()
'
'    Dim i_ua As Integer
'
'    With Application
'        .ScreenUpdating = False
'        .Calculation = xlCalculationManual
'        .EnableEvents = False
'    End With
'
'    For i_ua = 1 To num_ua
'        If Not IsEmpty(shtACPricer.Range("ul_code_list").Cells(i_ua, 1)) Then
'            shtACPricer.Range("ua_close_spot").Cells(i_ua, 2).Formula = "=Market!" & shtMarket.Range(CStr(shtACPricer.Range("ul_code_list").Cells(i_ua, 1))).Cells(4, 3).Address
'            shtACPricer.Range("ua_close_spot").Cells(i_ua, 3).Formula = "=Market!" & shtMarket.Range(CStr(shtACPricer.Range("ul_code_list").Cells(i_ua, 1))).Cells(4, 3).Offset(0, 5).Address
'            shtACPricer.Range("ua_close_spot").Cells(i_ua, 4).Formula = "=" & shtACPricer.Range("ua_close_spot").Cells(i_ua, 2).Address & "/" & shtACPricer.Range("ua_close_spot").Cells(i_ua, 3).Address & "-1"
'            shtACPricer.Range("ua_close_spot").Cells(i_ua, 5).Formula = "=Market!" & shtMarket.Range(CStr(shtACPricer.Range("ul_code_list").Cells(i_ua, 1))).Cells(5, 3).Address
'        End If
'    Next i_ua
'
'    With Application
'        .ScreenUpdating = True
'        .Calculation = xlCalculationAutomatic
'        .EnableEvents = True
'    End With
'
'End Sub

'기초자산 추가시 수정사항
'1. <Market> DB조회
'   시트 상에 기초자산 영역 삽입 및 이름 설정. 영역 내에 기초자산 이름 및 코트 Key-in
'   상관계수 행렬에 행열 삽입, 영역 내에 기초자산 코드 Key-in
'2. <Market-LocalVol> DB조회
'   시트 상에 기초자산 영역 삽입 및 이름 설정. 영역 내에 기초자산 이름 및 코트 Key-in
'   Flat vol. 시작 index 확인
'3. <Index_ELS> 계산
'   ul_ofs() 함수에 기초자산 추가
'   display_crossgamma() 함수에 상관계수 piar 추가
'   RangeClear, RangeClear2, pl_est_tot 범위 확대
'4. <Index_ELS> DV01입력
'   ua_idx_ccy 추가
'5. <Greek>
'   기초자산 추가
'7. <AC_Pricer> Calculate
'   cmd_calc_ac_price -> NUM_UA 변경
'   read_market_set 기초자산 추가