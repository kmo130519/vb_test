Option Explicit

'------------------------------------------------------------------------
' CLASS : clsVanillaOption
' DESC :Simple Vanilla Option Class. Contains pricing functionality in itself.
' Dependence: mdl_2_vanilla, mdl_common, clsRateCurve, clsDividendSchedule, clsGreeks
'------------------------------------------------------------------------

Private Const pl_currency As String = "KRW" ' PL Currency

'-------------------------------------------
' Deal information
'-------------------------------------------
Public ul_price As Double
Public k As Double 'Strike Price
Public call_put As String ' Call / Put Type
Public maturity_date As Date 'maturity date as date type
Public qty As Double ' Number of options
Public open_position As Double ' Number of options at the beginning of the day.
Public conversion_ratio As Double 'Conversion ratio. Maybe 500,000 for KOSPI2.
Public face_value As Double

Public bl_code As String

Public option_type As String ' European / American

Public option_on_futures As Boolean
Public futures_maturity As Date



Public ul_currency As String ' UL Currency. Optional.

'-------------------------------------------
' Booking information.
'-------------------------------------------
Public asset_code As String
Public fund_code_m As String
Public fund_code_c As String
Public ul_code As String
Public ul_name As String
Public trade_date As Date
Public ul_initial_price As Double ' This is NOT the current ul price, but the ul price at the deal inception.
Public deal_type As String  'Futures / Listed Option / ELW

Public underlying_type As String

Public unit_qty As Double

Public broker As String

Private t As Double

Private functions_ As New Functions

Public Function ul_code_as_array() As String()

    Dim rtn_array(1 To 1) As String
    
    
    If UCase(Left(ul_code, 2)) = "ES" Then
        rtn_array(1) = "SPX"
    Else
        rtn_array(1) = ul_code
    End If
    
    ul_code_as_array = rtn_array

End Function


'====================================================================================
' Sub: initialize
' DESC: Initializer
'====================================================================================
Public Sub initialize(ByVal maturity_date_in As Date, ByVal k_in As Double, ByVal qty_in As Double, ByVal open_position_in As Double, ByVal conversion_ratio_in As Double, ByVal call_put_in As String _
                     , Optional ByVal ul_currency_in As String = "KRW" _
                     , Optional ByVal option_type_in As String = "European" _
                     , Optional option_on_futures_in As Boolean = False _
                     , Optional futures_maturity_in As Date = -1 _
                     )
                     


    k = k_in
    call_put = call_put_in
    maturity_date = maturity_date_in
    qty = qty_in
    open_position = open_position_in
    conversion_ratio = conversion_ratio_in
    ul_currency = ul_currency_in
    option_type = option_type_in
    option_on_futures = option_on_futures_in
    
    If futures_maturity_in > 0 Then
        futures_maturity = futures_maturity_in
    Else
        futures_maturity = maturity_date
    End If


End Sub