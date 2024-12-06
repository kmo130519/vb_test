Option Explicit


Public asset_code As String
Public bl_code As String
Public last_price As Double
Public prev_price As Double
Public theo_price As Double
Public prev_theo_price As Double

Public maturity_date As Date

Public correct_data As Boolean


Public Function copy_obj() As clsQuote

    
    Dim rtn_obj As clsQuote
    
    Set rtn_obj = New clsQuote
    
    rtn_obj.asset_code = asset_code
    rtn_obj.bl_code = bl_code
    rtn_obj.last_price = last_price
    rtn_obj.prev_price = prev_price
    rtn_obj.theo_price = theo_price
    rtn_obj.prev_theo_price = prev_theo_price
    rtn_obj.maturity_date = maturity_date
    
    
    rtn_obj.correct_data = correct_data


    Set copy_obj = rtn_obj

End Function