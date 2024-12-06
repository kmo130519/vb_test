Option Explicit


Public tdate As Date
Public base_ccy As String
Public ccy As String
Public book_src As String
Public spot_position As Double
Public spot_position_pl As Double
Public fwd_position As Double
Public basic_rate As Double
Public book_rate As Double
Public book_amt As Double
Public spot_eval_pl As Double
Public spot_trd_pl As Double
Public fwd_eval_pl As Double
Public fwd_trd_pl As Double

Public fund_code As String

Public Sub update_evaluation(eval_rate As Double)

    basic_rate = eval_rate
    
    
    spot_eval_pl = spot_position_pl * (eval_rate - book_rate)


End Sub

'------------------------------------------------------------------------------
' Assumption: We cannot take negative fx position, meaning additional_position should be greater than or equal to - spot_position
'------------------------------------------------------------------------------
Public Sub update_spot_position(additional_position As Double, rate As Double, Optional update_pl_position As Boolean = True, Optional calc_trading_pl As Boolean = False)
        
    ' In case additional position. i.e. adding long position on long position
        
    If additional_position * spot_position_pl >= 0 Or Not calc_trading_pl Then
    
        spot_position = spot_position + additional_position
        
        If update_pl_position Then
            book_amt = book_amt + additional_position * rate
            spot_position_pl = spot_position_pl + additional_position
            
            If spot_position_pl <> 0 Then
                book_rate = book_amt / spot_position_pl
            Else
                book_rate = 0
            End If
            
        End If
        
    Else

        spot_position = spot_position + additional_position
        
        If update_pl_position Then
'            spot_position_pl = spot_position_pl + additional_position
'            book_amt = book_amt + additional_position * book_rate
'
'            spot_trd_pl = spot_trd_pl + additional_position * (book_rate - rate)
'
        ' In case canceling deal.
            Dim position_to_cancel As Double
            Dim new_position As Double
            
            position_to_cancel = min(Abs(additional_position), Abs(spot_position_pl)) * Sgn(additional_position)
            new_position = additional_position - position_to_cancel
            
            book_amt = book_amt + position_to_cancel * book_rate + new_position * rate
            spot_position_pl = spot_position_pl + additional_position
            
            spot_trd_pl = spot_trd_pl + position_to_cancel * (book_rate - rate)
        
        End If

        
        
    End If


End Sub