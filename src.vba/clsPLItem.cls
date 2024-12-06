
Option Explicit

Public pl_date As Date
Public book_date As Date
Public pl_item_title As String
Public crncy As String
Public qty As Double
Public book_value As Double
Public origin_book_value As Double
Public book_price As Double
Public evaluation As Double
Public trading_pl As Double
Public update_dt As Date


Public Function copy_obj() As clsPLItem

    Dim rtn_obj As clsPLItem
    
    Set rtn_obj = New clsPLItem
    
    rtn_obj.qty = qty
    rtn_obj.pl_date = pl_date
    rtn_obj.book_date = book_date
    rtn_obj.pl_item_title = pl_item_title
    rtn_obj.crncy = crncy
    rtn_obj.qty = qty
    rtn_obj.book_value = book_value
    rtn_obj.book_price = book_price
    rtn_obj.evaluation = evaluation
    rtn_obj.trading_pl = trading_pl
    rtn_obj.update_dt = update_dt
    

    Set copy_obj = rtn_obj

End Function

Public Sub update_values(additional_qty As Double, price As Double)
        
    ' In case additional position. i.e. adding long position on long position
    If qty * additional_qty >= 0 Then
    
        book_value = book_value + additional_qty * price
        qty = qty + additional_qty
    Else
    ' In case canceling deal.
        Dim qty_to_cancel As Double
        Dim new_qty As Double
        
        qty_to_cancel = min(Abs(additional_qty), Abs(qty)) * Sgn(additional_qty)
        new_qty = additional_qty - qty_to_cancel
        
        book_value = book_value + qty_to_cancel * book_price + new_qty * price
        qty = qty + additional_qty
        
        trading_pl = trading_pl + qty_to_cancel * (book_price - price)
        
    End If
    
    If qty <> 0 Then
        book_price = Abs(book_value) / qty
    End If
        

End Sub

Public Sub update_evaluation(eval_price As Double)

    evaluation = qty * eval_price

End Sub

Public Sub set_eval_date(pl_date_in As Date)

    pl_date = pl_date_in
    
End Sub