Option Explicit

Public deal_seq As Long
Public book_src As String
Public tdate As Date
Public ccy As String
Public ccy2 As String
Public sell_buy As Integer
Public valdate As Date
Public notional As Double
Public ref_rate As Double
Public trade_type As String



Public Function copy_obj() As clsCcyTrade

    Dim obj As clsCcyTrade
    
    Set obj = New clsCcyTrade
    
    obj.deal_seq = deal_seq
    obj.book_src = book_src
    obj.tdate = tdate
    obj.ccy = ccy
    obj.ccy2 = ccy2
    obj.sell_buy = sell_buy
    obj.valdate = valdate
    obj.notional = notional
    obj.ref_rate = ref_rate
    obj.trade_type = trade_type
    

    Set copy_obj = obj

End Function