Option ExplicitnnPublic deal_seq As LongnPublic book_src As StringnPublic tdate As DatenPublic ccy As StringnPublic ccy2 As StringnPublic sell_buy As IntegernPublic valdate As DatenPublic notional As DoublenPublic ref_rate As DoublenPublic trade_type As StringnnnnPublic Function copy_obj() As clsCcyTradenn    Dim obj As clsCcyTraden    n    Set obj = New clsCcyTraden    n    obj.deal_seq = deal_seqn    obj.book_src = book_srcn    obj.tdate = tdaten    obj.ccy = ccyn    obj.ccy2 = ccy2n    obj.sell_buy = sell_buyn    obj.valdate = valdaten    obj.notional = notionaln    obj.ref_rate = ref_raten    obj.trade_type = trade_typen    nn    Set copy_obj = objnnEnd Functionn