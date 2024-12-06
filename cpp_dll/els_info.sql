MERGE INTO RAS.RM_ELS_INFO USING dual ON (deal_name = :deal_name)
WHEN MATCHED THEN
    UPDATE
    SET     PAYOFF_DESC = :payoff_desc, 
            VALUE_DATE = :value_date, 
            ISSUE_DATE = :issue_date, 
            FUNDED_YN = :funded_yn, 
            CCY = :ccy, 
            NOTIONAL_LIMIT = :notional_limit, 
            NOTIONAL_EST = :notional_est, 
            FUNDING_SPRD = :funding_sprd, 
            DEAL_PRICE = :deal_price, 
            KI_SHIFT = :ki_shift, 
            STRIKE_SMOOTHING = :strike_smoothing, 
            THEO_PRICE_FR = :theo_price_fr, 
            THEO_PRICE_RM = :theo_price_rm, 
            QUOTE_DATE = :quote_date, 
            STATUS = :status, 
            INSERTTIME = SYSDATE
WHEN NOT MATCHED THEN
    INSERT (DEAL_NAME, PAYOFF_DESC, VALUE_DATE, ISSUE_DATE, FUNDED_YN, CCY, NOTIONAL_LIMIT, NOTIONAL_EST, FUNDING_SPRD, DEAL_PRICE, KI_SHIFT, STRIKE_SMOOTHING, THEO_PRICE_FR, THEO_PRICE_RM, QUOTE_DATE, STATUS, INSERTTIME)
    VALUES (:deal_name,
                   :payoff_desc,
                   :value_date,
                   :issue_date,
                   :funded_yn,
                   :ccy,
                   :notional_limit,
                   :notional_est,
                   :funding_sprd,
                   :deal_price,
                   :ki_shift,
                   :strike_smoothing,
                   :theo_price_fr,
                   :theo_price_rm,
                   :quote_date,
                   :status,
                   SYSDATE)