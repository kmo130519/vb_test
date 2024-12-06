select UL_CODE, EVAL_DATE, STRIKE, MATURITY_DATE, VOLATILITY, UPDATE_DT
from sps.ul_local_vol_surface
where  ul_code=:ul_code
and    eval_date=:tdate
order by 1, 3, 4
