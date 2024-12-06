select ul_code, eval_date, strike, maturity_date, volatility
from rcs.pml_local_vol_surface_st
where  eval_date = :tdate
and ul_code = :ul_code
and scenarioid = :scenarioid
order by 1, 3, 4