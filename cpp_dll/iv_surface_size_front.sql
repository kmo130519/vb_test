select ul_code,
       num_of_strikes,
       num_of_maturities
from   sps.ul_vol_surface_meta
where  eval_date = :tdate
and    ul_code = :ul_code