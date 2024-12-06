select *
from sps.ul_vol_surface
where  ul_code=:ul_code
and    eval_date=:tdate
order by 1, 3, 4