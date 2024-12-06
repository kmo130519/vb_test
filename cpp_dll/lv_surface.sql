select code, tdate, strike*ref_spot, mdate, lv
from rcs.pml_local_vol
where  tdate = :tdate
and code = :ul_code
and source = :source
order by 1, 3, 4