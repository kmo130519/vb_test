SELECT code,
       avg(vol)
from   rcs.pml_fx_vol_st
where  tdate = :tdate
and    code = :code
and    scenarioid = :scenarioid
group by code