SELECT code,
       avg(vol)
from   rcs.pml_fx_vol
where  tdate = :tdate
and    code = :code
group by code