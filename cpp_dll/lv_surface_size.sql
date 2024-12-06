select code,
       count(distinct(strike)),
       count(distinct(mdate))
from   rcs.pml_local_vol
where  tdate = :tdate
and    code = :ul_code
and    lv is not null
and    source = :source
group by code