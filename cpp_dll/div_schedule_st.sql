select ex_div_date,
       div
from   rcs.pml_div_schedule_st
where  tdate= :tdate
and    code = :ul_code
and    scenarioid = :scenarioid
and    div <> 0