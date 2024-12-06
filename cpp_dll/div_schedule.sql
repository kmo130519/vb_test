select ex_div_date,
       div
from   rcs.pml_div_schedule
where  tdate= :tdate
and    code = :ul_code
and    div <> 0