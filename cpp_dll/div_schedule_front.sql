select ex_div_date, nvl(dividend,0) dividend
from    sps.ul_div_schedule
where   eval_date = (select max(eval_date) from sps.ul_div_schedule where ul_code=:ul_code and eval_date <= :tdate) 
and     ex_div_date between :tdate and to_char(to_date(:tdate, 'YYYYMMDD')+1300, 'YYYYMMDD')
and     ul_code = :ul_code
group by ex_div_date, dividend order by 1 asc