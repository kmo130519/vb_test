select dividend from (
select row_number() over (order by eval_date desc) rnum,
       nvl(div_yield, 0)   dividend 
from   sps.ul_div_yield
where  eval_date <= :tdate
and    ul_code = :ul_code
) where rnum=1
