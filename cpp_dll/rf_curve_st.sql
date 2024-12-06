select vertex, to_date(termdate,'YYMMDD'), dcf 
from rcs.pml_rate_data_st 
where tdate= :tdate 
and rateid= :ccy 
and scenarioid = :scenarioid
order by 1 asc