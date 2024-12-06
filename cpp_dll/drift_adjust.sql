select term_date,
       adjust
from   sps.drift_adjust
where  ul_code=:ul_code
and    eval_date = (select max(eval_date)
        from   sps.ul_sabr_parameter
        where  ul_code=:ul_code
        and    eval_date<=:eval_date)
        
       