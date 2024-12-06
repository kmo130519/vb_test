select neutral_lambda from (
select row_number() over (order by eval_date desc) rnum,
       nvl(neutral_lambda, 0) neutral_lambda
from   sps.loc_corr_lambda_formula
where  eval_date <= :tdate
and    ul_code = :ul_code
) where rnum=1