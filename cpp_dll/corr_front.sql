select corr
from   (select row_number() over(
                order by eval_date desc) rnum,
               eval_date,
               nvl(corr, 0) corr
        from   sps.ul_corr
        where  ((ul1_code = :code1
                        and    ul2_code = :code2)
                or     (ul1_code = :code2
                        and    ul2_code = :code1))
        and    eval_date <= :tdate)
where  rnum = 1