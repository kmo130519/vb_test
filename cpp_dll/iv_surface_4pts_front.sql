select *
from sps.ul_vol_surface
where  ul_code=:ul_code
and    eval_date=:tdate
and    strike in (select strike
    from   (select strike, row_number() over(order by rn) rn2 from (select strike,
                   row_number() over(
                   order by strike/endprice-:k) as rn
            from   (select distinct strike, endprice
                    from sps.ul_vol_surface a, ras.if_index_data b
                    where  a.ul_code=:ul_code
                    and    a.eval_date=:tdate
                    and    a.ul_code=b.indexid
                    and    a.eval_date=b.tdate
                    and    a.strike/b.endprice>=:k)
    union all select strike,
                   row_number() over(
                   order by :k-strike/endprice) as rn
            from   (select distinct strike, endprice
                    from sps.ul_vol_surface a, ras.if_index_data b
                    where  a.ul_code=:ul_code
                    and    a.eval_date=:tdate
                    and    a.ul_code=b.indexid
                    and    a.eval_date=b.tdate
                    and    a.strike/b.endprice<:k)))
    where  rn2 in (1, 2))
and    maturity_date in (select maturity_date
    from   (select maturity_date, row_number() over(order by rn) rn2 from (select maturity_date,
                   row_number() over(
                   order by (to_date(maturity_date, 'yyyymmdd')-to_date(:tdate, 'yyyymmdd'))/365-:tau) as rn
            from   (select distinct maturity_date
                    from sps.ul_vol_surface
                    where  ul_code=:ul_code
                    and    eval_date=:tdate
                    and    (to_date(maturity_date, 'yyyymmdd')-to_date(:tdate, 'yyyymmdd'))/365 >=:tau)
    union all select maturity_date,
                   row_number() over(
                   order by :tau-(to_date(maturity_date, 'yyyymmdd')-to_date(:tdate, 'yyyymmdd'))/365) as rn
            from   (select distinct maturity_date
                    from sps.ul_vol_surface
                    where  ul_code=:ul_code
                    and    eval_date=:tdate
                    and    (to_date(maturity_date, 'yyyymmdd')-to_date(:tdate, 'yyyymmdd'))/365 <:tau)))
    where  rn2 in (1, 2))
order by 3, 4