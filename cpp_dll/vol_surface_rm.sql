select code, tdate, strike, mdate maturity_date, iv, lv
from rcs.pml_local_vol
where  code=:ul_code
and    tdate=:tdate
and    source=:source
and    strike in (select strike
    from   (select strike, row_number() over(order by rn) rn2 from (select strike,
                   row_number() over(
                   order by strike-:k) as rn
            from   (select distinct strike
                    from rcs.pml_local_vol
                    where  code=:ul_code
                    and    tdate=:tdate
                    and    source=:source
                    and    strike>=:k)
    union all select strike,
                   row_number() over(
                   order by :k-strike) as rn
            from   (select distinct strike
                    from rcs.pml_local_vol
                    where  code=:ul_code
                    and    tdate=:tdate
                    and    source=:source
                    and    strike<:k)))
    where  rn2 in (1, 2))
and    mdate in (select mdate
    from   (select mdate, row_number() over(order by rn) rn2 from (select mdate,
                   row_number() over(
                   order by (to_date(mdate, 'yyyymmdd')-to_date(:tdate, 'yyyymmdd'))/365-:tau) as rn
            from   (select distinct mdate
                    from rcs.pml_local_vol
                    where  code=:ul_code
                    and    tdate=:tdate
                    and    source=:source
                    and    (to_date(mdate, 'yyyymmdd')-to_date(:tdate, 'yyyymmdd'))/365 >=:tau)
    union all select mdate,
                   row_number() over(
                   order by :tau-(to_date(mdate, 'yyyymmdd')-to_date(:tdate, 'yyyymmdd'))/365) as rn
            from   (select distinct mdate
                    from rcs.pml_local_vol
                    where  code=:ul_code
                    and    tdate=:tdate
                    and    source=:source
                    and    (to_date(mdate, 'yyyymmdd')-to_date(:tdate, 'yyyymmdd'))/365 <:tau)))
    where  rn2 in (1, 2))
order by 3, 4