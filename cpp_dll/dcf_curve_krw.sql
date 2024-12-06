select to_date(termdate, 'yyyymmdd')-to_date(a.tdate, 'yyyymmdd') as day,
       to_date(termdate, 'yyyymmdd') as grid_date,
       rate,
       to_date(a.tdate, 'yyyymmdd') as tdate,
       b.name
from   ras.if_rate_data a,
       ras.if_rateid_info b
where  a.tdate=b.tdate
and    a.rateid=b.rateid
and    a.tdate in (select max(tdate)
        from   ras.if_day_info
        where  tdate < :tdate
        and    workind='Y')
and    a.rateid = :rateid
order by day asc 