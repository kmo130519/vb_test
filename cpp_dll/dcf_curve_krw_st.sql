select vertex as day,
       to_date(termdate, 'yyyymmdd') as grid_date,
       rate,
       to_date(a.tdate, 'yyyymmdd') as tdate,
       b.name
from   rcs.pml_rate_data_st a,
       ras.if_rateid_info b
where  a.tdate=b.tdate
and    a.rateid=b.rateid
and    a.tdate in (select max(tdate)
        from   ras.if_day_info
        where  tdate < :tdate
        and    workind='Y')
and    a.rateid = :rateid
and    scenarioid = :scenarioid
order by day asc 