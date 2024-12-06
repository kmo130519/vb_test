select day, grid_date, dcf, rnum from (
select to_date(grid_date, 'YYYYMMDD')-to_date(tdate, 'YYYYMMDD') as day,
       to_date(grid_date, 'YYYYMMDD') as grid_date,
       dcf ,
       row_number() over (partition by floor(grid_point*52)
        order by grid_point) rnum
from   spt.mmkt_ccy_dcf_rm
where  tdate= :tdate
and    ccy= :ccy
and    ccy = ccy2
and    to_date(grid_date, 'YYYYMMDD')-to_date(tdate, 'YYYYMMDD') <= 30
union all
select to_date(grid_date, 'YYYYMMDD')-to_date(tdate, 'YYYYMMDD') as day,
       to_date(grid_date, 'YYYYMMDD') as grid_date,
       dcf ,
       row_number() over (partition by floor(grid_point*12)
        order by grid_point) rnum
from   spt.mmkt_ccy_dcf_rm
where  tdate= :tdate
and    ccy= :ccy
and    ccy = ccy2
and    to_date(grid_date, 'YYYYMMDD')-to_date(tdate, 'YYYYMMDD') > 30
and    to_date(grid_date, 'YYYYMMDD')-to_date(tdate, 'YYYYMMDD') <= 90
union all
select to_date(grid_date, 'YYYYMMDD')-to_date(tdate, 'YYYYMMDD') as day,
       to_date(grid_date, 'YYYYMMDD') as grid_date,
       dcf ,
       row_number() over (partition by floor(grid_point*4)
        order by grid_point) rnum
from   spt.mmkt_ccy_dcf_rm
where  tdate= :tdate
and    ccy= :ccy
and    ccy = ccy2
and    to_date(grid_date, 'YYYYMMDD')-to_date(tdate, 'YYYYMMDD') > 90
and    to_date(grid_date, 'YYYYMMDD')-to_date(tdate, 'YYYYMMDD') <= 365*5 
) where rnum=1 
