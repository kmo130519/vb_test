select add_months(to_date(stnd_date, 'yyyymmdd'), mtrt_rmnn_mncn)-to_date(stnd_date, 'yyyymmdd') as day,
       add_months(to_date(stnd_date, 'yyyymmdd'), mtrt_rmnn_mncn) as grid_date,
       ert as rate,
       to_date(stnd_date, 'yyyymmdd') as tdate,
       trim(risk_crdt_grad) as name
from   bsys.tbfnia027l00@gdw
where  stnd_date in (select max(tdate)
        from   ras.if_day_info
        where  tdate < :tdate
        and    workind='Y')
and    crnc_cls_code='1'
and    kbp_bcdt_cls_code='2'
and    trim(risk_crdt_grad)= :rateid
order by day asc 