select nvl(time, 0) timetomaturity ,
       nvl(volatility, 0) volatility
from   sps.v_ul_vol_surf
where  (ul_code, fund_code_m, eval_date) in (select ul_code,
               fund_code_m ,
               max(eval_date)
        from   sps.v_ul_vol_surf
        where  eval_date <= :tdate
        and    ul_code = :code
        and    moneyness = 1
        and    fund_code_m = 'SA'
        group by ul_code, fund_code_m) order by time
