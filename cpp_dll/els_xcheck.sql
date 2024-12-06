--대사(vs E파): E파 DB로 deal_ticket 구성 vs 업무계 DB로 deal_ticket 구성

select i.deal_name,
       i.payoff_desc,
       --d.fund_code_m,
       --d.fund_code_c,
       d.note_currency,
       i.ccy,
       d.value_date,
       i.value_date,
       d.expiry_date,       
       d.settlement_date,
       i.issue_date,
       d.dummy_coupon,
       d.ki_barrier_yn,
       d.ki_touched_yn,
       d.ki_monitoring_freq,
       d.notional,       
       u.ul_code,
       u.reference_price,
       u.ki_barrier,
       s.call_date,
       s.coupon_on_call,
       s.strike,
       s.performance_type,
       s.strike_smoothing_width,
       e.call_date,
       e.coupon_on_call,
       e.strike,
       e.performance_type,
       e.strike_smoothing_width,
       e.ee_touched_yn,
       e.barrier_type       
from   sps.ac_deal d,
       sps.ac_underlying u,
       sps.ac_schedule s,
       sps.ac_ee_schedule e,
       ras.rm_els_info i
where  d.asset_code=replace(replace(replace(replace(i.deal_name,'OTC'),'ELS'),'A'),'B') --prefix 제거 후 join: E파 code는 OTC/ELS 무관하게 회차 숫자만 씀
and d.asset_code=u.asset_code
and d.asset_code=s.asset_code
and d.asset_code=e.asset_code(+)
and i.code='TBD'