select e.asset_code,
       round(e.greek_value/d.notional*decode(b2.DEAL_CLS_CODE,'1',-1,1)*b1.real_pblc_fcam, 5) price_theo_fr
from   spt.daily_closing_theo e,
       sps.ac_deal d,
       bsys.tbsimo100m00@gdw b1,
       bsys.tbsimo201m00@gdw b2
where  e.greek_cd='VALUE'
and    e.eval_date=:tdate
and    e.asset_code=:code
and    e.asset_code=d.asset_code
and    e.asset_code=b2.indv_iscd
and    b1.otc_fund_iscd = b2.otc_fund_iscd