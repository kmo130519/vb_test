select corr
from   rcs.pml_corr_beta_st
where  ((code1 = :code1 and code2 = :code2)
        or (code1 = :code2 and code2 = :code1))
and    tdate = :tdate
and    scenarioid = :scenarioid