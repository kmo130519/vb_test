select hvol from (
select row_number() over (order by tdate desc) rnum,
       tdate,
       decode(flag, 'U', upperlevel, 'H', hedgevol, 'L', lowerlevel) hvol
from   rcs.pml_hvol
where  tdate <= :tdate
and    code = :code
and    flag is not null
) where rnum=1