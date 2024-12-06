select endprice from 
(
select tdate, indexid code, endprice, scenarioid from rcs.pml_index_data_st
union all
select tdate, code, endprice, scenarioid from rcs.pml_stock_data_st
)
where tdate = :tdate
and code = :code
and scenarioid = :scenarioid