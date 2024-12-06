select endprice from 
(
select tdate, indexid code, endprice from ras.if_index_data
union all
select tdate, code, endprice from ras.if_stock_data
)
where tdate = :tdate
and code = :code