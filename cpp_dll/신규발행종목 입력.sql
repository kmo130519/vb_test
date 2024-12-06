MERGE
INTO   RAS.RM_ELS_INFO USING dual
ON (deal_name = :deal_name)
       WHEN MATCHED THEN
UPDATE
SET    code = :code,
       hedge_buf = :hedge_buf,
       inserttime = sysdate,
       indv_iscd = :indv_iscd,
       underling_type = :u_type,
       status = '발행 완료'
       WHEN NOT MATCHED THEN
INSERT (code,
               hedge_buf,
               inserttime,
               indv_iscd,
               underling_type,
               status)
VALUES (:code,
               :hedge_buf,
               sysdate,
               :indv_iscd,
               :u_type,
               '발행 완료')

