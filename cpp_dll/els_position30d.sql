--30�� �ѵ�����
select :tdate, sum(limit), sum(riskvalue) from
(
SELECT :tdate TDATE, 'CONFIRMED', NVL(SUM(T.PBLC_STCK_QTY*T.MLTL*DECODE(T.STLM_CRCD,'USD',1222.43,1)),0) LIMIT, NVL(SUM(T.PBLC_STCK_QTY*T.MLTL*DECODE(T.STLM_CRCD,'USD',1222.43,1)),0) RISKVALUE FROM
(
SELECT M.*, F.ENDPRICE FX FROM
(SELECT * FROM RMS.FX_DATA WHERE TDATE=:today) F,
(SELECT BD1.*
       , META_FOPCLS.CODE_VAL_NAME META_FOPCLS --FO_OTC
       , META_IOPCLS.CODE_VAL_NAME META_IOPCLS --IO_OTC
       , META_FPCLS.CODE_VAL_NAME META_FPCLS --FO��ǰ����
       , META_IPCLS.CODE_VAL_NAME META_IPCLS --IO��ǰ����
       , META_DCLS.CODE_VAL_NAME META_DCLS --�ŵ��ż�
       , META_HEDG.CODE_VAL_NAME META_HEDG --��������
       , META_PKIND.CODE_VAL_NAME META_PKIND --��ǰ����
       , META_UNAS.CODE_VAL_NAME META_UNAS --�����ڻ�����
       , META_FOBJ.CODE_VAL_NAME META_FOBJ --�ݵ�ŷ�����
       , META_HASET.CODE_VAL_NAME META_HASET --�����ڻ걸��
       , DECODE(BD1.PROD_KIND_CODE,'27',META_CLRD.CODE_VAL_NAME,NULL) META_CLRD --�����ȯ����
FROM   (
SELECT F.MANG_DPCD
      , G.DPNM --�μ���
      , B.PROD_FNCD --�ݵ��ڵ�
      , B.OTC_FUND_ISCD --��ǰ�ݵ��ڵ�
      , B.PROD_CLS_CODE FO_PROD_CLS_CODE
      , A.KOR_ISNM --��ǰ��
      , C.DEAL_CLS_CODE 
      , C.INDV_ISCD --���������ڵ�
      , C.PROD_CLS_CODE IO_PROD_CLS_CODE 
      , B.HEDG_TR_TYPE_CODE
      , B.HEDG_ISCD FO_HEDG_ISCD --������ǰ�ݵ��ڵ�
      , C.HEDG_ISCD IO_HEDG_ISCD --�������������ڵ�
      , B.FUND_TR_OBJT_CODE
      , B.HEDG_ASET_CLS_CODE
      , C.PROD_KIND_CODE
      , C.CLRD_TYPE_CODE
      , B.UNAS_TYPE_CODE
      , B.OTC_PROD_CLS_CODE FO_OTC_PROD_CLS_CODE
      , C.OTC_PROD_CLS_CODE IO_OTC_PROD_CLS_CODE     
      , C.STLM_CRCD --������ȭ
      ,
      -- ������ �����̸� ���ʹ������ ���, ������ ���� �ŷ����̸� �ܰ� �Ǵ� �ѹ����ܰ�������
      CASE WHEN C.PBLC_DATE<=:today THEN
        B.PBLC_STCK_QTY --���ʹ������
      ELSE
        NVL(D.RMND_QTY,C.TOTAL_PBLC_QTY) --�Ÿű��� �ܰ����
      END PBLC_STCK_QTY
      , B.FUND_PBLC_UNPR --����ܰ�(�����ǸŰ�) --�ݵ�����, ���� ����, �ŸŸ���,���������� ��ΰ�*10000              
      , C.MLTL --�׸�ܰ�(�¼�) --��������=����              
      , B.BRIF_EXPL --�������
      , B.FRST_STND_PRC_FIN_DATE --�ŷ���(���ʱ��ذ�������_��) --�ݵ�����, �������,�ŸŸ�����
      , C.STLM_DATE --�ŸŰ�����
      , C.PBLC_DATE --������
      , C.MTRT_DATE --������
      , C.EXCL_DATE --����������
FROM
        BSYS.TBSIMM100M00@GDW A --����⺻
      , BSYS.TBSIMO100M00@GDW B --����Ļ� �ݵ����� �⺻
      , BSYS.TBSIMO201M00@GDW C --����Ļ� �������� �⺻
      , BSYS.TBFNOM021L00@GDW D --�ܰ���
      , BSYS.TBFNPA001M00@GDW F
      , BSYS.TBCPPD001M00@GDW G      
WHERE
  A.ISCD=C.INDV_ISCD
  AND B.OTC_FUND_ISCD=C.OTC_FUND_ISCD  
  AND C.INDV_ISCD=D.ISCD(+)
  AND D.STND_DATE(+)=:today
  AND C.MTRT_DATE>:today
  AND (C.PBLC_DATE<=:today OR B.FRST_STND_PRC_FIN_DATE<=:today)
  AND (NVL(D.RMND_QTY,0)>0 OR B.FRST_STND_PRC_FIN_DATE=:today)
  AND F.FNCD=B.PROD_FNCD
  AND F.MANG_DPCD=G.DPCD
  AND F.MANG_DPCD='351'
  AND B.HEDG_TR_TYPE_CODE='1'
  AND B.PROD_CLS_CODE IN ('04','07','09','05') --2021.1.29 ('05' ���� �߰�. SF�ݵ��� equtiy swap)
  AND C.DEAL_CLS_CODE='1' --2021.1.29 �߰� (SF002 �ż� ����)
  AND C.PROD_KIND_CODE IN ('27')
  AND C.INDV_ISCD IN (SELECT ISCD FROM BSYS.TBSIMO200D00@GDW WHERE FUND_BTWN_DEAL_YN ='Y')
  AND B.FRST_STND_PRC_FIN_DATE BETWEEN TO_CHAR(TO_DATE(:tdate,'YYYYMMDD')-29,'YYYYMMDD') AND :today --30�� ���ʱ��ذ������� ����
) BD1,
 (SELECT  B.CODE_VAL
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'OTC_PROD_CLS_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_FOPCLS,
 (SELECT  B.CODE_VAL                 
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'OTC_PROD_CLS_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_IOPCLS,
 (SELECT  B.CODE_VAL                 
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'PROD_CLS_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_FPCLS,
 (SELECT  B.CODE_VAL                 
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'PROD_CLS_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_IPCLS ,
       (SELECT  B.CODE_VAL                 
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'DEAL_CLS_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_DCLS,
        (SELECT  B.CODE_VAL                 
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'HEDG_TR_TYPE_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_HEDG,
         (SELECT  B.CODE_VAL                 
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'PROD_KIND_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_PKIND,
          (SELECT  B.CODE_VAL                 
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'CLRD_TYPE_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_CLRD,
          (SELECT  B.CODE_VAL                 
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'UNAS_TYPE_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_UNAS,
           (SELECT  B.CODE_VAL                 
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'FUND_TR_OBJT_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_FOBJ,
            (SELECT  B.CODE_VAL                 
     ,  B.CODE_VAL_NAME            
     ,  B.CODE_VAL_SORT_SEQ        
     ,  B.CODE_VAL_EXPL            
  FROM  BSYS.TBCPPC001M02@GDW A             
     ,  BSYS.TBCPPC001M01@GDW B             
 WHERE  A.META_TERM_PHSC_NAME = 'HEDG_ASET_CLS_CODE'
   AND  A.CODE_ID = B.CODE_ID      
 ORDER  BY  B.CODE_VAL) META_HASET
WHERE BD1.FO_OTC_PROD_CLS_CODE=META_FOPCLS.CODE_VAL(+)
AND BD1.IO_OTC_PROD_CLS_CODE=META_IOPCLS.CODE_VAL(+)
AND BD1.FO_PROD_CLS_CODE=META_FPCLS.CODE_VAL(+)
AND BD1.IO_PROD_CLS_CODE=META_IPCLS.CODE_VAL(+)
AND BD1.DEAL_CLS_CODE = META_DCLS.CODE_VAL(+)
AND BD1.HEDG_TR_TYPE_CODE=META_HEDG.CODE_VAL(+)
AND BD1.PROD_KIND_CODE=META_PKIND.CODE_VAL(+)
AND BD1.CLRD_TYPE_CODE=META_CLRD.CODE_VAL(+)
AND BD1.UNAS_TYPE_CODE=META_UNAS.CODE_VAL(+)
AND BD1.FUND_TR_OBJT_CODE=META_FOBJ.CODE_VAL(+)
AND BD1.HEDG_ASET_CLS_CODE=META_HASET.CODE_VAL(+)
) M
WHERE 'FXKRW' || M.STLM_CRCD = F.CODE(+)
) T
union all
select VALUE_DATE TDATE, 'TBD',
       sum(NOTIONAL_LIMIT*DECODE(CCY,'USD',1222.43,1)) Limit,
       sum(NOTIONAL_EST*DECODE(CCY,'USD',1222.43,1)) riskvalue
from   ras.rm_els_info
where  VALUE_DATE between TO_CHAR(TO_DATE(:tdate,'YYYYMMDD')-29,'YYYYMMDD') and :tdate
and VALUE_DATE > :today
group by VALUE_DATE
)

       