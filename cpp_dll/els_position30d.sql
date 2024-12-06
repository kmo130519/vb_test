--30일 한도관리
select :tdate, sum(limit), sum(riskvalue) from
(
SELECT :tdate TDATE, 'CONFIRMED', NVL(SUM(T.PBLC_STCK_QTY*T.MLTL*DECODE(T.STLM_CRCD,'USD',1222.43,1)),0) LIMIT, NVL(SUM(T.PBLC_STCK_QTY*T.MLTL*DECODE(T.STLM_CRCD,'USD',1222.43,1)),0) RISKVALUE FROM
(
SELECT M.*, F.ENDPRICE FX FROM
(SELECT * FROM RMS.FX_DATA WHERE TDATE=:today) F,
(SELECT BD1.*
       , META_FOPCLS.CODE_VAL_NAME META_FOPCLS --FO_OTC
       , META_IOPCLS.CODE_VAL_NAME META_IOPCLS --IO_OTC
       , META_FPCLS.CODE_VAL_NAME META_FPCLS --FO상품구분
       , META_IPCLS.CODE_VAL_NAME META_IPCLS --IO상품구분
       , META_DCLS.CODE_VAL_NAME META_DCLS --매도매수
       , META_HEDG.CODE_VAL_NAME META_HEDG --헤지구분
       , META_PKIND.CODE_VAL_NAME META_PKIND --상품종류
       , META_UNAS.CODE_VAL_NAME META_UNAS --기초자산유형
       , META_FOBJ.CODE_VAL_NAME META_FOBJ --펀드거래목적
       , META_HASET.CODE_VAL_NAME META_HASET --헤지자산구분
       , DECODE(BD1.PROD_KIND_CODE,'27',META_CLRD.CODE_VAL_NAME,NULL) META_CLRD --조기상환유형
FROM   (
SELECT F.MANG_DPCD
      , G.DPNM --부서명
      , B.PROD_FNCD --펀드코드
      , B.OTC_FUND_ISCD --상품펀드코드
      , B.PROD_CLS_CODE FO_PROD_CLS_CODE
      , A.KOR_ISNM --상품명
      , C.DEAL_CLS_CODE 
      , C.INDV_ISCD --개별종목코드
      , C.PROD_CLS_CODE IO_PROD_CLS_CODE 
      , B.HEDG_TR_TYPE_CODE
      , B.HEDG_ISCD FO_HEDG_ISCD --헤지상품펀드코드
      , C.HEDG_ISCD IO_HEDG_ISCD --헤지개별종목코드
      , B.FUND_TR_OBJT_CODE
      , B.HEDG_ASET_CLS_CODE
      , C.PROD_KIND_CODE
      , C.CLRD_TYPE_CODE
      , B.UNAS_TYPE_CODE
      , B.OTC_PROD_CLS_CODE FO_OTC_PROD_CLS_CODE
      , C.OTC_PROD_CLS_CODE IO_OTC_PROD_CLS_CODE     
      , C.STLM_CRCD --결제통화
      ,
      -- 발행일 이후이면 최초발행수량 사용, 발행일 이전 거래일이면 잔고 또는 총발행잔고수량사용
      CASE WHEN C.PBLC_DATE<=:today THEN
        B.PBLC_STCK_QTY --최초발행수량
      ELSE
        NVL(D.RMND_QTY,C.TOTAL_PBLC_QTY) --매매기준 잔고수량
      END PBLC_STCK_QTY
      , B.FUND_PBLC_UNPR --발행단가(최종판매가) --펀드정보, 만원 기준, 매매목적,헤지수단은 장부가*10000              
      , C.MLTL --액면단가(승수) --개별정보=만원              
      , B.BRIF_EXPL --요약정보
      , B.FRST_STND_PRC_FIN_DATE --거래일(최초기준가결정일_끝) --펀드정보, 헤지대상,매매목적만
      , C.STLM_DATE --매매결제일
      , C.PBLC_DATE --발행일
      , C.MTRT_DATE --만기일
      , C.EXCL_DATE --만기지급일
FROM
        BSYS.TBSIMM100M00@GDW A --종목기본
      , BSYS.TBSIMO100M00@GDW B --장외파생 펀드종목 기본
      , BSYS.TBSIMO201M00@GDW C --장외파생 개별종목 기본
      , BSYS.TBFNOM021L00@GDW D --잔고내역
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
  AND B.PROD_CLS_CODE IN ('04','07','09','05') --2021.1.29 ('05' 복합 추가. SF펀드의 equtiy swap)
  AND C.DEAL_CLS_CODE='1' --2021.1.29 추가 (SF002 매수 제외)
  AND C.PROD_KIND_CODE IN ('27')
  AND C.INDV_ISCD IN (SELECT ISCD FROM BSYS.TBSIMO200D00@GDW WHERE FUND_BTWN_DEAL_YN ='Y')
  AND B.FRST_STND_PRC_FIN_DATE BETWEEN TO_CHAR(TO_DATE(:tdate,'YYYYMMDD')-29,'YYYYMMDD') AND :today --30일 최초기준가설정일 기준
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

       