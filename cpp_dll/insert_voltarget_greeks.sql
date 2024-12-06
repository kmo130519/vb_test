--RCS.PML_GREEK
DELETE FROM RCS.PML_GREEK WHERE TDATE=:tdate AND STK_CODE IN (SELECT A1.INDV_ISCD
        FROM   BSYS.TBSIMO201M00@GDW A1,
               BSYS.TBFNOM021L00@GDW A2,
               BSYS.TBSIMO203L00@GDW A3
        WHERE  A1.PROD_KIND_CODE = '02'
        AND    A1.INDV_ISCD = A2.ISCD
        AND    A1.INDV_ISCD = A3.INDV_ISCD
        AND    A2.STND_DATE = :tdate
        AND    A2.RMND_QTY > 0
        AND    A1.MTRT_DATE >= :tdate
        AND    A1.PROD_FNCD = 'SV001');

INSERT INTO RCS.PML_GREEK
SELECT :tdate TDATE,
       'SV001' FUND_CODE,
       D.ASSET_CODE STK_CODE,
       D.UL_CODE BASEASSET_CODE,
       '��ܿɼ�' GDS_TP,
       S.ENDPRICE CLOSE_AMT,
       TRUNC((0.75*D.TREE_DELTA + 0.25*D.SM_DELTA)/P.NOTIONAL,5) DELTA,
       TRUNC((0.75*G.TREE_GAMMA + 0.25*G.SM_GAMMA)/P.NOTIONAL,5) GAMMA,
       TRUNC(V.VEGA/P.NOTIONAL,5) VEGA,
       TRUNC((0.75*D.TREE_DELTA + 0.25*D.SM_DELTA)*S.ENDPRICE) DELTA_EXPOSURE,
       TRUNC((0.75*G.TREE_GAMMA + 0.25*G.SM_GAMMA)*(0.5*S.ENDPRICE*S.ENDPRICE*0.0001)) GAMMA_EXPOSURE,
       TRUNC(V.VEGA*0.01) VEGA_EXPOSURE,
       SYSDATE WORK_TIME,
       'SQL' WORK_TRM,
       'SQL' WORK_MEMB,
       NULL VEGA_VOL_EXPOSURE,
       NULL DELTA_EXPOSURE_SP,
       NULL GAMMA_EXPOSURE_SP,
       NULL VEGA_EXPOSURE_SP,
       NULL VEGA_VOL_EXPOSURE_SP,
       NULL DDELTADT,
       NULL DGAMMADT,
       TRUNC(T.THETA/P.NOTIONAL/365,5) THETA,
       '351' DEPT_CODE,
       NULL DV01
FROM   (SELECT TDATE,
               CODE,
               ENDPRICE
        FROM   RAS.IF_STOCK_DATA
        WHERE  TDATE=:tdate
        UNION ALL
        SELECT TDATE,
               INDEXID CODE,
               ENDPRICE
        FROM   RAS.IF_INDEX_DATA
        WHERE  TDATE=:tdate) S,
        (SELECT A1.INDV_ISCD, A2.RMND_QTY*A1.MLTL*DECODE(A1.DEAL_CLS_CODE,1,-1,1) NOTIONAL
        FROM   BSYS.TBSIMO201M00@GDW A1,
               BSYS.TBFNOM021L00@GDW A2,
               BSYS.TBSIMO203L00@GDW A3
        WHERE  A1.PROD_KIND_CODE = '02'
        AND    A1.INDV_ISCD = A2.ISCD
        AND    A1.INDV_ISCD = A3.INDV_ISCD
        AND    A2.STND_DATE = :tdate
        AND    A2.RMND_QTY > 0
        AND    A1.MTRT_DATE >= :tdate
        AND    A1.PROD_FNCD = 'SV001') P,
       (SELECT A.ASSET_CODE,
               A.UL_CODE,
               A.GREEK_VALUE TREE_DELTA,
               B.GREEK_VALUE SM_DELTA
        FROM   SPS.CL_DAILY_GREEKS A,
               SPS.CL_DAILY_GREEKS B
        WHERE  A.EVAL_DATE=:tdate
        AND    A.EVAL_DATE=B.EVAL_DATE
        AND    A.GREEK_CD='DELTA'
        AND    B.GREEK_CD='HEDGE_TARGET_DELTA'
        AND    A.ASSET_CODE=B.ASSET_CODE
        AND    A.UL_CODE=B.UL_CODE) D,
       (SELECT A.ASSET_CODE,
               A.UL_CODE,
               A.GREEK_VALUE TREE_GAMMA,
               B.GREEK_VALUE SM_GAMMA
        FROM   SPS.CL_DAILY_GREEKS A,
               SPS.CL_DAILY_GREEKS B
        WHERE  A.EVAL_DATE=:tdate
        AND    A.EVAL_DATE=B.EVAL_DATE
        AND    A.GREEK_CD='GAMMA'
        AND    B.GREEK_CD='HEDGE_TARGET_GAMMA'
        AND    A.ASSET_CODE=B.ASSET_CODE
        AND    A.UL_CODE=B.UL_CODE) G,
       (SELECT A.ASSET_CODE,
               A.UL_CODE,
               A.GREEK_VALUE VEGA
        FROM   SPS.CL_DAILY_GREEKS A
        WHERE  A.EVAL_DATE=:tdate
        AND    A.GREEK_CD='VEGA') V,
       (SELECT A.ASSET_CODE,
               A.UL_CODE,
               A.GREEK_VALUE THETA
        FROM   SPS.CL_DAILY_GREEKS A
        WHERE  A.EVAL_DATE=:tdate
        AND    TRIM(A.GREEK_CD)='THETA') T
WHERE  D.UL_CODE=S.CODE
AND D.ASSET_CODE=P.INDV_ISCD
AND D.ASSET_CODE=G.ASSET_CODE
AND D.ASSET_CODE=V.ASSET_CODE
AND D.ASSET_CODE=T.ASSET_CODE;

COMMIT;

--RAS.IF_OTC_TEMPLATE_FACTOR
DELETE FROM RAS.IF_OTC_TEMPLATE_FACTOR WHERE TDATE=:tdate AND CODE IN (SELECT A1.INDV_ISCD
        FROM   BSYS.TBSIMO201M00@GDW A1,
               BSYS.TBFNOM021L00@GDW A2,
               BSYS.TBSIMO203L00@GDW A3
        WHERE  A1.PROD_KIND_CODE = '02'
        AND    A1.INDV_ISCD = A2.ISCD
        AND    A1.INDV_ISCD = A3.INDV_ISCD
        AND    A2.STND_DATE = :tdate
        AND    A2.RMND_QTY > 0
        AND    A1.MTRT_DATE >= :tdate
        AND    A1.PROD_FNCD = 'SV001');

INSERT INTO RAS.IF_OTC_TEMPLATE_FACTOR
 (TDATE, CODE,
               FACTORID,
               FACTORPRICE,
               FACTORFIRSTSENSE,
               FACTORSECONDSENSE,
               FACTORVEGA,               
               WORK_TIME,
               WORK_TERMINAL,
               WORK_MEMBER)
SELECT TDATE,
       STK_CODE CODE,
       BASEASSET_CODE FACTORID,
       CLOSE_AMT FACTORPRICE,
       DELTA FACTORFIRSTSENSE,
       GAMMA FACTORSECONDSENSE,
       VEGA FACTORVEGA,
       SYSDATE WORK_TIME,
       'SQL' WORK_TERMINAL,
       'SQL' WORK_MEMBER
FROM   RCS.PML_GREEK
WHERE  TDATE=:tdate
AND    STK_CODE IN (SELECT A1.INDV_ISCD
        FROM   BSYS.TBSIMO201M00@GDW A1,
               BSYS.TBFNOM021L00@GDW A2,
               BSYS.TBSIMO203L00@GDW A3
        WHERE  A1.PROD_KIND_CODE = '02'
        AND    A1.INDV_ISCD = A2.ISCD
        AND    A1.INDV_ISCD = A3.INDV_ISCD
        AND    A2.STND_DATE = :tdate
        AND    A2.RMND_QTY > 0
        AND    A1.MTRT_DATE >= :tdate
        AND    A1.PROD_FNCD = 'SV001');

--RAS.IF_OTC_TEMPLATE_DATA
DELETE FROM RAS.IF_OTC_TEMPLATE_DATA WHERE TDATE=:tdate AND CODE IN (SELECT A1.INDV_ISCD
        FROM   BSYS.TBSIMO201M00@GDW A1,
               BSYS.TBFNOM021L00@GDW A2,
               BSYS.TBSIMO203L00@GDW A3
        WHERE  A1.PROD_KIND_CODE = '02'
        AND    A1.INDV_ISCD = A2.ISCD
        AND    A1.INDV_ISCD = A3.INDV_ISCD
        AND    A2.STND_DATE = :tdate
        AND    A2.RMND_QTY > 0
        AND    A1.MTRT_DATE >= :tdate
        AND    A1.PROD_FNCD = 'SV001');

INSERT INTO RAS.IF_OTC_TEMPLATE_DATA (TDATE, CODE, ENDPRICE) 
SELECT STND_DATE TDATE,
       ISCD CODE,
       ESTI_AFT_PBNT_SDPR ENDPRICE
FROM  BSYS.TBSIMO107L00@GDW
WHERE  STND_DATE=:tdate
AND    ISCD IN (SELECT A1.INDV_ISCD
        FROM   BSYS.TBSIMO201M00@GDW A1,
               BSYS.TBFNOM021L00@GDW A2,
               BSYS.TBSIMO203L00@GDW A3
        WHERE  A1.PROD_KIND_CODE = '02'
        AND    A1.INDV_ISCD = A2.ISCD
        AND    A1.INDV_ISCD = A3.INDV_ISCD
        AND    A2.STND_DATE = :tdate
        AND    A2.RMND_QTY > 0
        AND    A1.MTRT_DATE >= :tdate
        AND    A1.PROD_FNCD = 'SV001');

COMMIT;