SELECT A.FNCD, B.ISCD,
        B.BLMB_CODE,
        C.KOR_ISSU_ABWR_NAME,
        DECODE(FTOP_TYPE_CODE,1,'F',2,'C',3,'P','N/A'),
        EXER_PRC,
        trim(C.UNAS_ISCD) UNAS_ISCD,
        CRCD,
        TR_MLTL,
        TO_CHAR(TO_DATE(B.LAST_TR_DATE,'YYYYMMDD'),'YYYY-MM-DD') LAST_TR_DATE,
        TO_CHAR(TO_DATE(B.LAST_STLM_DATE,'YYYYMMDD'),'YYYY-MM-DD') LAST_STLM_DATE,
        A.QTY,
        FX.ENDPRICE FX,
        FT.ENDPRICE EXC,
        round(SE.GREEK_VALUE/FX.ENDPRICE,6) PRICE_ED,
        round(SR.GREEK_VALUE/FX.ENDPRICE,6) PRICE_RM
 FROM   (SELECT ISCD, FNCD, SUM(QTY) QTY FROM (SELECT ISCD, FNCD,
                SUM(DECODE(PSTN_CLS_CODE, 1, -1, 1)*RMND_QTY) QTY
         FROM   BSYS.TBFNEM017L00@GDW
         WHERE  STND_DATE = :tdate
         AND    (FNCD LIKE 'SC%' OR FNCD LIKE 'SA%' OR FNCD LIKE 'SF%' OR FNCD LIKE 'SV%' OR FNCD LIKE 'VP%')
         AND    RMND_QTY <> 0
         GROUP BY ISCD, FNCD 
--    If night_trade_included = True Then
         UNION ALL 
         SELECT ASSET_CODE ISCD, FUND_CODE FNCD, QTY FROM SPS.F_DERIVATIVES_DEAL WHERE TRADE_DATE=:tdate AND NIGHT_YN='Y' AND CONFIRM_YN='Y'
--    End If
         ) GROUP BY ISCD, FNCD) A,
        (SELECT ISCD, PROD_LKNG_CODE, LAST_TR_DATE, LAST_STLM_DATE, CRCD, TR_MLTL, EXER_PRC, FTOP_TYPE_CODE, BLMB_CODE
         FROM   BSYS.TBSIMD200M00@GDW
         WHERE  LAST_TR_DATE >= :tdate
         AND    BLMB_CODE IS NOT NULL) B,
         BSYS.TBSIMM100M00@GDW C,

        (select CODE, ENDPRICE from RAS.IF_FX_DATA where tdate=:tdate) FX,

--        (select * from RCS.PML_FX_DATA_ST where tdate=:tdate and scenarioid='" & scenarioID & "') FX,

        (select CODE, ENDPRICE from RAS.BL_FUTURES_DATA where tdate=:tdate) FT,
        (select ASSET_CODE, GREEK_CD, GREEK_VALUE from spt.daily_closing_theo where GREEK_CD='VALUE' AND EVAL_DATE=:tdate) SE,
        (select ASSET_CODE, GREEK_CD, GREEK_VALUE from spt.daily_closing_theo_rm where GREEK_CD='VALUE' AND EVAL_DATE=:tdate) SR
 WHERE  A.ISCD = B.ISCD
 AND    B.ISCD = C.ISCD
 AND    B.PROD_LKNG_CODE = '0002'
 AND    B.ISCD = FT.CODE(+)
 AND    B.ISCD = SE.ASSET_CODE(+)
 AND    B.ISCD = SR.ASSET_CODE(+)
 AND    FX.CODE(+)=concat('FXKRW',CRCD)
 AND    (trim(C.UNAS_ISCD) in ('SX5E','SPX','NKY') OR trim(C.UNAS_ISCD) like 'ES%')
 UNION ALL
 SELECT A.FNCD, B.ISCD,
        B.BLMB_CODE,
        C.KOR_ISSU_ABWR_NAME,
        DECODE(FTOP_TYPE_CODE,1,'F',2,'C',3,'P','N/A'),
        EXER_PRC,
        trim(C.UNAS_ISCD) UNAS_ISCD,
        CRCD,
        TR_MLTL,
        TO_CHAR(TO_DATE(B.LAST_TR_DATE,'YYYYMMDD'),'YYYY-MM-DD') LAST_TR_DATE,
        TO_CHAR(TO_DATE(B.LAST_STLM_DATE,'YYYYMMDD'),'YYYY-MM-DD') LAST_STLM_DATE,
        A.QTY,
        FX.ENDPRICE FX,
        FT.ENDPRICE EXC,
        round(SE.GREEK_VALUE/FX.ENDPRICE,6) PRICE_ED,
        round(SR.GREEK_VALUE/FX.ENDPRICE,6) PRICE_RM
 FROM   (SELECT ISCD, FNCD, SUM(QTY) QTY FROM (SELECT ISCD, FNCD,
                SUM(DECODE(PSTN_CLS_CODE, 1, -1, 1)*RMND_QTY) QTY
         FROM   BSYS.TBFNEM017L00@GDW
         WHERE  STND_DATE = :tdate
         AND    (FNCD LIKE 'SC%' OR FNCD LIKE 'SA%' OR FNCD LIKE 'SF%' OR FNCD LIKE 'SV%' OR FNCD LIKE 'VP%')
         AND    RMND_QTY <> 0
         GROUP BY ISCD, FNCD 
--    If night_trade_included = True Then
         UNION ALL 
         SELECT ASSET_CODE ISCD, FUND_CODE FNCD, QTY FROM SPS.F_DERIVATIVES_DEAL WHERE TRADE_DATE=:tdate AND NIGHT_YN='Y' AND CONFIRM_YN='Y'
--    End If
         ) GROUP BY ISCD, FNCD) A,
        (SELECT ISCD, PROD_LKNG_CODE, LAST_TR_DATE, LAST_STLM_DATE, CRCD, TR_MLTL, EXER_PRC, FTOP_TYPE_CODE, BLMB_CODE
         FROM   BSYS.TBSIMD200M00@GDW
         WHERE  LAST_TR_DATE > :tdate
         AND    BLMB_CODE IS NOT NULL) B,
         BSYS.TBSIMM100M00@GDW C,

        (select CODE, ENDPRICE from RAS.IF_FX_DATA where tdate=:tdate) FX,

--        (select * from RCS.PML_FX_DATA_ST where tdate=:tdate and scenarioid='" & scenarioID & "') FX,

        (select CODE, ENDPRICE from RAS.BL_FUTURES_DATA where tdate=:tdate) FT,
        (select ASSET_CODE, GREEK_CD, GREEK_VALUE from spt.daily_closing_theo where GREEK_CD='VALUE' AND EVAL_DATE=:tdate) SE,
        (select ASSET_CODE, GREEK_CD, GREEK_VALUE from spt.daily_closing_theo_rm where GREEK_CD='VALUE' AND EVAL_DATE=:tdate) SR
 WHERE  A.ISCD = B.ISCD
 AND    B.ISCD = C.ISCD
 AND    B.PROD_LKNG_CODE = '0002'
 AND    B.ISCD = FT.CODE(+)
 AND    B.ISCD = SE.ASSET_CODE(+)
 AND    B.ISCD = SR.ASSET_CODE(+)
 AND    FX.CODE(+)=concat('FXKRW',CRCD)
 AND    trim(C.UNAS_ISCD) ='HSCEI'
 UNION ALL
 SELECT NVL(A.FNCD, ' ') FNCODE,
        NVL(C.STND_ISCD, ' ') ISCD,
        null BLMB_CODE,
        C.KOR_ISSU_ABWR_NAME,
        DECODE(B.FTOP_TYPE_CODE, '1','F','2', 'C', '3', 'P', 'N/A') FCP,
        NVL(B.EXER_PRC, 0) EXER_PRC,
        NVL(B.UNAS_ISCD, ' ') UNAS_ISCD,
        'KRW' CRCD,
        NVL(B.TR_MLTL, 0) TR_MLTL,
        TO_CHAR(TO_DATE(B.LAST_TR_DATE,'YYYYMMDD'),'YYYY-MM-DD') LAST_TR_DATE,
        TO_CHAR(TO_DATE(B.LAST_STLM_DATE,'YYYYMMDD'),'YYYY-MM-DD') LAST_STLM_DATE,
        NVL(A.RMND_QTY, 0)*DECODE(A.PSTN_CLS_CODE, '1', -1, 1) QTY,
        1 FX,
        FT.endprice EXC,
        round(SE.GREEK_VALUE,6) PRICE_ED,
        round(SR.GREEK_VALUE,6) PRICE_RM
 FROM   BSYS.TBFNEM007L00@GDW A,
        BSYS.TBSIMD100M00@GDW B,
        BSYS.TBSIMM100M00@GDW C,
        (select TDATE, CODE, ENDPRICE from ras.if_futures_data where tdate=:tdate union all
        select TDATE, CODE, ENDPRICE from ras.if_option_data where tdate=:tdate) FT,
        (select ASSET_CODE, GREEK_CD, GREEK_VALUE from spt.daily_closing_theo where GREEK_CD='VALUE' AND EVAL_DATE=:tdate) SE,
        (select ASSET_CODE, GREEK_CD, GREEK_VALUE from spt.daily_closing_theo_rm where GREEK_CD='VALUE' AND EVAL_DATE=:tdate) SR
 WHERE  A.STND_DATE=:tdate
 AND    B.LAST_TR_DATE>=:tdate
 AND    B.ISCD=A.ISCD
 AND    C.ISCD=A.ISCD
 AND    A.RMND_QTY<>0
 AND    C.STND_ISCD = FT.CODE(+)
 AND    C.STND_ISCD = SE.ASSET_CODE(+)
 AND    C.STND_ISCD = SR.ASSET_CODE(+)
 AND    (FNCD LIKE 'SC%' OR FNCD LIKE 'SA%' OR FNCD LIKE 'SF%' OR FNCD LIKE 'SV%' OR FNCD LIKE 'VP%')
