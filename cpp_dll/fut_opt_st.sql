SELECT A.FNCD,
       L.STND_ISCD,
       L.BLMB_CODE,
       L.KOR_ISNM,
       TRIM(L.UNAS_ISCD) UNAS_ISCD,
       DECODE(L.FTOP_TYPE_CODE, 1, 'F', 2, 'C', 3, 'P', 'N/A') FTOP_TYPE_CODE,
       L.EXER_PRC,
       TO_CHAR(TO_DATE(L.LAST_TR_DATE, 'YYYYMMDD'), 'YYYY-MM-DD') LAST_TR_DATE,
       TO_CHAR(TO_DATE(L.LAST_STLM_DATE, 'YYYYMMDD'), 'YYYY-MM-DD') LAST_STLM_DATE,
       L.CRCD,
       NVL(FX.ENDPRICE, 1) FX,
       L.TR_MLTL,
       A.NIGHT_YN,
       A.RMND_QTY,
       FT.ENDPRICE PRICE_EXC, --거래소 정산가
       ROUND(SE.GREEK_VALUE/NVL(FX.ENDPRICE, 1), 6) PRICE_ED --이론가
FROM   (SELECT ISCD,
               FNCD,
               NIGHT_YN,
               SUM(DECODE(PSTN_CLS_CODE, 1, -1, 1)*RMND_QTY) RMND_QTY
        FROM   (
                --KRX
                SELECT ISCD,
                       FNCD,
                       'N' NIGHT_YN,
                       PSTN_CLS_CODE,
                       RMND_QTY
                FROM   BSYS.TBFNEM007L00@GDW
                WHERE  STND_DATE = :tdate
                --해외
                UNION ALL
                SELECT ISCD,
                       FNCD,
                       'N' NIGHT_YN,
                       PSTN_CLS_CODE,
                       RMND_QTY
                FROM   BSYS.TBFNEM017L00@GDW
                WHERE  STND_DATE = :tdate                
                --야간거래
                UNION ALL
                SELECT ASSET_CODE ISCD,
                       FUND_CODE FNCD,
                       NIGHT_YN,
                       '-1' PSTN_CLS_CODE,
                       QTY RMND_QTY
                FROM   SPS.F_DERIVATIVES_DEAL
                WHERE  TRADE_DATE =:tdate
                AND    NIGHT_YN ='Y'
                AND    CONFIRM_YN ='Y'                
                )
        WHERE RMND_QTY <> 0
        GROUP BY ISCD, FNCD, NIGHT_YN ) A, --잔고
       (--KRX
        SELECT A.ISCD,
               B.STND_ISCD,
               A.ISCD BLMB_CODE,
               B.KOR_ISNM,
               A.UNAS_ISCD, --기초자산 종목코드
               A.SCRT_GRP_CLS_CODE, --증권그룹구분
               A.FTOP_TYPE_CODE, --장내파생유형
               A.EXER_PRC, --행사가격
               A.TR_MLTL, --거래승수
               A.LAST_TR_DATE, --최종거래일자
               A.LAST_STLM_DATE, --최종결제일자
               'KRW' CRCD,
               NULL PROD_LKNG_CODE,
               NULL OVRS_LSTN_EXCH_CODE
        FROM   BSYS.TBSIMD100M00@GDW A,
               BSYS.TBSIMM100M00@GDW B
        WHERE  A.LAST_TR_DATE >= :tdate
        AND    A.ISCD=B.ISCD
        --해외
        UNION ALL
        SELECT A.ISCD,
               B.STND_ISCD,
               A.BLMB_CODE,
               B.KOR_ISNM,
               B.UNAS_ISCD,
               A.SCRT_GRP_CLS_CODE, --증권그룹구분
               A.FTOP_TYPE_CODE, --장내파생유형
               A.EXER_PRC, --행사가격
               A.TR_MLTL, --거래승수
               A.LAST_TR_DATE, --최종거래일자
               A.LAST_STLM_DATE, --최종결제일자
               A.CRCD, --통화코드
               A.PROD_LKNG_CODE, --상품기호
               A.OVRS_LSTN_EXCH_CODE --해외상장거래소
        FROM   BSYS.TBSIMD200M00@GDW A,
               BSYS.TBSIMM100M00@GDW B
        WHERE  A.LAST_STLM_DATE >= :tdate
        AND    A.BLMB_CODE IS NOT NULL
        AND    A.ISCD=B.ISCD ) L, --종목정보
       (SELECT DISTINCT FNCD FROM BSYS.TBFNPA001M00@GDW WHERE MANG_DPCD = '351') F,
       (SELECT CODE, ENDPRICE FROM RCS.PML_FX_DATA_ST WHERE TDATE=:tdate AND SCENARIOID=:scenarioid) FX,
       (SELECT CODE,
               ENDPRICE
        FROM   RAS.BL_FUTURES_DATA
        WHERE  TDATE = :tdate
        UNION ALL
        SELECT CODE, ENDPRICE FROM
        (SELECT CODE,
               ENDPRICE
        FROM   RAS.IF_FUTURES_DATA
        WHERE  TDATE = :tdate
        UNION ALL
        SELECT CODE,
               ENDPRICE
        FROM   RAS.IF_OPTION_DATA
        WHERE  TDATE = :tdate)
        WHERE CODE NOT IN (SELECT CODE
        FROM   RAS.BL_FUTURES_DATA
        WHERE  TDATE = :tdate)) FT,
       (SELECT ASSET_CODE,
               GREEK_VALUE
        FROM   SPT.DAILY_CLOSING_THEO
        WHERE  GREEK_CD = 'VALUE'
        AND    EVAL_DATE = :tdate) SE
WHERE A.FNCD = F.FNCD
AND A.ISCD = L.ISCD
AND L.STND_ISCD = FT.CODE(+)
AND L.STND_ISCD = SE.ASSET_CODE(+)
AND CONCAT('FXKRW', L.CRCD) = FX.CODE(+)
ORDER BY L.CRCD, 5, 6, 8, L.EXER_PRC