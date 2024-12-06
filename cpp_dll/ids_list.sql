-- ELS : FO = IO
-- ELS_H : FO = IO(FRN) + IO(ESWAP)
-- FRN_T : FO = IO(FRN)
-- ESWAP_T : FO = IO(ESWAP)
-- ��������

--�Ǹ� �ŷ����� ������ ���� ������ 
-- 1) ��ҵ� �ŷ� �Ǵ�
-- 2) �ſ� ������ �Ǵ� ���ο����� ���� �������� �ŷ�
--    ����. DLB(hi-five) FO261420A95D(�Ｚ���� ��Ź)

-- ��ȯ����(RDMP_CMPT_YN) �ʵ�� ����Ȯ�ϹǷ� ������� ���� -> �������(PBLC_STCK_QTY) ���
--��. OTC����(04 swap)�� �ŷ���� �ʵ尡 ���� NVL(FD.RDMP_CMPT_YN, ' ') NOT IN ('C') ���ǿ��� ���͵��� ����.
--��. (07 ELS) ���� �߿��� ��ҵǾ������� RDMP_CMPT_YN �ʵ忡 �ݿ����� ���� ��찡 ����.

-- ��ŷ����
-- 1. FO420920B27T ���� IO122721B31T �ſ�OTC931ȣBXXXX ... �ݵ������ڵ尡 FO420520C27T�� �߸� ��. �����ʿ�
-- 2. �ſ�OTC931ȣA ���� FO440121A26T �ſ�OTC931ȣA[F��] ... swap-swap �ε� F�� �ڵ� ������. ���εǴ� IO �ڵ� ����. �����ʿ�
-- 3. �ſ�OTC1027ȣA ���� FO440122A27T �ſ�OTC1027ȣA[F��] ... swap-swap �ε� F�� �ڵ� ������. ���εǴ� IO �ڵ� ����. �����ʿ�
-- 4. �ſ�OTC897ȣC ���� FO420520D97T �ſ�OTC897ȣCT�� 3�� IO ���� ���ε�. ... ��� ����. ELN���� swap-swap ���� �Һи�. FO440120C97T �ſ�OTC897ȣC[F��] ... ���εǴ� IO �ڵ� ����. �����ʿ�
-- 5. �ſ�OTC941ȣA ���� FO420421A41T �ſ�OTC941ȣA[����] ... ��������ݵ������ڵ� ���� FO420521B40T ->FO420521C41T


SELECT 
        EE.HEDG_TR_TYPE_CODE,
        --EE.PROD_CLS_CODE,
        --H.CODE_VAL_NAME "��������",
        N.CODE_VAL_NAME PROD_CLS_CODE_D,
        EE.STND_ISCD,
        EE.PROD_FNCD,        
        EE.ELS_FO_D,
        EE.ELS_IO_D,        
        EE.ELS_FO_H,
        N2.CODE_VAL_NAME PROD_CLS_CODE_H,
        EE.ESWAP_IO_H,
        EE.FRN_IO_H,        
        'SF003' ESWAP_FNCD,
        EE.ELS_FO_H ESWAP_FO_T, --H�ڵ� �ݺ�
        EE.ESWAP_IO_H ESWAP_IO_T, --H�ڵ� �ݺ�    
        EE.FRN_FO_T,
        EE.FRN_IO_T,        
        EE.KOR_ISNM,
        EE.STLM_CRCD,
        EE.REAL_PBLC_FCAM,
        EE.DEAL_CLS_CODE,
        EE.PBLC_STCK_QTY,
        QT.RMND_QTY,
        QT.RMND_QTY * EE.REAL_PBLC_FCAM  "���ݾ�CCY",
        QT.RMND_QTY * EE.REAL_PBLC_FCAM * RMS.GET_FXRATE(:tdate, 'FXKRW' || EE.STLM_CRCD)  "���ݾ�KRW",
        TO_CHAR(TO_DATE(EE.FRST_STND_PRC_FIN_DATE, 'YYYYMMDD'), 'YYYY-MM-DD') FRST_STND_PRC_FIN_DATE,
        TO_CHAR(TO_DATE(EE.PBLC_DATE, 'YYYYMMDD'), 'YYYY-MM-DD') PBLC_DATE, --"������", --[v]
        TO_CHAR(TO_DATE(EE.MTRT_DATE, 'YYYYMMDD'), 'YYYY-MM-DD') MTRT_DATE, --"������", --[v]
        TO_DATE(EE.MTRT_DATE, 'YYYYMMDD') - TO_DATE(:tdate, 'YYYYMMDD') "�����ϼ�",
        EE.FUND_PBLC_UNPR, --���డ(������ ���, �ټ��� ��Ʈ �������� �߸� �Էµ�)
        EE.PERC_APLY_THPR*EE.REAL_PBLC_FCAM, --��ΰ�
        HF.HEDGE_BUF,
        HF.HEDGE_BUF_H,
        M.CODE_VAL_NAME "�����ȯ����", --[v]
        U.CODE_VAL || ' ' || U.CODE_VAL_NAME,
        UA.NUM_UA,
        UA.UNAS_NAME1 "�����ڻ�1", --[v]
       UA.UNAS_NAME2 "�����ڻ�2", --[v]
       UA.UNAS_NAME3 "�����ڻ�3", --[v]
       UA.UNAS_NAME4 "�����ڻ�4", --[v]
       UA.UNAS_ISCD1 "�����ڻ��ڵ�1", --[v]
       UA.UNAS_ISCD2 "�����ڻ��ڵ�2", --[v]
       UA.UNAS_ISCD3 "�����ڻ��ڵ�3", --[v]
       UA.UNAS_ISCD4 "�����ڻ��ڵ�4", --[v]    
       UA.UNAS_INTL_PRC1 "���ʱ��ذ�1", --[v]
       UA.UNAS_INTL_PRC2 "���ʱ��ذ�2", --[v]
       UA.UNAS_INTL_PRC3 "���ʱ��ذ�3", --[v]
       UA.UNAS_INTL_PRC4 "���ʱ��ذ�4", --[v]
       --AC.NUM_AC,
       AC.UNAS_SDRT1 "�����ȯ������1", --[v]
       AC.UNAS_SDRT2 "�����ȯ������2", --[v]
       AC.UNAS_SDRT3 "�����ȯ������3", --[v]
       AC.UNAS_SDRT4 "�����ȯ������4", --[v]
       AC.UNAS_SDRT5 "�����ȯ������5", --[v]
       AC.UNAS_SDRT6 "�����ȯ������6", --[v]
       AC.UNAS_SDRT7 "�����ȯ������7", --[v]
       AC.UNAS_SDRT8 "�����ȯ������8", --[v]
       AC.UNAS_SDRT9 "�����ȯ������9", --[v]
       AC.UNAS_SDRT10 "�����ȯ������10", --[v]
       AC.UNAS_SDRT11 "�����ȯ������11", --[v]
       AC.UNAS_SDRT12 "�����ȯ������12", --[v]
--       AC.AVRG_APLY_YN1 "��հ���1", --[v]
--       AC.AVRG_APLY_YN2 "��հ���2", --[v]
--       AC.AVRG_APLY_YN3 "��հ���3",--[v]
--       AC.AVRG_APLY_YN4 "��հ���4",--[v]
--       AC.AVRG_APLY_YN5 "��հ���5",--[v]
--       AC.AVRG_APLY_YN6 "��հ���6",--[v]
--       AC.AVRG_APLY_YN7 "��հ���7",--[v]
--       AC.AVRG_APLY_YN8 "��հ���8",--[v]
--       AC.AVRG_APLY_YN9 "��հ���9",--[v]
--       AC.AVRG_APLY_YN10 "��հ���10",--[v]
--       AC.AVRG_APLY_YN11 "��հ���11",--[v]
--       AC.AVRG_APLY_YN12 "��հ���12",--[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE1, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����1", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE2, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����2", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE3, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����3", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE4, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����4", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE5, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����5", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE6, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����6", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE7, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����7", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE8, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����8", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE9, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����9", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE10, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����10", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE11, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����11", --[v]
       TO_CHAR(TO_DATE(AC.TRTH_CLRD_DTRM_DATE12, 'YYYYMMDD'), 'YYYY-MM-DD') "��ȯ����12", --[v]   
       AC.CLRD_ERT1 "��ȯ��1",--[v]
       AC.CLRD_ERT2 "��ȯ��2",--[v]
       AC.CLRD_ERT3 "��ȯ��3",--[v]
       AC.CLRD_ERT4 "��ȯ��4",--[v]
       AC.CLRD_ERT5 "��ȯ��5",--[v]
       AC.CLRD_ERT6 "��ȯ��6",--[v]
       AC.CLRD_ERT7 "��ȯ��7",--[v]
       AC.CLRD_ERT8 "��ȯ��8",--[v]
       AC.CLRD_ERT9 "��ȯ��9",--[v]
       AC.CLRD_ERT10 "��ȯ��10",--[v]
       AC.CLRD_ERT11 "��ȯ��11",--[v]
       AC.CLRD_ERT12 "��ȯ��12",--[v]
       EE.KI "���κ�����",--[v]
       EE.HIT_DATE "������ġ��",      --[v] 
       MTH.BONS_CUPN_STA_SDRT "�����޺�����",--[v]
       MTH.BONS_CUPN_INRT "��������",--[v]
       --E.EE_NUM,
       E.CLRD_BARR_VAL1 "���ڵ庣����1",--[v]
       E.CLRD_BARR_VAL2 "���ڵ庣����2",--[v]
       E.CLRD_BARR_VAL3 "���ڵ庣����3",--[v]
       TO_CHAR(TO_DATE(E.CALC_FIN_DATE1, 'YYYYMMDD'), 'YYYY-MM-DD') "���ڵ屸��1", --[v]   
       TO_CHAR(TO_DATE(E.CALC_FIN_DATE2, 'YYYYMMDD'), 'YYYY-MM-DD') "���ڵ屸��2", --[v]   
       TO_CHAR(TO_DATE(E.CALC_FIN_DATE3, 'YYYYMMDD'), 'YYYY-MM-DD') "���ڵ屸��3", --[v]   
       E.CLRD_INRT1 "���ڵ��ȯ��1",--[v]
       E.CLRD_INRT2 "���ڵ��ȯ��2",--[v]
       E.CLRD_INRT3 "���ڵ��ȯ��3", --[v]
       E.CLRD_BARR_HIT_YN1 "���ڵ���ġ1",--[v]
       E.CLRD_BARR_HIT_YN2 "���ڵ���ġ2",--[v]
       E.CLRD_BARR_HIT_YN3 "���ڵ���ġ3",--[v]
       NVL(CP_D.TR_CNRP_NAME, EE.SALE_CHAN_EXPL) SALE_CHAN_NAME,
       CP_H.TR_CNRP_NAME,
       EE.PRC_APLY_CLS_CODE
FROM   
    (
    SELECT M.*,
           --E.PROD_CLS_CODE_H,
           E.ESWAP_FNCD,
           E.ESWAP_FO_T,
           E.ESWAP_IO_T,
           F.FRN_FO_T,
           F.FRN_IO_T
    FROM   (SELECT FD.HEDG_TR_TYPE_CODE,
                   FD.PROD_CLS_CODE,
                   FD.PROD_FNCD, 
                   I.STND_ISCD,
                   FD.OTC_FUND_ISCD ELS_FO_D,
                   ID.INDV_ISCD ELS_IO_D,
                   IH27.PROD_CLS_CODE PROD_CLS_CODE_H, 
                   FD.HEDG_ISCD ELS_FO_H,
                   ID.HEDG_ISCD ESWAP_IO_H,
                   IH.INDV_ISCD FRN_IO_H,
                   I.KOR_ISNM,
                   ID.CLRD_TYPE_CODE,
                   FD.STLM_CRCD,
                   FD.RDMP_CMPT_YN,
                   ID.DEAL_CLS_CODE,
                   ID.PRC_APLY_CLS_CODE, 
                   FD.REAL_PBLC_FCAM, 
                   FD.PBLC_STCK_QTY,
                   FD.FUND_PBLC_UNPR,
                   ID.PERC_APLY_THPR,
                   FD.FRST_STND_PRC_FIN_DATE,
                   ID.PBLC_DATE, 
                   ID.MTRT_DATE,               
                   ID.CLRD_DATE,
                   ID.UNAS_CHOC_MTHD_CODE, 
                   DECODE(ID.BARR_VAL1,0,'',ID.BARR_VAL1) KI,
                   TO_CHAR(TO_DATE(ID.BARR_HIT_DATE, 'YYYYMMDD'), 'YYYY-MM-DD') HIT_DATE,  
                   FD.SALE_CHAN_EXPL,
                   FD.HEDG_TR_CNRP_EXPL
            FROM   BSYS.TBSIMM100M00@GDW I, --���� ����
                   BSYS.TBSIMO100M00@GDW FD, --OTC �ݵ����� ����: SP�� ELS �ŵ�
                   BSYS.TBSIMO201M00@GDW ID, --OTC �������� ����: SP�� ELS �ŵ�                   
                   BSYS.TBSIMO201M00@GDW IH, --OTC �������� ����: SP�� FRN �ż�
                                      BSYS.TBSIMO201M00@GDW IH27 --OTC �������� ����: SP�� Equity Swap �ż�    
            WHERE  FD.MTRT_DATE > :tdate
            AND    NVL(FD.RDMP_CMPT_YN,' ') NOT IN ('C')
            AND    FD.PBLC_STCK_QTY <> 0
            AND    (FD.PBLC_DATE <= :tdate
                    OR     FD.FRST_STND_PRC_FIN_DATE <= :tdate)            
            AND    FD.OTC_FUND_ISCD = I.ISCD
            AND    ID.OTC_FUND_ISCD = FD.OTC_FUND_ISCD
            AND    ID.PROD_KIND_CODE = '27'
            AND    IH.OTC_FUND_ISCD(+) = FD.HEDG_ISCD
            AND    IH.PROD_KIND_CODE(+) = '00'
            AND    IH27.OTC_FUND_ISCD(+) = FD.HEDG_ISCD
            AND    IH27.PROD_KIND_CODE(+) = '27'
            ) M,
           (SELECT FD.OTC_FUND_ISCD ELS_FO_D,
                   FD.HEDG_ISCD,
                   ITE.PROD_FNCD ESWAP_FNCD,
                   ITE.OTC_FUND_ISCD ESWAP_FO_T,
                   ITE.INDV_ISCD ESWAP_IO_T
            FROM   BSYS.TBSIMO100M00@GDW FD, --OTC �ݵ����� ����: SP�� ELS �ŵ�
                   BSYS.TBSIMO100M00@GDW FT, --OTC �ݵ����� ����: E��, F�� �ż� ���ΰŷ�
                   BSYS.TBSIMO201M00@GDW ID, --OTC �������� ����: SP�� ELS �ŵ�       
                   BSYS.TBSIMO201M00@GDW ITE --OTC �������� ����: E�� Equity Swap �ŵ�
            WHERE  FD.MTRT_DATE > :tdate
            AND    NVL(FD.RDMP_CMPT_YN,' ') NOT IN ('C')
            AND    FD.PBLC_STCK_QTY <> 0
            AND    (FD.PBLC_DATE <= :tdate OR FD.FRST_STND_PRC_FIN_DATE <= :tdate)
            AND    ID.OTC_FUND_ISCD = FD.OTC_FUND_ISCD
            AND    ID.PROD_KIND_CODE = '27' --
            AND    FD.HEDG_ISCD = FT.HEDG_ISCD
            AND    FD.OTC_FUND_ISCD <> FT.OTC_FUND_ISCD
            AND    ITE.OTC_FUND_ISCD = FT.OTC_FUND_ISCD
            AND    ITE.PROD_KIND_CODE = '27'
                   ) E,
           (SELECT FD.OTC_FUND_ISCD ELS_FO_D,
                   FD.HEDG_ISCD,                   
                   ITF.OTC_FUND_ISCD FRN_FO_T,
                   ITF.INDV_ISCD FRN_IO_T
            FROM   BSYS.TBSIMO100M00@GDW FD, --OTC �ݵ����� ����: SP�� ELS �ŵ�
                   BSYS.TBSIMO100M00@GDW FT, --OTC �ݵ����� ����: E��, F�� �ż� ���ΰŷ�
                   BSYS.TBSIMO201M00@GDW ID, --OTC �������� ����: SP�� ELS �ŵ�       
                   BSYS.TBSIMO201M00@GDW ITF --OTC �������� ����: F�� FRN �ŵ�(��� unfunded swap)
            WHERE  FD.MTRT_DATE > :tdate
            AND    NVL(FD.RDMP_CMPT_YN,' ') NOT IN ('C')
            AND    FD.PBLC_STCK_QTY <> 0
            AND    (FD.PBLC_DATE <= :tdate
                    OR     FD.FRST_STND_PRC_FIN_DATE <= :tdate)
            AND    ID.OTC_FUND_ISCD = FD.OTC_FUND_ISCD
            AND    ID.PROD_KIND_CODE = '27'
            AND    FD.HEDG_ISCD = FT.HEDG_ISCD
            AND    FD.OTC_FUND_ISCD <> FT.OTC_FUND_ISCD
            AND    ITF.OTC_FUND_ISCD = FT.OTC_FUND_ISCD
            AND    ITF.PROD_KIND_CODE = '00'
                   ) F
    WHERE  M.ELS_FO_D = E.ELS_FO_D(+)
    ANd    M.ELS_FO_D = F.ELS_FO_D(+)
    order by 1, 2, 3 
    ) EE,
    (SELECT INDV_ISCD,
            MAX(RNUM) NUM_UA,
               MAX(DECODE(RNUM,1, KOR_ISSU_ABWR_NAME)) UNAS_NAME1,
               MAX(DECODE(RNUM,2, KOR_ISSU_ABWR_NAME)) UNAS_NAME2,
               MAX(DECODE(RNUM,3, KOR_ISSU_ABWR_NAME)) UNAS_NAME3,
               MAX(DECODE(RNUM,4, KOR_ISSU_ABWR_NAME)) UNAS_NAME4,                 
               MAX(DECODE(RNUM,1, UNAS_ISCD)) UNAS_ISCD1,
               MAX(DECODE(RNUM,2, UNAS_ISCD)) UNAS_ISCD2,
               MAX(DECODE(RNUM,3, UNAS_ISCD)) UNAS_ISCD3,
               MAX(DECODE(RNUM,4, UNAS_ISCD)) UNAS_ISCD4,            
               MAX(DECODE(RNUM,1, UNAS_INTL_PRC)) UNAS_INTL_PRC1,
               MAX(DECODE(RNUM,2, UNAS_INTL_PRC)) UNAS_INTL_PRC2,
               MAX(DECODE(RNUM,3, UNAS_INTL_PRC)) UNAS_INTL_PRC3,
               MAX(DECODE(RNUM,4, UNAS_INTL_PRC)) UNAS_INTL_PRC4
        FROM(
        SELECT ROW_NUMBER() OVER(PARTITION BY E1.INDV_ISCD
                ORDER BY E1.UNAS_ISCD) AS RNUM,
               INDV_ISCD,
               UNAS_INTL_PRC,
               BARR_VAL, --�����ڻ꺰 KI barrier
               BARR_HIT_CLS_CODE,  --�����ڻ꺰 KI Hit flag
               DECODE(E1.UNAS_ISCD,'NIKKEI225','NKY',E1.UNAS_ISCD) UNAS_ISCD,
               KOR_ISSU_ABWR_NAME
        FROM   BSYS.TBSIMO202D00@GDW E1,
               BSYS.TBSIMM100M00@GDW E2
        WHERE  E1.UNAS_ISCD = E2.STND_ISCD) GROUP BY INDV_ISCD) UA, --�����ڻ�
       (SELECT M2.CODE_VAL,
               M2.CODE_VAL_NAME
        FROM   BSYS.TBCPPC001M02@GDW M1,
               BSYS.TBCPPC001M01@GDW M2
        WHERE  M1.META_TERM_PHSC_NAME = 'PROD_CLS_CODE'
        AND    M1.CODE_ID=M2.CODE_ID) N, --��ǰ����
        (SELECT M2.CODE_VAL,
               M2.CODE_VAL_NAME
        FROM   BSYS.TBCPPC001M02@GDW M1,
               BSYS.TBCPPC001M01@GDW M2
        WHERE  M1.META_TERM_PHSC_NAME = 'PROD_CLS_CODE'
        AND    M1.CODE_ID=M2.CODE_ID) N2, --��ǰ����
       (SELECT M2.CODE_VAL,
               M2.CODE_VAL_NAME
        FROM   BSYS.TBCPPC001M02@GDW M1,
               BSYS.TBCPPC001M01@GDW M2
        WHERE  M1.META_TERM_PHSC_NAME = 'CLRD_TYPE_CODE'
        AND    M1.CODE_ID=M2.CODE_ID) M, --�����ȯ����
               (SELECT M2.CODE_VAL,
               M2.CODE_VAL_NAME
        FROM   BSYS.TBCPPC001M02@GDW M1,
               BSYS.TBCPPC001M01@GDW M2
        WHERE  M1.META_TERM_PHSC_NAME = 'HEDG_TR_TYPE_CODE'
        AND    M1.CODE_ID=M2.CODE_ID) H, --�����ȯ����
       (SELECT M2.CODE_VAL,
               M2.CODE_VAL_NAME
        FROM   BSYS.TBCPPC001M02@GDW M1,
               BSYS.TBCPPC001M01@GDW M2
        WHERE  M1.META_TERM_PHSC_NAME = 'UNAS_CHOC_MTHD_CODE'
        AND    M1.CODE_ID=M2.CODE_ID) U, --�����ڻ�������
    BSYS.TBFNOM021L00@GDW QT,    
               (SELECT INDV_ISCD,
                       --MAX(SRNO) NUM_AC,
                       MAX(DECODE(SRNO,1, UNAS_SDRT1)) UNAS_SDRT1,
                       MAX(DECODE(SRNO,2, UNAS_SDRT1)) UNAS_SDRT2,
                       MAX(DECODE(SRNO,3, UNAS_SDRT1)) UNAS_SDRT3,
                       MAX(DECODE(SRNO,4, UNAS_SDRT1)) UNAS_SDRT4,
                       MAX(DECODE(SRNO,5, UNAS_SDRT1)) UNAS_SDRT5,
                       MAX(DECODE(SRNO,6, UNAS_SDRT1)) UNAS_SDRT6,
                       MAX(DECODE(SRNO,7, UNAS_SDRT1)) UNAS_SDRT7,
                       MAX(DECODE(SRNO,8, UNAS_SDRT1)) UNAS_SDRT8,
                       MAX(DECODE(SRNO,9, UNAS_SDRT1)) UNAS_SDRT9,                       
                       MAX(DECODE(SRNO,10, UNAS_SDRT1)) UNAS_SDRT10,                       
                       MAX(DECODE(SRNO,11, UNAS_SDRT1)) UNAS_SDRT11,                       
                       MAX(DECODE(SRNO,12, UNAS_SDRT1)) UNAS_SDRT12,                       
                       MAX(DECODE(SRNO,1, AVRG_APLY_YN)) AVRG_APLY_YN1,
                       MAX(DECODE(SRNO,2, AVRG_APLY_YN)) AVRG_APLY_YN2,
                       MAX(DECODE(SRNO,3, AVRG_APLY_YN)) AVRG_APLY_YN3,       
                       MAX(DECODE(SRNO,4, AVRG_APLY_YN)) AVRG_APLY_YN4,       
                       MAX(DECODE(SRNO,5, AVRG_APLY_YN)) AVRG_APLY_YN5,       
                       MAX(DECODE(SRNO,6, AVRG_APLY_YN)) AVRG_APLY_YN6,       
                       MAX(DECODE(SRNO,7, AVRG_APLY_YN)) AVRG_APLY_YN7,       
                       MAX(DECODE(SRNO,8, AVRG_APLY_YN)) AVRG_APLY_YN8,       
                       MAX(DECODE(SRNO,9, AVRG_APLY_YN)) AVRG_APLY_YN9,                       
                       MAX(DECODE(SRNO,10, AVRG_APLY_YN)) AVRG_APLY_YN10,                       
                       MAX(DECODE(SRNO,11, AVRG_APLY_YN)) AVRG_APLY_YN11,                       
                       MAX(DECODE(SRNO,12, AVRG_APLY_YN)) AVRG_APLY_YN12,                       
                       MAX(DECODE(SRNO,1, CLRD_ERT)) CLRD_ERT1,
                       MAX(DECODE(SRNO,2, CLRD_ERT)) CLRD_ERT2,
                       MAX(DECODE(SRNO,3, CLRD_ERT)) CLRD_ERT3,
                       MAX(DECODE(SRNO,4, CLRD_ERT)) CLRD_ERT4,
                       MAX(DECODE(SRNO,5, CLRD_ERT)) CLRD_ERT5,
                       MAX(DECODE(SRNO,6, CLRD_ERT)) CLRD_ERT6,
                       MAX(DECODE(SRNO,7, CLRD_ERT)) CLRD_ERT7,
                       MAX(DECODE(SRNO,8, CLRD_ERT)) CLRD_ERT8,
                       MAX(DECODE(SRNO,9, CLRD_ERT)) CLRD_ERT9,                      
                       MAX(DECODE(SRNO,10, CLRD_ERT)) CLRD_ERT10,                      
                       MAX(DECODE(SRNO,11, CLRD_ERT)) CLRD_ERT11,                      
                       MAX(DECODE(SRNO,12, CLRD_ERT)) CLRD_ERT12,                      
                       MAX(DECODE(SRNO,1, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE1,
                       MAX(DECODE(SRNO,2, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE2,
                       MAX(DECODE(SRNO,3, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE3,
                       MAX(DECODE(SRNO,4, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE4,
                       MAX(DECODE(SRNO,5, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE5,
                       MAX(DECODE(SRNO,6, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE6,
                       MAX(DECODE(SRNO,7, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE7,
                       MAX(DECODE(SRNO,8, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE8,
                       MAX(DECODE(SRNO,9, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE9,
                       MAX(DECODE(SRNO,10, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE10,
                       MAX(DECODE(SRNO,11, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE11,
                       MAX(DECODE(SRNO,12, TRTH_CLRD_DTRM_DATE)) TRTH_CLRD_DTRM_DATE12--,
--                       MAX(DECODE(SRNO,1, CLRD_DTRM_DATE)) CLRD_DTRM_DATE1,
--                       MAX(DECODE(SRNO,2, CLRD_DTRM_DATE)) CLRD_DTRM_DATE2,
--                       MAX(DECODE(SRNO,3, CLRD_DTRM_DATE)) CLRD_DTRM_DATE3,
--                       MAX(DECODE(SRNO,4, CLRD_DTRM_DATE)) CLRD_DTRM_DATE4,
--                       MAX(DECODE(SRNO,5, CLRD_DTRM_DATE)) CLRD_DTRM_DATE5,
--                       MAX(DECODE(SRNO,6, CLRD_DTRM_DATE)) CLRD_DTRM_DATE6,
--                       MAX(DECODE(SRNO,7, CLRD_DTRM_DATE)) CLRD_DTRM_DATE7,
--                       MAX(DECODE(SRNO,8, CLRD_DTRM_DATE)) CLRD_DTRM_DATE8,
--                       MAX(DECODE(SRNO,9, CLRD_DTRM_DATE)) CLRD_DTRM_DATE9,
--                       MAX(DECODE(SRNO,10, CLRD_DTRM_DATE)) CLRD_DTRM_DATE10,
--                       MAX(DECODE(SRNO,11, CLRD_DTRM_DATE)) CLRD_DTRM_DATE11,
--                       MAX(DECODE(SRNO,12, CLRD_DTRM_DATE)) CLRD_DTRM_DATE12,
--                       MAX(DECODE(SRNO,1, CLRD_DATE)) CLRD_DATE1,
--                       MAX(DECODE(SRNO,2, CLRD_DATE)) CLRD_DATE2,
--                       MAX(DECODE(SRNO,3, CLRD_DATE)) CLRD_DATE3,
--                       MAX(DECODE(SRNO,4, CLRD_DATE)) CLRD_DATE4,
--                       MAX(DECODE(SRNO,5, CLRD_DATE)) CLRD_DATE5,
--                       MAX(DECODE(SRNO,6, CLRD_DATE)) CLRD_DATE6,
--                       MAX(DECODE(SRNO,7, CLRD_DATE)) CLRD_DATE7,
--                       MAX(DECODE(SRNO,8, CLRD_DATE)) CLRD_DATE8,
--                       MAX(DECODE(SRNO,9, CLRD_DATE)) CLRD_DATE9,
--                       MAX(DECODE(SRNO,10, CLRD_DATE)) CLRD_DATE10,
--                       MAX(DECODE(SRNO,11, CLRD_DATE)) CLRD_DATE11,
--                       MAX(DECODE(SRNO,12, CLRD_DATE)) CLRD_DATE12                       
                FROM(
                SELECT INDV_ISCD, SRNO, UNAS_SDRT1, AVRG_APLY_YN, CLRD_ERT, TRTH_CLRD_DTRM_DATE, CLRD_DTRM_DATE, CLRD_DATE FROM BSYS.TBSIMO203D00@GDW

                ) GROUP BY INDV_ISCD ) AC,
                (SELECT INDV_ISCD,
                       --MAX(RNUM) EE_NUM,
                       MAX(DECODE(RNUM,1, CLRD_BARR_VAL)) CLRD_BARR_VAL1,
                       MAX(DECODE(RNUM,2, CLRD_BARR_VAL)) CLRD_BARR_VAL2,
                       MAX(DECODE(RNUM,3, CLRD_BARR_VAL)) CLRD_BARR_VAL3,
                       MAX(DECODE(RNUM,1, CLRD_BARR_HIT_YN)) CLRD_BARR_HIT_YN1,
                       MAX(DECODE(RNUM,2, CLRD_BARR_HIT_YN)) CLRD_BARR_HIT_YN2,
                       MAX(DECODE(RNUM,3, CLRD_BARR_HIT_YN)) CLRD_BARR_HIT_YN3,
                       MAX(DECODE(RNUM,1, CLRD_INRT)) CLRD_INRT1,
                       MAX(DECODE(RNUM,2, CLRD_INRT)) CLRD_INRT2,
                       MAX(DECODE(RNUM,3, CLRD_INRT)) CLRD_INRT3,
                       MAX(DECODE(RNUM,1, CALC_FIN_DATE)) CALC_FIN_DATE1,
                       MAX(DECODE(RNUM,2, CALC_FIN_DATE)) CALC_FIN_DATE2,
                       MAX(DECODE(RNUM,3, CALC_FIN_DATE)) CALC_FIN_DATE3
                FROM   (
                SELECT ROW_NUMBER() OVER(PARTITION BY INDV_ISCD ORDER BY CALC_FIN_DATE) AS RNUM,
                       INDV_ISCD,
                       CLRD_BARR_VAL,
                       CLRD_BARR_HIT_YN,
                       CLRD_INRT,
                       CALC_FIN_DATE
                FROM BSYS.TBSIMO227L00@GDW WHERE CALC_FIN_DATE is not null) GROUP BY INDV_ISCD) E,    --EARLY EXIT
               (SELECT INDV_ISCD, BONS_CUPN_STA_SDRT, BONS_CUPN_INRT FROM BSYS.TBSIMO210L00@GDW) MTH,
               RAS.RM_ELS_INFO HF,
    (SELECT TI.FUNDCODE,
           TI.CPARTYCODE,
           CI.TR_CNRP_NAME
    FROM   RMS.CREDIT_OTCTRADING_INFO TI,
           BSYS.TBFNIB001M00@GDW  CI
    WHERE CI.RISK_TR_CNRP_CODE=TI.CPARTYCODE
    AND TI.COMM2 <> '0899'
    ) CP_D,
    (SELECT TI.FUNDCODE,
           TI.CPARTYCODE,
           CI.TR_CNRP_NAME
    FROM   RMS.CREDIT_OTCTRADING_INFO TI,
           BSYS.TBFNIB001M00@GDW  CI
    WHERE CI.RISK_TR_CNRP_CODE=TI.CPARTYCODE
    AND TI.COMM2 <> '0899'
    ) CP_H
WHERE EE.ELS_FO_D = CP_D.FUNDCODE(+)
AND  EE.ELS_FO_H = CP_H.FUNDCODE(+)
AND  QT.STND_DATE(+)= :tdate
--AND  (NVL(QT.RMND_QTY,0) > 0 OR EE.FRST_STND_PRC_FIN_DATE = :tdate OR NVL(EE.CLRD_DATE,'0') >= :tdate)
AND  (NVL(QT.RMND_QTY,0) > 0 OR (EE.FRST_STND_PRC_FIN_DATE <= :tdate AND EE.PBLC_DATE >= :tdate) OR NVL(EE.CLRD_DATE,'0') >= :tdate) --������ ���� �� ���ذ� ���� ���� ���� �߰� 2024.03.29
AND  QT.ISCD(+)=EE.ELS_IO_D
--AND EE.DEAL_CLS_CODE=1
AND EE.PROD_FNCD IN ('SF002')
AND EE.ELS_IO_D IN (SELECT ISCD FROM BSYS.TBSIMO200D00@GDW WHERE FUND_BTWN_DEAL_YN ='Y')
--AND EE.ELS_IO_D IN (SELECT ISCD FROM BSYS.TBSIMO200D00@GDW WHERE FUND_BTWN_DEAL_YN ='N')
--AND EE.ELS_IO_D NOT IN (SELECT ISCD FROM BSYS.TBSIMO200D00@GDW WHERE FUND_BTWN_DEAL_YN ='Y')
AND EE.ELS_IO_D = UA.INDV_ISCD
AND EE.ELS_IO_D = AC.INDV_ISCD
AND    EE.CLRD_TYPE_CODE = M.CODE_VAL(+)
AND    EE.PROD_CLS_CODE = N.CODE_VAL(+)
AND    EE.PROD_CLS_CODE_H = N2.CODE_VAL(+)
AND    EE.HEDG_TR_TYPE_CODE = H.CODE_VAL(+)
AND    EE.UNAS_CHOC_MTHD_CODE = U.CODE_VAL(+)
AND    EE.ELS_IO_D= MTH.INDV_ISCD(+)
AND    EE.ELS_IO_D= E.INDV_ISCD(+)
--AND    EE.ESWAP_IO_T=HF.INDV_ISCD(+)
AND    nvl(EE.ESWAP_IO_T,EE.ESWAP_IO_H)=HF.INDV_ISCD(+)
--ORDER BY DECODE(HEDG_TR_TYPE_CODE, '1', 1,'2', 2), DECODE(N.CODE_VAL_NAME, 'ELS', 1,'ELN', 1,'ELB', 1, '����', 2), FRST_STND_PRC_FIN_DATE, STLM_CRCD, KOR_ISNM