
CREATE TABLE VIC_ROI.TBL_TEMPLATEMAPPING
(
  TEMPLATENAME  VARCHAR2(1000 BYTE),
  PRODUCT       VARCHAR2(50 BYTE),
  TERM_MIN      INTEGER,
  TERM_MAX      INTEGER
)
TABLESPACE VIC_DATA
PCTUSED    40
PCTFREE    10
INITRANS   1
MAXTRANS   255
STORAGE    (
            INITIAL          64K
            MINEXTENTS       1
            MAXEXTENTS       2147483645
            PCTINCREASE      0
            FREELISTS        1
            FREELIST GROUPS  1
            BUFFER_POOL      DEFAULT
           )
LOGGING 
NOCOMPRESS 
NOCACHE
NOPARALLEL
MONITORING;



Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease term 48 and less.tem', 'ELTOOL', 0, 48);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease term 49 to 60.tem', 'ELTOOL', 49, 60);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease term 61 to 72.tem', 'ELTOOL', 61, 72);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease greater than 72.tem', 'ELTOOL', 73, 999);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease term 48 and less.tem', 'MENQSI', 0, 48);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease term 49 to 60.tem', 'MENQSI', 49, 60);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease term 61 to 72.tem', 'MENQSI', 61, 72);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease greater than 72.tem', 'MENQSI', 73, 999);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease term 48 and less.tem', 'MEOQSI', 0, 48);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease term 49 to 60.tem', 'MEOQSI', 49, 60);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease term 61 to 72.tem', 'MEOQSI', 61, 72);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease greater than 72.tem', 'MEOQSI', 73, 999);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Muni term 48 and less.tem', 'MEQMUN', 0, 48);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Muni term 49 to 60.tem', 'MEQMUN', 49, 60);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Muni term 61 to 72.tem', 'MEQMUN', 61, 72);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Muni  greater than 72.tem', 'MEQMUN', 73, 999);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Loan term 48 and less.tem', 'MEREG', 0, 48);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Loan term 49 to 60.tem', 'MEREG', 49, 60);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Loan term 61 to 72.tem', 'MEREG', 61, 72);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Loan greater than 72.tem', 'MEREG', 73, 999);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease term 48 and less.tem', 'OPERLS', 0, 48);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease term 49 to 60.tem', 'OPERLS', 49, 60);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease term 61 to 72.tem', 'OPERLS', 61, 72);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease greater than 72.tem', 'OPERLS', 73, 999);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease term 48 and less.tem', 'SGLINV', 0, 48);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease term 49 to 60.tem', 'SGLINV', 49, 60);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease term 61 to 72.tem', 'SGLINV', 61, 72);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Lease greater than 72.tem', 'SGLINV', 73, 999);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease term 48 and less.tem', 'QTOOL', 0, 48);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease term 49 to 60.tem', 'QTOOL', 49, 60);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease term 61 to 72.tem', 'QTOOL', 61, 72);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF quasi (non-tax) lease greater than 72.tem', 'QTOOL', 73, 999);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Muni term 48 and less.tem', 'METEXM', 0, 48);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Muni term 49 to 60.tem', 'METEXM', 49, 60);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Muni term 61 to 72.tem', 'METEXM', 61, 72);
Insert into VIC_ROI.TBL_TEMPLATEMAPPING
   (TEMPLATENAME, PRODUCT, TERM_MIN, TERM_MAX)
 Values
   ('E:\internalsites\SuperTRUMP\PRMTemplates\VF\VF Muni  greater than 72.tem', 'METEXM', 73, 999);
COMMIT;

