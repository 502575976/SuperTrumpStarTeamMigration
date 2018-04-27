
CREATE TABLE VIC_ROI.TBL_DEPRICIATION
(
  DEPRECIATION_TYPE  INTEGER,
  METHOD             VARCHAR2(500 BYTE)
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



Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (2, 'DB 150');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (5, '"Str line, to end"');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (8, 'None');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (10, 'Feed from Stream 17');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (11, 'ACRS 5 yr');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (12, 'Pickle');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (16, 'None');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (17, 'n/a');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (18, 'MACRS');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (33, '"Str line, to end"');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (91, '"MACRS (with ""Additional 1st year %"" of 30%)"');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (92, '"DB 150 (with ""Additional 1st year %"" of 30%)"');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (93, '"Str line, to end (with ""Additional 1st year %"" of 30%)"');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (94, '"MACRS (with ""Additional 1st year %"" of 50%)"');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (95, '"DB 150 (with ""Additional 1st year %"" of 50%)"');
Insert into VIC_ROI.TBL_DEPRICIATION
   (DEPRECIATION_TYPE, METHOD)
 Values
   (96, '"Str line, to end (with ""Additional 1st year %"" of 50%)"');
COMMIT;

