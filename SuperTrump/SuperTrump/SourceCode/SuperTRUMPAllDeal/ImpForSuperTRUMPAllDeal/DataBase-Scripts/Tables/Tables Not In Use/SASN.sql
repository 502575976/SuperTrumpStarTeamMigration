
CREATE TABLE VIC_ROI.SASN
(
  CDR_TYPE  VARCHAR2(3 BYTE),
  REC_TYPE  VARCHAR2(3 BYTE)
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



Insert into VIC_ROI.SASN
   (CDR_TYPE, REC_TYPE)
 Values
   ('1', '2  ');
Insert into VIC_ROI.SASN
   (CDR_TYPE, REC_TYPE)
 Values
   ('3', '4');
Insert into VIC_ROI.SASN
   (CDR_TYPE, REC_TYPE)
 Values
   ('5', '6');
Insert into VIC_ROI.SASN
   (CDR_TYPE, REC_TYPE)
 Values
   ('7', '8');
Insert into VIC_ROI.SASN
   (CDR_TYPE, REC_TYPE)
 Values
   ('1', '2  ');
Insert into VIC_ROI.SASN
   (CDR_TYPE, REC_TYPE)
 Values
   ('3', '4');
Insert into VIC_ROI.SASN
   (CDR_TYPE, REC_TYPE)
 Values
   ('7', '8');
COMMIT;

