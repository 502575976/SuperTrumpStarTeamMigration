CREATE OR REPLACE procedure GetAccountDetail (Creation_date_Check IN DATE default null,curAccountDetail OUT SYS_REFCURSOR)
IS
BEGIN
OPEN curAccountDetail FOR select * from tbl_accountschedule_detail,TBL_PMSDATA,TBL_TEMPLATEMAPPING where
     tbl_accountschedule_detail.Location=TBL_PMSDATA.PMS_Location
  AND
  (TBL_TEMPLATEMAPPING.Product=tbl_accountschedule_detail.Product and TBL_TEMPLATEMAPPING.TERM_MIN<=tbl_accountschedule_detail.TERM and TBL_TEMPLATEMAPPING.TERM_MAX>=tbl_accountschedule_detail.TERM)
  AND tbl_accountschedule_detail.Process_flag=0;
END;
/