CREATE OR REPLACE procedure GetStreamDetail (Creation_date_Check IN DATE default null,curAccountDetail OUT SYS_REFCURSOR)
IS
BEGIN
OPEN curAccountDetail FOR
select * from TBL_STREAM_DETAIL where
     TBL_STREAM_DETAIL.ACCOUNT_SCHEDULE_NBR in ( select ACCOUNT_SCHEDULE_NBR  from tbl_accountschedule_detail where Process_flag=0);
END;
/
