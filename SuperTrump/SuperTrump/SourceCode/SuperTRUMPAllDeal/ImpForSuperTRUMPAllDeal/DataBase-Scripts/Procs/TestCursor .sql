CREATE OR REPLACE procedure TestCursor (Creation_date_Check IN DATE default null,AccountScheduleFeed OUT SYS_REFCURSOR,StreamFeed OUT SYS_REFCURSOR,AssetLevelFeed OUT SYS_REFCURSOR,ProductMapping OUT SYS_REFCURSOR,TemplateMapping OUT SYS_REFCURSOR,Depriciation OUT SYS_REFCURSOR)
IS
BEGIN

OPEN AccountScheduleFeed FOR select * from tbl_accountschedule_detail;

OPEN StreamFeed FOR
select * from TBL_STREAM_DETAIL where
     TBL_STREAM_DETAIL.ACCOUNT_SCHEDULE_NBR in ( select ACCOUNT_SCHEDULE_NBR  from tbl_accountschedule_detail );

OPEN AssetLevelFeed FOR
select * from TBL_ASSET_DETAIL where
     ACCOUNT_SCHEDULE_NBR in ( select ACCOUNT_SCHEDULE_NBR  from tbl_accountschedule_detail );

OPEN ProductMapping FOR
SELECT * FROM TBL_PMSDATA;

OPEN TemplateMapping FOR
select * from TBL_TEMPLATEMAPPING ;

OPEN Depriciation FOR
SELECT * FROM TBL_DEPRICIATION;

END;
/
