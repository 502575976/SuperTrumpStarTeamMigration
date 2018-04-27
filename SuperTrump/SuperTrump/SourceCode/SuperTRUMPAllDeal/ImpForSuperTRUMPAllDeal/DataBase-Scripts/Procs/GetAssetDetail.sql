CREATE OR REPLACE procedure GetAssetDetail (Creation_date_Check IN DATE default null,curAccountDetail OUT SYS_REFCURSOR)
IS
BEGIN
OPEN curAccountDetail FOR
select * from TBL_ASSET_DETAIL,TBL_DEPRICIATION where
     TBL_ASSET_DETAIL.ACCOUNT_SCHEDULE_NBR in ( select ACCOUNT_SCHEDULE_NBR  from tbl_accountschedule_detail where Process_flag=0)
  AND TBL_DEPRICIATION.Depreciation_Type=TBL_ASSET_DETAIL.Depreciation_Type;
END;
/