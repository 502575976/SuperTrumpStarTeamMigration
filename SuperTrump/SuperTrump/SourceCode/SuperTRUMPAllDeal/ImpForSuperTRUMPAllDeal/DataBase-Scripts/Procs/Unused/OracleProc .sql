CREATE OR REPLACE PROCEDURE OracleProc (curResults    OUT SYS_REFCURSOR)
IS
-- etc ---
BEGIN
OPEN curResults FOR Select * From tbl_asset_detail;
END;
/
