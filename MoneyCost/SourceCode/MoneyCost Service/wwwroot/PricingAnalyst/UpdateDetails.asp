<!--#include virtual="/common/ssi/CefCommonFunc.asp"-->
<%
	On Error Resume Next

	'declare local variables
	Dim lstrUSER_SSOID				'to store logged-in user's SSO ID
	Dim llSQ_MC_ID					'to store selected MC File ID
	Dim lstrPageInfoRequestXML		'to store request XML for Page Info i.e. logged-in user's details and all MC File list
	Dim lstrPageInfoResponseXML		'to store response XML of Page Info i.e. logged-in user's details and all MC File list
	Dim lstrMCFileDetailRequestXML	'to store request XML for getting MC File details
	Dim lstrMCFileDetailResponseXML	'to store response XML of MC File details
	Dim lobjMoneyCostService		'to get handle of BSMoneyCost.IMoneyCostService class object
	Dim lobjMoneyCostMgr			'to get handle of BSMoneyCost.IMoneyCostMgr class object
	Dim lobjPageInfoDOM				'DOM object to load Page Info XML
	Dim lobjMCFileDetailDOM			'DOM object to load MC File details XML
	Dim lobjMCFileNodeList			'Node List object for MC_FILESet
	Dim lobjMCFileDetailNodeList	'Node List object for MC_FILE_DETAIL node set
	Dim lbShowMCFileDetailFlag		'boolean variable to check whether MC File Details XML is loaded properly or not
	Dim liCount						'for loop counter
	Dim lobjMCFileDetailElement		'single Node element object for MC_FILE_DETAIL node set
	Dim lstrSQ_INDEX_IDList			'comma separated list of SQ_INDEX_IDs of all index rate data shown, for selected MC File
	Dim larrSQ_INDEX_ID				'array for all SQ_INDEX_IDs of all index rate data shown, for selected MC File
	Dim lstrAdderRateCommon			'to store common Adder Rate, updateable to all index data
	Dim lstrEffectiveDateCommon		'to store common Effective Date, updateable to all index data
	Dim larrAdderRate				'array of all Adder Rates for all index data
	Dim larrEffectiveDate			'array of all Effective Dates for all index data
	Dim lstrUpdateDetailRequestXML	'dynamically build request XML for update index details
	Dim lstrOrigMCFileDetailXML		'to store original (as per existing database) of MC File details
	Dim lstrErrorMsg				'to store error message
	Dim lstrOldAdderRate			'to store existing Adder Rate for any particular index data row
	Dim lstrOldEffectiveDate		'to store existing Effective Date for any particular index data row
	Dim lstrUpdateDetailResponseXML	'to store response XML from update MC File details
	Dim lstrUpdateConfMsg			'to store confirmation message, after updation.
	Dim lstrUSER_GESSOUID			'to store logged-in user's GESSOUID, from server variable

	'Error constaints
	Const cERR_ROOT_TAG	= "ERROR_DETAILS"
	Const cERR_SRC		= "ERROR_SOURCE"
	Const cERR_NBR		= "ERROR_NUMBER"
	Const cERR_DESC		= "ERROR_DESCRIPTION"
	Const cERR_ShowUser	= "ERROR_SHOW_USER"

	lstrUSER_GESSOUID = Trim(Request.ServerVariables("HTTP_GESSOUID"))

	lbShowMCFileDetailFlag = False

	llSQ_MC_ID = Trim(Request.Form.Item("cboMoneyCostFile"))
	lstrSQ_INDEX_IDList = Trim(Request.Form.Item("hdnSQ_INDEX_IDList"))

	'create instances for object variables
	Set lobjMoneyCostService = Server.CreateObject("BSMoneyCost.IMoneyCostService")
	Set lobjMoneyCostMgr = Server.CreateObject("BSMoneyCost.IMoneyCostMgr")
	Set lobjPageInfoDOM = Server.CreateObject("MSXML2.DOMDocument.4.0")
	Set lobjMCFileDetailDOM = Server.CreateObject("MSXML2.DOMDocument.4.0")

	'build request XML for fetching all MC File list
	lstrPageInfoRequestXML = "<MC_FILES_REQUEST>" & _
								 "<USER_GESSOUID>" & lstrUSER_GESSOUID & "</USER_GESSOUID>" & _
							 "</MC_FILES_REQUEST>"

	'fetch MC File list for logged-in user
	lstrPageInfoResponseXML = lobjMoneyCostService.GetMCFiles(lstrPageInfoRequestXML)

	RedirectOnErrorMoneyCost(lstrPageInfoResponseXML)

	'load Page Info response XML into DOM object
	If Not lobjPageInfoDOM.loadXML(lstrPageInfoResponseXML) Then Response.End

	lstrUSER_SSOID = lobjPageInfoDOM.getElementsByTagName("uid").item(0).text

	Set lobjMCFileNodeList = lobjPageInfoDOM.selectNodes("/USER_MC_FILE_RESPONSE/MC_FILE_RESPONSE/MC_FILESet")

	'if user clicks on "Save" button
	If UCase(Trim(Request.Form.Item("btnUpdate"))) = "SAVE" Then
		'get data into local variables
		lstrOrigMCFileDetailXML = Trim(Request.Form.Item("hdnMCFileDetailXML"))
		lstrAdderRateCommon = Trim(Request.Form.Item("txtAdderRateCommon"))
		lstrEffectiveDateCommon = Trim(Request.Form.Item("txtEffectiveDateCommon"))

		'check, whether authorised user has logged-in
		If lstrUSER_GESSOUID = "" Then
			lstrErrorMsg = "lstrErrorMsg & Unauthorised logged-in user.<BR>"
		End If

		'check, whether user has selected any Money Cost File
		If llSQ_MC_ID = "" Then
			lstrErrorMsg = lstrErrorMsg & "Please select a Money Cost file.<BR>"
		End If

		'check, whether valid XML found for existing MC File details
		If Not lobjMCFileDetailDOM.loadXML(lstrOrigMCFileDetailXML) Then
			lstrErrorMsg = lstrErrorMsg & "Invalid MC File detail XML.<BR>"
		End If

		'if any data row found in MC file detail, store into array
		If lstrSQ_INDEX_IDList <> "" Then
			larrSQ_INDEX_ID = Split(lstrSQ_INDEX_IDList, ",")
		'else get error message
		Else
			lstrErrorMsg = lstrErrorMsg & "No data found for updation.<BR>"
		End If

		If Trim(Request.Form("txtAdderRate")) <> "" Then
			larrAdderRate = Split(Trim(Request.Form("txtAdderRate")), ",")
		End If

		If Trim(Request.Form("txtEffectiveDate")) <> "" Then
			larrEffectiveDate = Split(Trim(Request.Form("txtEffectiveDate")), ",")
		End If

		'build request XML for updation, dynamically
		lstrUpdateDetailRequestXML = "<UPDATE_MC_FILE_DETAIL_REQUEST>" & _
										"<USER_SSOID>" & lstrUSER_SSOID & "</USER_SSOID>" & _
										"<MC_FILE_DETAILSet>"

		If lstrErrorMsg = "" Then
			'loop for each index row
			For liCount = 0 To UBound(larrSQ_INDEX_ID)
				lstrOldAdderRate = ""
				lstrOldEffectiveDate = ""

				'search particular MC_FILE_DETAIL node set in existing MC File detail XML
				Set lobjMCFileDetailElement = lobjMCFileDetailDOM.selectSingleNode("/MC_FILE_DETAIL_RESPONSE/MC_FILE_DETAILSet/MC_FILE_DETAIL[SQ_INDEX_ID=" & larrSQ_INDEX_ID(liCount) & "]")

'				'if existing node set found
'				If Not lobjMCFileDetailElement Is Nothing Then
'					'get existing Adder Rate for the row
'					If lobjMCFileDetailElement.getElementsByTagName("AMT_ADDER").length > 0 Then
'						lstrOldAdderRate = lobjMCFileDetailElement.getElementsByTagName("AMT_ADDER").Item(0).Text
'					End If
'
'					'get existing Effective Date for the row
'					If lobjMCFileDetailElement.getElementsByTagName("DATE_EFFECTIVE").length > 0 Then
'						lstrOldEffectiveDate = lobjMCFileDetailElement.getElementsByTagName("DATE_EFFECTIVE").Item(0).Text
'					End If
'				End If

				'check validation for Adder Rate
				If Trim(larrAdderRate(liCount)) = "" Then
					If InStr(lstrErrorMsg, "Please update Adder Rate.") <= 0 Then
						lstrErrorMsg = lstrErrorMsg & "Please update Adder Rate.<BR>"
					End If
				ElseIf IsNumeric(Trim(larrAdderRate(liCount))) = False Then
					If InStr(lstrErrorMsg, "Adder Rate entered is not valid numeric") <= 0 Then
						lstrErrorMsg = lstrErrorMsg & "Adder Rate entered is not valid numeric.<BR>"
					End If
				ElseIf Abs(CDbl(Trim(larrAdderRate(liCount)))) > 99.899999 Then
					If InStr(lstrErrorMsg, "Adder Rate entered should be between -99.899999 and +99.899999") <= 0 Then
						lstrErrorMsg = lstrErrorMsg & "Adder Rate entered should be between -99.899999 and +99.899999<BR>"
					End If
				End If

				'check validation for Effective Date
				If Trim(larrEffectiveDate(liCount)) = "" Then
					If InStr(lstrErrorMsg, "Please update Effective Date") <= 0 Then
						lstrErrorMsg = lstrErrorMsg & "Please update Effective Date.<BR>"
					End If
				ElseIf CheckDate(Trim(larrEffectiveDate(liCount))) = False Then
					If InStr(lstrErrorMsg, "Effective Date entered is not valid.") <= 0 Then
						lstrErrorMsg = lstrErrorMsg & "Effective Date entered is not valid.<BR>"
					End If
				End If

				'if no error found
				If lstrErrorMsg = "" Then
					'dynamically build request XMl for updation
					lstrUpdateDetailRequestXML = lstrUpdateDetailRequestXML & _
											"<MC_FILE_DETAIL>" & _
												"<SQ_INDEX_ID>" & Trim(larrSQ_INDEX_ID(liCount)) & "</SQ_INDEX_ID>" & _
												"<AMT_ADDER>" & Trim(larrAdderRate(liCount)) & "</AMT_ADDER>" & _
												"<DATE_EFFECTIVE>" & Trim(larrEffectiveDate(liCount)) & "</DATE_EFFECTIVE>" & _
											"</MC_FILE_DETAIL>"

'												"<AMT_ADDER_OLD>" & Trim(lstrOldAdderRate) & "</AMT_ADDER_OLD>" & _
'												"<DATE_EFFECTIVE_OLD>" & Trim(lstrOldEffectiveDate) & "</DATE_EFFECTIVE_OLD>" & _
				End If
			Next
		End If

		'if no error found
		If lstrErrorMsg = "" Then
			lstrUpdateDetailRequestXML = lstrUpdateDetailRequestXML & _
										"</MC_FILE_DETAILSet>" & _
									"</UPDATE_MC_FILE_DETAIL_REQUEST>"

			lstrUpdateDetailResponseXML = lobjMoneyCostMgr.UpdateMCDetails(lstrUpdateDetailRequestXML)

			RedirectOnErrorMoneyCost(lstrUpdateDetailResponseXML)

			If InStr(lstrUpdateDetailResponseXML, "<STATUS>SUCCESS</STATUS>") > 0 Then
				lstrUpdateConfMsg = "Data has been saved successfully."
			End If

			lstrAdderRateCommon = ""
			lstrEffectiveDateCommon = ""
		End If
	End If

	'if user has selected any MC File, get details of selected MC File
	If llSQ_MC_ID <> "" Then
		'build request XML for fetching MC File details
		lstrMCFileDetailRequestXML = "<MC_FILE_DETAIL_REQUEST>" & _
										"<SQ_MC_ID>" & llSQ_MC_ID & "</SQ_MC_ID>" & _
									"</MC_FILE_DETAIL_REQUEST>"

		'call GetMCFileDetails() method to fetch details of selected MC File
		lstrMCFileDetailResponseXML = lobjMoneyCostService.GetMCFileDetails(lstrMCFileDetailRequestXML)

		RedirectOnErrorMoneyCost(lstrMCFileDetailResponseXML)

		'load response XML into DOM object
		If lobjMCFileDetailDOM.loadXML(lstrMCFileDetailResponseXML) Then lbShowMCFileDetailFlag = True
	End If
%>
<html><head><title>Update: Money Cost Details</title>
<style type=text/css>
.main
{
    FONT-SIZE: 11px;
    COLOR: #000000;
    FONT-FAMILY: Arial, Helvetica, sans-serif
}
.buttons
{
    BORDER-RIGHT: #000000 1px solid;
    BORDER-TOP: #000000 1px solid;
    FONT-WEIGHT: normal;
    FONT-SIZE: 11px;
    MARGIN: 1px 2px;
    BORDER-LEFT: #000000 1px solid;
    COLOR: #000000;
    BORDER-BOTTOM: #000000 1px solid;
    FONT-FAMILY: Arial, Helvetica, sans-serif;
    BACKGROUND-COLOR: #e2e2e2;
    font-color: #000000
}
.error
{
    FONT-SIZE: 11px;
    COLOR: #red;
    FONT-FAMILY: Arial, Helvetica, sans-serif
}
</style>

<script language="JavaScript">
function fncUpdateAll()
{
	/*On click of "Apply All" button, call this method*/
	var lstrCommonAdderRate;
	var lstrCommonEffeciveDate;
	var liCounter;

	//get common Adder Rate & Effective Date into local variables
	lstrCommonAdderRate = document.frmMoneyCost.txtAdderRateCommon.value;
	lstrCommonEffeciveDate = document.frmMoneyCost.txtEffectiveDateCommon.value;

	if(document.frmMoneyCost.txtAdderRate != null){
		//loop for all MC File detail data row set and update common value for Adder Rate
		//and Effective Date with common values
		for(liCounter = 0; liCounter < document.frmMoneyCost.txtAdderRate.length; liCounter++){
			document.frmMoneyCost.txtAdderRate[liCounter].value = lstrCommonAdderRate;
			document.frmMoneyCost.txtEffectiveDate[liCounter].value = lstrCommonEffeciveDate;
		}
	}
}

function fncCancel()
{
	/* On click of "Cancel" button, ask for confirmation. If confirmed, load page as default and unsave all changed data*/
	if(confirm("Any changes made will not be saved. Are you sure to continue?")){
		document.frmMoneyCost.hdnSQ_INDEX_IDList.value = "";
		document.frmMoneyCost.cboMoneyCostFile.value = "";
		document.frmMoneyCost.submit();
	}
}

function fncSetDefaultDateCommon(astrCurrentValue, astrControlName, astrCurrentDate)
{
	/* function to set current date as default date as Effective Date, in "Apply All" section*/
	if(fncAllTrim(astrCurrentValue) != '')
		astrControlName.value = astrCurrentDate;
}

function fncSetDefaultDate(astrCurrentValue, astrCurrentDate)
{
	/* function to set current date as default date as Effective Date, in detailed section*/
	var liCounter;

	for(liCounter = 0; liCounter < document.frmMoneyCost.txtEffectiveDate.length; liCounter++)
		if(fncAllTrim(astrCurrentValue) != '')
			document.frmMoneyCost.txtEffectiveDate[liCounter].value = astrCurrentDate;
}

function fncAllTrim(astrRequestString)
{
	/* function to trim the string from left side as well from right side */
	while(astrRequestString.substring(0, 1) == ' '){
		astrRequestString = astrRequestString.substring(1, astrRequestString.length);
	}

	while(astrRequestString.substring(astrRequestString.length-1, astrRequestString.length) == ' '){
		astrRequestString = astrRequestString.substring(0, astrRequestString.length-1);
	}

	return astrRequestString;
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frmMoneyCost" action="UpdateDetails.asp" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	  <td width="24%" align="right" bgcolor="396797"><img src="../images/top_lft.jpg" width="279" height="98"></td>  
      <td width="76%" bgcolor="396797"><img src="../images/top_bar.jpg" height="97"></td>
</tr>
<tr bgcolor="04204E">
	<td colspan="2" height=5></td>
</tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
	<td><img src="../images/moneyupdate_title.gif" alt="Money Cost Update" width="158" height="18"></td>
</tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
	<td bgcolor="cccccc">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="main">
		<tr bgcolor=ffffff>
			<td height="30" colspan="2"><strong>&nbsp;&nbsp;Welcome <%=lobjPageInfoDOM.getElementsByTagName("givenname").item(0).text%>&nbsp;<%=lobjPageInfoDOM.getElementsByTagName("sn").item(0).text%></strong></td>
		</tr>
<%	If lstrErrorMsg <> "" Then%>
		<tr bgcolor=ffffff>
			<td height="30" colspan="2" class="error"><%=lstrErrorMsg%></td>
		</tr>
<%	ElseIf lstrUpdateConfMsg <> "" Then%>
		<tr bgcolor=ffffff>
			<td height="30" colspan="2"><%=lstrUpdateConfMsg%></td>
		</tr>
<%	End If%>
		<tr bgcolor=ffffff>
			<td width="14%"><strong>&nbsp;&nbsp;Money Cost File</strong></td>
            <td width="86%">
				<select name="cboMoneyCostFile" onChange="document.frmMoneyCost.submit();" class="main">
					<option value="">Select a Money Cost File</option>
<%					'use function from CefCommonFunc.asp file to build combo
					Response.Write GetSelectedOptionTag(lobjMCFileNodeList.item(0).xml, llSQ_MC_ID)%>
				</select>
			</td>
		</tr>
		<tr bgcolor=ffffff>
			<td colspan="2"><br>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td bgcolor="eeeeee">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" class="main">
						<tr bgcolor=eeeeee>
							<td colspan="5"><strong>&nbsp;&nbsp;Apply All</strong></td>
						</tr>
						<tr bgcolor=ffffff>
							<td width="12%"><strong>&nbsp;&nbsp;Adder Rate</strong></td>
							<td width="21%" bgcolor="ffffff"><input type="text" name="txtAdderRateCommon" class="main" VALUE="<%=lstrAdderRateCommon%>" MAXLENGTH=9 OnBlur="JavaScript:fncSetDefaultDateCommon(this.value, document.frmMoneyCost.txtEffectiveDateCommon, '<%=FormatDateSpecial(Date())%>');"></td>
							<td width="15%" bgcolor="ffffff"><strong>&nbsp;&nbsp;Effective Date</strong></td>
							<td width="21%" bgcolor="ffffff"><input type="text" name="txtEffectiveDateCommon" class="main" VALUE="<%=lstrEffectiveDateCommon%>" MAXLENGTH=10></td>
							<td width="31%" bgcolor="ffffff"><input type="button" name="btnUpdateAll" value="Apply All" class="buttons" onClick="JavaScript:fncUpdateAll();" title="Apply All"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
				<br>
			</td>
		</tr>
		<tr bgcolor=ffffff>
			<td valign="top"><strong>&nbsp;&nbsp;Index Rates</strong></td>
            <td bgcolor="eeeeee">
				<table width="100%" border="0" cellspacing="0" cellpadding="0" class="main">
                <tr>
					<td width="10%"><strong>Col Position</strong></td>
					<td width="20%"><strong>Code</strong></td>
					<td width="19%"><strong>Description</strong></td>
					<td width="15%"><strong>Adder</strong></td>
					<td width="23%"><strong>Effective Date</strong></td>
				</tr>
<%	'if MC File detail XML is valid
	If lbShowMCFileDetailFlag = True Then
		'get handle of all MC_FILE_DETAIL node set into Node List object
		Set lobjMCFileDetailNodeList = lobjMCFileDetailDOM.selectNodes("/MC_FILE_DETAIL_RESPONSE/MC_FILE_DETAILSet/MC_FILE_DETAIL")

		'initialize comma separated SQ_INDEX_ID list as blank
		lstrSQ_INDEX_IDList = ""

		'loop for each MC File detail node set
		For liCount = 0 To lobjMCFileDetailNodeList.length - 1
			'get handle for single node of MC_FILE_DETAIL node set
			Set lobjMCFileDetailElement = lobjMCFileDetailNodeList.item(liCount)

			'get SQ_INDEX_ID as comma separated for all index row set
			lstrSQ_INDEX_IDList = lstrSQ_INDEX_IDList & lobjMCFileDetailElement.getElementsByTagName("SQ_INDEX_ID").item(0).text & ","
%>
				<tr bgcolor=ffffff> 
					<td height="15"><%=lobjMCFileDetailElement.getElementsByTagName("MC_FILE_COL_POSITION").item(0).text%></td>
					<td><%=lobjMCFileDetailElement.getElementsByTagName("INDEX_CODE").item(0).text%></td>
					<td><%=lobjMCFileDetailElement.getElementsByTagName("DESCRIPTION").item(0).text%></td>
					<td width="15%"><input type="text" name="txtAdderRate" class="main" VALUE="<%=lobjMCFileDetailElement.getElementsByTagName("AMT_ADDER").item(0).text%>" MAXLENGTH=9 OnBlur="JavaScript:fncSetDefaultDate(this.value, '<%=FormatDateSpecial(Date())%>');"></td>
					<td><input type="text" name="txtEffectiveDate" class="main" VALUE="<%=lobjMCFileDetailElement.getElementsByTagName("DATE_EFFECTIVE").item(0).text%>" MAXLENGTH=10></td>
                </tr>
<%		Next

		'remove last comma, if present
		If InStr(lstrSQ_INDEX_IDList, ",") > 0 Then
			lstrSQ_INDEX_IDList = Mid(lstrSQ_INDEX_IDList, 1, Len(lstrSQ_INDEX_IDList) - 1)
		End If
	End If
%>
				</table>
			</td>
		</tr>
        </table>
	</td>
</tr>
<tr align="center" bgcolor=ffffff>
	<td height="40" colspan="2"><input type="submit" name="btnUpdate" value="Save" class="buttons" title="Save">
        &nbsp;&nbsp;&nbsp;<input type="button" name="Update All3" value="Cancel" class="buttons" onClick="JavaScript:fncCancel();" title="Cancel"> 
        <INPUT TYPE="hidden" NAME="hdnSQ_INDEX_IDList" SIZE=1 VALUE="<%=lstrSQ_INDEX_IDList%>"> 
		<INPUT TYPE="hidden" NAME="hdnMCFileDetailXML" VALUE="<%=lstrMCFileDetailResponseXML%>"> 
	</td>
</tr>
</table>
</form>
</body>
</html>
<%
	'clear all object and array variables from memory
	Set larrAdderRate = Nothing
	Set larrEffectiveDate = Nothing
	Set larrSQ_INDEX_ID = Nothing
	Set lobjMCFileDetailDOM = Nothing
	Set lobjMCFileDetailElement = Nothing
	Set lobjMCFileDetailNodeList = Nothing
	Set lobjMCFileNodeList = Nothing
	Set lobjMoneyCostMgr = Nothing
	Set lobjMoneyCostService = Nothing
	Set lobjPageInfoDOM = Nothing
	
'================================================================
'METHOD  : RedirectOnErrorMoneyCost
'PURPOSE : To redict the user to error page if any error is found.
'PARMS   :
'          astrReturnString [String] = Return value from Business component
'RETURN  : none
'================================================================
Function RedirectOnErrorMoneyCost(byval astrReturnString)
dim lobjErrorDOM
dim lstrErrorSource
dim lstrErrorNumber
dim lbErrorOccured
dim lstrErrorDescription
dim lstrUserErrorDescription

lbErrorOccured = false
If astrReturnString <> "" then
	If InStr(1, astrReturnString, "</" & cERR_ROOT_TAG &">") > 0 then
		lbErrorOccured = true
		set lobjErrorDOM = Server.CreateObject("MSXML2.DOMDocument.4.0")

		if lobjErrorDOM.loadXML(astrReturnString) then
			if lobjErrorDOM.getElementsByTagName(cERR_ROOT_TAG).length > 0 then
				lstrErrorSource = lobjErrorDOM.getElementsByTagName(cERR_SRC).item(0).text
				lstrErrorNumber = lobjErrorDOM.getElementsByTagName(cERR_NBR).item(0).text
				lstrErrorDescription = lobjErrorDOM.getElementsByTagName(cERR_DESC).item(0).text
				if lobjErrorDOM.getElementsByTagName(cERR_ShowUser).length > 0 then
					lstrUserErrorDescription = lobjErrorDOM.getElementsByTagName(cERR_ShowUser).item(0).text
				end if
			end if
		end if
		set lobjErrorDOM = nothing
	end if
end if

if astrReturnString = "" and err.number <> 0 then
	lbErrorOccured = true
	lstrErrorSource = err.Source
	lstrErrorNumber = err.number
	lstrErrorDescription = err.Description
	lstrUserErrorDescription = "An internal error occured while performing the requested operaton"
end if

if lbErrorOccured then
	Session("number") = lstrErrorNumber
	Session("source") = Request.ServerVariables("SERVER_NAME") &  Request.ServerVariables("URL") & " - " &  lstrErrorSource
	Session("description") = lstrErrorDescription
	Session("UserFriendlyErrorDescription") = lstrUserErrorDescription
	Session("Email") = Request.ServerVariables("HTTP_EMAIL")

	Response.Redirect "../error.asp"
	Response.End
end if
End Function
%>