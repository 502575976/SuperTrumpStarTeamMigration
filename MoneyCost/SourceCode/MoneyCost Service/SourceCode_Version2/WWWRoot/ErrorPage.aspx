<%@ Page Language="VB" AutoEventWireup="false" Inherits="MoneyCostWeb.ErrorPage" Codebehind="ErrorPage.aspx.vb" %>
<%@ Register TagPrefix="top" TagName="Top" Src="~/Common/SSI/Top.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
		<title>ErrorPage</title>
		<link href='<%=ResolveUrl("~/css/MoneyCost.css")%>' type ="text/css" rel="stylesheet" />
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<script type="text/javascript">
	
	function retry11()
     {
	    window.history.back(-1);
        return false;
      }
</script>
    <form id="frmError" runat="server">
        <table  cellspacing="0" cellpadding="0" width="100%" align="center"  border="0">
				<tr>
					<td align="center"><top:top id="Top1" runat="server"></top:top></td>
				</tr>
		</table>
		<table height="50%" cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
				<tr>
					<td height="10"></td>
				</tr>
				<tr>
					<td width="5%">&nbsp;</td>
					<td width="91%"><b><font size="3" face="Arial, Helvetica, sans-serif">ERROR	DESCRIPTION&nbsp; :&nbsp;
								<br />
								<br />
							</font></b><font size="2" face="Arial, Helvetica, sans-serif"><b>
								<asp:Label ID="lblErrorDescription" Runat="Server"></asp:Label>
								<input type="hidden" id="txtErrorDesc" name="txtErrorDesc" runat="server" /> </b>
						</font>
						<br />
								<br />
					</td>
					<td width="4%">&nbsp;</td>
				</tr>
			</table>
			<table cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td valign="top">
					    <asp:Button ID="btnRetry" runat="server" CssClass="txtbuttons" Text="Retry" Width="65px" />                      
					</td>
				</tr>
			</table>
    </form>
</body>
</html>
