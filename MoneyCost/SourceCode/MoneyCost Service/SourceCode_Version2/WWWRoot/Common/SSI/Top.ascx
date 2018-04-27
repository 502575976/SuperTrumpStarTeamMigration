<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="Top.ascx.vb" Inherits="MoneyCostWeb.Top" %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	  <td width="24%" align="right" bgcolor="396797"><img  alt="" runat="server" src="~/images/top_lft.jpg"  height="98" /></td>  
      <td width="76%" bgcolor="396797"><img alt="" runat="server" src="~/images/top_bar.jpg" height="97" /></td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr bgcolor="04204E">
	<td colspan="5" height="5"></td>
</tr>
<tr bgcolor="04204E">
<td>&nbsp;</td>
<td>
<a href='../PricingAnalyst/UpdateDetails.aspx' style="color:White">Update Money Cost</a>
</td>
<td>
<%  If UCase(Request.ServerVariables("HTTP_CEFMONEYCOSTROLE")).Contains("ADMIN") Then%>
<a href='../PricingAnalyst/UpdateMCFile.aspx' style="color:White">Update MC File Information</a>
<%  End If%>
</td>
<td>
<%  If UCase(Request.ServerVariables("HTTP_CEFMONEYCOSTROLE")).Contains("ADMIN") Then%>
<a href='../PricingAnalyst/UpdateMCSecurity.aspx' style="color:White">Update MC Security</a>
<%end if %>
</td>
<td>
<%  If UCase(Request.ServerVariables("HTTP_CEFMONEYCOSTROLE")).Contains("ADMIN") Then%>
<a href='../PricingAnalyst/UpdateCsvFiles.aspx' style="color:White">Update CSV File</a>
<%end if %>
</td>
<td>
<a href='../logout.aspx' style="color:White">Logout</a>
</td>
</tr>
<tr>
<td>&nbsp;</td></tr>
<tr>
	<td colspan="4"><img id="Img1" src="~/images/moneyupdate_title.gif"  runat="server" alt="Money Cost Update" width="158" height="18" /></td>
</tr>
</table>