<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="UpdateMCSecurity.aspx.vb" Inherits="MoneyCostWeb.UpdateMCSecurity" %>

<%@ Register Src="~/Common/SSI/Top.ascx" TagName="Top"  TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<link id="Link1" href= "~/CSS/MoneyCost.css" runat="server" type ="text/css" rel="stylesheet" />
    <title>Update: Money Cost MC File</title>
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
</head>
<body topmargin=0 leftmargin=0>
    <form id="frmMoneyCost" runat="server">
   <table class="tdbg1" cellSpacing="0" cellPadding="0" width=100%  align="center" border="0">
       <tr>
          <td>
                <uc1:Top id="Top1" runat="server">
                </uc1:Top>
          </td>
        </tr>
        <tr>
          <td align=center>
               <asp:Label ID="lblErrorMessage" Font-Bold="true" ForeColor="red" runat="server" Text=""  Font-Names="Geneva, Arial, Helvetica, sans-serif"></asp:Label>
               <asp:HiddenField ID="hSSOId" Value="" runat="server" />
          </td>
           </tr>   
   </table>
      <br /> 
          
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="main">
		   <tr bgcolor=ffffff>
			    <td height="30" colspan="2"><strong>&nbsp;&nbsp;Welcome </strong>
			       <asp:Label id="lblUserName" runat="server"></asp:Label>
			    </td>
		    </tr>

		    <tr bgcolor=ffffff>
			    <td colspan="2">
			    </td>
		    </tr>
		    <tr>
	             <td class="text11normal"><asp:label id="lblNoMCDetail" runat="server"></asp:label></td>
             </tr>
		</table>
		<table id="tblIndexRate" runat="server"   border="0"  style="width:70%; height: 16px;" >
             <tr><td><hr /></td></tr>
             
             <tr>
                <td class="text11normal">
                    <asp:label id="lblTableHeder" Font-Bold="true" Font-Underline="true" runat="server" Text="MC Security"></asp:label>
                </td>
             </tr>
             <tr>
                <td class="text11normal">
                    <asp:TextBox ID="txtSSOID" name="txtSSOID" runat="server" CssClass=""></asp:TextBox> &nbsp; <asp:Button  ID="btnSSOID" name="btnSSOID" runat="server" Text="Check Security" />
                </td>
                <td>
                <div id=chkbox runat=server>
                <b>All</b> <asp:CheckBox runat=server ID=chkAllFile name=chkAllFile AutoPostBack=true />
                </div>
                </td>
             </tr>
             <tr>
                <td colspan=2 class="txt11normal" >
                       <asp:datagrid id="grdMCDetail" runat="server" BorderWidth="1px"  DataKeyField="SQ_MC_ID"
	                        AutoGenerateColumns="False"  bodyheight="100" CellSpacing="2"  AlternatingItemStyle-CssClass="textnormalAlt"
	                        CellPadding="2"  GridLines="Both" Width="100%">
	                        <FooterStyle CssClass="lightbg"></FooterStyle>
	                        <ItemStyle  CssClass="text11normal"  VerticalAlign="Top" HorizontalAlign="Left" ></ItemStyle>
	                        <HeaderStyle cssClass="TextBigHD"  VerticalAlign="Top" HorizontalAlign="Left"></HeaderStyle>
	                        <Columns>	            
	                            <asp:BoundColumn  DataField="SQ_MC_ID" Visible="False"></asp:BoundColumn>	        
	                            <asp:BoundColumn  DataField="MC_Security" Visible="False"></asp:BoundColumn>	        
	                            <asp:BoundColumn  DataField="Description" HeaderText="Description"></asp:BoundColumn>	
	                            <asp:TemplateColumn HeaderText="MC File" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Left">
									<ItemTemplate>
									<asp:CheckBox ID="chkMCFile"  name="chkMCFile" runat="server" checked='<%#eval("MC_Security")%>' />
									</ItemTemplate>
								</asp:TemplateColumn>
	                           </Columns>
                            </asp:datagrid>                       
                     </td>
               </tr>                            
	    </table>
	    <table id="tblButtons" cellspacing="1" cellpadding="1" border="0" runat="server" style="width:91%; height: 16px;" >
            <tr>
                <td style="height: 25px;" align="center" valign="bottom">
                    <asp:button id="btnSave" Runat="Server" CssClass="txtbuttons" Text="Save" Width="120px"></asp:button>
                    <asp:button id="btnCancel" Runat="Server" CssClass="txtbuttons" Text="Cancel" Width="120px"></asp:button>
                 </td>                            
            </tr>
        </table>
    </form>
</body>
</html>
