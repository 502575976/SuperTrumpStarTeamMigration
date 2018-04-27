<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="UpdateMCFile.aspx.vb" Inherits="MoneyCostWeb.UpdateMCFile" %>
<%@ Register Src="~/Common/SSI/Top.ascx" TagName="Top"  TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<link id="Link1" href= "~/CSS/MoneyCost.css" runat="server" type ="text/css" rel="stylesheet" />
    <title>Update: Money Cost MC File</title>
    <style type="text/css">
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
		<table id="tblIndexRate" runat="server"   border="0"  style="width:100%; height: 16px;" >
             <tr><td><hr /></td></tr>
             <tr>
                <td class="text11normal">
                    <asp:label id="lblTableHeder" Font-Bold="true" Font-Underline="true" runat="server" Text="MC Files"></asp:label>
                </td>
             </tr>
             <tr>
                <td class="txt11normal" >
                       <asp:datagrid id="grdMCDetail" runat="server" BorderWidth="0px"  DataKeyField="SQ_MC_ID"
	                        AutoGenerateColumns="False"  bodyheight="100" CellSpacing="2"  AlternatingItemStyle-CssClass="textnormalAlt"
	                        CellPadding="2"  GridLines="Both" Width="980px">
	                        <FooterStyle CssClass="lightbg"></FooterStyle>
	                        <ItemStyle  CssClass="text11normal"  VerticalAlign="Top" HorizontalAlign="Left" ></ItemStyle>
	                        <HeaderStyle cssClass="TextBigHD"  VerticalAlign="Top" HorizontalAlign="Left"></HeaderStyle>
	                        <Columns>	            
	                            <asp:BoundColumn  DataField="SQ_MC_ID" Visible="False"></asp:BoundColumn>	        
	                            <asp:BoundColumn  DataField="Description" HeaderText="Description"></asp:BoundColumn>	
	                            <asp:BoundColumn  DataField="MC_CODE" HeaderText="MC Code"></asp:BoundColumn>	
	                            <asp:TemplateColumn HeaderText="Currency Code" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Left">
									<ItemTemplate>
									<asp:TextBox runat="server" CssClass="txtbox" Width="40" id="txtCurrencyCode" align="middle" name="txtCurrencyCode" Text='<%#eval("Currency_Code")%>'></asp:TextBox>
									</ItemTemplate>
								</asp:TemplateColumn>
	                            <asp:TemplateColumn HeaderText="IND Active" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Left">
									<ItemTemplate>
									<asp:TextBox runat="server" CssClass="txtbox" Width="20" id="txtINDActive" align="middle" name="txtINDActive" Text='<%#eval("IND_Active")%>'></asp:TextBox>
									</ItemTemplate>
								</asp:TemplateColumn>
	                            <asp:TemplateColumn HeaderText="Days to Skip" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Left">
									<ItemTemplate>
									<asp:TextBox runat="server" Width="20" CssClass="txtbox" id="txtDaysToSkip" align="middle" name="txtDaysToSkip" Text='<%#eval("Days_To_Skip")%>'></asp:TextBox>
									</ItemTemplate>
								</asp:TemplateColumn>
								<asp:TemplateColumn  HeaderText="Frequency" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Left">
									<ItemTemplate>
									    <asp:TextBox runat="server"  CssClass="txtbox" Width="40" id="txtFrequency" align="middle" name="txtFrequency" Text='<%#eval("Frequency")%>'></asp:TextBox>
									</ItemTemplate>
								</asp:TemplateColumn>	    	                            	                            	                        
								<asp:TemplateColumn  HeaderText="Frequency Count" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Left">
									<ItemTemplate>
									    <asp:TextBox runat="server" CssClass="txtbox" Width="40" id="txtFrequencyCount" align="middle" name="txtFrequencyCount" Text='<%#eval("Frequency_Count")%>'></asp:TextBox>
									</ItemTemplate>
								</asp:TemplateColumn>	 
										<asp:TemplateColumn   HeaderText="Process Date" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Left">
									<ItemTemplate >
									    <asp:TextBox runat="server" CssClass="txtbox"  Width="100" id="txtProcessDate" align="middle" name="txtProcessDate"  Text='<%#eval("Last_Schedule_Process_Date", "{0:MM/dd/yyyy}")%>'></asp:TextBox>
									</ItemTemplate>
								</asp:TemplateColumn>	   	                            	                            	                        
										<asp:TemplateColumn  HeaderText="Business Contact" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Left">
									<ItemTemplate>
									    <asp:TextBox runat="server" CssClass="txtbox" Width="150" id="txtBusinessContact" align="middle" name="txtBusinessContact" Text='<%#eval("Business_Contact")%>'></asp:TextBox>
									</ItemTemplate>
								</asp:TemplateColumn>	   	                            	                            	                        
							<asp:TemplateColumn  HeaderText="FTP Directory" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Left">
									<ItemTemplate>
									    <asp:TextBox runat="server" CssClass="txtbox" Width="150" id="txtFTPDirectory" align="middle" name="txtFTPDirectory" Text='<%#eval("FTP_Directory")%>'></asp:TextBox>
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
