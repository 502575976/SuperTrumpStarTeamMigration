<%@ Page Language="C#" AutoEventWireup="true" CodeFile="STPerTest.aspx.cs" Inherits="STPerTest" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>SuperTRUMP WebService Test</title>
    <style type="text/css">
        .text {
	        font-family: Arial, Helvetica, sans-serif;
	        font-size: 12px;
	        font-style: normal;
	        font-weight: bold;
        }

        .button {
	        font-family: Arial, Helvetica, sans-serif;
	        font-weight: bold;
	        color: #FFFFFF;
	        border: 1pt solid #cdcdcd;
	        font-size: 11px;
	        cursor:hand;
	        background-color: #7AA6BA;
        }
    </style>   
</head>
<body>
    <form id="form1" runat="server">    
        <table width="70%" border="1" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr style="background-image: url(bg.gif); "> 
                            <td>
                                <img src="images/top_new.jpg" alt="top"  height="121"/>
                            </td>
                        </tr>                        
                    </table> 
                </td>
            </tr>                       
            <tr>
                <td align="center">
                    <table cellspacing="1" cellpadding="1" width="80%" border="0">    
                        <tr>
                            <td align="center" height="15">&nbsp;</td>
                        </tr>
                        <tr> 
                            <td valign="top">
                                <asp:Label ID="lblWSDL" runat="server" CssClass="text">WEB SERVICE</asp:Label>
                            </td>
                            <td colspan="3">
                                <asp:DropDownList ID="cboWSDL" runat="server" CssClass="text" width="645px"></asp:DropDownList>                            		
		                    </td>
                        </tr>   
                        <tr> 
                            <td colspan="4" height="5"></td>    
                        </tr>
                        <tr> 
                            <td valign="top"></td>
                            <td colspan="3" height="170">
                                <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                                    <tr> 
                                        <td align="left"> 
                                            <asp:Button id="btnConvertPRMToXML" runat="server" Text="ConvertPRMToXML" 
                                                CssClass="button" onclick="btnConvertPRMToXML_Click" />                                            
                                        </td>
                                        <td align="left"> 
                                            <asp:Button id="btnGeneratePRMFiles" runat="server" Text="GeneratePRMFiles" 
                                                CssClass="button" onclick="btnGeneratePRMFiles_Click" />                                            
                                        </td>
                                        <td align="left"> 
                                            <asp:Button id="btnGetAmortizationSchedule" runat="server" 
                                                Text="GetAmortizationSchedule" CssClass="button" 
                                                onclick="btnGetAmortizationSchedule_Click" />                                            
                                        </td>
                                    </tr>
                                    <tr> 
                                        <td align="left" height="8">&nbsp;</td>
                                        <td align="left" height="8">&nbsp;</td>
                                        <td align="left" height="8">&nbsp;</td>
                                    </tr>
                                    <tr> 
                                        <td align="left">
                                            <asp:Button id="btnGetPricingReports" runat="server" Text="GetPricingReports" 
                                                CssClass="button" onclick="btnGetPricingReports_Click" />                                            
                                        </td>
                                        <td align="left">
                                            <asp:Button id="btnGetPRMParams" runat="server" Text="GetPRMParams" 
                                                CssClass="button" onclick="btnGetPRMParams_Click" />                                            
                                        </td>
                                        <td align="left">
                                            <asp:Button id="btnModifyPRMFiles" runat="server" Text="ModifyPRMFiles" 
                                                CssClass="button" onclick="btnModifyPRMFiles_Click" />                                            
                                        </td>
                                    </tr>
                                    <tr> 
                                        <td align="left" height="8">&nbsp;</td>
                                        <td align="left" height="8">&nbsp;</td>
                                        <td align="left" height="8">&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td align="left">
                                            <asp:Button id="btnProcessPricingRequest" runat="server" 
                                                Text="ProcessPricingRequest" CssClass="button" 
                                                onclick="btnProcessPricingRequest_Click" />                                            
                                        </td>
                                        <td align="left">
                                            <asp:Button id="btnProcessMQMessage" runat="server" Text="ProcessMQMessage" 
                                                CssClass="button" onclick="btnProcessMQMessage_Click" />                                            
                                        </td>
                                        <td align="left">
                                            <asp:Button id="btnRunAdhoc" runat="server" Text="RunAdhoc" CssClass="button" 
                                                onclick="btnRunAdhoc_Click" />                                            
                                        </td>
		                            </tr>
                                    <tr> 
                                        <td align="left" height="8">&nbsp;</td>
                                        <td align="left" height="8">&nbsp;</td>
                                        <td align="left" height="8">&nbsp;</td>
                                    </tr>
		                            <tr>
                                        <td align="left">
                                            <asp:Button id="btnPing" runat="server" Text="Ping" CssClass="button" 
                                                onclick="btnPing_Click" />                                            
                                        </td>
                                        <td align="left">
                                            <asp:Button id="btnTest" runat="server"  Text="Test" CssClass="button" 
                                                onclick="btnTest_Click" />                                            
                                        </td>
                                        <td>
                                        </td>
                                    </tr>
                                </table> 
                            </td>
                        </tr>
                        <tr> 
                            <td colspan="4" height="10">
                            </td>
                        </tr>
                        <tr> 
                            <td valign="top">
                                <asp:Label ID="lblOUTPUT" runat="server" CssClass="text">OUTPUT XML:</asp:Label>
                            </td>
                            <td colspan="3">
                                <asp:TextBox Rows="8" Columns="58" Width="645px" Height="200px" ID="txtOutputXml" runat="server" TextMode="MultiLine" CssClass="text"></asp:TextBox>                                                                
                            </td>
                        </tr>
                        <tr  height="10"> 
                            <td colspan="4">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                            <td width="7%" bgcolor="7AA6BA">
                                <img src="images/gelogo.gif" alt="logo" width="65" height="33"/>                                
                            </td>
                            <td width="93%" bgcolor="7AA6BA">
                                <div align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif"><strong>&nbsp;</strong></font>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
       </table>
   </form>
</body>
</html>
