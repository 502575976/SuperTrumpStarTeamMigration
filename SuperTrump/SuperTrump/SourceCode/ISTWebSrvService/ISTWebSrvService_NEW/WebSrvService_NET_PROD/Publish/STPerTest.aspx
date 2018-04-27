<%@ page language="C#" autoeventwireup="true" inherits="STPerTest, App_Web_ix58c0lg" %>

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
                            <td colspan="2" height="5"></td>    
                        </tr>
                        <tr> 
                            <td valign="top"></td>
                            <td colspan="2" height="50">
                                <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">                                    
		                            <tr>
                                        <td align="left">
                                            <asp:Button id="btnPing" runat="server" Text="Ping" CssClass="button" 
                                                onclick="btnPing_Click" />                                            
                                        </td>
                                        <td align="right">
                                            <asp:Button id="btnTest" runat="server"  Text="Test" CssClass="button" 
                                                onclick="btnTest_Click" />                                            
                                        </td>
                                    </tr>
                                </table> 
                            </td>
                        </tr>
                        <tr> 
                            <td colspan="2" height="10">
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
