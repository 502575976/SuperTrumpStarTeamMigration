<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="UpdateCsvFiles.aspx.vb" Inherits="MoneyCostWeb.UpdateCsvFiles" %>
<%@ Register Src="~/Common/SSI/Top.ascx" TagName="Top"  TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>upload CSV file</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table align="center" border="0" cellpadding="0" cellspacing="0" class="tdbg1" width="100%">
            <tr>
                <td>
                    <uc1:top id="Top1" runat="server"></uc1:top>
                </td>
            </tr>
            <tr>
                <td align=center>
                    <asp:Label ID="lblErrorMessage" runat="server" Font-Bold="true" Font-Names="Geneva, Arial, Helvetica, sans-serif"
                        ForeColor="red" Text=""></asp:Label>
                    <asp:HiddenField ID="hSSOId" runat="server" Value="" />
                </td>
            </tr>
        </table>
        <br />
        <table align="center" border="0" cellpadding="0" cellspacing="0" class="main" width="100%">
            <tr bgcolor="#ffffff">
                <td colspan="2" height="30">
                    <strong>&nbsp; Welcome </strong>
                    <asp:Label ID="lblUserName" runat="server"></asp:Label>
                </td>
            </tr>
            <tr bgcolor="#ffffff">
                <td colspan="2">
                    <asp:Label ID="Label2" runat="server" Font-Bold="True" Font-Names="Geneva,Arial,Helvetica,sans-serif"
                        Font-Size="Smaller" ForeColor="Red" Text="Note:- Please take backup before copying any file"
                        Width="513px"></asp:Label></td>
            </tr>
            <tr>
                <td class="text11normal">
                    <asp:Label ID="lblNoMCDetail" runat="server"></asp:Label></td>
            </tr>
        </table>
        <table id="tblIndexRate" runat="server" border="0" style="width: 70%; height: 16px">
            <tr>
                <td>
                    <hr />
                </td>
            </tr>
            <tr>
                <td class="text11normal">
                    <asp:Label ID="lblTableHeder" runat="server" Font-Bold="True" Font-Underline="True"
                        Text="Select CSV File Path"></asp:Label>
                    </td>
            </tr>
            <tr>
                <td class="text11normal" style="height: 24px">
                    <asp:FileUpload ID="cntrlFileUpload" runat="server" Width="428px" /></td>
                    
                    
                <td style="height: 24px">
                    <div id="chkbox" runat="server">
                    </div>
                </td>
            </tr>
            <tr>
                <td class="txt11normal" colspan="2">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Underline="True" Text="Select Destination File Path"
                        Width="249px"></asp:Label>
                    </td>
            </tr>
            <tr>
                <td class="txt11normal" colspan="2">
                    <asp:TextBox ID="txtDestinationPath" runat="server" Width="465px"></asp:TextBox>
                    <asp:Label ID="Label3" runat="server" Font-Bold="True" Font-Names="Geneva,Arial,Helvetica,sans-serif"
                        Font-Size="Smaller" ForeColor="Red" Text="(Use share drive path)"></asp:Label></td>
            </tr>
        </table>
        <table id="tblButtons" runat="server" border="0" cellpadding="1" cellspacing="1"
            style="width: 91%; height: 16px">
            <tr>
                <td align="center" style="height: 25px" valign="bottom">
                    <asp:Button ID="btnSave" runat="Server" CssClass="txtbuttons" Text="Copy" Width="120px" />
                    <asp:Button ID="btnCancel" runat="Server" CssClass="txtbuttons" Text="Cancel" Width="120px" />
                </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
