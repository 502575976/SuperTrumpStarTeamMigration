using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

public partial class STPerTest : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        txtOutputXml.Text = string.Empty;
        if (!IsPostBack)
        {
            FillComboWSDL();
        }        
    }
    void FillComboWSDL()
    {
        DataSet ds=null;
        try
        {
            ds = new DataSet();
            ds.ReadXml(Server.MapPath("WSDL.xml"));
            cboWSDL.Items.Clear();
            cboWSDL.DataTextField = "text";
            cboWSDL.DataValueField  = "value";        
            cboWSDL.DataSource = ds;
            cboWSDL.DataBind();
            cboWSDL.Items.Insert(0, new ListItem("---  .NET Webservice  ---", "0"));
        }
        catch(Exception ex)
        {
            txtOutputXml.Text = ex.Message.ToString();
        }
        finally
        {
            if (ds != null)
            {
                ds.Dispose();
                ds = null;
            }            
        }        
    }

    void InvokeSTWebSvc(string astrMethod)
    {
        ISTWebSrvService.cSTWebSrvService lobjcSTWebSrvService;
        try
        {
            if (cboWSDL.SelectedValue == "0")
            {
                txtOutputXml.Text = "Please Select AtLeast one WSDL (DEV/QA/PROD)";
            }
            else
            {
                lobjcSTWebSrvService = new ISTWebSrvService.cSTWebSrvService();
                txtOutputXml.Text = lobjcSTWebSrvService.InvokeSTWebSvc(astrMethod, cboWSDL.SelectedValue);
            }
        }
        catch (Exception ex)
        {
            txtOutputXml.Text = ex.Message.ToString();
        }
        finally
        {
            lobjcSTWebSrvService = null;
        }
    }
    protected void btnConvertPRMToXML_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("0");
    }
    protected void btnGeneratePRMFiles_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("1");
    }
    protected void btnGetAmortizationSchedule_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("2");
    }
    protected void btnGetPricingReports_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("3");
    }
    protected void btnGetPRMParams_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("4");
    }
    protected void btnModifyPRMFiles_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("5");
    }
    protected void btnProcessPricingRequest_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("6");
    }
    protected void btnProcessMQMessage_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("7");
    }
    protected void btnRunAdhoc_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("10");
    }
    protected void btnPing_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("8");
    }
    protected void btnTest_Click(object sender, EventArgs e)
    {
        InvokeSTWebSvc("9");
    }
}
