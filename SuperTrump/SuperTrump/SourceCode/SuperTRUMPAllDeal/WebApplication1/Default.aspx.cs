using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using BSVICROIBL.BSVICROIBL;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using System.Text;

namespace WebApplication1
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //#region CODE FOR XSL TRANSFORMATION 
            //DirectoryInfo dInfo = new DirectoryInfo(@"D:\StarTeam\Transform\InputXml\");
            //if (Directory.Exists(@"D:\StarTeam\Transform\InputXml\"))
            //{
            //    FileInfo[] fInfo = dInfo.GetFiles("*.xml");
            //    System.Xml.Xsl.XslCompiledTransform xslt = default(System.Xml.Xsl.XslCompiledTransform);
            //    xslt = new System.Xml.Xsl.XslCompiledTransform();
            //    if (File.Exists(@"D:\StarTeam\Transform\Xslt\LeaseXSLT.xsl"))
            //    {
            //        xslt.Load(@"D:\StarTeam\Transform\Xslt\LeaseXSLT.xsl");
            //    }
            //    foreach (var fi in fInfo)
            //    {
            //        xslt.Transform(@"D:\StarTeam\Transform\InputXml\" + fi.Name, @"D:\StarTeam\Transform\OutputXml\" + fi.Name + ".xml");
            //    }
            //} 
            //#endregion





            ////WebApplication1.ServiceReference2.ServiceEntity objEntity = new WebApplication1.ServiceReference2.ServiceEntity();
            //ServiceReference2.SuperTRUMPAllDealServiceClient obj1 = new ServiceReference2.SuperTRUMPAllDealServiceClient();

            ////objEntity = obj1.GetServiceCredentials();



            //obj1.ClientCredentials.Windows.ClientCredential.UserName = "992000016";
            //obj1.ClientCredentials.Windows.ClientCredential.Password = "gf6eD8xp";
            //obj1.ClientCredentials.Windows.ClientCredential.Domain = "Comfin";

            ////string strXML = obj1.GenerateInputXML();
            ////string strXML = obj1.ExecuteServiceFlow();
            ////Response.Write(strXML);
            ////string strXML11 = (obj1.TestWCFService("Test").ToString());

            //string processName = "WebDev.WebServer";

            //foreach (Process process in Process.GetProcessesByName(processName))
            //{
            //    Response.Write("Killing ASP.NET worker process (Process ID:" + process.Id + ")");
            //    process.Kill();
            //}
            //obj1 = null;



            //SuperTRUMPCommon.BSVICROICommon aaa = new SuperTRUMPCommon.BSVICROICommon();
            //aaa.ReadConfigurationFileValue("MailTo");

            //try
            //{
            //    StreamReader objInput = new StreamReader(@"D:\StarTeam\InputDAT\Account_Schedule.dat", System.Text.Encoding.Default);
            //    StringBuilder sb = new StringBuilder();
            //    Int32 Id = 1;
            //    while (objInput.EndOfStream == false)
            //    {
            //        sb.Append(objInput.ReadLine());
            //        Id++;
            //    }

            //    //string contents = objInput.ReadToEnd().Trim();
            //    //string[] split = System.Text.RegularExpressions.Regex.Split(contents, "\\s+", RegexOptions.None);
            //    //foreach (string s in split)
            //    //{
            //    //    Console.WriteLine(s);
            //    //}
            //    Response.Write(sb.ToString());
            //    Response.End();
            //}
            //catch (Exception ex)
            //{
            //    Response.Write(ex.Message);
            //    Response.End();
            //}

            //cSTForAllDealsBLMGR objforalldealblmgr = new cSTForAllDealsBLMGR();
            //string strdemo = objforalldealblmgr.FillDataInXmlFormat();


            SuperTRUMPAllDeal.SuperTRUMPAllDealService StObj = new SuperTRUMPAllDeal.SuperTRUMPAllDealService();
            string strReturn = StObj.ExecuteServiceFlow();


            //int iii = objForAllDealBLMgr.MapLocation("\\\\cmfciohpapwx.comfin.ge.com\\supertrumpftp", "992000016", "gf6eD8xp");
            //objForAllDealBLMgr.SendNotificationByMail("demo");
            //objForAllDealBLMgr = null;

            //BSVICROIBL.BSVICROIBL.cSTForAllDealsBLMGR  ss = new cSTForAllDealsBLMGR();
            //ss.GetFTPInputFiles();


            //// **** Discuss
            //BSVICROIBL.BSVICROIBL.cSTForAllDealsBLMGR objMgr = new BSVICROIBL.BSVICROIBL.cSTForAllDealsBLMGR();
            ////objMgr.FillDataInXml();
            //DirectoryInfo dInfo = new DirectoryInfo(@"D:\StarTeam\Transform\OutputXml\");
            //if (Directory.Exists(@"D:\StarTeam\Transform\OutputXml\"))
            //{
            //    FileInfo[] fInfo = dInfo.GetFiles("*.xml");
            //    foreach (var fi in fInfo)
            //    {
            //        XmlDocument xmlDoc = new XmlDocument();
            //        xmlDoc.Load(@"D:\StarTeam\Transform\OutputXml\" + fi.Name);
            //        string strOutXml = xmlDoc.OuterXml;
            //        strOutXml = strOutXml.Substring(strOutXml.IndexOf("<PRM"));
            //        SuperTRUMP.ISuperTrumpServiceSoapPort obj = new SuperTRUMP.ISuperTrumpServiceSoapPort();
            //        string strSuperTRUMPOutXml = string.Empty;
            //        strSuperTRUMPOutXml = obj.RunAdHocXMLInOutQuery(strOutXml.Trim());
            //        //xmlDoc = new XmlDocument();
            //        //xmlDoc.LoadXml(strSuperTRUMPOutXml);
            //        //xmlDoc.Save(@"D:\StarTeam\Transform\RS\" + fi.Name);
            //    }
            //}



        }
    }
}
