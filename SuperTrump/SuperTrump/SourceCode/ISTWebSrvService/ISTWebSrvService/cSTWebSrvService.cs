using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Configuration;
using System.Net;
using System.Data;

namespace ISTWebSrvService
{    
    interface ISTWebSrvService
    {
        string InvokeSTWebSvc(string astrMethod, string astrWSDL,string astrWSDLPath);
        DataSet CompareSTVersions(string astrPRMFileName, string astrPRMFilePath, string astrWSDLTo, string astrWSDLFrom);
    }    
    public class cSTWebSrvService:ISTWebSrvService
    {        

            public string InvokeSTWebSvc(string astrMethod, string astrWSDL, string astrWSDLPath)
            {
                BSCEFSuperTrump.ISuperTrumpServiceSoapPort lobjISuperTrumpServiceSoapPort=null;
                BSCEFSuperTrump.IClientServiceSoapPort lobjIClientServiceSoapPort=null;               
                XmlDocument lobjXMLSTDoc;
                XmlDocument lobjXMLSTErrorDoc;
                XmlNode lobjXMLSTErrorDocNode;
                DirectoryInfo lobjDirInfo;
                string lstrFileName = string.Empty;
                string lstrFolderName = string.Empty;
                string lstrOUTXML = string.Empty;
                string lstrInPutPath = string.Empty;
                string lstrOutPutPath = string.Empty;
                string lstrINXML = string.Empty;
                string lstrOUTFileName = string.Empty;
                string lstrStatus = string.Empty;
                NetworkCredential iobjNetworkCredential;
                DataSet dsWSDL = null;
                DataView dv = null;
                try
                {
                    lobjISuperTrumpServiceSoapPort = new BSCEFSuperTrump.ISuperTrumpServiceSoapPort();
                    lobjIClientServiceSoapPort = new BSCEFSuperTrump.IClientServiceSoapPort();

                    //Change by Sumit
                    dsWSDL = new DataSet();
                    dsWSDL.ReadXml(astrWSDLPath);                    
                    dv = new DataView(dsWSDL.Tables[0]);
                    dv.RowFilter = "value ='" + astrWSDL + "'";
                    //string strUser = dv[0]["Uid"].ToString();
                    

                    if (astrMethod == "81" || astrMethod == "91")
                    {                     
                        lobjISuperTrumpServiceSoapPort.Abort();
                        lobjIClientServiceSoapPort.Abort();
                        //iobjNetworkCredential = new NetworkCredential();
                        //iobjNetworkCredential.UserName = "";
                        //iobjNetworkCredential.Password = "";

                        lobjISuperTrumpServiceSoapPort = new BSCEFSuperTrump.ISuperTrumpServiceSoapPort();
                        lobjIClientServiceSoapPort = new BSCEFSuperTrump.IClientServiceSoapPort();


                        lobjISuperTrumpServiceSoapPort.Credentials = System.Net.CredentialCache.DefaultCredentials;


                        lobjIClientServiceSoapPort.Credentials = System.Net.CredentialCache.DefaultCredentials;
                    }
                    else
                    {
                        if (string.Compare(dv[0]["Uid"].ToString(), string.Empty) != 0)
                        {
                            iobjNetworkCredential = new NetworkCredential();
                            iobjNetworkCredential.UserName = dv[0]["Uid"].ToString();
                            iobjNetworkCredential.Password = dv[0]["Pwd"].ToString();
                            lobjISuperTrumpServiceSoapPort.UseDefaultCredentials = false;
                            lobjISuperTrumpServiceSoapPort.Credentials = iobjNetworkCredential;

                            lobjIClientServiceSoapPort.UseDefaultCredentials = false;
                            lobjIClientServiceSoapPort.Credentials = iobjNetworkCredential;
                        }


                        //if (astrWSDL.LastIndexOf("https://bscefsupertrump.qa.comfin.ge.com") > 0 || astrWSDL.LastIndexOf("http://3.239.150.22/") > 0 || astrWSDL.LastIndexOf("http://3.239.150.81/") > 0)
                        //{
                        //    iobjNetworkCredential = new NetworkCredential();
                        //    iobjNetworkCredential.UserName = ConfigurationManager.AppSettings["USERQA"];
                        //    iobjNetworkCredential.Password = ConfigurationManager.AppSettings["USERKEYQA"];
                        //    lobjISuperTrumpServiceSoapPort.UseDefaultCredentials = false;
                        //    lobjISuperTrumpServiceSoapPort.Credentials = iobjNetworkCredential;

                        //    lobjIClientServiceSoapPort.UseDefaultCredentials = false;
                        //    lobjIClientServiceSoapPort.Credentials = iobjNetworkCredential;
                        //}

                        //if (astrWSDL.LastIndexOf("https://bscefsupertrump.comfin.ge.com") > 0)
                        //{
                        //    iobjNetworkCredential = new NetworkCredential();
                        //    iobjNetworkCredential.UserName = ConfigurationManager.AppSettings["USERPROD"];
                        //    iobjNetworkCredential.Password = ConfigurationManager.AppSettings["USERKEYPROD"];
                        //    lobjISuperTrumpServiceSoapPort.UseDefaultCredentials = false;
                        //    lobjISuperTrumpServiceSoapPort.Credentials = iobjNetworkCredential;

                        //    lobjIClientServiceSoapPort.UseDefaultCredentials = false;
                        //    lobjIClientServiceSoapPort.Credentials = iobjNetworkCredential;
                        //}
                    }
                        
                                       
                    switch (astrMethod)
                    {
                        case "0":
                            lstrFolderName = "ConvertPRMToXML";
                            break;
                        case "1":
                            lstrFolderName = "GeneratePRMFiles";
                            break;
                        case "2":
                            lstrFolderName = "GetAmortizationSchedule";
                            break;
                        case "3":
                            lstrFolderName = "GetPricingReports";
                            break;
                        case "4":
                            lstrFolderName = "GetPRMParams";
                            break;
                        case "5":
                            lstrFolderName = "ModifyPRMFiles";
                            break;
                        case "6":
                            lstrFolderName = "ProcessPricingRequest";
                            break;
                        case "7":
                            lstrFolderName = "ProcessMQMessage";
                            break;
                        case "8":
                            lobjIClientServiceSoapPort.Url = astrWSDL;                            
                            lstrOUTXML = lobjIClientServiceSoapPort.Ping();
                            lobjIClientServiceSoapPort.UseDefaultCredentials = true; 
                            return lstrOUTXML;                            
                        case "9":
                            lobjISuperTrumpServiceSoapPort.Url = astrWSDL;
                            lstrOUTXML = lobjISuperTrumpServiceSoapPort.Test();
                            lobjISuperTrumpServiceSoapPort.UseDefaultCredentials = true;
                            return lstrOUTXML;
                        case "10":
                            lstrFolderName = "RunAdhocXMLQuery";
                            break;
                        case "81":                                                       
                            lobjIClientServiceSoapPort.Url = astrWSDL;
                            lstrOUTXML = lobjIClientServiceSoapPort.Ping();                            
                            return lstrOUTXML;
                        case "91":                            
                            lobjISuperTrumpServiceSoapPort.Url = astrWSDL;
                            lstrOUTXML = lobjISuperTrumpServiceSoapPort.Test();                            
                            return lstrOUTXML;
                    }
                    
                    lstrInPutPath =  ConfigurationManager.AppSettings["INPUT_FILE_PATH"] + "\\" + lstrFolderName;
                    lstrOutPutPath = ConfigurationManager.AppSettings["OUTPUT_FILE_PATH"] + "\\" + lstrFolderName;

                    lobjDirInfo = new DirectoryInfo(lstrInPutPath);
                    lobjXMLSTDoc = new XmlDocument();
                                       
                    if (lobjDirInfo.GetFiles().Count() <= 0)
                    {
                        lstrStatus = "There is no file exist in " + lstrInPutPath + " Folder";
                        return lstrStatus;
                    }
                    foreach (FileInfo lobjFileInfo in lobjDirInfo.GetFiles())
                    {
                        try
                        {

                            lstrFileName = lobjFileInfo.Name;
                            lobjXMLSTDoc.Load(lstrInPutPath + "\\" + lstrFileName);
                            lstrINXML = lobjXMLSTDoc.InnerXml;
                            switch (astrMethod)
                            {
                                case "0":                                   
                                    lobjISuperTrumpServiceSoapPort.Url = astrWSDL;
                                    lstrOUTXML = lobjISuperTrumpServiceSoapPort.ConvertPRMToXML(lstrINXML);
                                    break;
                                case "1":
                                    lobjISuperTrumpServiceSoapPort.Url = astrWSDL;
                                    lstrOUTXML = lobjISuperTrumpServiceSoapPort.GeneratePRMFiles(lstrINXML);
                                    break;
                                case "2":
                                    lobjISuperTrumpServiceSoapPort.Url = astrWSDL;
                                    lstrOUTXML = lobjISuperTrumpServiceSoapPort.GetAmortizationSchedule(lstrINXML);
                                    break;
                                case "3":
                                    lobjISuperTrumpServiceSoapPort.Url = astrWSDL;
                                    lstrOUTXML = lobjISuperTrumpServiceSoapPort.GetPricingReports(lstrINXML);
                                    break;
                                case "4":
                                    lobjISuperTrumpServiceSoapPort.Url = astrWSDL;
                                    lstrOUTXML = lobjISuperTrumpServiceSoapPort.GetPRMParams(lstrINXML);
                                    break;
                                case "5":
                                    lobjISuperTrumpServiceSoapPort.Url = astrWSDL;
                                    lstrOUTXML = lobjISuperTrumpServiceSoapPort.ModifyPRMFiles(lstrINXML);
                                    break;
                                case "6":
                                    lobjIClientServiceSoapPort.Url = astrWSDL;
                                    lstrOUTXML = lobjIClientServiceSoapPort.ProcessPricingRequest(lstrINXML);
                                    break;
                                case "7":
                                    lobjIClientServiceSoapPort.Url = astrWSDL;
                                    lstrOUTXML = lobjIClientServiceSoapPort.ProcessMQMessage(lstrINXML);
                                    break;                                
                                case "10":
                                    lobjISuperTrumpServiceSoapPort.Url = astrWSDL;
                                    lstrOUTXML = lobjISuperTrumpServiceSoapPort.RunAdHocXMLInOutQuery(lstrINXML);
                                    break;
                            }
                            lstrOUTFileName = lstrFileName + "_OUT.xml";
                            lstrOUTFileName = lstrOutPutPath + "\\" + lstrOUTFileName;
                            lobjXMLSTErrorDoc = new XmlDocument();
                            lobjXMLSTErrorDoc.LoadXml(lstrOUTXML);
                            lobjXMLSTErrorDocNode = lobjXMLSTErrorDoc.SelectSingleNode("ERROR_DESC");
                            if (lobjXMLSTErrorDocNode != null)
                            {
                                if (lobjXMLSTErrorDocNode.InnerText.Length > 0)
                                {
                                    lstrStatus = lstrStatus + lobjXMLSTErrorDocNode.InnerText;
                                }
                            }                            
                            if(SaveOutput(lstrOUTFileName,lstrOUTXML))
                            {
                                lstrStatus = lstrStatus + lstrFileName + " - Successfully Done" + "\n";
                            }                            
                        }
                        catch(Exception ex)
                        {
                            lstrStatus = lstrStatus + lstrFileName + " - Failed, " + ex.Source.ToString() + ":-" + ex.Message.ToString();
                        }                        
                    }
                    if (lstrStatus == string.Empty)
                    {
                        lstrStatus = "There is no file exist in " + lstrInPutPath + " Folder";
                    }
                    return lstrStatus;
                }
                catch (Exception ex)
                {
                    return "Error occurred - " + ex.Message.ToString();
                }
                finally
                {
                    if (lobjISuperTrumpServiceSoapPort != null)
                    {
                        lobjISuperTrumpServiceSoapPort.Dispose();                       
                        
                    }
                    if (lobjIClientServiceSoapPort != null)
                    {
                        lobjIClientServiceSoapPort.Dispose();                        
                    }                    
                    lobjXMLSTDoc=null;
                    lobjDirInfo=null;
                    lobjXMLSTErrorDoc = null;
                    lobjXMLSTErrorDocNode = null;
                    iobjNetworkCredential = null;
                }
            }
            private bool SaveOutput(string astrFileName, string astrData)
            {
                StreamWriter lobjStream;
                try
                {
                    if (File.Exists(astrFileName))
                    {
                        File.Delete(astrFileName);                        
                    }
                   
                    lobjStream = File.CreateText(astrFileName);
                    lobjStream.WriteLine(astrData);
                    lobjStream.Close();                               
                    return true;
                }
                catch (Exception ex)
                {
                    throw ex;                    
                }
                finally
                {
                    lobjStream = null;
                }
            }
        private string GetBase64File(string astrFilePath)
        {
            FileStream lobjFS=default(FileStream);
            BinaryReader lobjBinReader=default(BinaryReader);
            Byte[] larrFile;
            string lstrResult = string.Empty; 
            try
            {
                lobjFS=new FileStream(astrFilePath, FileMode.Open, FileAccess.Read);
                lobjBinReader=new BinaryReader(lobjFS);
                larrFile=lobjBinReader.ReadBytes(Convert.ToInt16(lobjFS.Length));
                lstrResult = Convert.ToBase64String(larrFile);
                return lstrResult;
            }
            catch(Exception ex)
            {
                return lstrResult;
            }
            finally
            {
                lobjFS=null;
                lobjBinReader=null;
            }
        }
        private int  FindinDataSet(DataTable  aobjDT, string strValue)
        {
            int intCounter=0;
            try
            {
                foreach (DataRow dtRow in aobjDT.Rows)
                {                   
                    if (dtRow["XML_NODE"].ToString().ToUpper().Trim() == strValue.ToUpper().Trim())
                    {
                        return intCounter;
                    }
                    intCounter++;
                }
                return -1;
            }
            catch (Exception ex)
            {
                return intCounter;
            }
            finally
            {
                
            }
        }
        public DataSet CompareSTVersions(string astrPRMFileName, string astrPRMFilePath, string astrWSDLTo, string astrWSDLFrom)
            {
                string lobjXMLTo = string.Empty;
                string lobjXMLFrom = string.Empty;
                DataSet lobjComparedDS = null;
                DataTable lobjComparedDT=null;
                DataColumn lobjDTCol1 = null;
                DataColumn lobjDTCol2 = null;
                DataColumn lobjDTCol3 = null;
                string lstrPRM_FILE_XML = string.Empty;
                XmlDocument lobjXMLDoc = null;
                BSCEFSuperTrump.ISuperTrumpServiceSoapPort lobjISuperTrumpServiceSoapPortFrom = null;
                BSCEFSuperTrump.ISuperTrumpServiceSoapPort lobjISuperTrumpServiceSoapPortTo = null;

                DataSet lobjDSFrom = null;
                DataSet lobjDSTo = null;
                DataRow dtRow = null;
                int intCountRow = 0;
                NetworkCredential iobjNetworkCredentialFrom;
                NetworkCredential iobjNetworkCredentialTo;

                String lstrErrorMessage = string.Empty;
                try
                {

                    lobjComparedDS = new DataSet();

                    //lstrPRM_FILE_XML = "<PRM_FILE_LIST><PRM_FILE><FILE_NAME></FILE_NAME><FILE_DATA xmlns:dt='urn:schemas-microsoft-com:datatypes' dt:dt='bin.base64'></FILE_DATA></PRM_FILE></PRM_FILE_LIST>";
                    lobjXMLDoc = new XmlDocument();
                    //lobjXMLDoc.LoadXml(lstrPRM_FILE_XML);

                    //astrPRMFileName.Replace("&", "");
                    //astrPRMFileName.Replace("'", "");
                    //astrPRMFileName.Replace("*", "");
                    //astrPRMFileName.Replace("!", "");
                    //astrPRMFileName.Replace("\\", "");
                    //astrPRMFileName.Replace("/", "");
                    //astrPRMFileName.Replace("^", "");
                    //astrPRMFileName.Replace("?", "");
                    //astrPRMFileName.Replace("|", "");
                    //astrPRMFileName.Replace(":", "");
                    //astrPRMFileName.Replace(";", "");

                    //lobjXMLDoc.SelectSingleNode("//PRM_FILE_LIST/PRM_FILE/FILE_NAME").InnerText = astrPRMFileName;
                    //lobjXMLDoc.SelectSingleNode("//PRM_FILE_LIST/PRM_FILE/FILE_DATA").InnerText = GetBase64File(astrPRMFilePath) + "=";
                    lobjXMLDoc.Load(astrPRMFilePath);

                    lobjISuperTrumpServiceSoapPortFrom = new BSCEFSuperTrump.ISuperTrumpServiceSoapPort();
                    lobjISuperTrumpServiceSoapPortTo = new BSCEFSuperTrump.ISuperTrumpServiceSoapPort();

                    if (astrWSDLFrom.LastIndexOf("https://bscefsupertrump.qa.comfin.ge.com") > 0)
                    {
                        iobjNetworkCredentialFrom = new NetworkCredential();
                        iobjNetworkCredentialFrom.UserName = ConfigurationManager.AppSettings["USERQA"];
                        iobjNetworkCredentialFrom.Password = ConfigurationManager.AppSettings["USERKEYQA"];
                        lobjISuperTrumpServiceSoapPortFrom.UseDefaultCredentials = false;
                        lobjISuperTrumpServiceSoapPortFrom.Credentials = iobjNetworkCredentialFrom;
                    }

                    if (astrWSDLTo.LastIndexOf("https://bscefsupertrump.qa.comfin.ge.com") > 0)
                    {
                        iobjNetworkCredentialTo = new NetworkCredential();
                        iobjNetworkCredentialTo.UserName = ConfigurationManager.AppSettings["USERQA"];
                        iobjNetworkCredentialTo.Password = ConfigurationManager.AppSettings["USERKEYQA"];
                        lobjISuperTrumpServiceSoapPortTo.UseDefaultCredentials = false;
                        lobjISuperTrumpServiceSoapPortTo.Credentials = iobjNetworkCredentialTo;
                    }

                    if (astrWSDLFrom.LastIndexOf("https://bscefsupertrump.comfin.ge.com") > 0)
                    {
                        iobjNetworkCredentialFrom = new NetworkCredential();
                        iobjNetworkCredentialFrom.UserName = ConfigurationManager.AppSettings["USERPROD"];
                        iobjNetworkCredentialFrom.Password = ConfigurationManager.AppSettings["USERKEYPROD"];
                        lobjISuperTrumpServiceSoapPortFrom.UseDefaultCredentials = false;
                        lobjISuperTrumpServiceSoapPortFrom.Credentials = iobjNetworkCredentialFrom;
                    }

                    if (astrWSDLTo.LastIndexOf("https://bscefsupertrump.comfin.ge.com") > 0)
                    {
                        iobjNetworkCredentialTo = new NetworkCredential();
                        iobjNetworkCredentialTo.UserName = ConfigurationManager.AppSettings["USERPROD"];
                        iobjNetworkCredentialTo.Password = ConfigurationManager.AppSettings["USERKEYPROD"];
                        lobjISuperTrumpServiceSoapPortTo.UseDefaultCredentials = false;
                        lobjISuperTrumpServiceSoapPortTo.Credentials = iobjNetworkCredentialTo;
                    }




                    lobjISuperTrumpServiceSoapPortFrom.Url = astrWSDLFrom;
                    lobjXMLFrom = lobjISuperTrumpServiceSoapPortFrom.ConvertPRMToXML(lobjXMLDoc.OuterXml);

                    lobjDSFrom = new DataSet();                    
                    lobjDSFrom.EnforceConstraints = false;                    
                    //DataSet.EnforceConstraints

                    lobjDSFrom.ReadXml(new StringReader(lobjXMLFrom));



                    lobjISuperTrumpServiceSoapPortTo.Url = astrWSDLTo;
                    lobjXMLTo = lobjISuperTrumpServiceSoapPortTo.ConvertPRMToXML(lobjXMLDoc.OuterXml);

                    lobjDSTo = new DataSet();
                    lobjDSFrom.EnforceConstraints = false;
                    lobjDSTo.ReadXml(new StringReader(lobjXMLTo));

                    lobjComparedDS=new DataSet();
                    foreach (DataTable dtTable in lobjDSFrom.Tables)
                    {
                        lobjComparedDT=new DataTable(dtTable.TableName);
                        lobjDTCol1 = new DataColumn("XML_NODE");                       
                        lobjDTCol2 = new DataColumn("XMLFROMVALUE");                        
                        lobjDTCol3 = new DataColumn("XMLTOVALUE");                        
                        lobjComparedDT.Columns.Add(lobjDTCol1);
                        lobjComparedDT.Columns.Add(lobjDTCol2);
                        lobjComparedDT.Columns.Add(lobjDTCol3);                        
                        foreach (DataColumn dtCol in dtTable.Columns)
                        {
                           dtRow = lobjComparedDT.NewRow();
                           dtRow["XML_NODE"] = dtCol.ColumnName;
                           lobjComparedDT.Rows.Add(dtRow);
                           lobjComparedDT.Rows[lobjComparedDT.Rows.Count - 1]["XMLFROMVALUE"] = dtTable.Rows[0][dtCol.ColumnName];                           
                        }
                        lobjComparedDS.Tables.Add(lobjComparedDT);
                    }                    

                    
                    foreach (DataTable dtTable in lobjDSTo.Tables)
                    {
                        if (lobjComparedDS.Tables.Contains(dtTable.TableName))
                        {
                            foreach (DataColumn dtCol in dtTable.Columns)
                            {
                                intCountRow = FindinDataSet(lobjComparedDS.Tables[dtTable.TableName], dtCol.ColumnName);

                                if (intCountRow != -1)
                                {
                                    lobjComparedDS.Tables[dtTable.TableName].Rows[intCountRow]["XMLTOVALUE"] = dtTable.Rows[0][dtCol.ColumnName];
                                }
                                else
                                {                                   
                                    dtRow = lobjComparedDT.NewRow();
                                    dtRow["XML_NODE"] = dtCol.ColumnName;
                                    dtRow["XMLTOVALUE"] = dtTable.Rows[0][dtCol.ColumnName];
                                    lobjComparedDT.Rows.Add(dtRow);                                 
                                }
                            }
                        }
                        else
                        {
                            lobjComparedDT = new DataTable(dtTable.TableName);
                            lobjDTCol1 = new DataColumn("XML_NODE");
                            lobjDTCol1 = new DataColumn("XML_NODE");
                            lobjDTCol2 = new DataColumn("XMLFROMVALUE");                            
                            lobjDTCol3 = new DataColumn("XMLTOVALUE");                            
                            lobjComparedDT.Columns.Add(lobjDTCol1);
                            lobjComparedDT.Columns.Add(lobjDTCol2);
                            lobjComparedDT.Columns.Add(lobjDTCol3);
                            foreach (DataColumn dtCol in dtTable.Columns)
                            {
                                dtRow = lobjComparedDT.NewRow();
                                dtRow["XML_NODE"] = dtCol.ColumnName;
                                lobjComparedDT.Rows.Add(dtRow);
                                lobjComparedDT.Rows[lobjComparedDT.Rows.Count - 1]["XMLTOVALUE"] = dtTable.Rows[0][dtCol.ColumnName];
                            }
                            lobjComparedDS.Tables.Add(lobjComparedDT);
                        }
                    }
                    if (string.IsNullOrEmpty(cSTWebSrvError.Error_Message.ToString()) == true)
                    {
                        cSTWebSrvError.Error_Message = string.Empty; 
                    }
                    return lobjComparedDS;
                }
                catch (Exception ex)
                {
                    cSTWebSrvError.Error_Message = cSTWebSrvError.Error_Message.ToString() + "<li>" + astrPRMFileName + "<ul><li>Error Message:- " + ex.Message.ToString() + "</li></ul></li>";                    
                    return lobjComparedDS;
                }
                finally
                {
                   lobjComparedDS = null;
                   lobjComparedDT = null;
                   lobjDTCol1 = null;
                   lobjDTCol2 = null;
                   lobjDTCol3 = null;
                   lobjXMLDoc = null;
                   lobjISuperTrumpServiceSoapPortFrom = null;
                   lobjISuperTrumpServiceSoapPortTo = null;
                   lobjDSFrom = null;
                   lobjDSTo = null;
                   dtRow = null;
                   iobjNetworkCredentialTo = null;
                   iobjNetworkCredentialFrom = null;
                }
            }
    }
    public static class cSTWebSrvError
    {
        public static string gstrError=string.Empty;
        public static string Error_Message
        {
            get
            {
                return gstrError;
            }
            set
            {
                gstrError = value;
            }
        }
    }
}
