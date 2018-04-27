using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Configuration;

namespace ISTWebSrvService
{

    interface ISTWebSrvService
    {
        string InvokeSTWebSvc(string astrMethod, string astrWSDL);
    }
    public class cSTWebSrvService:ISTWebSrvService
    {      
            public string InvokeSTWebSvc(string astrMethod, string astrWSDL)
            {
                BSCEFSuperTrump.ISuperTrumpServiceSoapPort lobjISuperTrumpServiceSoapPort;
                BSCEFSuperTrump.IClientServiceSoapPort lobjIClientServiceSoapPort;
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
                try
                {
                    lobjISuperTrumpServiceSoapPort = new BSCEFSuperTrump.ISuperTrumpServiceSoapPort();
                    lobjIClientServiceSoapPort = new BSCEFSuperTrump.IClientServiceSoapPort();                  
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
                            return lstrOUTXML;                            
                        case "9":
                            lobjISuperTrumpServiceSoapPort.Url = astrWSDL;
                            lstrOUTXML = lobjISuperTrumpServiceSoapPort.Test();
                            return lstrOUTXML;
                        case "10":
                            lstrFolderName = "RunAdhocXMLQuery";
                            break;
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
                    lobjISuperTrumpServiceSoapPort=null;
                    lobjIClientServiceSoapPort=null;
                    lobjXMLSTDoc=null;
                    lobjDirInfo=null;
                    lobjXMLSTErrorDoc = null;
                    lobjXMLSTErrorDocNode = null;
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
    }
}
