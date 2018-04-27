using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.Transactions;
using BSVICROIEntity.BSVICROIEntity;
using BSVICROIBL.BSVICROIBL;
using BSVICROIEntity;

namespace SuperTRUMPAllDeal
{
    
    // NOTE: If you change the class name "SuperTRUMPWCFService" here, you must also update the reference to "SuperTRUMPWCFService" in App.config.
    [ServiceBehavior(TransactionAutoCompleteOnSessionClose = true,TransactionIsolationLevel = IsolationLevel.Serializable)]
    public class SuperTRUMPAllDealService : ISuperTRUMPAllDealService
    {       
        [OperationBehavior(TransactionAutoComplete = true, TransactionScopeRequired = true)]
        public string TestSuperTRUMPAllDealService(string message)
        {
            cSTForAllDealsBLMGR objForAllDealBLMgr = new cSTForAllDealsBLMGR();

            return string.Format("Message Received {0}:{1}", DateTime.Now, message + " - " + objForAllDealBLMgr.TestServiceinBL());
        }

        [OperationBehavior(TransactionAutoComplete = true, TransactionScopeRequired = true)]
        public string LoadDWData()
        {
            cSTForAllDealsEntity obj1 = new cSTForAllDealsEntity();

            cSTForAllDealsBLMGR objForAllDealBLMgr = new cSTForAllDealsBLMGR();
            try
            {
                objForAllDealBLMgr.ExportDATFileDataInDatabase();
                return "done";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                objForAllDealBLMgr = null;
                obj1 = null;
            }

        }
       
        [OperationBehavior(TransactionAutoComplete = true, TransactionScopeRequired = true)]
        public string GenerateInputXML()
        {
            cSTForAllDealsEntity obj1 = new cSTForAllDealsEntity();

            cSTForAllDealsBLMGR objForAllDealBLMgr = new cSTForAllDealsBLMGR();
            try
            {

                return objForAllDealBLMgr.FillDataInXmlFormat();
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
            finally
            {
                objForAllDealBLMgr = null;
                obj1 = null;
            }
        }

       
        public string   ExecuteServiceFlow()
        {          
            cSTForAllDealsBLMGR objForAllDealBLMgr = new cSTForAllDealsBLMGR();

            try
            {

                if (objForAllDealBLMgr.GetFTPInputFiles().ToString().Contains("Error") != true)
                {
                    objForAllDealBLMgr.SendNotificationByMail("Execute GetFTPInputFiles function successfully", false);
                }
                else
                {
                    objForAllDealBLMgr.SendNotificationByMail("Error to Execute GetFTPInputFiles function", true);
                }


                if (objForAllDealBLMgr.ExportDATFileDataInDatabase().ToString().Contains("ERROR") != true)
                {
                    objForAllDealBLMgr.SendNotificationByMail("Execute ExportDATFileDataInDatabase function successfully", false);
                }
                else
                {
                    objForAllDealBLMgr.SendNotificationByMail("Error to Execute ExportDATFileDataInDatabase function", true);
                }


                if (objForAllDealBLMgr.FillDataInXmlFormat().ToString().Contains("DONE SUCCESSFULLY") == true)
                {
                    objForAllDealBLMgr.SendNotificationByMail("Execute FillDataInXml function successfully", false);
                }
                else
                {
                    objForAllDealBLMgr.SendNotificationByMail("Error to Execute FillDataInXml function", true);
                }


                return "Done";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                objForAllDealBLMgr = null;
               
            }
        }

        //public ServiceEntity GetServiceCredentials()
        //{
        //    cSTForAllDealsBLMGR objForAllDealBLMgr = new cSTForAllDealsBLMGR();
        //    try
        //    {
        //        return objForAllDealBLMgr.GetServiceCredentials();
        //    }
        //    catch (Exception)
        //    {
        //        return null;
        //        throw;
        //    }
        //    finally
        //    {
        //        objForAllDealBLMgr = null;
        //    }
            

               

        //}


    }



    

}
