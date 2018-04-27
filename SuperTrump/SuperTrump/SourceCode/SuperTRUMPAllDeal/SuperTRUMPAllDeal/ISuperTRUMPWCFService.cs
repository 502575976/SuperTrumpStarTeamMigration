using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using BSVICROIEntity.BSVICROIEntity;
using BSVICROIEntity;

namespace SuperTRUMPAllDeal
{
    // NOTE: If you change the interface name "ISuperTRUMPWCFService" here, you must also update the reference to "ISuperTRUMPWCFService" in App.config.
    [ServiceContract(Namespace = "http://www.garimasikarwar.com")]
    public interface ISuperTRUMPAllDealService
    {       
        [OperationContract]
        [TransactionFlow(TransactionFlowOption.Allowed)]
        string TestSuperTRUMPAllDealService(string myValue);

        [OperationContract]
        [TransactionFlow(TransactionFlowOption.Allowed)]
        string LoadDWData();

        [OperationContract]
        [TransactionFlow(TransactionFlowOption.Allowed)]
        string GenerateInputXML();
      
        [OperationContract]
        [TransactionFlow(TransactionFlowOption.Allowed)]
        string ExecuteServiceFlow();

        //[OperationContract]
        //[TransactionFlow(TransactionFlowOption.Allowed)]
        //ServiceEntity GetServiceCredentials();
    }

   

}
