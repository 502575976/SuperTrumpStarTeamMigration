using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OracleClient;
using System.Data;
using Microsoft.Win32;



namespace BSVICROIDBADAPTOR
{
   public class SetLogConfiguration
    {
        public log4net.ILog SetLog4Net()
        {
            try
            {
                RegistryKey regKey = default(RegistryKey);
                //regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\FacilitySettings\\SuperTRUMPForAllDeal", true);
                if (log4net.LogManager.GetRepository().Configured  == false)
                {
                    
                    //log4net.Config.XmlConfigurator.ConfigureAndWatch(new System.IO.FileInfo(regKey.GetValue("SuperTRUMPForAllDealLogFile").ToString()));

                    log4net.Config.XmlConfigurator.ConfigureAndWatch(new System.IO.FileInfo("E:\\internalsites\\cef\\SuperTrumpAllDeal\\ConfigXml\\log4net_DLL.config"));


                }
                log4net.ILog STLogger = log4net.LogManager.GetLogger("SuperTRUMPForAllDeal");
                return STLogger;
            }
            catch (Exception ex)
            {
                throw;
                return null;
            }
        }

    }
}
