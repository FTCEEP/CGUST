using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NLog;

namespace Fpg.Cim.Adapter.ModbusTCP.AdapterLogs
{
    class AdapterLogs
    {
        private static Logger gLogger = NLog.LogManager.GetCurrentClassLogger();

        public class MS_SQL
        {
            private int giLogWriteTimes = int.Parse(System.Configuration.ConfigurationSettings.AppSettings.Get("LogWriteTimes"));
            public MS_SQL()
            {
                if (giLogWriteTimes == 0)
                    giLogWriteTimes = 3;
            }
            public void Write(string sName, string sContents, DateTime dtTime)
            {
                EEP_Client_WS.EEP_Client_WS wsEEP_Client_WS = new EEP_Client_WS.EEP_Client_WS();
                bool blLog = false;
                try
                {
                    for (int i = 1; i <= giLogWriteTimes; i++)
                    {
                        blLog = wsEEP_Client_WS.WrtieAdapterLogs(sName, sContents, dtTime);
                        if (blLog)
                            break;
                    }

                }
                catch (Exception ex)
                {
                    gLogger.ErrorException("AdapterLogs.Write", ex);
                }
            }
        }
    }
}
