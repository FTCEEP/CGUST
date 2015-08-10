using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;

using NLog;

namespace Fpg.Cim.Adapter.ModbusTCP.Program
{
    static class Program
    {
        private static Logger gLogger = NLog.LogManager.GetCurrentClassLogger();
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
            { 
                new Service1.Service1() 
            };
            ServiceBase.Run(ServicesToRun);

            //EquipmentState.EquipmentState objEquipmentState = new EquipmentState.EquipmentState();
            //objEquipmentState.Start();

            //UpdateServiceTime.UpdateServiceTime objUpdateServiceTime = new UpdateServiceTime.UpdateServiceTime();
            //objUpdateServiceTime.Update_Time();

            //gLogger.Trace(string.Format("程式開始 {0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")));
            //do
            //{

            //    gLogger.Trace(string.Format("程式開始 {0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")));
           //     ModbusTCP.ModbusTCP objModbusTCP = new ModbusTCP.ModbusTCP();
            //    objModbusTCP.Start();
            //    gLogger.Trace(string.Format("程式結束 {0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")));
            //    gLogger.Trace("");
            //    System.Threading.Thread.Sleep(1000 * 30);
           //   ACChange.ACChange objACChange = new ACChange.ACChange();
           // objACChange.Start();
            //    //i++;
            //    //System.Threading.Thread.Sleep(1000 * 60 * 10);

            //} while (true);
            //gLogger.Trace(string.Format("程式結束 {0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")));
            //gLogger.Trace("");
        }
    }
}
