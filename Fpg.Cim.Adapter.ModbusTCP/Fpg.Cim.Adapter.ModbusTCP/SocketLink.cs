using System;
using System.Collections.Generic;
using System.Text;

using NLog;
using System.Net.Sockets;
using System.Configuration;

namespace Fpg.Cim.Adapter.ModbusTCP.SocketLink
{
    class SocketLink
    {
        private static Logger gLogger = NLog.LogManager.GetCurrentClassLogger();
        private int giReceiveTimeout = Convert.ToInt32(ConfigurationManager.AppSettings["ReceiveTimeout"].ToString()); //從AppConfig讀取「接收DAE的TimeOut」參數
        private int giSendTimeout = Convert.ToInt32(ConfigurationManager.AppSettings["ReceiveTimeout"].ToString()); //從AppConfig讀取「接收DAE的TimeOut」參數
        private string gsProgrammerMail = ConfigurationManager.AppSettings.Get("ProgrammerMail"); //從AppConfig讀取「ProgrammerMail」參數

        /// <summary>
        /// 連線或斷線至ModBusToEthernet設備
        /// iStatus = 0，則連線
        /// iStatus = 1，則斷線
        /// </summary>
        /// <param name="iStatus">0為連線、1為斷線</param>
        /// <param name="sIP">設備的IP位址</param>
        /// <param name="sIPPort">設備的IP_Port</param>
        public bool LinkEQP(int iStatus, string sIP, string sIPPort, Socket gWorkSocket)
        {
            bool blFlag = false;
            try
            {
                switch (iStatus)
                {
                    case 0:
                        // Set the timeout for synchronous receive methods to 
                        // 1 second (1000 milliseconds.)
                        gWorkSocket.ReceiveTimeout = giReceiveTimeout * 1000;

                        // Set the timeout for synchronous send methods
                        // to 1 second (1000 milliseconds.) 
                        gWorkSocket.SendTimeout = giSendTimeout * 1000;

                        if (gWorkSocket.Connected == true)
                        {
                            gWorkSocket.Shutdown(SocketShutdown.Both);
                            gWorkSocket.Close();
                        }
                        else if (gWorkSocket.Connected == false)
                        {
                            gWorkSocket.Connect(sIP, Convert.ToInt32(sIPPort));
                            if ((gWorkSocket.Connected == true))
                            {
                                //System.Threading.Thread.Sleep(150);
                                gLogger.Trace("ModbusRTU--成功連結 to：" + sIP);
                            }                         
                        }
                        blFlag = true;
                        break;
                    case 1:
                        if ((gWorkSocket.Connected == true))
                        {
                            gWorkSocket.Shutdown(SocketShutdown.Both);
                            gWorkSocket.Close();
                            //System.Threading.Thread.Sleep(150);
                            gLogger.Trace("ModbusRTU--取消連線 to：" + sIP);
                        }
                        blFlag = false;
                        break;
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("ModbusTCP.LinkEQP：", ex);
                if (gWorkSocket.Connected)
                {
                    gWorkSocket.Shutdown(SocketShutdown.Both);
                    gWorkSocket.Close();
                }
                throw ex;
            }
            return blFlag;
        }
    }
}
