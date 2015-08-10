using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.NetworkInformation;
using NLog;
using System.Configuration;
using System.Data;
using System.Net.Sockets;

namespace Fpg.Cim.Adapter.ModbusTCP.EquipmentState
{
    class EquipmentState
    {
        private static Logger gLogger = NLog.LogManager.GetCurrentClassLogger();
        private static int giEquipmenStateCheck = Int16.Parse(ConfigurationManager.AppSettings.Get("EquipmenStateCheck"));
        private string gsEquipmentsFilePath = ConfigurationManager.AppSettings.Get("EquipmentsFilePath");
        private string gsCompanyID = ConfigurationManager.AppSettings.Get("CompanyID");

        public void Start()
        {
            do
            {
                DataSet dtsEquipments = new DataSet("Equipments");
                EEP_Client_WS.EEP_Client_WS wsEEP_Client_WS = new EEP_Client_WS.EEP_Client_WS();
                bool blFlag = false;
                try
                {
                    dtsEquipments.ReadXml(@gsEquipmentsFilePath);
                    DataTable dtDistinct = dtsEquipments.Tables[0].DefaultView.ToTable(true, new string[] { "CompanyID", "EquipmentID", "IP", "Port" });
                    foreach (DataRow dtwEquipment in dtDistinct.Rows)
                    {
                        //blFlag = PingIt(dtwEquipment["IP"].ToString());
                        blFlag = PingHost(dtwEquipment["IP"].ToString(), Int16.Parse(dtwEquipment["Port"].ToString()));
                        if (!blFlag)
                            wsEEP_Client_WS.Set_AlarmOccur(gsCompanyID
                                , dtwEquipment["EquipmentID"].ToString()
                                , dtwEquipment["EquipmentID"].ToString() + "(" + dtwEquipment["IP"].ToString() + ":"+ dtwEquipment["Port"].ToString() + ")" + ".State"
                                , dtwEquipment["EquipmentID"].ToString() + "(" + dtwEquipment["IP"].ToString() + ":"+ dtwEquipment["Port"].ToString() + ")" + ".State", DateTime.Now);
                    }
                }
                catch (Exception ex)
                {
                    gLogger.ErrorException("EquipmentState.Start", ex);
                }
                finally
                {
                    if (wsEEP_Client_WS != null) { wsEEP_Client_WS.Dispose(); wsEEP_Client_WS = null; }
                    if (dtsEquipments != null) { dtsEquipments.Dispose(); dtsEquipments = null; }
                }

                System.Threading.Thread.Sleep(1000 * giEquipmenStateCheck);
            } while (true);
        }

        private bool PingIt(string sIP)
        {
            bool blFlag = false;
            try
            {
                Ping p = new Ping();
                var ip = sIP;
                PingReply r = p.Send(ip);
                if (r.Status == IPStatus.Success)
                {
                    gLogger.Trace(string.Format("IP:{0} pint test ok！", ip));
                    blFlag = true;
                }
                else
                {
                    gLogger.Trace(string.Format("IP:{0} pint test failed！", ip));
                    blFlag = false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return blFlag;
        }

        public static bool PingHost(string _HostURI, int _PortNumber)
        {
            try
            {
                TcpClient client = new TcpClient(_HostURI, _PortNumber);
                client.Close();
                return true;
            }
            catch (Exception ex)
            {
                gLogger.Error("Error pinging host:'" + _HostURI + ":" + _PortNumber.ToString() + "' " + ex.Message);
                return false;
            }
        }
    }
}
