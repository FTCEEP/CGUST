using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.Data;
using NLog;

namespace Fpg.Cim.Adapter.ModbusTCP.Equipment
{
    class Equipment
    {
        private static Logger gLogger = NLog.LogManager.GetCurrentClassLogger();
        private static int giDBCallTimes = 3;
        public class ModbusTCP
        {
            private EEP_Client_WS.EEP_Client_WS wsEEP_Client_WS;
            AdapterLogs.AdapterLogs.MS_SQL objMS_SQL = null;
            Common.Common objCommon = null;
            Socket objSocket = null;
            private string gsCompanyID = string.Empty;
            private string gsEquipmentIDs = string.Empty;
            private string gsIP = string.Empty;
            private int giPort = 0;
            private int giEquipmentStateTime = 50;
            public ModbusTCP(string sCompanyID, string sEquipmentIDs, int iDBCallTimes, int iEquipmentStateTime)
            {
                gsCompanyID = sCompanyID;
                gsEquipmentIDs = sEquipmentIDs;
                giDBCallTimes = iDBCallTimes;
                giEquipmentStateTime = iEquipmentStateTime;
                wsEEP_Client_WS = new EEP_Client_WS.EEP_Client_WS();
                objMS_SQL = new AdapterLogs.AdapterLogs.MS_SQL();
                objCommon = new Common.Common();
            }

            public void Start()
            {
                do
                {
                    DataSet dtsEquipments = new DataSet("Equipments");
                    try
                    {
                        dtsEquipments = objCommon.ByteToDataset(wsEEP_Client_WS.Get_Equipments(gsCompanyID, Assemble()));
                        if (dtsEquipments.Tables[0].Rows.Count > 0)
                        {
                            foreach (DataRow dtwEquipments in dtsEquipments.Tables[0].Rows)
                            {
                                StateCheck(dtwEquipments["CompanyID"].ToString()
                                    , dtwEquipments["EquipmentID"].ToString()
                                    , dtwEquipments["IP"].ToString()
                                    , int.Parse(dtwEquipments["Port"].ToString()));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        gLogger.ErrorException("Equipment.ModbusTCP.Start", ex);
                        objMS_SQL.Write("Equipment.ModbusTCP.Start", ex.Message, DateTime.Now);
                    }
                    System.Threading.Thread.Sleep(giEquipmentStateTime * 60 * 1000);
                } while (true);
            }

            private void StateCheck(string sCompanyID, string sEquipmentID, string sIP, int iPort)
            {
                try
                {
                    objSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                    objSocket.Connect(gsIP, giPort);
                    if (objSocket.Connected)
                        for (int i = 0; i < giDBCallTimes; i++)
                        {
                            bool blFlag = wsEEP_Client_WS.Update_EquipmentsState(sCompanyID, sEquipmentID
                                , sIP, iPort, DateTime.Now);
                            if (blFlag)
                                break;
                        }
                }
                catch (Exception ex)
                {
                    gLogger.ErrorException("Equipment.ModbusTCP.StateCheck", ex);
                    objMS_SQL.Write("Equipment.ModbusTCP.StateCheck", ex.Message, DateTime.Now);
                    throw ex;
                }
                finally
                {
                    if (objSocket.Connected)
                        objSocket.Close(1);
                }
            }

            private string Assemble()
            {
                string[] sEquipmentIDs = null;
                string sNewEquipmentIDs = string.Empty;
                try
                {
                    if (string.IsNullOrEmpty(gsEquipmentIDs))
                        return null;
                    sEquipmentIDs = gsEquipmentIDs.Split(',', ';');
                    foreach (string sEquipmentID in sEquipmentIDs)
                    {
                        sNewEquipmentIDs += string.Format("'{0}',", sEquipmentID);
                    }
                    sNewEquipmentIDs = sNewEquipmentIDs.TrimEnd(',');
                }
                catch (Exception ex)
                {
                    gLogger.ErrorException("Equipment.ModbusTCP.Assemble", ex);
                    throw ex;
                }
                return sNewEquipmentIDs;
            }
        }
    }
}
