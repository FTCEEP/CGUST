using System;
using System.Collections.Generic;
using System.Text;

using NLog;
using System.Configuration;
using System.Data;
using System.Net.Sockets;
using System.Threading;
using System.Data.SqlClient;

namespace Fpg.Cim.Adapter.ModbusTCP.ACChange
{
    class ACChange
    {
        private static Logger g_oLogger = NLog.LogManager.GetCurrentClassLogger();
        private string g_sProjectID = ConfigurationManager.AppSettings["ProjectID"].ToString(); //從AppConfig讀取「ProjectID」參數
        private string g_sCompanyID = ConfigurationManager.AppSettings["CompanyID"].ToString(); //從AppConfig讀取「CompanyID」參數
        private string g_sIP = ConfigurationManager.AppSettings.Get("IP"); //從AppConfig讀取「IP」參數
        private string g_sPort = ConfigurationManager.AppSettings.Get("Port"); //從AppConfig讀取「Port」參數
        private string g_sProgrammerMail = ConfigurationManager.AppSettings.Get("ProgrammerMail"); //從AppConfig讀取「ProgrammerMail」參數
        private int g_iReTryTimes = Convert.ToInt16(ConfigurationManager.AppSettings.Get("ReTryTimes")); //從AppConfig讀取「ReTryTimes」參數
        private int g_iCheckMeterParaTime = Convert.ToInt16(ConfigurationManager.AppSettings.Get("CheckMeterParaTime")); //從AppConfig讀取「CheckMeterParaTime」參數
        private string g_sLocalModbusParaPath = ConfigurationManager.AppSettings.Get("LocalModbusParaPath"); //從AppConfig讀取「LocalModbusParaPath」參數
        private string g_sSplitSymbol = ConfigurationManager.AppSettings.Get("SplitSymbol"); //從AppConfig讀取「SplitSymbol」參數
        private string g_sEquipmentsFilePath = ConfigurationManager.AppSettings.Get("EquipmentsFilePath");
        private string g_sPointsFilePath = ConfigurationManager.AppSettings.Get("PointsFilePath");
        private string g_sEquipmentType = ConfigurationManager.AppSettings.Get("EquipmentType");
        private int g_iRunCount = 0;
        private string sConnString = ConfigurationManager.ConnectionStrings["EEP_CGUST"].ConnectionString;

        SocketLink.SocketLink objSocketLink = null;
        Socket gWorkSocket = null;
        //104.2.4新增ping設備ip
        System.Net.NetworkInformation.Ping ping = new System.Net.NetworkInformation.Ping();
        DBConnection Utility = new DBConnection();

        public void Start()
        {
            EEP_Client_WS.EEP_Client_WS wsEEP_Client_WS = new EEP_Client_WS.EEP_Client_WS();
            DataTable dttIPPort = new DataTable(); //存放ModbusTCP之設備IP與Port
            DataTable dttACList = new DataTable();//存放空調狀況改變排程作業
            DataTable dttACListPara = new DataTable();
            DataTable dttExist = null;
            DataRow[] dtwPointsDatafilter = null;
            DataSet dtsPointsVlue = null; //存放讀取設備的讀值
            DataEncryption.DataEncryption objDataEncryption = new DataEncryption.DataEncryption();
            //DataSet dtsEquipments = new DataSet("Equipments");
            //DataSet dtsPoints = new DataSet("Points");
            DataSet dtsPointsData = new DataSet("PointsData");
            int iCount = 0;////
            int iInsertCount = 0;
            int retAffectedRows = 0;
            string sSQLCmd = "";
            string sSQLCmd2 = "";
            string sTableName = "";
            string sTableName2 = "";
            try
            {
                try
                {
                    //SqlConnection conn = new SqlConnection(sConnString);
                    //conn.Open();
                    //dtsEquipments.ReadXml(@gsEquipmentsFilePath);
                    //dtsPoints.ReadXml(@gsPointsFilePath);
                    sTableName = "ACCommandList";
                    sSQLCmd = string.Format("SELECT * FROM {0} Where convert(varchar (10), RecTime,120 ) = convert(varchar (10), getdate(),120  ) and Result IS NULL ", sTableName);
                    sSQLCmd2 = string.Format("IF EXISTS(SELECT name FROM sys.objects WHERE name = '{0}') SELECT 'True'", sTableName);
                    DataTable dataTable2 = Utility.execDataTable(sSQLCmd2, sConnString);
                    if (dataTable2.Rows.Count > 0)
                    {
                        dttACList = Utility.execDataTable(sSQLCmd, sConnString);
                    }

                    //SqlCommand cmd = new SqlCommand("Select * From  ACStatusPara", conn);



                    //dttACPara.Load(cmd.ExecuteReader());
                    //cmd.Dispose();
                    //conn.Close();
                    //conn.Dispose();


                    //dtsPointsData = wsEEP_Client_WS.Get_PointsData_ModbusTCP(g_sCompanyID, "", "");
                }
                catch (Exception ex)
                {
                    g_oLogger.ErrorException("讀取異常-- ", ex);
                }

                //if (dtsEquipments.Tables[0].Rows.Count == 0)
                if (dttACList.Rows.Count == 0)
                {
                    g_oLogger.Info("---------------ModbusTCP之設備檔異常，請檢查(DB)檔案---------------");
                }
                //if (dtsEquipments.Tables[0].Rows.Count > 0)
                if (dttACList.Rows.Count > 0)
                {
                    //dttIPPort = dtsEquipments.Tables[0].DefaultView.ToTable(true, new string[] { "IP", "Port" });
                    //dttIPPort = dttACPara.DefaultView.ToTable(true, new string[] { "IP", "Port" });
                    
                    
                    sTableName = "ACCommandList";
                    sTableName2 = "ACStatusPara";
                    sSQLCmd = string.Format("SELECT AL.[RecTime] ,AL.[PointID] ,AL.[SetMode] ,AL.[Result] ,AL.[FinishTime] ,AP.[IP] ,AP.[DeviceID] FROM {0} as AL, {1} as AP Where convert(varchar (10), AL.RecTime,120 ) = convert(varchar (10), getdate(),120 ) And AL.PointID = AP.PointID", sTableName, sTableName2);
                    sSQLCmd2 = string.Format("IF EXISTS(SELECT name FROM sys.objects WHERE name = '{0}') SELECT 'True'", sTableName2);
                    DataTable dataTable3 = Utility.execDataTable(sSQLCmd2, sConnString);
                    if (dataTable3.Rows.Count > 0)
                    {
                        dttACListPara = Utility.execDataTable(sSQLCmd, sConnString);
                    }
                    dttIPPort = dttACListPara.DefaultView.ToTable(true, new string[] { "IP" });
                     

                    foreach (DataRow dtwIPPort in dttIPPort.Rows)
                    {
                        DataTable dttPointsValue = new DataTable("ACRealTimeData");
                        //dttPointsValue.Columns.Add("PointID", typeof(string));
                        //dttPointsValue.Columns.Add("Status", typeof(string));
                        //dttPointsValue.Columns.Add("RecTime", typeof(DateTime));
                        string sIP = dtwIPPort["IP"].ToString();
                        //string sPort = dtwIPPort["Port"].ToString();
                        string sPort = "502";
                        objSocketLink = new SocketLink.SocketLink();
                        gWorkSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                        //104.2.4 ping不到設備就往下一個ip設備詢問
                        g_oLogger.Debug(sIP + "進行連線------------------------------------------------------");
                        if (ping.Send(sIP, 1000).Status != System.Net.NetworkInformation.IPStatus.Success)
                        {
                            g_oLogger.Debug(sIP + "的設備連線TimeOut-----------------------------------------");
                            continue;
                        }
                        objSocketLink.LinkEQP(0, sIP, sPort, gWorkSocket);
                        //dtwPointsDatafilter = dtsEquipments.Tables[0].Select("IP = '" + sIP + "' and Port = '" + sPort + "'"); //dtsEquipments.Tables[0].Select("IP = '" + sIP + "' and Port = '" + sPort + "'", "Address DESC");
                        dtwPointsDatafilter = dttACListPara.Select("IP = '" + sIP + "'");
                        foreach (DataRow dtw in dtwPointsDatafilter)
                        {
                            try
                            {
                                if (gWorkSocket.Connected) //0:表連線，1:表斷線
                                {
                                    string sCommand = "";
                                    byte[] byReceived = null;
                                    bool bResult = false;

                                    switch (dtw["SetMode"].ToString())
                                    {
                                        //卡機模式
                                        case "Auto":
                                        sCommand = string.Format("0,0,0,0,0,6,{0},5,0,10,0,0", Convert.ToDouble(dtw["DeviceID"]));
                                        byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                        if (byReceived != null)
                                        { g_oLogger.Debug("Auto模式設定：第一階段設定成功"); }
                                        else
                                        { g_oLogger.Debug("Auto模式設定：第一階段設定失敗"); }

                                        sCommand = string.Format("0,0,0,0,0,6,{0},5,0,2,0,0", Convert.ToDouble(dtw["DeviceID"]));
                                        byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                        if (byReceived != null)
                                        { g_oLogger.Debug("Auto模式設定：第二階段設定成功"); }
                                        else
                                        { g_oLogger.Debug("Auto模式設定：第二階段設定失敗"); }

                                        sCommand = string.Format("0,0,0,0,0,6,{0},5,0,10,255,0", Convert.ToDouble(dtw["DeviceID"]));
                                        byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                        if (byReceived != null)
                                        { g_oLogger.Debug("Auto模式設定：第三階段設定成功"); }
                                        else
                                        { g_oLogger.Debug("Auto模式設定：第三階段設定失敗"); }

                                        //驗證是否設定成功
                                        sCommand = string.Format("0,0,0,0,0,6,{0},3,0,3,0,1", Convert.ToDouble(dtw["DeviceID"]));
                                        byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                        if (byReceived != null)
                                        {
                                            switch (byReceived[10].ToString("00"))
                                            {
                                                //卡機模式
                                                case "00":
                                                    g_oLogger.Debug("模式設定ok：現在是Auto模式");
                                                    bResult = true;
                                                    break;
                                                //強開模式
                                                case "161":
                                                    g_oLogger.Debug("模式設定Fail：現在是FOn模式");
                                                    break;
                                                //強關模式
                                                case "128":
                                                    g_oLogger.Debug("模式設定Fail：現在是FOff模式");
                                                    break;
                                            }
                                        }
                                        break;

                                        //強開模式
                                        case "FOn":
                                        sCommand = string.Format("0,0,0,0,0,6,{0},5,0,10,0,0", Convert.ToDouble(dtw["DeviceID"]));
                                        byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                        if (byReceived != null)
                                        { g_oLogger.Debug("FOn模式設定：第一階段設定成功"); }
                                        else
                                        { g_oLogger.Debug("FOn模式設定：第一階段設定失敗"); }

                                        sCommand = string.Format("0,0,0,0,0,6,{0},5,0,2,255,0", Convert.ToDouble(dtw["DeviceID"]));
                                        byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                        if (byReceived != null)
                                        { g_oLogger.Debug("FOn模式設定：第二階段設定成功"); }
                                        else
                                        { g_oLogger.Debug("FOn模式設定：第二階段設定失敗"); }


                                        //驗證是否設定成功
                                        sCommand = string.Format("0,0,0,0,0,6,{0},3,0,3,0,1", Convert.ToDouble(dtw["DeviceID"]));
                                        byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                        if (byReceived != null)
                                        {
                                            switch (byReceived[10].ToString("00"))
                                            {
                                                //卡機模式
                                                case "00":
                                                    g_oLogger.Debug("模式設定Fail：現在是Auto模式");
                                                    break;
                                                //強開模式
                                                case "161":
                                                    g_oLogger.Debug("模式設定ok：現在是FOn模式");
                                                    bResult = true;
                                                    break;
                                                //強關模式
                                                case "128":
                                                    g_oLogger.Debug("模式設定Fail：現在是FOff模式");
                                                    break;
                                            }
                                        }

                                            break;

                                        //強關模式
                                        case "FOff":
                                        sCommand = string.Format("0,0,0,0,0,6,{0},5,0,10,0,0", Convert.ToDouble(dtw["DeviceID"]));
                                        byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                        if (byReceived != null)
                                        { g_oLogger.Debug("FOff模式設定：第一階段設定成功"); }
                                        else
                                        { g_oLogger.Debug("FOff模式設定：第一階段設定失敗"); }

                                        sCommand = string.Format("0,0,0,0,0,6,{0},5,0,2,0,0", Convert.ToDouble(dtw["DeviceID"]));
                                        byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                        if (byReceived != null)
                                        { g_oLogger.Debug("FOff模式設定：第二階段設定成功"); }
                                        else
                                        { g_oLogger.Debug("FOff模式設定：第二階段設定失敗"); }


                                        //驗證是否設定成功
                                        sCommand = string.Format("0,0,0,0,0,6,{0},3,0,3,0,1", Convert.ToDouble(dtw["DeviceID"]));
                                        byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                        if (byReceived != null)
                                        {
                                            switch (byReceived[10].ToString("00"))
                                            {
                                                //卡機模式
                                                case "00":
                                                    g_oLogger.Debug("模式設定Fail：現在是Auto模式");
                                                    break;
                                                //強開模式
                                                case "161":
                                                    g_oLogger.Debug("模式設定Fail：現在是FOn模式");
                                                    break;
                                                //強關模式
                                                case "128":
                                                    g_oLogger.Debug("模式設定ok：現在是FOff模式");
                                                    bResult = true;
                                                    break;
                                            }

                                        }


                                            break;


                                        default:
                                            bResult = false;
                                            g_oLogger.Debug("模式設定參數錯誤");
                                            break;
                                    
                                    }                                    
                           
                                    //byReceived = ModbusTCP_Operate(sCommand, gWorkSocket, sIP, 1);
                                    //if (byReceived == null)
                                    //    continue;

                                    iCount++;
                                    if (iCount % 100 == 0)
                                    {
                                        objSocketLink.LinkEQP(1, sIP, sPort, gWorkSocket);
                                        if (objSocketLink != null) { objSocketLink = null; }

                                        objSocketLink = new SocketLink.SocketLink();
                                        gWorkSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                                        objSocketLink.LinkEQP(0, sIP, sPort, gWorkSocket);
                                    }


                                    sTableName = "ACCommandList";
                                    sSQLCmd = string.Format("Update {0} Set Result = '{1}' ,  FinishTime ='{2}' Where RecTime ='{3}' and PointID ='{4}'", sTableName, bResult.ToString(), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), Convert.ToDateTime(dtw["RecTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff"), dtw["PointID"]);
                                    sSQLCmd2 = string.Format("IF EXISTS(SELECT name FROM sys.objects WHERE name = '{0}') SELECT 'True'", sTableName);
                                    DataTable dataTable4 = Utility.execDataTable(sSQLCmd2, sConnString);
                                    if (dataTable3.Rows.Count > 0)
                                    {

                                        iInsertCount += Utility.execNonQuery(sSQLCmd, sConnString);
                                        g_oLogger.Debug("更新任務狀態：第" + iInsertCount.ToString() +"次");
                                    }











                                    //SolveValue(ref dttPointsValue, int.Parse(dtw["Address"].ToString()), byReceived, dtsPoints.Tables[0], sIP, int.Parse(sPort), g_sEquipmentType);
                                    //SolveValue(ref dttPointsValue, int.Parse(dtw["Address"].ToString()), byReceived, dtsPointsData.Tables[0], sIP, int.Parse(sPort), g_sEquipmentType, dtw["PointID"].ToString(), dtw["EquipmentID"].ToString());
                                    //新增dtw參數
                                    //SolveValue(ref dttPointsValue, int.Parse(dtw["Address"].ToString()), byReceived, dtsPointsData.Tables[0], sIP, int.Parse(sPort), g_sEquipmentType, dtw["PointID"].ToString(), dtw["EquipmentID"].ToString(),dtw);
                                    //SolveValueAC(ref dttPointsValue, byReceived, dtw["PointID"].ToString());

                                    //for (int i = 0; i < dttPointsValue.Rows.Count; i++)
                                    //{
                                    //    gLogger.Trace(string.Format("{0}, {1}, {2}, {3}"
                                    //        , dttPointsValue.Rows[i][0].ToString()
                                    //        , dttPointsValue.Rows[i][1].ToString()
                                    //        , dttPointsValue.Rows[i][2].ToString()
                                    //        , ((DateTime)dttPointsValue.Rows[i][3]).ToString("yyyy-MM-dd HH:mm:ss.fff")));
                                    //}
                                }
                            }
                            catch (Exception ex)
                            {
                                g_oLogger.ErrorException(dtw["PointID"].ToString() + " 讀取異常", ex);
                                gWorkSocket.Shutdown(SocketShutdown.Both);
                                System.Threading.Thread.Sleep(500);
                                gWorkSocket.Disconnect(false);
                                gWorkSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                                objSocketLink.LinkEQP(0, sIP, sPort, gWorkSocket);
                            }
                        }

                        objSocketLink.LinkEQP(1, sIP, sPort, gWorkSocket);

                        dtsPointsVlue = new DataSet("PointsVlue");
                        dtsPointsVlue.Tables.Add(dttPointsValue);
                        //dtsPointsVlue.WriteXml(@"C:\EEP System\HeatPump\Adapter\Data.xml");
                        //if (!wsEEP_Client_WS.Set_PointValueToRealTimeTable(objDataEncryption.DataSetToByte(dtsPointsVlue), g_sProjectID))
                        //{
                        //    CSV.CSV objCSV = new CSV.CSV();
                        //    objCSV.WriteToCSV(dttPointsValue, Convert.ToDateTime(dttPointsValue.Rows[0]["RecTime"]).ToString("yyyyMMdd"));
                        //    if (objCSV != null) { objCSV = null; }
                        //}
                       



                        g_iRunCount++;
                        g_oLogger.Debug("Count:" + g_iRunCount.ToString());
                        g_oLogger.Debug(DateTime.Now + ": Insert RealTime Table");

                        //Service alive 狀態紀錄
                        string sDefTableName = "ACServiceAlive";
                        string sCmdExistSQL = string.Format("IF EXISTS(SELECT name FROM sys.objects WHERE name = '{0}') SELECT 'True'", sDefTableName);
                        dttExist = Utility.execDataTable(sCmdExistSQL, sConnString);
                        if (dttExist.Rows.Count == 0)
                        {
                            string cmdSQL = string.Format("CREATE TABLE [dbo].[{0}]([ServiceName] " +
                            "[nvarchar](20) NOT NULL, " +
                            "[RecTime] [datetime] NULL, CONSTRAINT " +
                            "[PK_ACServiceAlive] PRIMARY KEY CLUSTERED ([ServiceName] ASC " +
                            ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, " +
                            "ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
                            , sDefTableName);
                            int i = Utility.execNonQuery(cmdSQL, sConnString);

                            g_oLogger.Trace("產生新的ACServiceAlive的Table");
                        }

                        try
                        {
                            string cmdSQL = "";
                            cmdSQL = string.Format("DELETE {0} WHERE ServiceName = '{1}'", sDefTableName, "ACChange");
                            Utility.execNonQuery(cmdSQL, sConnString);

                            cmdSQL = "";
                            cmdSQL = string.Format("INSERT INTO {0} (ServiceName , RecTime) " +
                                "VALUES ('{1}', '{2}');"
                                , sDefTableName
                                , "ACChange"
                                , DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));

                            Utility.execNonQuery(cmdSQL, sConnString);

                        }
                        catch (Exception ex)
                        {
                            g_oLogger.ErrorException("insert ACServiceAliveTable異常：", ex);
                        }





                    }
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusRTU.讀取設備參數階段", ex);
            }
            finally
            {
                if (gWorkSocket != null && gWorkSocket.Connected) { gWorkSocket.Shutdown(SocketShutdown.Both); gWorkSocket.Close(1); ; gWorkSocket = null; }
                if (objSocketLink != null) { objSocketLink = null; }
                if (dtsPointsData != null) { dtsPointsData.Clear(); dtsPointsData.Dispose(); dtsPointsData = null; }
                if (dtsPointsVlue != null) { dtsPointsVlue.Clear(); dtsPointsVlue.Dispose(); dtsPointsVlue = null; }
                if (wsEEP_Client_WS != null) { wsEEP_Client_WS.Dispose(); wsEEP_Client_WS = null; }
                if (Utility != null) { Utility = null; }
            }
        }



        public byte[] ModbusTCP_Operate(string sCommand, Socket gWorkSocket, string sIP, int iLength)
        {
            string[] sSend = null;
            byte[] SendBuffer = null;
            string sReceived = string.Empty;
            bool blRead = false;
            byte[] ReceiveBuffer = null;
            try
            {
                sSend = sCommand.Split(',');
                SendBuffer = new byte[sSend.Length];

                for (int i = 0; (i <= (sSend.Length - 1)); i++)
                {
                    SendBuffer[i] = Convert.ToByte(sSend[i]);
                }
                int iReceiveCount = 1;

                do
                {
                    sReceived = "";
                    this.Send(gWorkSocket, SendBuffer, 0, SendBuffer.Length, 10000);

                    ReceiveBuffer = new byte[2048];
                    int iReceiveWord = 0;
                    if (SendBuffer[7] == 1 || SendBuffer[7] == 2)
                    {
                        iReceiveWord = SendBuffer[11];
                    }
                    else if (SendBuffer[7] == 3 || SendBuffer[7] == 4)
                    {
                        iReceiveWord = SendBuffer[11] * 2;
                    }



                    this.Receive(gWorkSocket, ref ReceiveBuffer, 0, iReceiveWord + 9, 10000);
                    for (int i = 0; i < ReceiveBuffer.Length; i++)
                    {
                        sReceived += ReceiveBuffer[i].ToString();
                        sReceived += ',';
                    }

                    g_oLogger.Trace("ModbusTCP.Modbus 第" + iReceiveCount.ToString() + "次接收到的字串：" + sReceived.TrimEnd(','));
                    if (SendBuffer[7] == 1 || SendBuffer[7] == 2)
                    {
                        iLength = 1;
                    }
                    else if (SendBuffer[7] == 3 || SendBuffer[7] == 4)
                    {
                        iLength = iLength * 2;
                    }
                    if (ReceiveBuffer[8] == iLength)
                    {
                        blRead = true;
                    }
                    else
                    {
                        iReceiveCount = iReceiveCount + 1;
                        if (iReceiveCount > g_iReTryTimes)
                        {
                            blRead = true;
                            g_oLogger.Error(string.Format("ModbusTCP命令連續讀取{0}次錯誤，請確認ModbusTCP指令{1}，IP:{2}", iReceiveCount, sCommand, sIP));
                            return null;
                        }
                    }
                } while (!blRead);
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.Read", ex);
            }
            return ReceiveBuffer;
        }

        private void RecordingLogs(DataRow dtwPointValue)
        {
            try
            {
                for (int i = 0; i < dtwPointValue.ItemArray.Length; i++)
                {
                    switch (i)
                    {
                        case 0:
                            g_oLogger.Info("EquipmentID- " + dtwPointValue[i].ToString());
                            break;
                        case 1:
                            g_oLogger.Info("PointID- " + dtwPointValue[i].ToString());
                            break;
                        case 2:
                            g_oLogger.Info("PointValue- " + dtwPointValue[i].ToString());
                            break;
                        case 3:
                            g_oLogger.Info("RecTime- " + Convert.ToDateTime(dtwPointValue[i]).ToString("yyyy-MM-dd HH:mm:ss"));
                            break;
                        default:
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.RecordingLogs", ex);
            }
        }

        /// <summary>
        /// 解析長庚冷氣電錶狀態
        /// </summary>
        /// <returns></returns>
        private void SolveValueAC(ref DataTable dttPointsValue, byte[] bReceive, string sPointID)
        {
            try
            {
                string sRecevice = bReceive[10].ToString("00");
                string sStatus = "";
                DataRow dtwPointsValue = null;
                g_oLogger.Trace("狀態：" + sRecevice);
                //回來狀態是10進位
                switch (sRecevice)
                {
                    //卡機模式
                    case "00":
                        sStatus = "Auto";
                        break;
                    //強開模式
                    case "161":
                        sStatus = "FOn";
                        break;
                    //強關模式
                    case "128":
                        sStatus = "FOff";
                        break;

                    default:
                        g_oLogger.Info(string.Format("{0}設備解析失敗，請確認!", sPointID));
                        break;
                }

                dtwPointsValue = dttPointsValue.NewRow();
                dtwPointsValue["PointID"] = sPointID;
                dtwPointsValue["Status"] = sStatus;
                dtwPointsValue["RecTime"] = DateTime.Now;
                dttPointsValue.Rows.Add(dtwPointsValue);

            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.SolveValueAC", ex);
            }
        }






        /// <summary>
        /// 解析台隆控制器
        /// </summary>
        /// <returns></returns>
        private void SolveValue(ref DataTable dttPointsValue, int iAddress, byte[] bReceive, DataTable dttPoints, string sIP, int iPort, string sEquipmentType, string sPointID, string sEquipmentID, DataRow dtw)
        {
            try
            {
                switch (sEquipmentType)
                {
                    case "永宏":
                        switch (bReceive[7])
                        {
                            case 1:
                                switch (10)
                                //switch (iAddress)
                                {
                                    case 10:
                                    case 11:
                                    case 12:
                                    case 13:
                                    case 14:
                                    case 28:
                                    case 29:
                                    case 36:
                                    case 37:
                                    case 38:
                                    case 40:
                                    case 41:
                                    case 42:
                                    case 43:
                                    case 44:
                                    case 45:
                                    case 48:
                                    case 49:
                                    case 50:
                                    case 51:
                                    case 52:
                                    case 53:
                                    case 56:
                                    case 58:
                                    case 59:
                                    case 60:
                                    case 61:
                                    case 62:
                                    case 63:
                                    case 64:
                                    case 67:
                                    case 80:
                                    case 81:
                                    case 88:
                                    case 89:
                                    case 90:
                                    case 92:
                                    case 93:
                                    case 94:
                                    case 95:
                                    case 96:
                                    case 97:
                                    case 100:
                                    case 101:
                                    case 102:
                                    case 103:
                                    case 104:
                                    case 105:
                                    case 108:
                                    case 110:
                                    case 111:
                                    case 112:
                                    case 113:
                                    case 114:
                                    case 115:
                                    case 116:
                                    case 119:
                                    case 147:
                                    case 148:
                                    case 155:
                                    case 156:
                                    case 157:
                                    case 159:
                                    case 160:
                                    case 161:
                                    case 162:
                                    case 163:
                                    case 164:
                                    case 167:
                                    case 168:
                                    case 169:
                                    case 170:
                                    case 171:
                                    case 172:
                                    case 175:
                                    case 177:
                                    case 178:
                                    case 179:
                                    case 180:
                                    case 181:
                                    case 182:
                                    case 183:
                                    case 186:
                                    case 199:
                                    case 200:
                                    case 207:
                                    case 208:
                                    case 209:
                                    case 211:
                                    case 212:
                                    case 213:
                                    case 214:
                                    case 215:
                                    case 216:
                                    case 219:
                                    case 220:
                                    case 221:
                                    case 222:
                                    case 223:
                                    case 224:
                                    case 227:
                                    case 229:
                                    case 230:
                                    case 231:
                                    case 232:
                                    case 233:
                                    case 234:
                                    case 235:
                                    case 238:
                                    case 393:
                                    case 6:
                                    case 7:
                                    case 8:
                                    case 9:
                                    case 125:
                                    case 126:
                                    case 127:
                                    case 128:
                                    case 16:
                                    case 17:
                                    case 18:
                                    case 19:
                                    case 20:
                                    case 21:
                                    case 22:
                                    case 23:
                                    case 24:
                                    case 25:
                                    case 26:
                                    case 27:
                                    case 68:
                                    case 69:
                                    case 70:
                                    case 71:
                                    case 72:
                                    case 73:
                                    case 74:
                                    case 75:
                                    case 76:
                                    case 77:
                                    case 78:
                                    case 79:
                                    case 135:
                                    case 136:
                                    case 137:
                                    case 138:
                                    case 139:
                                    case 140:
                                    case 141:
                                    case 142:
                                    case 143:
                                    case 144:
                                    case 145:
                                    case 146:
                                    case 187:
                                    case 188:
                                    case 189:
                                    case 190:
                                    case 191:
                                    case 192:
                                    case 193:
                                    case 194:
                                    case 195:
                                    case 196:
                                    case 197:
                                    case 198:
                                    case 1:
                                    case 2:
                                    case 3:
                                    case 4:
                                    case 5:
                                    case 120:
                                    case 121:
                                    case 122:
                                    case 123:
                                    case 124:

                                      //  HeatPump_DI(ref dttPointsValue, dttPoints, bReceive, sIP, iPort, iAddress, sPointID, sEquipmentID);
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            //case 4:
                            case 3:
                                //switch (170)
                                //switch (iAddress) 
                                switch (dtw["DataType"].ToString())
                                {
                                    //case 200:  case 405:
                                    //    HeatPump_ReadMultipleInputRegisters_PaddleWheel(ref dttPointsValue, dttPoints, bReceive, sIP, iPort, iAddress, sPointID, sEquipmentID);
                                    //    break;
                                    //case 170:  case 188:  case 375: case 393: case 97: case 302: case 139: case 344: case 129: case 334: case 163: case 368:
                                    //case 196:  case 198:  case 401: case 403:
                                    case "Uint32":
                                    case "Float":
                                       // ReadMultipleInputRegisters_2(ref dttPointsValue, dttPoints, bReceive, sIP, iPort, iAddress, sPointID, sEquipmentID, dtw);
                                        break;
                                    //case 413: case 414: case 435: case 436: case 457: case 458: case 479: case 480:
                                    //case 28:  case 31:  case 32:  case 33:  case 34:  case 35:  case 36:  case 37:  case 38:  case 59:  case 62:  case 63:
                                    //case 64:  case 65:  case 66:  case 67:  case 68:  case 69:  case 233: case 236: case 237: case 238: case 239: case 240:  case 241: case 242: case 243: case 264:
                                    //case 267: case 268: case 269: case 270: case 271: case 272: case 273: case 274: case 42:  case 46:  case 73:  case 77:   case 247: case 251: case 278: case 282:
                                    //case 165: case 370:
                                    case "Uint16":
                                        //HeatPump_ReadMultipleInputRegisters_1(ref dttPointsValue, dttPoints, bReceive, sIP, iPort, iAddress, sPointID, sEquipmentID, dtw);
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            default:
                                break;
                        }
                        break;
                    case "TCM3001":
                        switch (iAddress)
                        {
                            case 257:
                                TCM3001(ref dttPointsValue, dttPoints, bReceive, sIP, iPort, iAddress, sPointID);
                                break;
                            default:
                                break;
                        }
                        break;
                    case "TA-LR6641":
                        switch (iAddress)
                        {
                            case 257:
                              //  Catalog1(ref dttPointsValue, dttPoints, bReceive, sIP, iPort, iAddress, sPointID);
                                break;
                            case 513:
                               // Catalog2_1(ref dttPointsValue, dttPoints, bReceive, sIP, iPort, iAddress, sPointID);
                                break;
                            case 541:
                              //  Catalog2_2(ref dttPointsValue, dttPoints, bReceive, sIP, iPort, iAddress, sPointID);
                                break;
                            case 769:
                            case 783:
                            case 797:
                            case 811:
                            case 825:
                            case 839:
                            case 853:
                            case 867:
                            case 881:
                            case 895:
                                if (bReceive[11] == 0 && bReceive[12] == 0)
                                    break;
                                StatusLog(ref dttPointsValue, dttPoints, bReceive, sIP, iPort, iAddress, sPointID);
                                break;
                            case 1025:
                            case 1037:
                            case 1049:
                            case 1061:
                            case 1073:
                            case 1085:
                            case 1097:
                            case 1109:
                            case 1121:
                            case 1133:
                            case 1145:
                            case 1157:
                            case 1169:
                            case 1181:
                            case 1193:
                            case 1205:
                            case 1217:
                            case 1229:
                            case 1241:
                            case 1253:
                                if (bReceive[11] == 0 && bReceive[12] == 0)
                                    break;
                                EventLog(ref dttPointsValue, dttPoints, bReceive, sIP, iPort, iAddress, sPointID);
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
                        g_oLogger.Info(string.Format("無{0}設備的解析程式，請確認!", sEquipmentType));
                        break;
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.SolveValue", ex);
            }
        }

        //104.1.7 新增轉換正確值與四捨五入程式
        /// <summary>
        /// 補償值轉換
        /// </summary>
        /// <param name="sRowValue">要補償之數值</param>
        /// <param name="dtwPointValue">補償所需之參數資料</param>
        /// <returns></returns>
        public double CompensateValue(string sRowValue, DataRow dtwPointValue)
        {
            double dValue = 0.0;
            try
            {
                switch (dtwPointValue["DataType"].ToString())
                {
                    case "Boolean":
                        dValue = Convert.ToDouble(sRowValue);
                        break;
                    case "Float":
                    case "Uint32":
                    case "Uint16":
                        //dValue = C1Round(((Convert.ToDouble(sRowValue) - Convert.ToDouble(dtwPointValue["Raw_Min"])) * (Convert.ToDouble(dtwPointValue["EU_Max"]) - Convert.ToDouble(dtwPointValue["EU_Min"])) / (Convert.ToDouble(dtwPointValue["Raw_Max"]) - Convert.ToDouble(dtwPointValue["Raw_Min"])) + Convert.ToDouble(dtwPointValue["EU_Min"])), 2);
                        dValue = ((Convert.ToDouble(sRowValue) - Convert.ToDouble(dtwPointValue["Raw_Min"])) * (Convert.ToDouble(dtwPointValue["EU_Max"]) - Convert.ToDouble(dtwPointValue["EU_Min"])) / (Convert.ToDouble(dtwPointValue["Raw_Max"]) - Convert.ToDouble(dtwPointValue["Raw_Min"])) + Convert.ToDouble(dtwPointValue["EU_Min"]));

                        break;
                    case null:
                        dValue = Convert.ToDouble(sRowValue); ;
                        break;
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("CompensateValue-- ", ex);
                throw ex;
            }
            return dValue;
        }

        /// <summary>
        /// 四捨五入
        /// </summary>
        /// <param name="value"></param>
        /// <param name="digit"></param>
        /// <returns></returns>
        public static double C1Round(double value, int digit)
        {
            try
            {
                double vt = Math.Pow(10, digit);
                double vx = value * vt;

                vx += 0.5;
                return (Math.Floor(vx) / vt);
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("C1Round-- ", ex);
                throw ex;
            }
        }



        private void StatusLog(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            EM3001.EM3001 objEM3001 = new EM3001.EM3001();
            try
            {
                byte bEventType = bReceive[9];
                DateTime dtRecTime = new DateTime(2000 + bReceive[11], bReceive[12], bReceive[13], bReceive[14], bReceive[15], bReceive[16]);
                int iOutsideTemperature = BitConverter.ToInt16(bReceive, 17) / 10;
                int iInsideTemperature = BitConverter.ToInt16(bReceive, 19) / 10;
                int iVoltage48 = BitConverter.ToInt16(bReceive, 27) / 10;
                string sFireAlarm = objEM3001.ConvertToBinary(bReceive[29])[3];
                string sHeatExchangerAlarm = objEM3001.ConvertToBinary(bReceive[29])[2];
                string sAirAlarm = objEM3001.ConvertToBinary(bReceive[29])[1];
                string sStifledAlarm = objEM3001.ConvertToBinary(bReceive[29])[0];
                string sFanSwitch = objEM3001.ConvertToBinary(bReceive[30])[3];
                string sAlarmSwitch = objEM3001.ConvertToBinary(bReceive[30])[2];
                string sAirASwitch = objEM3001.ConvertToBinary(bReceive[30])[1];
                string sAirBSwitch = objEM3001.ConvertToBinary(bReceive[30])[0];
                //string sAO = objEM3001.AnalogOutput(bReceive[31]);
                string sTemperatureControlMode = objEM3001.TemperatureControlMode(bReceive[32]);
                string sTemperatureControlDays = objEM3001.AssociateDays(bReceive[33]);
                string sAirAssociateDays = objEM3001.AirAssociateDays(bReceive[34]);
                string sSDState = objEM3001.SDState(bReceive[35]);


                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}", sIP, iPort));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (dtwData["PointID"].ToString().Split('.')[dtwData["PointID"].ToString().Split('.').Length - 1].ToLower())
                    {
                        case "eventtype":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = objEM3001.EventType(bEventType);
                            break;
                        case "outsidetemperature":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = iOutsideTemperature.ToString("0.00");
                            break;
                        case "insidetemperature":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = iInsideTemperature.ToString("0.00");
                            break;
                        case "voltage48":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = iVoltage48.ToString("0.00");
                            break;
                        case "firealarm":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sFireAlarm;
                            break;
                        case "heatexchangeralarm":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sHeatExchangerAlarm;
                            break;
                        case "airalarm":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sAirAlarm;
                            break;
                        case "stifledalarm":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sStifledAlarm;
                            break;
                        case "fanswitch":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sFanSwitch;
                            break;
                        case "alarmswitch":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sAlarmSwitch;
                            break;
                        case "airaswitch":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sAirASwitch;
                            break;
                        case "airbswitch":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sAirBSwitch;
                            break;
                        case "temperaturecontrolmode":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sTemperatureControlMode;
                            break;
                        case "temperaturecontroldays":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sTemperatureControlDays;
                            break;
                        case "airassociatedays":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sAirAssociateDays;
                            break;
                        case "sdstate":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sSDState;
                            break;
                        default:
                            break;
                    }
                    dtwPointsValue["RecTime"] = dtRecTime;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.StatusLog", ex);
                throw ex;
            }
        }

        private void EventLog(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            EM3001.EM3001 objEM3001 = new EM3001.EM3001();
            try
            {
                byte bEventType = bReceive[9];
                DateTime dtRecTime = new DateTime(2000 + bReceive[11], bReceive[12], bReceive[13], bReceive[14], bReceive[15], bReceive[16]);
                int iOutsideTemperature = BitConverter.ToInt16(bReceive, 17) / 10;
                int iInsideTemperature = BitConverter.ToInt16(bReceive, 19) / 10;
                int iVoltage48 = BitConverter.ToInt16(bReceive, 27) / 10;
                string sFireAlarm = objEM3001.ConvertToBinary(bReceive[29])[3];
                string sHeatExchangerAlarm = objEM3001.ConvertToBinary(bReceive[29])[2];
                string sAirAlarm = objEM3001.ConvertToBinary(bReceive[29])[1];
                string sStifledAlarm = objEM3001.ConvertToBinary(bReceive[29])[0];
                string sFanSwitch = objEM3001.ConvertToBinary(bReceive[30])[3];
                string sAlarmSwitch = objEM3001.ConvertToBinary(bReceive[30])[2];
                string sAirASwitch = objEM3001.ConvertToBinary(bReceive[30])[1];
                string sAirBSwitch = objEM3001.ConvertToBinary(bReceive[30])[0];
                //string sAO = objEM3001.AnalogOutput(bReceive[31]);
                string sTemperatureControlMode = objEM3001.TemperatureControlMode(bReceive[32]);
                //string sTemperatureControlDays = objEM3001.AssociateDays(bReceive[33]);
                //string sAirAssociateDays = objEM3001.AirAssociateDays(bReceive[34]);
                //string sSDState = objEM3001.SDState(bReceive[35]);


                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}:{2}", sIP, iPort, iAddress));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (dtwData["PointID"].ToString().Split('.')[dtwData["PointID"].ToString().Split('.').Length - 1].ToLower())
                    {

                        case "eventtype":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = objEM3001.EventType(bEventType);
                            break;
                        case "outsidetemperature":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = iOutsideTemperature.ToString("0.00");
                            break;
                        case "insidetemperature":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = iInsideTemperature.ToString("0.00");
                            break;
                        case "voltage48":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = iVoltage48.ToString("0.00");
                            break;
                        case "firealarm":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sFireAlarm;
                            break;
                        case "heatexchangeralarm":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sHeatExchangerAlarm;
                            break;
                        case "airalarm":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sAirAlarm;
                            break;
                        case "stifledalarm":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sStifledAlarm;
                            break;
                        case "fanswitch":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sFanSwitch;
                            break;
                        case "alarmswitch":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sAlarmSwitch;
                            break;
                        case "airaswitch":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sAirASwitch;
                            break;
                        case "airbswitch":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sAirBSwitch;
                            break;
                        case "temperaturecontrolmode":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = sTemperatureControlMode;
                            break;
                        default:
                            break;
                    }
                    dtwPointsValue["RecTime"] = dtRecTime;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.EventLog", ex);
                throw ex;
            }
        }

        private void ControlMode(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            EM3001.EM3001 objEM3001 = new EM3001.EM3001();
            try
            {
                byte bControlMode = bReceive[9];

                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}:{2}", sIP, iPort, iAddress));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (bControlMode)
                    {

                        case 1:
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = "ManualOpen";
                            break;
                        case 2:
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = "ManualStop";
                            break;
                        case 4:
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = "Automatic";
                            break;
                        default:
                            break;
                    }
                    dtwPointsValue["RecTime"] = DateTime.Now;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.ControlMode", ex);
                throw ex;
            }
        }

        private void WorkingHours(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            EM3001.EM3001 objEM3001 = new EM3001.EM3001();
            try
            {
                UInt32 uiSystemWorkingHour = BitConverter.ToUInt32(bReceive, 9);
                UInt32 uiAirFanerWorkingHour = BitConverter.ToUInt32(bReceive, 13);
                UInt32 uiAirConditionerAWorkingHour = BitConverter.ToUInt32(bReceive, 17);
                UInt32 uiAirConditionerBWorkingHour = BitConverter.ToUInt32(bReceive, 20);

                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}:{2}", sIP, iPort, iAddress));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (dtwData["PointID"].ToString().Split('.')[2].ToLower())
                    {

                        case "systemworkinghour":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = uiSystemWorkingHour;
                            break;
                        case "airfanerworkinghour":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = uiAirFanerWorkingHour;
                            break;
                        case "airconditioneraworkinghour":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = uiAirConditionerAWorkingHour;
                            break;
                        case "airconditionerbworkinghour":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = uiAirConditionerBWorkingHour;
                            break;
                        default:
                            g_oLogger.Error("請確認PointID之內容、格式是否正確");
                            break;
                    }
                    dtwPointsValue["RecTime"] = DateTime.Now;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.WorkingHours", ex);
                throw ex;
            }
        }

        private void CabinTempSetting(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            EM3001.EM3001 objEM3001 = new EM3001.EM3001();
            try
            {
                float fHHTempereture = BitConverter.ToSingle(bReceive, 9);
                float fHTempereture = BitConverter.ToSingle(bReceive, 13);
                float fLTempereture = BitConverter.ToSingle(bReceive, 17);
                float fLLTempereture = BitConverter.ToSingle(bReceive, 20);
                float fDiffTempereture = BitConverter.ToSingle(bReceive, 24);
                float fInsideOffsetTempereture = BitConverter.ToSingle(bReceive, 28);

                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}:{2}", sIP, iPort, iAddress));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (dtwData["PointID"].ToString().Split('.')[2].ToLower())
                    {

                        case "hhtempereture":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fHHTempereture.ToString("0.00");
                            break;
                        case "htempereture":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fHTempereture.ToString("0.00");
                            break;
                        case "ltempereture":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fLTempereture.ToString("0.00");
                            break;
                        case "lltempereture":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fLLTempereture.ToString("0.00");
                            break;
                        case "difftempereture":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fDiffTempereture.ToString("0.00");
                            break;
                        case "offsettempereture":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fInsideOffsetTempereture.ToString("0.00");
                            break;
                        default:
                            g_oLogger.Error("請確認PointID之內容、格式是否正確");
                            break;
                    }
                    dtwPointsValue["RecTime"] = DateTime.Now;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.CabinTempSetting", ex);
                throw ex;
            }
        }

        private void PSMSetting(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            EM3001.EM3001 objEM3001 = new EM3001.EM3001();
            try
            {
                byte bCoworkerModeDay = bReceive[9];
                byte bAirConditionerModeDay = bReceive[10];

                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}:{2}", sIP, iPort, iAddress));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (dtwData["PointID"].ToString().Split('.')[2].ToLower())
                    {

                        case "coworkermodeday":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = bCoworkerModeDay;
                            break;
                        case "airconditionermodeday":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = bAirConditionerModeDay;
                            break;
                        default:
                            g_oLogger.Error("請確認PointID之內容、格式是否正確");
                            break;
                    }
                    dtwPointsValue["RecTime"] = DateTime.Now;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.PSMSetting", ex);
                throw ex;
            }
        }

        private void LWTSetting(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            EM3001.EM3001 objEM3001 = new EM3001.EM3001();
            try
            {
                byte bAirFanerMinTime = bReceive[9];
                byte bAirConditionerMinTime = bReceive[10];

                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}:{2}", sIP, iPort, iAddress));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (dtwData["PointID"].ToString().Split('.')[2].ToLower())
                    {

                        case "airfanermintime":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = bAirFanerMinTime;
                            break;
                        case "airconditionermintime":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = bAirConditionerMinTime;
                            break;
                        default:
                            g_oLogger.Error("請確認PointID之內容、格式是否正確");
                            break;
                    }
                    dtwPointsValue["RecTime"] = DateTime.Now;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.LWTSetting", ex);
                throw ex;
            }
        }

        private void SystemTime(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            EM3001.EM3001 objEM3001 = new EM3001.EM3001();
            try
            {
                DateTime dtSystem = new DateTime(2000 + bReceive[9], bReceive[10], bReceive[11]
                    , bReceive[12], bReceive[13], bReceive[14]);

                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}:{2}", sIP, iPort, iAddress));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (dtwData["PointID"].ToString().Split('.')[2].ToLower())
                    {

                        case "systemtime":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = dtSystem.ToString("yyyy-MM-dd HH:mm:ss.fff");
                            break;
                        default:
                            g_oLogger.Error("請確認PointID之內容、格式是否正確");
                            break;
                    }
                    dtwPointsValue["RecTime"] = DateTime.Now;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.SystemTime", ex);
                throw ex;
            }
        }

        private void LVSSetting(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            EM3001.EM3001 objEM3001 = new EM3001.EM3001();
            try
            {
                float fHighVoltage = BitConverter.ToSingle(bReceive, 9);
                float fLowVoltage = BitConverter.ToSingle(bReceive, 13);
                float fStopVoltage = BitConverter.ToSingle(bReceive, 17);

                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}:{2}", sIP, iPort, iAddress));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (dtwData["PointID"].ToString().Split('.')[2].ToLower())
                    {

                        case "highvoltage":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fHighVoltage.ToString("0.00"); ;
                            break;
                        case "lowvoltage":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fLowVoltage.ToString("0.00"); ;
                            break;
                        case "stopvoltage":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fStopVoltage.ToString("0.00"); ;
                            break;
                        default:
                            g_oLogger.Error("請確認PointID之內容、格式是否正確");
                            break;
                    }
                    dtwPointsValue["RecTime"] = DateTime.Now;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.LVSSetting", ex);
                throw ex;
            }
        }

        private void AirConditionerForceOPSetting(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            EM3001.EM3001 objEM3001 = new EM3001.EM3001();
            try
            {
                byte bContinousNoTurnOnDay = bReceive[9];
                byte bForceTurnOnTime = bReceive[10];

                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}:{2}", sIP, iPort, iAddress));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (dtwData["PointID"].ToString().Split('.')[2].ToLower())
                    {

                        case "continousnoturnonday":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = bContinousNoTurnOnDay;
                            break;
                        case "forceturnontime":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = bForceTurnOnTime;
                            break;
                        default:
                            g_oLogger.Error("請確認PointID之內容、格式是否正確");
                            break;
                    }
                    dtwPointsValue["RecTime"] = DateTime.Now;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.AirConditionerForceOPSetting", ex);
                throw ex;
            }
        }

        private DataRow[] SelectPoints(DataTable dttPoints, string sIP, int iPort, int iAddress, string sPointID)
        {
            DataRow[] dtwPoints = null;
            try
            {
                dtwPoints = dttPoints.Select(string.Format("IP = '{0}' AND Port = '{1}' AND Address = '{2}' AND PointID = '{3}'", sIP, iPort, iAddress, sPointID));
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.SelectPoints", ex);
                throw ex;
            }
            return dtwPoints;
        }





      

        private void Send(Socket socket, byte[] buffer, int offset, int size, int timeout)
        {
            int startTickCount = Environment.TickCount;
            int sent = 0;  // how many bytes is already sent
            do
            {
                if (Environment.TickCount > startTickCount + timeout)
                    throw new Exception("Timeout.");
                try
                {
                    sent += socket.Send(buffer, offset + sent, size - sent, SocketFlags.None);
                }
                catch (SocketException ex)
                {
                    if (ex.SocketErrorCode == SocketError.WouldBlock ||
                        ex.SocketErrorCode == SocketError.IOPending ||
                        ex.SocketErrorCode == SocketError.NoBufferSpaceAvailable)
                    {
                        // socket buffer is probably full, wait and try again
                        Thread.Sleep(30);
                    }
                    else
                        throw ex;  // any serious error occurr
                }
            } while (sent < size);
        }

        public void Receive(Socket socket, ref byte[] buffer, int offset, int size, int timeout)
        {
            int startTickCount = Environment.TickCount;
            int received = 0;  // how many bytes is already received
            do
            {
                if (Environment.TickCount > startTickCount + timeout)
                    throw new Exception("Timeout.");
                try
                {
                    received += socket.Receive(buffer, offset + received, size - received, SocketFlags.None);
                }
                catch (SocketException ex)
                {
                    if (ex.SocketErrorCode == SocketError.WouldBlock ||
                        ex.SocketErrorCode == SocketError.IOPending ||
                        ex.SocketErrorCode == SocketError.NoBufferSpaceAvailable)
                    {
                        // socket buffer is probably empty, wait and try again
                        Thread.Sleep(30);
                    }
                    else
                        throw ex;  // any serious error occurr
                }
            } while (received < size);
            Array.Resize(ref buffer, received);
        }

        private void TCM3001(ref DataTable dttPointsValue, DataTable dttPoints, byte[] bReceive, string sIP, int iPort, int iAddress, string sPointID)
        {
            try
            {
                float fPhaseA_IRMS = BitConverter.ToSingle(bReceive, 9);
                float fPhaseA_VRMS = BitConverter.ToSingle(bReceive, 13);
                float fPhaseA_PF = BitConverter.ToSingle(bReceive, 17);
                float fPhaseA_Hz = BitConverter.ToSingle(bReceive, 21);
                float fPhaseA_kW = BitConverter.ToSingle(bReceive, 25);
                float fPhaseA_kWh = BitConverter.ToSingle(bReceive, 29);
                float fPhaseA_VAr = BitConverter.ToSingle(bReceive, 33);
                float fPhaseA_VArh = BitConverter.ToSingle(bReceive, 37);
                float fPhaseA_VA = BitConverter.ToSingle(bReceive, 41);
                float fPhaseA_VAh = BitConverter.ToSingle(bReceive, 45);

                float fPhaseB_IRMS = BitConverter.ToSingle(bReceive, 49);
                float fPhaseB_VRMS = BitConverter.ToSingle(bReceive, 53);
                float fPhaseB_PF = BitConverter.ToSingle(bReceive, 57);
                float fPhaseB_Hz = BitConverter.ToSingle(bReceive, 61);
                float fPhaseB_kW = BitConverter.ToSingle(bReceive, 65);
                float fPhaseB_kWh = BitConverter.ToSingle(bReceive, 69);
                float fPhaseB_VAr = BitConverter.ToSingle(bReceive, 73);
                float fPhaseB_VArh = BitConverter.ToSingle(bReceive, 77);
                float fPhaseB_VA = BitConverter.ToSingle(bReceive, 81);
                float fPhaseB_VAh = BitConverter.ToSingle(bReceive, 85);

                float fPhaseC_IRMS = BitConverter.ToSingle(bReceive, 89);
                float fPhaseC_VRMS = BitConverter.ToSingle(bReceive, 93);
                float fPhaseC_PF = BitConverter.ToSingle(bReceive, 97);
                float fPhaseC_Hz = BitConverter.ToSingle(bReceive, 101);
                float fPhaseC_kW = BitConverter.ToSingle(bReceive, 105);
                float fPhaseC_kWh = BitConverter.ToSingle(bReceive, 109);
                float fPhaseC_VAr = BitConverter.ToSingle(bReceive, 113);
                float fPhaseC_VArh = BitConverter.ToSingle(bReceive, 117);
                float fPhaseC_VA = BitConverter.ToSingle(bReceive, 121);
                float fPhaseC_VAh = BitConverter.ToSingle(bReceive, 125);

                DataRow[] dtwPoint = SelectPoints(dttPoints, sIP, iPort, iAddress, sPointID);
                if (dtwPoint.Length == 0)
                {
                    g_oLogger.Trace(string.Format("請確認點位資料- {0}:{1}:{2}", sIP, iPort, iAddress));
                    return;
                }
                foreach (DataRow dtwData in dtwPoint)
                {
                    DataRow dtwPointsValue = null;
                    switch (dtwData["PointID"].ToString().Split('.')[dtwData["PointID"].ToString().Split('.').Length - 1].ToLower())
                    {

                        case "phasea_irms":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseA_IRMS.ToString("0.00");
                            break;
                        case "phasea_vrms":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseA_VRMS.ToString("0.00");
                            break;
                        case "phasea_pf":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseA_PF.ToString("0.00");
                            break;
                        case "phasea_fz":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseA_Hz.ToString("0.00");
                            break;
                        case "phasea_kw":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseA_kW.ToString("0.00");
                            break;
                        case "phasea_kwh":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseA_kWh.ToString("0.00");
                            break;
                        case "phasea_var":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseA_VAr.ToString("0.00");
                            break;
                        case "phasea_varh":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseA_VArh.ToString("0.00");
                            break;
                        case "phasea_va":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseA_VA.ToString("0.00");
                            break;
                        case "phasea_vah":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseA_VAh.ToString("0.00");
                            break;

                        case "phaseb_irms":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseB_IRMS.ToString("0.00");
                            break;
                        case "phaseb_vrms":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseB_VRMS.ToString("0.00");
                            break;
                        case "phaseb_pf":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseB_PF.ToString("0.00");
                            break;
                        case "phaseb_fz":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseB_Hz.ToString("0.00");
                            break;
                        case "phaseb_kw":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseB_kW.ToString("0.00");
                            break;
                        case "phaseb_kwh":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseB_kWh.ToString("0.00");
                            break;
                        case "phaseb_var":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseB_VAr.ToString("0.00");
                            break;
                        case "phaseb_varh":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseB_VArh.ToString("0.00");
                            break;
                        case "phaseb_va":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseB_VA.ToString("0.00");
                            break;
                        case "phaseb_vah":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseB_VAh.ToString("0.00");
                            break;

                        case "phasec_irms":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseC_IRMS.ToString("0.00");
                            break;
                        case "phasec_vrms":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseC_VRMS.ToString("0.00");
                            break;
                        case "phasec_pf":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseC_PF.ToString("0.00");
                            break;
                        case "phasec_fz":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseC_Hz.ToString("0.00");
                            break;
                        case "phasec_kw":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseC_kW.ToString("0.00");
                            break;
                        case "phasec_kwh":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseC_kWh.ToString("0.00");
                            break;
                        case "phasec_var":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseC_VAr.ToString("0.00");
                            break;
                        case "phasec_varh":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseC_VArh.ToString("0.00");
                            break;
                        case "phasec_va":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseC_VA.ToString("0.00");
                            break;
                        case "phasec_vah":
                            dtwPointsValue = dttPointsValue.NewRow();
                            dtwPointsValue["EquipmentID"] = dtwData["EquipmentID"].ToString();
                            dtwPointsValue["PointID"] = dtwData["PointID"].ToString();
                            dtwPointsValue["PointValue"] = fPhaseC_VAh.ToString("0.00");
                            break;
                        default:
                            g_oLogger.Error("請確認PointID之內容、格式是否正確");
                            break;
                    }
                    dtwPointsValue["RecTime"] = DateTime.Now;
                    dttPointsValue.Rows.Add(dtwPointsValue);
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("ModbusTCP.TCM3001", ex);
                throw ex;
            }
        }





    }
}
