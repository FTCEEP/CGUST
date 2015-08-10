using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using NLog;
using System.Timers;
using System.Configuration;
using System.Reflection;
using System.Threading;

namespace Fpg.Cim.Adapter.ModbusTCP.Service1
{
    public partial class Service1 : ServiceBase
    {
        System.Timers.Timer g_oTimerModbusTCP = new System.Timers.Timer();
        System.Timers.Timer g_oTimerUpdateService = new System.Timers.Timer();
        System.Timers.Timer g_oTimerACChange = new System.Timers.Timer();

        private static Logger g_oLogger = NLog.LogManager.GetCurrentClassLogger();
        private static string g_sCompanyID = ConfigurationManager.AppSettings.Get("CompanyID");
        private static string g_sProjectID = ConfigurationManager.AppSettings.Get("ProjectID");
        private static string g_sEquipmentIDs = ConfigurationManager.AppSettings.Get("EquipmentIDs");
        private static string g_sIP = ConfigurationManager.AppSettings.Get("IP");
        private static int g_iPort = int.Parse(ConfigurationManager.AppSettings.Get("Port"));
        private static string g_sProgrammerMail = ConfigurationManager.AppSettings.Get("ProgrammerMail");
        private static int g_iReTryTimes = int.Parse(ConfigurationManager.AppSettings.Get("ReTryTimes"));
        private static int g_iCheckMeterParaTime = int.Parse(ConfigurationManager.AppSettings.Get("CheckMeterParaTime"));
        private static string g_sSplitSymbol = ConfigurationManager.AppSettings.Get("SplitSymbol");
        private static int g_iReceiveTimeout = int.Parse(ConfigurationManager.AppSettings.Get("ReceiveTimeout"));
        private static string g_sEquipmentsFilePath = ConfigurationManager.AppSettings.Get("EquipmentsFilePath");
        private static string g_sPointsFilePath = ConfigurationManager.AppSettings.Get("PointsFilePath");
        private static int g_iDBCallTimes = int.Parse(ConfigurationManager.AppSettings.Get("DBCallTimes"));
        private static int g_iEquipmentStateTime = int.Parse(ConfigurationManager.AppSettings.Get("EquipmentStateTime"));
        private static int g_iSystemRunTimeRecord = int.Parse(ConfigurationManager.AppSettings.Get("SystemRunTimeRecord"));
        private static string g_sEquipmentType = ConfigurationManager.AppSettings.Get("EquipmentType");
        private static int g_iEquipmenStateCheck = Int16.Parse(ConfigurationManager.AppSettings.Get("EquipmenStateCheck"));

        public Service1()
        {
            InitializeComponent();
            RecordingSystemLog();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                // 設定objTimerModbusRTU的Handler
                g_oTimerModbusTCP.Elapsed += new ElapsedEventHandler(objTimerModbusTCP_Elapsed);
                // objTimerModbusRTU的時間間隔(單位: ms) 5s1筆
                g_oTimerModbusTCP.Interval = 5 * 1000 * 1 ;
                // 啟動objTimerModbusRTU
                g_oTimerModbusTCP.Start();
                g_oLogger.Trace(string.Format("{0} 線程開始", "ModbusTCP"));

                // 設定
                g_oTimerACChange.Elapsed += new ElapsedEventHandler(objTimerACChange_Elapsed);
                // objTimerACChange的時間間隔(單位: ms) 2s1筆
                g_oTimerACChange.Interval = 2 * 1000 * 1;
                // 啟動obj
                g_oTimerACChange.Start();
                g_oLogger.Trace(string.Format("{0} 線程開始", "ACChange"));


                //Equipment.Equipment.ModbusTCP objEquipment = new Equipment.Equipment.ModbusTCP(
                //    g_sCompanyID, g_sEquipmentIDs, g_iDBCallTimes, g_iEquipmentStateTime);
                //Thread objEquipmentThread = new Thread(new ThreadStart(objEquipment.Start));
                //objEquipmentThread.IsBackground = true;
                //objEquipmentThread.Start();
                //g_oLogger.Trace(string.Format("{0} 線程開始", "Equipment"));

                // 設定objTimerUpdateService的Handler
                //g_oTimerUpdateService.Elapsed += new ElapsedEventHandler(objTimerUpdateService_Elapsed);
                // objTimerUpdateService的時間間隔(單位: ms)
                //g_oTimerUpdateService.Interval = g_iSystemRunTimeRecord * 1000;
                // 啟動objTimerUpdateService
                //g_oTimerUpdateService.Start();
                //g_oLogger.Trace(string.Format("{0} 線程開始", "UpdateService"));

                //g_oLogger.Info("Service OnStart");
                //g_oLogger.Info("Version：" + FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion.ToString());
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("Service1.OnStart", ex);
                throw ex;
            }
        }

        protected override void OnStop()
        {
            try
            {
                g_oLogger.Trace("Window Service 結束");
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("Service1.OnStop", ex);
                throw ex;
            }
        }

        private void objTimerModbusTCP_Elapsed(object sender, EventArgs e)
        {
            ModbusTCP.ModbusTCP objModbusTCP = new ModbusTCP.ModbusTCP();
            try
            {
                // Do something here!
                g_oLogger.Info("TimerModbusTCP Run Time--" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                objModbusTCP.Start();
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("Service1.objTimerModbusTCP_Elapsed", ex);
            }
            finally
            {
                if (objModbusTCP != null) { objModbusTCP = null; }
            }
        }

        /// <summary>
        ///104.8.5新增ACChange空調更改程式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void objTimerACChange_Elapsed(object sender, EventArgs e)
        {
            ACChange.ACChange objACChange = new ACChange.ACChange();
            try
            {
                // Do something here!
                g_oLogger.Info("TimerACChange Run Time--" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                objACChange.Start();
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("Service1.objTimerACChange_Elapsed", ex);
            }
            finally
            {
                if (objACChange != null) { objACChange = null; }
            }
        }


        public void RecordingSystemLog()
        {
            Dictionary<string, string> SystemLogs = new Dictionary<string, string>();
            try
            {
                SystemLogs.Add("CompanyID", g_sCompanyID);
                SystemLogs.Add("ProjectID", g_sProjectID);
                SystemLogs.Add("EquipmentIDs", g_sEquipmentIDs);
                SystemLogs.Add("IP", g_sIP);
                SystemLogs.Add("Port", g_iPort.ToString());
                SystemLogs.Add("ProgrammerMail", g_sProgrammerMail);
                SystemLogs.Add("ReTryTimes", g_iReTryTimes.ToString());
                SystemLogs.Add("DBCallTimes", g_iDBCallTimes.ToString());
                SystemLogs.Add("CheckMeterParaTime", g_iCheckMeterParaTime.ToString());
                SystemLogs.Add("SplitSymbol", g_sSplitSymbol);
                SystemLogs.Add("ReceiveTimeout", g_iReceiveTimeout.ToString());
                SystemLogs.Add("EquipmentsFilePath", g_sEquipmentsFilePath);
                SystemLogs.Add("PointsFilePath", g_sPointsFilePath);
                SystemLogs.Add("EquipmentStateTime", g_iEquipmentStateTime.ToString());
                SystemLogs.Add("SystemRunTimeRecord", g_iSystemRunTimeRecord.ToString());
                SystemLogs.Add("EquipmentType", g_sEquipmentType);
                SystemLogs.Add("EquipmenStateCheck", g_iEquipmenStateCheck.ToString());

                g_oLogger.Info("------------------系統參數--------------------");
                foreach (string sKey in SystemLogs.Keys)
                {
                    g_oLogger.Info(string.Format("系統App.config設定值{0}: {1}"
                        , sKey
                        , SystemLogs[sKey]));
                }
                g_oLogger.Info("------------------系統參數--------------------");

            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("RecordingSystemLog--", ex);
            }
        }

        private void objTimerUpdateService_Elapsed(object sender, EventArgs e)
        {
            UpdateServiceTime.UpdateServiceTime objUpdateServiceTime = new UpdateServiceTime.UpdateServiceTime();
            try
            {
                g_oLogger.Info("UpdateServiceTime Run Time--" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                objUpdateServiceTime.Update_Time();
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("Service1.objTimerUpdateService_Elapsed", ex);
            }
            finally
            {
                if (objUpdateServiceTime != null) { objUpdateServiceTime = null; }
            }
        }
    }
}
