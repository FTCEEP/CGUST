using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Configuration;
using NLog;

namespace Fpg.Cim.Adapter.ModbusTCP.UpdateServiceTime
{
    class UpdateServiceTime
    {
        private string gsProjectID = ConfigurationManager.AppSettings.Get("ProjectID"); //從AppConfig讀取「ProjectID」參數
        private int giSystemRunTimeRecord = Convert.ToInt16(ConfigurationManager.AppSettings.Get("SystemRunTimeRecord")); //從AppConfig讀取「SystemRunTimeRecord」參數
        private string gsIP = ConfigurationManager.AppSettings.Get("IP");
        private string gsPort = ConfigurationManager.AppSettings.Get("Port");
        //DB Connection string
        private string gsEEPDB = ConfigurationManager.ConnectionStrings["EEP_Client"].ToString();
        // Nlog
        private static Logger gLogger = NLog.LogManager.GetCurrentClassLogger();

        /// <summary>
        /// 紀錄WindowServic運作時間
        /// 用來判斷WindowService是否有正常作動
        /// </summary>
        public void Update_Time()
        {
            DBConnection cDBConnection;
            bool blFlag;
            string cmdSQL = string.Empty;
            string sServiceName = gsProjectID + ".ModbusTCP.Adapter(" + gsIP + ":" + gsPort + ")";
            try
            {
                cDBConnection = new DBConnection();
                blFlag = false;
                blFlag = CheckEEP_ExecutingServiceStateTable(gsProjectID);
                if (blFlag == false)
                {
                    gLogger.Error("產生EEP_ExecutingServiceState Table失敗");
                }
                cmdSQL = string.Format("SELECT * FROM EEP_ExecutingServiceState WHERE ProjectID = '{0}' AND ServiceName = '{1}'", gsProjectID, sServiceName);
                DataTable dttExecutingServiceState = cDBConnection.execDataTable(cmdSQL, gsEEPDB);
                if (dttExecutingServiceState.Rows.Count > 0)
                {
                    cmdSQL = string.Format("UPDATE EEP_ExecutingServiceState SET RecTime = '{2}', AlarmIntervalSecond = '{3}' WHERE ProjectID = '{0}' AND ServiceName = '{1}'", gsProjectID, sServiceName, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), dttExecutingServiceState.Rows[0]["AlarmIntervalSecond"].ToString());
                    cDBConnection.execNonQuery(cmdSQL, gsEEPDB);
                    gLogger.Trace("更新WindowService監控時間成功");
                }
                else
                {
                    cmdSQL = string.Format("INSERT INTO EEP_ExecutingServiceState (ProjectID, ServiceName, RecTime, AlarmIntervalSecond) VALUES ('{0}', '{1}', '{2}', '{3}')", gsProjectID, sServiceName, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), giSystemRunTimeRecord);
                    cDBConnection.execNonQuery(cmdSQL, gsEEPDB);
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("UpdateServiceTime.Update_Time", ex);
            }
        }

        private bool CheckEEP_ExecutingServiceStateTable(string sProjectID)
        {
            bool blExistence = false;
            DBConnection cDBConnection;
            try
            {
                cDBConnection = new DBConnection();
                DataTable dtt = new DataTable();
                DataRow[] dtwSelection;
                string cmdSQL = "select name from SysObjects order by name";
                dtt = cDBConnection.execDataTable(cmdSQL, gsEEPDB);
                dtwSelection = dtt.Select("name = 'EEP_ExecutingServiceState'");
                switch (dtwSelection.Length)
                {
                    case 1: // have being the table.
                        blExistence = true;
                        gLogger.Trace("已經有EEP_ExecutingServiceState的Table");
                        break;
                    case 0://Create a new table
                        cmdSQL = "CREATE TABLE dbo.EEP_ExecutingServiceState (ProjectID NVARCHAR (50) NOT NULL, ServiceName NVARCHAR (50) NOT NULL, RecTime DATETIME NOT NULL, AlarmIntervalMinute INT NOT NULL, CONSTRAINT PK_EEP_ExecutingServiceState PRIMARY KEY (ProjectID, ServiceName));";
                        //cmdSQL = string.Format("CREATE TABLE [dbo].[RMC_HistoryData_{0}_{1}]([EquipmentID] [nvarchar](50) NULL, [PointID] [nvarchar](50) NULL, [PointValue] [nvarchar](50) NULL, [RecTime] [datetime] NULL) ON [PRIMARY]", sProjectID, DateTime.Now.ToString("yyyyMM"));
                        cDBConnection.execNonQuery(cmdSQL, gsEEPDB);
                        blExistence = true;
                        gLogger.Trace("產生新的EEP_ExecutingServiceState的Table");
                        break;
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("UpdateServiceTime.CheckEEP_ExecutingServiceStateTable", ex);
            }
            return blExistence;
        }
    }
}
