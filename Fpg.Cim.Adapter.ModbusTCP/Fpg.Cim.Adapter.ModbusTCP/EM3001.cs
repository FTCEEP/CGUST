using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NLog;

namespace Fpg.Cim.Adapter.ModbusTCP.EM3001
{
    class EM3001
    {
        private Logger gLogger = NLog.LogManager.GetCurrentClassLogger();
        public string[] ConvertToBinary(byte byData)
        {
            string[] sValue = null;
            try
            {
                switch (Convert.ToString(byData, 2))
                {
                    case "0":
                        sValue = new string[4] { "0", "0", "0", "0" };
                        break;
                    case "1":
                        sValue = new string[4] { "0", "0", "0", "1" };
                        break;
                    case "10":
                        sValue = new string[4] { "0", "0", "1", "0" };
                        break;
                    case "11":
                        sValue = new string[4] { "0", "0", "1", "1" };
                        break;
                    case "100":
                        sValue = new string[4] { "0", "1", "0", "0" };
                        break;
                    case "101":
                        sValue = new string[4] { "0", "1", "0", "1" };
                        break;
                    case "111":
                        sValue = new string[4] { "0", "1", "1", "1" };
                        break;
                    case "1000":
                        sValue = new string[4] { "1", "0", "0", "0" };
                        break;
                    case "1001":
                        sValue = new string[4] { "1", "0", "0", "1" };
                        break;
                    case "1010":
                        sValue = new string[4] { "1", "0", "1", "0" };
                        break;
                    case "1011":
                        sValue = new string[4] { "1", "0", "1", "1" };
                        break;
                    case "1100":
                        sValue = new string[4] { "1", "1", "0", "0" };
                        break;
                    case "1101":
                        sValue = new string[4] { "1", "1", "0", "1" };
                        break;
                    case "1110":
                        sValue = new string[4] { "1", "1", "1", "0" };
                        break;
                    case "1111":
                        sValue = new string[4] { "1", "1", "1", "1" };
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return sValue;
        }

        public string AnalogOutput(byte byData)
        {
            string sValue;
            try
            {
                switch (Convert.ToInt32(byData).ToString("X"))
                {
                    case "1":
                        sValue = "0";
                        break;
                    case "2":
                        sValue = "2.5";
                        break;
                    case "4":
                        sValue = "5";
                        break;
                    case "8":
                        sValue = "7.5";
                        break;
                    case "10":
                        sValue = "10";
                        break;
                    default:
                        sValue = "0";
                        break;
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("EM3001.AnalogOutput", ex);
                throw ex;
            }
            return sValue;
        }

        public string AirConditionerSwapDuty(byte byData)
        {
            string sValue;
            try
            {
                switch (Convert.ToInt32(byData).ToString("X"))
                {
                    case "0":
                        sValue = "Disable";
                        break;
                    default:
                        sValue = Convert.ToInt32(byData).ToString("X");
                        break;
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("EM3001.AnalogOutput", ex);
                throw ex;
            }
            return sValue;
        }

        public string TemperatureControlMode(byte byData)
        {
            string sValue;
            try
            {
                switch (byData)
                {
                    case 1:
                        sValue = "Single";
                        break;
                    case 2:
                        sValue = "Associate";
                        break;
                    default:
                        sValue = "";
                        break;
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("EM3001.TemperatureControlMode", ex);
                throw ex;
            }
            return sValue;
        }

        public string AssociateDays(byte byData)
        {
            string sValue;
            try
            {
                switch (byData)
                {
                    case 0:
                        sValue = "InActive";
                        break;
                    default:
                        sValue = byData.ToString();
                        break;
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("EM3001.AssociateDays", ex);
                throw ex;
            }
            return sValue;
        }

        public string AirAssociateDays(byte byData)
        {
            string sValue;
            try
            {
                switch (byData)
                {
                    case 0:
                        sValue = "InActive";
                        break;
                    default:
                        sValue = byData.ToString();
                        break;
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("EM3001.AssociateDays", ex);
                throw ex;
            }
            return sValue;
        }

        public string SDState(byte byData)
        {
            string sValue;
            try
            {
                switch (byData)
                {
                    case 0:
                        sValue = "No";
                        break;
                    case 1:
                        sValue = "Yes";
                        break;
                    default:
                        sValue = byData.ToString();
                        break;
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("EM3001.SDState", ex);
                throw ex;
            }
            return sValue;
        }

        public string EventType(int iEventType)
        {
            string sEventType = string.Empty;
            try
            {
                switch (iEventType)
                {
                    case 0:
                        sEventType = "每分鐘的Status log";
                        break;
                    case 1:
                        sEventType = "風機告警產生";
                        break;
                    case 101:
                        sEventType = "風機告警解除";
                        break;
                    case 2:
                        sEventType = "空調告警產生";
                        break;
                    case 102:
                        sEventType = "空調告警解除";
                        break;
                    case 3:
                        sEventType = "濾網告警產生";
                        break;
                    case 103:
                        sEventType = "濾網告警解除";
                        break;
                    case 4:
                        sEventType = "火警告警產生";
                        break;
                    case 104:
                        sEventType = "火警告警解除";
                        break;
                    case 5:
                        sEventType = "超高溫告警產生";
                        break;
                    case 105:
                        sEventType = "超高溫告警解除";
                        break;
                    case 6:
                        sEventType = "室內溫度線告警產生";
                        break;
                    case 106:
                        sEventType = "室內溫度線告警解除";
                        break;
                    case 7:
                        sEventType = "室外溫度線告警產生";
                        break;
                    case 107:
                        sEventType = "室外溫度線告警解除";
                        break;
                    case 8:
                        sEventType = "48V高壓告警產生";
                        break;
                    case 108:
                        sEventType = "48V高壓告警解除";
                        break;
                    case 9:
                        sEventType = "48V低壓告警產生";
                        break;
                    case 109:
                        sEventType = "48V低壓告警解除";
                        break;
                    case 10:
                        sEventType = "48V停止壓告警產生";
                        break;
                    case 110:
                        sEventType = "48V停止壓告警解除";
                        break;
                    case 200:
                        sEventType = "開機紀錄";
                        break;
                    case 201:
                        sEventType = "SD card回存開始";
                        break;
                    case 202:
                        sEventType = "SD card回存結束";
                        break;
                    case 210:
                        sEventType = "參數改變: 系統時間設定";
                        break;
                    case 211:
                        sEventType = "參數改變: 室內溫度設定";
                        break;
                    case 212:
                        sEventType = "參數改變: 室外溫度設定";
                        break;
                    case 213:
                        sEventType = "參數改變: 溫度控制模式設定";
                        break;
                    case 214:
                        sEventType = "參數改變: 48V電壓設定";
                        break;
                    case 215:
                        sEventType = "參數改變: 空調強制運行設定";
                        break;
                    case 216:
                        sEventType = "參數改變: 空調輪替模式設定";
                        break;
                    case 217:
                        sEventType = "參數改變: 空調輪替週期設定";
                        break;
                    case 218:
                        sEventType = "參數改變: 空調開關模式設定";
                        break;
                    case 219:
                        sEventType = "參數改變: 節能驗證設定";
                        break;
                    case 220:
                        sEventType = "參數改變: 最運行時間設定";
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("EM3001.EventType", ex);
                throw ex;
            }
            return sEventType;
        }

        public string ControlMode(byte byData)
        {
            string sValue;
            try
            {
                switch (byData)
                {
                    case 0:
                        sValue = "All Open";
                        break;
                    case 2:
                        sValue = "All Close";
                        break;
                    case 4:
                        sValue = "Auto";
                        break;
                    default:
                        sValue = byData.ToString();
                        break;
                }
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("EM3001.SDState", ex);
                throw ex;
            }
            return sValue;
        }
    }
}
