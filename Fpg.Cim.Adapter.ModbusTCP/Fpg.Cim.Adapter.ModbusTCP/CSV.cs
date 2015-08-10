using System;
using System.Collections.Generic;
using System.Text;

using System.Data;
using NLog;
using System.IO;
using System.Configuration;

namespace Fpg.Cim.Adapter.CSV
{
    class CSV
    {
        private string g_sCSVFilePath = ConfigurationManager.AppSettings.Get("CSVFilePath");
        private static Logger g_oLogger = NLog.LogManager.GetCurrentClassLogger();

        class Parameter
        {
            public DataTable dttRealTimeData;
            public string sFilePath = string.Empty;
            public StreamWriter sw;
            public FileStream fs;
        }

        public void WriteToCSV(DataTable dttRealTime, string sFilePath)
        {
            Parameter objParameter = new Parameter();
            objParameter.dttRealTimeData = dttRealTime.Copy();
            objParameter.sFilePath = @g_sCSVFilePath + sFilePath + ".csv";
            try
            {
                if (CheckFile(objParameter))
                {
                    objParameter.sw = new StreamWriter(objParameter.fs, Encoding.Default);
                    int intColCount = objParameter.dttRealTimeData.Columns.Count;
                    if (objParameter.fs.Length == 0)
                    {
                        if (objParameter.dttRealTimeData.Columns.Count > 0)
                            objParameter.sw.Write(objParameter.dttRealTimeData.Columns[0]);
                        for (int i = 1; i < objParameter.dttRealTimeData.Columns.Count; i++)
                            objParameter.sw.Write("," + objParameter.dttRealTimeData.Columns[i]);
                        objParameter.sw.Write(objParameter.sw.NewLine);
                    }
                    
                    foreach (DataRow dtw in objParameter.dttRealTimeData.Rows)
                    {
                        if (objParameter.dttRealTimeData.Columns.Count > 0 && !Convert.IsDBNull(dtw[0]))
                            objParameter.sw.Write(Convert.ToString(dtw[0]));
                        for (int i = 1; i < intColCount; i++)
                            objParameter.sw.Write("," + Convert.ToString(dtw[i]));
                        objParameter.sw.Write(objParameter.sw.NewLine);
                    }
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("WriteToCSV-- ", ex);
            }
            finally
            {
                if (objParameter.sw != null) { objParameter.sw.Flush(); objParameter.sw.Close(); objParameter.sw.Dispose(); objParameter.sw = null; }
                if (objParameter.fs!= null) { objParameter.fs.Close(); objParameter.fs.Dispose(); objParameter.fs = null; }
                if (objParameter.dttRealTimeData != null) { objParameter.dttRealTimeData.Clear(); objParameter.dttRealTimeData.Dispose(); objParameter.dttRealTimeData = null; }
            }
        }

        private bool CheckFile(Parameter objParameter)
        {
            bool blExistence = false;
            try
            {
                if (File.Exists(objParameter.sFilePath))
                {
                    objParameter.fs = new FileStream(objParameter.sFilePath, FileMode.Append);
                    blExistence = true;
                }
                else
                {
                    objParameter.fs = new FileStream(objParameter.sFilePath, FileMode.Create);
                    blExistence = true;
                }
            }
            catch (Exception ex)
            {
                g_oLogger.ErrorException("CheckFile-- ", ex);
            }
            return blExistence;
        }
    }
}
