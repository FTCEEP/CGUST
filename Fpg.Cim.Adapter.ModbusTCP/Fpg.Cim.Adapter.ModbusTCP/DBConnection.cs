using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;

using NLog;

class DBConnection
{
    private static Logger gLogger = NLog.LogManager.GetCurrentClassLogger(); // Nlog

    /// <summary>
    /// 
    /// </summary>
    public DBConnection()
    {

    }
    /// <summary>
    /// 下SQL Command取回資料
    /// </summary>
    /// <param name="sSqlCmd">SQL Commnad String</param>
    /// <param name="sConnStr1">SQL Connection String</param>
    /// <returns>回傳經由SQL Command Select後的結果，傳回DataTable</returns>
    public DataTable execDataTable(string sSqlCmd, string sConnStr1)
    {
        SqlConnection conDB = new SqlConnection();
        SqlDataAdapter sqlDataAdap = new SqlDataAdapter();
        DataTable dttReturnData = new DataTable();
        try
        {
            conDB = new SqlConnection(sConnStr1);
            sqlDataAdap = new SqlDataAdapter(sSqlCmd, conDB);
            sqlDataAdap.Fill(dttReturnData);
        }
        catch (Exception ex)
        {
            gLogger.ErrorException("DBConnection.execDataTable", ex);
            throw ex;
        }
        finally
        {
            if (sqlDataAdap != null)
            {
                sqlDataAdap.Dispose();
            }
            if (conDB != null)
            {
                if (conDB.State == ConnectionState.Open)
                {
                    conDB.Close();
                }
                conDB.Dispose();
            }
        }
        return dttReturnData;
    }

    /// <summary>
    /// 下SQL Command輸入、更新或刪除資料
    /// </summary>
    /// <param name="iSqlCmd">SQL Commnad String</param>
    /// <param name="sConnStr1">SQL Connection String</param>
    /// <returns>受影響的資料列數</returns>
    public int execNonQuery(string iSqlCmd, string sConnStr1)
    {
        SqlConnection myConn = new SqlConnection();
        SqlCommand myCmd = new SqlCommand();
        int retAffectedRows;

        try 
        {	       
            //string myConnStr2 = "";
            //if (sConnStr1.ToString().Length == 0)	
            //{
            //    string sDefaultConn = System.Web.Configuration.WebConfigurationManager.AppSettings["WebDB"].ToString();
            //    myConnStr2 = System.Web.Configuration.WebConfigurationManager.ConnectionStrings[sDefaultConn].ConnectionString;
            //    myConn = new SqlConnection(myConnStr2);
            //}
            //else
            //{
            //    myConn = new SqlConnection(sConnStr1);
            //}
            myConn = new SqlConnection(sConnStr1);
            myConn.Open();
            myCmd = new SqlCommand(iSqlCmd, myConn);
            retAffectedRows = myCmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            gLogger.ErrorException("DBConnection.execNonQuery", ex);
            throw ex;
        }
        finally
        {
            myConn.Close();
            if (myCmd != null)
            {
		        myCmd.Dispose();
            }
            if (myConn != null)
            {
		        if (myConn.State == ConnectionState.Open)
                {
		            myConn.Close();
                }
                myConn.Dispose();
            }
        }
            return retAffectedRows;
    }
}
