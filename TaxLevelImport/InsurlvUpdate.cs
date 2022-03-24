using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
namespace TaxLevelImport
{
    public class InsurlvUpdate
    {
        System.Data.IDbConnection _conn;
        public InsurlvUpdate(System.Data.IDbConnection conn)
        {
            _conn = conn;
        }
        public bool ChecInsurlv()
        {
            JBTools.Security.SalaryEncrypt se = new JBTools.Security.SalaryEncrypt();
            SqlConnection conn = new SqlConnection(_conn.ConnectionString);
            string cmd = string.Format("SELECT * FROM INSURLV WHERE AMT={0}", 24000);
            string cmd1 = string.Format("UPDATE INSURLV SET LFF_DATEL = '{0}', LFF_DATEH = '{1}', LFF_DATER = '{2}', KEY_MAN = '{3}', KEY_DATE = '{4}' WHERE AMT ={5}",
                             "2021/12/31", "2021/12/31", "2021/12/31", "JB", DateTime.Now.ToString("yyyy-MM-dd hh: mm:ss"), 24000);
            //string cmd1 = string.Format("INSERT INTO INSURLV SELECT {0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',{9}",
            //                23800, "2020/01/01", "9999/12/31", "2020/01/01", "9999/12/31", "JB", DateTime.Now.ToString("yyyy-MM-dd hh: mm:ss"), "2020/01/01", "9999/12/31", 0.0469);
            //string cmd  = "SELECT * FROM INSURLV WHERE AMT=23100";
            //string cmd1 = "INSERT INTO INSURLV SELECT 23100,'2019/01/01','9999/12/31','2019/01/01','9999/12/31','JB','2019/01/3','2019/01/01','9999/12/31',0.0469";
            //string cmd2 = "UPDATE INSURLV SET LFF_DATEH='2017/12/31' WHERE AMT<22000 AND '2018/01/01' BETWEEN EFF_DATEH AND LFF_DATEH AND AMT!=0";
            //string cmd2 = "UPDATE INSURLV SET LFF_DATEL = '2017/12/31', LFF_DATEH='2017/12/31', LFF_DATER='2017/12/31' WHERE AMT IN (21009, 21900)";
            using (conn)
            {
                if (conn.State != ConnectionState.Open) conn.Open();
                try
                {
                    SqlCommand sqlcmd = new SqlCommand(cmd, conn);
                    var sql = sqlcmd.ExecuteScalar();
                    if (sql != null)
                    {
                        sqlcmd = new SqlCommand(cmd1, conn);
                        sqlcmd.ExecuteNonQuery();
                    }
                    //sqlcmd = new SqlCommand(cmd2, conn);
                    //sqlcmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    JBModule.Message.TextLog.path = @"C:\TEMP\Error\";
                    JBModule.Message.TextLog.WriteLog("執行更新時發生錯誤：" + ex.Message);
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 如果無此級距就新增投保級距
        /// </summary>
        /// <returns></returns>
        public bool addNewInsurlv(decimal SUPPLEMIN,double EFF_RATE,string updateYear)
        {
            string updateDate = new DateTime(int.Parse(updateYear),1,1).ToString("yyyy-MM-dd");
            JBTools.Security.SalaryEncrypt se = new JBTools.Security.SalaryEncrypt();
            SqlConnection conn = new SqlConnection(_conn.ConnectionString);
            string cmd = string.Format("SELECT * FROM INSURLV WHERE AMT={0}", SUPPLEMIN);
            string cmd1 = string.Format("INSERT INSURLV ([AMT], [EFF_DATEL], [LFF_DATEL], [EFF_DATEH], [LFF_DATEH], [KEY_MAN], [KEY_DATE], [EFF_DATER], [LFF_DATER], [EFF_RATE]) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')",
                             SUPPLEMIN, updateDate, "9999-12-31", updateDate, "9999-12-31", "JB", updateDate, updateDate, "9999-12-31", EFF_RATE);

            using (conn)
            {
                if (conn.State != ConnectionState.Open) conn.Open();
                try
                {
                    SqlCommand sqlcmd = new SqlCommand(cmd, conn);
                    var sql = sqlcmd.ExecuteScalar();
                    if (sql == null)
                    {
                        sqlcmd = new SqlCommand(cmd1, conn);
                        sqlcmd.ExecuteNonQuery();
                    }
                    //sqlcmd = new SqlCommand(cmd2, conn);
                    //sqlcmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    JBModule.Message.TextLog.path = @"C:\TEMP\Error\";
                    JBModule.Message.TextLog.WriteLog("執行投保級距新增時發生錯誤：" + ex.Message);
                    return false;
                }
            }
            return true;
        }
    }
}
