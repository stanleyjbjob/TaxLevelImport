using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
namespace TaxLevelImport
{
    public class LaborUpdate
    {
        System.Data.IDbConnection _conn;
        public LaborUpdate(System.Data.IDbConnection conn)
        {
            _conn = conn;
        }
        public int CheckLaborMax()
        {
            //JBModule.Data.CEncrypt.USER_ID = "JBADMIN";
            //JBModule.Data.CDecryp.USER_ID = "JBADMIN";
            JBTools.Security.SalaryEncrypt se = new JBTools.Security.SalaryEncrypt();
            SqlConnection conn = new SqlConnection(_conn.ConnectionString);
            //conn.ConnectionString = conn.ConnectionString;
            string cmd = "SELECT * FROM INSLAB WHERE '20160501' BETWEEN IN_DATE AND OUT_DATE AND FA_IDNO=''";
            string mode = "1";
            SqlDataAdapter sa = new SqlDataAdapter("SELECT * FROM INSLAB WHERE 1=0", conn);
            using (conn)
            {
                if (conn.State != ConnectionState.Open) conn.Open();
                SqlCommandBuilder scb = new SqlCommandBuilder(sa);
                var trans = conn.BeginTransaction();
                try
                {
                    SqlCommand sqlcmd = new SqlCommand("UPDATE INSURLV SET EFF_DATEL='20160501' WHERE AMT=45800", conn);
                    sqlcmd.Transaction = trans;
                    sqlcmd.ExecuteNonQuery();
                    sa.SelectCommand.Transaction = trans;
                    var dr = sa.SelectCommand.ExecuteReader();
                    DataTable dt = new DataTable();
                    dt.Load(dr);
                    if (dt.Columns.Contains("LBDATE"))
                    {
                        cmd = "SELECT * FROM INSLAB WHERE '20160501' BETWEEN LBDATE AND LEDATE AND FA_IDNO=''";
                        mode = "2";
                    }
                    string updateCommand = "", insertCommand = "", insertCommand1 = "";
                    updateCommand = "UPDATE INSLAB SET ";
                    insertCommand = "INSERT INTO INSLAB (";
                    insertCommand1 = " VALUES (";
                    SqlCommand insCmd = new SqlCommand();
                    SqlCommand updateCmd = new SqlCommand();
                    List<SqlParameter> parmList = new List<SqlParameter>();
                    List<SqlParameter> parmList1 = new List<SqlParameter>();
                    foreach (DataColumn dc in dt.Columns)
                    {
                        updateCommand += dc.ColumnName + "=@" + dc.ColumnName + ",";
                        insertCommand += dc.ColumnName + ",";
                        insertCommand1 += "@" + dc.ColumnName + ",";
                        SqlParameter parm = new SqlParameter();
                        parm.ParameterName = "@" + dc.ColumnName;
                        parm.SourceColumn = dc.ColumnName;
                        parmList.Add(parm);
                        SqlParameter parm1 = new SqlParameter();
                        parm1.ParameterName = "@" + dc.ColumnName;
                        parm1.SourceColumn = dc.ColumnName;
                        parmList1.Add(parm1);
                    }
                    insertCommand = insertCommand.Substring(0, insertCommand.Length - 1) + ")";
                    insertCommand1 = insertCommand1.Substring(0, insertCommand1.Length - 1) + ")";
                    updateCommand = updateCommand.Substring(0, updateCommand.Length - 1);
                    insCmd.CommandText = insertCommand + " " + insertCommand1;
                    updateCmd.CommandText = updateCommand + " WHERE NOBR=@NOBR AND FA_IDNO=@FA_IDNO AND IN_DATE=@IN_DATE";
                    if (mode == "2")
                        updateCmd.CommandText = updateCommand + " WHERE NOBR=@NOBR AND FA_IDNO=@FA_IDNO AND LBDATE=@LBDATE";
                    sa = new SqlDataAdapter(cmd, conn);
                    scb = new SqlCommandBuilder(sa);

                    sa.SelectCommand.Transaction = trans;
                    dr = sa.SelectCommand.ExecuteReader();
                    dt.Load(dr);
                    insCmd.Parameters.AddRange(parmList.ToArray());
                    updateCmd.Parameters.AddRange(parmList1.ToArray());
                    //sa.InsertCommand = scb.GetInsertCommand();
                    sa.InsertCommand = insCmd;
                    sa.UpdateCommand = updateCmd;
                    //sa.UpdateCommand = scb.GetUpdateCommand();

                    sa.InsertCommand.Transaction = trans;
                    sa.UpdateCommand.Transaction = trans;
                    List<DataRow> rows = new List<DataRow>();
                    int count = 0;
                    foreach (DataRow r in dt.Rows)
                    {
                        decimal r_amt = se.Decode(Convert.ToDecimal(r["r_amt"]));
                        decimal h_amt = se.Decode(Convert.ToDecimal(r["h_amt"]));
                        string nobr = r["nobr"].ToString();
                        if (r_amt > 43900M)
                        {
                            var ri = dt.NewRow();
                            //複製資料
                            ri.ItemArray = r.ItemArray;
                            bool IsNeedAddnew = false;
                            count++;
                            if (mode == "1")
                            {
                                if (Convert.ToDateTime(r["in_date"]) < new DateTime(2016, 5, 1))//如果是5/1就不新增一筆
                                {
                                    //忽略月中異動的可能性
                                    //上筆失效
                                    r["out_date"] = new DateTime(2016, 4, 30);
                                    if (dt.Columns.Contains("rout_date"))
                                        r["rout_date"] = new DateTime(2016, 4, 30);
                                    //下筆生效
                                    ri["in_date"] = new DateTime(2016, 5, 1);
                                    IsNeedAddnew = true;
                                }
                                else
                                    r["l_amt"] = se.Encode(45800M);
                            }
                            else if (mode == "2")
                            {
                                if (Convert.ToDateTime(r["LBDATE"]) < new DateTime(2016, 5, 1))//如果是5/1就不新增一筆
                                {
                                    //上筆失效
                                    //r["LBDATE"] = new DateTime(2016, 4, 30);
                                    r["LEDATE"] = new DateTime(2016, 4, 30);
                                    //下筆生效
                                    ri["LBDATE"] = new DateTime(2016, 5, 1);
                                    IsNeedAddnew = true;
                                }
                                else
                                    r["l_amt"] = se.Encode(45800M);
                            }
                            if (IsNeedAddnew)
                            {
                                ri["l_amt"] = se.Encode(45800M);
                                ri["code"] = "2";//調整
                                ri["key_date"] = DateTime.Now;
                                ri["key_man"] = "JB";
                                rows.Add(ri);
                            }
                        }
                        else if (r_amt == 0 && h_amt > 43900M)
                        {
                            var ri = dt.NewRow();
                            //複製資料
                            ri.ItemArray = r.ItemArray;
                            bool IsNeedAddnew = false;
                            count++;
                            if (mode == "1")
                            {
                                if (Convert.ToDateTime(r["in_date"]) < new DateTime(2016, 5, 1))//如果是5/1就不新增一筆
                                {
                                    //忽略月中異動的可能性
                                    //上筆失效
                                    r["out_date"] = new DateTime(2016, 4, 30);
                                    if (dt.Columns.Contains("rout_date"))
                                        r["rout_date"] = new DateTime(2016, 4, 30);
                                    //下筆生效
                                    ri["in_date"] = new DateTime(2016, 5, 1);
                                    IsNeedAddnew = true;
                                }
                                else
                                    r["l_amt"] = se.Encode(45800M);
                            }
                            else if (mode == "2")
                            {
                                if (Convert.ToDateTime(r["LBDATE"]) < new DateTime(2016, 5, 1))//如果是5/1就不新增一筆
                                {
                                    //上筆失效
                                    //r["LBDATE"] = new DateTime(2016, 4, 30);
                                    r["LEDATE"] = new DateTime(2016, 4, 30);
                                    //下筆生效
                                    ri["LBDATE"] = new DateTime(2016, 5, 1);
                                    IsNeedAddnew = true;
                                }
                                else
                                    r["l_amt"] = se.Encode(45800M);
                            }
                            if (IsNeedAddnew)
                            {
                                ri["l_amt"] = se.Encode(45800M);
                                ri["code"] = "2";//調整
                                ri["key_date"] = DateTime.Now;
                                ri["key_man"] = "JB";
                                rows.Add(ri);
                            }
                        }
                    }
                    foreach (var it in rows)
                        dt.Rows.Add(it);
                    sa.Update(dt);
                    SqlCommand updatecmd = new SqlCommand("UPDATE INSLAB SET  KEY_DATE=DATEADD(millisecond,-1*DATEPART(millisecond, KEY_DATE), KEY_DATE) "
                        + "  WHERE DATEPART(millisecond, KEY_DATE)>0 and KEY_DATE>='2016/05/01'", conn);
                    updatecmd.Transaction = trans;
                    updatecmd.ExecuteNonQuery();
                    trans.Commit();
                    return count;
                }
                catch (Exception ex)
                {
                    JBModule.Message.TextLog.path = @"C:\TEMP\Error\";
                    JBModule.Message.TextLog.WriteLog("執行更新時發生錯誤：" + ex.Message);
                    return -1;
                }
            }
        }
    }
}
