using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Deployment.Application;
namespace TaxLevelImport
{
    public partial class Form1 : Form
    {

        bool IsFoxpro = false;
        string YEAR = "2022";
        string FileName = "Tax2022.xlsx";//所得稅級距表檔名
        string Dir = "";
        string MainFile = "Main.ini";
        int AMT_L = 25250;//自2022年(民國111年)1月1日起，基本工資月薪將調漲為25,250元，時薪調漲為168元
        int SUPPLEMIN = 25250;//基本工資自2022年(民國111年)1月1日起從每月24,000元調整為25,250元，「勞工保險投保薪資分級表」配合基本工資調整同步修正。原分級表第1級24,0000元刪除，原第2級25,250元遞移為第1級，其餘投保薪資金額均未修正並依序遞移，修正後分級表為14級。
        double LastSUPPLEINSLABRATE = 0.0191;//調整前補充保險費率
        double ThisSUPPLEINSLABRATE = 0.0211;//調整後補充保險費率，自2021年(民國110年)1月1日起由1.91%調漲為2.11%
        double compersoncnt = 1.58;//平均眷口數，自2020年(民國109)年1月1日起調整平均眷口數為0.58人
        double LastNORMALRATE = 0.1;//調整前勞工保險普通事故保險費率
        double ThisNORMALRATE = 0.105;//調整後勞工保險普通事故保險費率，自2021年(民國110年)1月1日起由10%調整為10.5%
        double EFF_RATE = 0.0517;//公司健保費率，自2021年(民國110年)起費率調整為5.17%
        int FORSALBASD = 37875;//因配合基本薪資調整，非居住者基本工資1.5倍金額由36000->37875
        List<string> Supplemin_M_FORMAT = new List<string>() { "50" };
        int SUPPLEMIN1 = 20000;//單次給付金達20000元
        List<string> Supplemin_M_FORMAT1 = new List<string>() { "9A", "9B", "51", "52", "54", "5A", "5B", "5C" };

        DateTime ExpireDate = new DateTime(2022, 12, 31);//效期(避免提早執行)
        public Form1()
        {
            InitializeComponent();
        }
        [STAThread]
        private void Form1_Load(object sender, EventArgs e)
        {
            //if (ApplicationDeployment.IsNetworkDeployed == true)
            //{
            //    ApplicationDeployment thisDeployment = ApplicationDeployment.CurrentDeployment;
            //    this.Text = "檢查更新中...";
            //    if (thisDeployment.CheckForUpdate() == true)
            //    {
            //        if (MessageBox.Show("偵測到有新版本，是否需要更新?", "是否更新", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            //        {
            //            this.Text = "更新中...";
            //            thisDeployment.Update();
            //            MessageBox.Show("更新完成，需要重新啟動程式!");
            //            Application.Restart();
            //        }
            //        else
            //        {
            //            this.Text = Application.ProductName + " " + Application.ProductVersion;
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show("目前版本已為最新版!");
            //    }
            //}
            //else
            //    MessageBox.Show("此版本無網路更新");

            Dir = Directory.GetCurrentDirectory();
            string constr = GetConnectionString();
            Application.DoEvents();
            if (constr.Trim().Length == 0)
            {
                RefreshState("無法取得連線資訊", 1);
                return;
            }
            BW.RunWorkerAsync(constr);
            //BW_DoWork(null,null);
        }
        public static string StateMsg = "";
        public string GetConnectionString()
        {
            JBModule.Message.TextLog.path = @"C:\TEMP\Error\";

            string Dir = Directory.GetCurrentDirectory();
            //string FileName = "Tax2016.xls";
            //string FullName = Dir + @"\" + FileName;
            //string MainFile = "";
            DateTime ExpireDate = new DateTime(2017, 5, 31);
            string Server = "", DataBase = "", Id = "jb", Pwd = "JB8421";
            //string Server = "", DataBase = "", Id = "jb", Pwd = "JB8421";
            try
            {
                RefreshState("檢查HR程式版本", 10);
                //bool IsDoNet = false;
                //if (System.IO.File.Exists(@"C:\Temp\App.config"))
                //    IsDoNet = true;

                #region dotNet版本

                {
                    #region 讀取JBHR的connectionString
                    string connectionString = "";
                    try
                    {
                        //System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                        string strPath = @"temp\App.Config";
                        bool isFound = false;
                        string[] drives = Directory.GetLogicalDrives();
                        foreach (string drive in drives)
                        {
                            if (File.Exists(drive + strPath))
                            {
                                StreamReader sr = File.OpenText(drive + strPath);
                                string readStr = sr.ReadToEnd();
                                sr.Close();
                                string[] pp = readStr.Split(new string[] { "connectionString=\"", "\" />" }, StringSplitOptions.RemoveEmptyEntries);

                                foreach (var item in pp)
                                {
                                    if (item.Substring(0, 12) == "Data Source=")
                                    {
                                        connectionString = JBModule.Data.CDecryp.ConnectString(item);
                                        isFound = true;
                                        break;
                                    }
                                }
                                break;
                            }
                        }
                        if (isFound)
                        {
                            RefreshState("JBHR系統", 14);
                            SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder(connectionString);
                            Server = sb.DataSource;
                            DataBase = sb.InitialCatalog;
                            Id = sb.UserID;
                            Pwd = sb.Password;
                            //RefreshState("主機為：" + Server);
                        }
                        else
                        {
                            //RefreshState("找不到相關程式連線字串", 14);
                            //return;
                        }
                    }
                    catch (Exception ex)
                    {
                        //RefreshState("找不到相關程式連線字串", 14);
                        //return;
                    }

                    #endregion
                    #region 讀取dotNet的connectionString
                    if (connectionString.Length == 0)//耐落正常會走這一個連線規則
                    {
                        try
                        {
                            //System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                            bool isFound = false;
                            string[] drives = Directory.GetLogicalDrives();
                            //依序掃描磁碟跟目錄搜尋JBHR.CONNECTION.STR，來判斷是否有連線資訊
                            foreach (string drive in drives)
                            {
                                if (File.Exists(drive + "JBHR.CONNECTION.STR"))
                                {
                                    StreamReader sr = File.OpenText(drive + "JBHR.CONNECTION.STR");
                                    string readStr = JBModule.Data.CDecryp.Text(sr.ReadLine());
                                    sr.Close();
                                    string[] pp = readStr.Split(new string[] { "[=]" }, StringSplitOptions.RemoveEmptyEntries);
                                    if (pp.Length == 2)
                                    {
                                        connectionString = pp[1];
                                        isFound = true;
                                        break;
                                    }
                                }
                            }
                            if (!isFound)
                            {
                                foreach (string drive in drives)
                                {
                                    if (!File.Exists(drive + "JBHR.CONNECTION.STR"))
                                    {
                                        try
                                        {
                                            File.Copy(Directory.GetCurrentDirectory() + "JBHR.CONNECTION.STR", drive + "JBHR.CONNECTION.STR");
                                        }
                                        catch (Exception ex)
                                        {
                                            continue;
                                        }

                                        StreamReader sr = File.OpenText(drive + "JBHR.CONNECTION.STR");
                                        string readStr = JBModule.Data.CDecryp.Text(sr.ReadLine());
                                        sr.Close();
                                        string[] pp = readStr.Split(new string[] { "[=]" }, StringSplitOptions.RemoveEmptyEntries);
                                        if (pp.Length == 2)
                                        {
                                            connectionString = pp[1];
                                            isFound = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (isFound)
                            {
                                RefreshState("MircroSoft.Net", 14);
                                SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder(connectionString);
                                Server = sb.DataSource;
                                DataBase = sb.InitialCatalog;
                                Id = sb.UserID;
                                Pwd = sb.Password;
                                //RefreshState("主機為：" + Server);
                            }
                        }
                        catch (Exception ex)
                        {
                            RefreshState("找不到相關程式連線字串", 14);
                            return "";
                        }
                    }
                    #endregion
                    #region FoxPro版本
                    if (connectionString.Trim().Length == 0)//還是找不到
                    {
                        OpenFileDialog ofd = new OpenFileDialog();
                        ofd.Filter = "*.ini|*.ini";
                        MessageBox.Show("請選擇Main.ini的檔案位置");
                        if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            //Dir = ofd.FileName;
                            //Dir = pDir.FullName;
                            MainFile = ofd.FileName;
                            if (File.Exists(MainFile))
                            {
                                RefreshState("Visual FoxPro", 12);
                                IsFoxpro = true;
                                StreamReader sr = File.OpenText(MainFile);
                                //string host = "", db = "";
                                while (sr.Peek() > 0)
                                {
                                    string readStr = sr.ReadLine();
                                    if (readStr.IndexOf("SERVER=") != -1) Server = readStr.Split('=')[1];
                                    if (readStr.IndexOf("DATABASE=") != -1) DataBase = readStr.Split('=')[1];
                                }
                                sr.Close();
                            }
                        }
                    }
                    #endregion
                }
                #endregion

                #region 取得連線
                SqlConnectionStringBuilder sbb = new SqlConnectionStringBuilder();
                sbb.DataSource = Server.Trim();
                sbb.InitialCatalog = DataBase.Trim();
                sbb.UserID = Id.Trim();
                sbb.Password = Pwd.Trim();

                HrDBDataContext db = new HrDBDataContext();
                db.Connection.ConnectionString = sbb.ConnectionString;
                if (DataBase.Trim().Length == 0) db.Connection.ConnectionString = "";

                try
                {
                    db.Connection.Open();
                }
                catch (Exception ex)
                {
                    JBHR.U_SETDB frm = new JBHR.U_SETDB();
                    JBHR.U_SETDB.Server = Server;
                    JBHR.U_SETDB.DataBase = DataBase;
                    JBHR.U_SETDB.UserId = Id;
                    JBHR.U_SETDB.PassWord = Pwd;
                    frm.ShowDialog();

                    sbb = new SqlConnectionStringBuilder();
                    sbb.DataSource = JBHR.U_SETDB.Server;
                    sbb.InitialCatalog = JBHR.U_SETDB.DataBase;
                    sbb.UserID = JBHR.U_SETDB.UserId;
                    sbb.Password = JBHR.U_SETDB.PassWord;
                    if (db.Connection.State == ConnectionState.Open) db.Connection.Close();
                    db.Connection.ConnectionString = sbb.ConnectionString;
                    try
                    {
                        db.Connection.Open();
                    }
                    catch (Exception ex1)
                    {
                        RefreshState("連線失敗", 22);
                        JBModule.Message.TextLog.WriteLog(ex1);
                        return "";
                    }
                }
                #endregion
                RefreshState("主機位置為" + sbb.DataSource, 16);
                RefreshState("資料庫位置為" + sbb.InitialCatalog, 16);
                return sbb.ConnectionString;
            }
            catch (Exception ex)
            {
                RefreshState("程式發生錯誤", toolStripProgressBar1.Value);
                RefreshState(ex.Message, toolStripProgressBar1.Value);
                JBModule.Message.TextLog.WriteLog(ex);
            }
            return "";
        }
        private void BW_DoWork(object sender, DoWorkEventArgs e)
        {
            string Title = "1. 111年所得稅級距表  2. 保險級距更新  3.非居住者基本工資";
            //string Title = "1. 111年所得稅級距表  2. 50格式起扣點  3. 勞保普通事故保險費率  4. 保險級距更新  5. 健保費率  6. 補充保費費率";
            string version = "非安裝版本";
            JBModule.Message.TextLog.path = System.IO.Directory.GetCurrentDirectory() + @"\";
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                version = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            if (MessageBox.Show("目前自動更新程式版本為：" + version + Environment.NewLine + "此次更新內容為：" + Title + Environment.NewLine + "是否要執行?", "提示", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
            {
                //bool IsFoxpro = false;
                string FullName = Dir + @"\" + FileName;
                //string MainFile = "Main.ini";
               
                string Server = "", DataBase = "", Id = "jb", Pwd = "JB8421";
                try
                {
                    RefreshState("檢查程式是否有效，有效日期至" + ExpireDate.ToString("yyyy/MM/dd"), 2);
                    if (DateTime.Today <= ExpireDate)
                        RefreshState("OK", 4);
                    else
                    {
                        RefreshState("程式不可使用，請洽傑報資訊取得最新修正程式", 4);
                        return;
                    }

                    HrDBDataContext db = new HrDBDataContext(e.Argument.ToString());
                    var ds = new DataSet();
                    var dt = new DataTable();
                    if (ConfigSetting.AppSettingValue("NeedTaxLvl") == "1")
                    {
                        RefreshState("檢查檔案是否存在", 6);

                        if (File.Exists(FullName))
                        {
                            RefreshState("OK", 8);
                        }
                        else
                        {
                            RefreshState("找不到檔案" + Environment.NewLine + "(" + FullName + ")", 8);
                            return;
                        }


                        #region 讀取Excel檔

                        if (ConfigSetting.AppSettingValue("NeedTaxLvl") == "1")
                        {
                            RefreshState("讀取Excel檔案", 18);
                            ds = JBModule.Data.CNPOI.ReadExcelToDataSet(FullName, JBTools.IO.LoadExcelColumnNameStyle.DefinedColumn);
                            dt = ds.Tables[0];
                            RefreshState("OK", 20);
                        }
                    }
                        #endregion
                    if(true)//小狐狸必執行此檔做初始化
                    {
                        try
                        {
                            string DeleteCommand = "DELETE  FROM TAXLVL WHERE SUBSTRING(YEAR,3,1)=' '";
                            db.ExecuteCommand(DeleteCommand, new object[] { });
                        }
                        catch { }
                    }
                    
                    #region 更新所得稅級距表
                    if (ConfigSetting.AppSettingValue("NeedTaxLvl") == "1")
                    {
                        var _YEAR = YEAR;
                        if (IsFoxpro)
                        {
                            _YEAR = (Convert.ToInt32(YEAR) - 1911).ToString();
                        }
                        RefreshState("清空相同年度級距表", 24);
                        string DeleteCommand = "DELETE TAXLVL WHERE YEAR=" + _YEAR;
                        db.ExecuteCommand(DeleteCommand, new object[] { });
                        RefreshState("OK", 26);

                        RefreshState("開始匯入資料", 28);
                        int i = 0, count = dt.Rows.Count;
                        foreach (DataRow it in dt.Rows)
                        {
                            i++;
                            BW.ReportProgress(i * 60 / count + 30);
                            if (IsFoxpro)
                            {
                                TAXLVL1 r = new TAXLVL1();
                                r.AMT_H = Convert.ToDouble(it["AMT_H"]);
                                r.AMT_L = Convert.ToDouble(it["AMT_L"]);
                                r.KEY_DATE = Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                                r.KEY_MAN = "JB";
                                r.PER0 = Convert.ToDouble(it["per0"]);
                                r.PER1 = Convert.ToDouble(it["per1"]);
                                r.PER2 = Convert.ToDouble(it["per2"]);
                                r.PER3 = Convert.ToDouble(it["per3"]);
                                r.PER4 = Convert.ToDouble(it["per4"]);
                                r.PER5 = Convert.ToDouble(it["per5"]);
                                r.PER6 = Convert.ToDouble(it["per6"]);
                                r.PER7 = Convert.ToDouble(it["per7"]);
                                r.PER8 = Convert.ToDouble(it["per8"]);
                                r.PER9 = Convert.ToDouble(it["per9"]);
                                r.PER10 = Convert.ToDouble(it["per10"]);
                                //r.PER11 = Convert.ToDecimal(it["per11"]);
                                r.YEAR = _YEAR;
                                db.TAXLVL1.InsertOnSubmit(r);
                            }
                            else
                            {
                                TAXLVL r = new TAXLVL();
                                r.AMT_H = Convert.ToDecimal(it["AMT_H"]);
                                r.AMT_L = Convert.ToDecimal(it["AMT_L"]);
                                r.KEY_DATE = Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                                r.KEY_MAN = "JB";
                                r.PER0 = Convert.ToDecimal(it["per0"]);
                                r.PER1 = Convert.ToDecimal(it["per1"]);
                                r.PER2 = Convert.ToDecimal(it["per2"]);
                                r.PER3 = Convert.ToDecimal(it["per3"]);
                                r.PER4 = Convert.ToDecimal(it["per4"]);
                                r.PER5 = Convert.ToDecimal(it["per5"]);
                                r.PER6 = Convert.ToDecimal(it["per6"]);
                                r.PER7 = Convert.ToDecimal(it["per7"]);
                                r.PER8 = Convert.ToDecimal(it["per8"]);
                                r.PER9 = Convert.ToDecimal(it["per9"]);
                                r.PER10 = Convert.ToDecimal(it["per10"]);
                                r.PER11 = Convert.ToDecimal(it["per11"]);
                                r.YEAR = _YEAR;
                                db.TAXLVL.InsertOnSubmit(r);
                            }
                        }
                        db.SubmitChanges();
                        RefreshState("級距表匯入完成", 40);
                        string UpdateTaxLvCode = "UPDATE TAXLVL SET AMT_H=9999999999 WHERE AMT_H=500000.00 AND YEAR=" + _YEAR;
                        RefreshState("修改所得稅最大級距", 45);

                        try
                        {
                            var cc = db.ExecuteCommand(UpdateTaxLvCode, new object[] { });
                            RefreshState("修改所得稅最大級距完成，共更新" + cc.ToString() + "筆資料", 47);
                        }
                        catch (Exception eex)
                        {
                            RefreshState("修改所得稅最大級距失敗", 47);
                            RefreshState(eex.Message, 48);
                            JBModule.Message.TextLog.WriteLog(eex);
                        }

                        if (!IsFoxpro)
                        {
                            string Update50FORMAT = string.Format("UPDATE YRFORMAT SET SUPPLEMIN = {0} where M_FORMAT='50'", SUPPLEMIN);
                            string Update50FORMAT_FAX = string.Format("UPDATE YRFORMAT SET AMT_L = {0} where M_FORMAT='50'", AMT_L);
                            RefreshState("修改50格式起扣點", 45);

                            try
                            {
                                int cc = 0;
                                try
                                {
                                    cc = db.ExecuteCommand(Update50FORMAT_FAX, new object[] { });
                                }
                                catch
                                {
                                    cc = db.ExecuteCommand(Update50FORMAT, new object[] { });
                                }
                                RefreshState("修改50格式起扣點完成，共更新" + cc.ToString() + "筆資料", 47);
                            }
                            catch (Exception eex)
                            {
                                RefreshState("修改50格式起扣點失敗", 47);
                                RefreshState(eex.Message, 48);
                                JBModule.Message.TextLog.WriteLog(eex);
                            }

                            //只有AP要更新
                            try
                            {
                                string UpdateFIXRATE = string.Format("UPDATE YRFORMAT SET FIXRATE = {0}", ThisSUPPLEINSLABRATE);
                                var cd = db.ExecuteCommand(UpdateFIXRATE, new object[] { });
                                RefreshState(string.Format("修改補充保費費率為{0}", ThisSUPPLEINSLABRATE), 48);
                                RefreshState(string.Format("修改補充保費費率為{0}完成，共更新" + cd.ToString() + "筆資料", ThisSUPPLEINSLABRATE), 49);
                            }
                            catch { } 
                        }

                    }

                    int percent = 50;
                    #endregion
                    #region 更新補充保費扣繳門檻
                    if (ConfigSetting.AppSettingValue("Supplemin") == "1")
                    {
                        if (IsFoxpro)
                        {
                            string Sqlstring = "IN (";
                            foreach (var item in Supplemin_M_FORMAT)
                                Sqlstring += string.Format("'{0}',", item.ToString());
                            Sqlstring = Sqlstring.Remove(Sqlstring.Length - 1) + ")";
                            string UpdateHarCodeSys = string.Format("update YRFOMAT set AMT_L = {0} where M_FORMAT {1}", SUPPLEMIN, Sqlstring);
                            Sqlstring = "IN (";
                            foreach (var item in Supplemin_M_FORMAT1)
                                Sqlstring += string.Format("'{0}',", item.ToString());
                            Sqlstring = Sqlstring.Remove(Sqlstring.Length - 1) + ")";
                            string UpdateHarCodeSys1 = string.Format("update YRFOMAT set AMT_L = {0} ,AMT_H = 10000000 where M_FORMAT {1}", SUPPLEMIN1, Sqlstring);
                            RefreshState(string.Format("更新補充保費扣繳門檻( {0} )", SUPPLEMIN), percent += 2);
                            try
                            {
                                var cc = db.ExecuteCommand(UpdateHarCodeSys, new object[] { });
                                cc += db.ExecuteCommand(UpdateHarCodeSys1, new object[] { });
                                RefreshState("更新補充保費扣繳門檻，共更新" + cc.ToString() + "筆資料", percent += 2);
                            }
                            catch (Exception eex)
                            {
                                RefreshState("更新補充保費扣繳門檻失敗", percent += 2);
                                RefreshState(eex.Message, percent);
                                JBModule.Message.TextLog.WriteLog(eex);
                            }
                        }
                        else
                        {
                            string Sqlstring = "IN (";
                            foreach (var item in Supplemin_M_FORMAT)
                                Sqlstring += string.Format("'{0}',", item.ToString());
                            Sqlstring = Sqlstring.Remove(Sqlstring.Length - 1) + ")";
                            string UpdateHarCodeSys = string.Format("update YRFORMAT set Supplemin = {0} where M_FORMAT {1}", SUPPLEMIN, Sqlstring);
                            Sqlstring = "IN (";
                            foreach (var item in Supplemin_M_FORMAT1)
                                Sqlstring += string.Format("'{0}',", item.ToString());
                            Sqlstring = Sqlstring.Remove(Sqlstring.Length - 1) + ")";
                            string UpdateHarCodeSys1 = string.Format("update YRFORMAT set Supplemin = {0},SuppleMAX = 10000000 where M_FORMAT {1}", SUPPLEMIN1, Sqlstring);
                            RefreshState(string.Format("更新補充保費扣繳門檻( {0} )", SUPPLEMIN), percent += 2);
                            try
                            {
                                var cc = db.ExecuteCommand(UpdateHarCodeSys, new object[] { });
                                cc += db.ExecuteCommand(UpdateHarCodeSys1, new object[] { });
                                RefreshState("更新補充保費扣繳門檻，共更新" + cc.ToString() + "筆資料", percent += 2);
                            }
                            catch (Exception eex)
                            {
                                RefreshState("更新補充保費扣繳門檻失敗", percent += 2);
                                RefreshState(eex.Message, percent);
                                JBModule.Message.TextLog.WriteLog(eex);
                            }
                        }
                    }
                    #endregion
                    #region 更新補充保費率
                    if (ConfigSetting.AppSettingValue("SupplementaryPre") == "1")
                    {
                        if (IsFoxpro)
                        {
                            //string UpdateHarCodeSys = "ALTER TABLE YRFOMAT ALTER COLUMN TAXRATE decimal(18, 4)";
                            //string UpdateHarCodeSys2 = "update YRFOMAT set TAXRATE = 0.0191 where TAXRATE = 0.02";
                            //RefreshState("更新補充保費率( 2% -> 1.91% )", percent += 2);
                            //try
                            //{
                            //    var cc = db.ExecuteCommand(UpdateHarCodeSys, new object[] { });
                            //    cc = db.ExecuteCommand(UpdateHarCodeSys2, new object[] { });
                            //    RefreshState("更新補充保費率，共更新" + cc.ToString() + "筆資料", percent += 2);
                            //}
                            //catch (Exception eex)
                            //{
                            //    RefreshState("更新補充保費率失敗", percent += 2);
                            //    RefreshState(eex.Message, percent);
                            //    JBModule.Message.TextLog.WriteLog(eex);
                            //}
                        }
                        else
                        {
                            string UpdateHarCodeSys = string.Format("update U_SYS5 set SUPPLEINSLABRATE = {0}", ThisSUPPLEINSLABRATE);
                            string UpdateHarCodeSys2 = string.Format("update YRFORMAT set FIXRATE = {0} where FIXRATE = {1}", ThisSUPPLEINSLABRATE, LastSUPPLEINSLABRATE);
                            RefreshState(string.Format("更新補充保費率( {0}% -> {1}% )", LastSUPPLEINSLABRATE * 100, ThisSUPPLEINSLABRATE * 100), percent += 2);
                            try
                            {
                                var cc = db.ExecuteCommand(UpdateHarCodeSys, new object[] { });
                                cc += db.ExecuteCommand(UpdateHarCodeSys2, new object[] { });
                                RefreshState("更新補充保費率，共更新" + cc.ToString() + "筆資料", percent += 2);
                            }
                            catch (Exception eex)
                            {
                                RefreshState("更新補充保費率失敗", percent += 2);
                                RefreshState(eex.Message, percent);
                                JBModule.Message.TextLog.WriteLog(eex);
                            }
                        }
                    }
                    #endregion
                    #region 更新平均眷口數
                    if (ConfigSetting.AppSettingValue("AveragePerNo") == "1")
                    {
                        if (!IsFoxpro)
                        {
                            string UpdateHarCodeSys = string.Format("update u_sys5 set compersoncnt={0}", compersoncnt);
                            RefreshState(string.Format("更新平均眷口數( {0}% )", compersoncnt), percent += 2);
                            try
                            {
                                var cc = db.ExecuteCommand(UpdateHarCodeSys, new object[] { });
                                RefreshState("更新平均眷口數，共更新" + cc.ToString() + "筆資料", percent += 2);
                            }
                            catch (Exception eex)
                            {
                                RefreshState("更新平均眷口數失敗", percent += 2);
                                RefreshState(eex.Message, percent);
                                JBModule.Message.TextLog.WriteLog(eex);
                            }
                        }
                    }
                    #endregion
                    #region 更新勞保普通費率

                    if (ConfigSetting.AppSettingValue("LaborIns") == "1")
                    {
                        //if (!IsFoxpro)
                        {
                            string UpdateLarCode = string.Format("update LARCODE set NORMALRATE={0} where NORMALRATE={1}", ThisNORMALRATE, LastNORMALRATE);
                            RefreshState(string.Format("更新勞保普通費率( {0}% => {1}% )", LastNORMALRATE * 100, ThisNORMALRATE * 100), percent += 2);
                            try
                            {
                                var cc = db.ExecuteCommand(UpdateLarCode, new object[] { });
                                RefreshState("更新勞保普通費率完成，共更新" + cc.ToString() + "筆資料", percent += 2);
                            }
                            catch (Exception eex)
                            {
                                RefreshState("更新勞保普通費率失敗", percent += 2);
                                RefreshState(eex.Message, percent);
                                JBModule.Message.TextLog.WriteLog(eex);
                            }
                        }
                    }
                    #endregion
                    #region 更新公司健保費率
                    if (ConfigSetting.AppSettingValue("UpdateHealthRate") == "1")
                    {

                        string UpdateHarCode = string.Format("update INSURLV set EFF_RATE={0}", EFF_RATE);
                        RefreshState(string.Format("更新健保費率( {0}% )", EFF_RATE * 100), percent += 2);
                        try
                        {
                            var cc = db.ExecuteCommand(UpdateHarCode, new object[] { });
                            RefreshState("更新健保費率完成，共更新" + cc.ToString() + "筆資料", percent += 2);
                        }
                        catch (Exception eex)
                        {
                            RefreshState("更新健保費率失敗", percent += 2);
                            RefreshState(eex.Message, percent);
                            JBModule.Message.TextLog.WriteLog(eex);
                        }
                        if (!IsFoxpro)
                        {
                            string UpdateHarCodeSys = string.Format("update U_SYS5 set HEACOMPRATE={0}", EFF_RATE);
                            RefreshState(string.Format("更新公司健保費率( {0}% )", EFF_RATE * 100), percent += 2);
                            try
                            {
                                var cc = db.ExecuteCommand(UpdateHarCodeSys, new object[] { });
                                RefreshState("更新公司健保費率完成，共更新" + cc.ToString() + "筆資料", percent += 2);
                            }
                            catch (Exception eex)
                            {
                                RefreshState("更新健保費率失敗", percent += 2);
                                RefreshState(eex.Message, percent);
                                JBModule.Message.TextLog.WriteLog(eex);
                            }
                        }
                    }
                    #endregion
                    #region 更新MAIN.INI設定
                    //RefreshState("UpdateMain.ini == " + ConfigSetting.AppSettingValue("UpdateMain.ini"), percent += 2);
                    if (ConfigSetting.AppSettingValue("UpdateMain.ini") == "1")
                    {
                        if (IsFoxpro)
                        {
                            //RefreshState("File.Exists(MainFile)=" + File.Exists(MainFile).ToString() + " " + MainFile.ToString(), percent += 2);
                            if (File.Exists(MainFile))
                            {
                                try
                                {
                                    RefreshState("更新MAIN.INI設定", percent += 2);
                                    StreamReader sr = new StreamReader(MainFile, System.Text.Encoding.Default);
                                    List<string> fullText = new List<string>();
                                    while (sr.Peek() > 0)
                                    {
                                        string readStr = sr.ReadLine();
                                        if (readStr.IndexOf("SUPPLERATE=") != -1) fullText.Add(string.Format("SUPPLERATE={0}", ThisSUPPLEINSLABRATE));//補充保費費率
                                        else if (readStr.IndexOf("COMPERSONCNT=") != -1) fullText.Add(string.Format("COMPERSONCNT={0}", compersoncnt));//平均眷口數
                                        else if (readStr.IndexOf("COMPRATE=") != -1) fullText.Add(string.Format("COMPRATE={0}",EFF_RATE));//健保費率
                                        else if (readStr.IndexOf("FORSALBASD=") != -1) fullText.Add(string.Format("FORSALBASD={0}", FORSALBASD));//非居住者基本工資1.5倍金額
                                        else fullText.Add(readStr);
                                    }
                                    sr.Close();
                                    StreamWriter sw = new StreamWriter(MainFile, false, Encoding.Default);
                                    foreach (var it in fullText)
                                    {
                                        sw.WriteLine(it);
                                    }
                                    sw.Close();
                                    RefreshState("更新MAIN.INI設定完成", percent += 2);
                                }
                                catch (Exception eex)
                                {
                                    RefreshState("更新MAIN.INI設定失敗", percent += 2);
                                    RefreshState(eex.Message, percent);
                                    JBModule.Message.TextLog.WriteLog(eex);
                                }
                            }
                        }
                    }
                    #endregion
                    #region 勞保投保金額更新
                    if (ConfigSetting.AppSettingValue("UpdateLaborMax") == "1")
                    {
                        JBModule.Message.TextLog.WriteLog("開始進行-勞保投保金額更新");

                        LaborUpdate lu = new LaborUpdate(new SqlConnection(e.Argument.ToString()));
                        int i = lu.CheckLaborMax();
                        JBModule.Message.TextLog.WriteLog("完成-勞保投保金額更新，共影響" + i.ToString() + "筆資料");
                        RefreshState("完成-勞保投保金額更新，共影響" + i.ToString() + "筆資料", 99);
                        if (i < 0)
                        {
                            RefreshState(@"更新異常，請至 C:\TEMP\Error\Log\ 查閱相關訊息", 99);
                            return;
                        }
                    }
                    #endregion

                    #region 保險級距更新
                    if (ConfigSetting.AppSettingValue("UpdateInsurlv") == "1")
                    {
                        JBModule.Message.TextLog.WriteLog(string.Format("開始進行-保險級距{0}失效更新", SUPPLEMIN));

                        InsurlvUpdate lu = new InsurlvUpdate(new SqlConnection(e.Argument.ToString()));
                        bool i = lu.ChecInsurlv();
                        JBModule.Message.TextLog.WriteLog("完成-保險級距{0}失效更新", SUPPLEMIN);
                        RefreshState("完成-保險級距更新", 99);
                        if (!i)
                        {
                            RefreshState(@"更新異常，請至 C:\TEMP\Error\Log\ 查閱相關訊息", 99);
                            return;
                        }
                    }
                    #endregion

                    #region 保險級距新增
                    if (ConfigSetting.AppSettingValue("InsertInsurlv") == "1")
                    {
                        JBModule.Message.TextLog.WriteLog(string.Format("開始進行-保險級距{0}新增", SUPPLEMIN));

                        InsurlvUpdate lu = new InsurlvUpdate(new SqlConnection(e.Argument.ToString()));
                        bool b = lu.addNewInsurlv(SUPPLEMIN, EFF_RATE, YEAR);
                        JBModule.Message.TextLog.WriteLog("完成-保險級距{0}新增", SUPPLEMIN);
                        RefreshState("完成-保險級距新增", 99);
                        if (!b)
                        {
                            RefreshState(@"更新異常，請至 C:\TEMP\Error\Log\ 查閱相關訊息", 99);
                            return;
                        }
                    }
                    #endregion

                    #region 非居住者基本工資1.5倍
                    if (ConfigSetting.AppSettingValue("ForeignerTaxAmt") == "1")
                    {
                        if (!IsFoxpro)
                        {
                            JBModule.Message.TextLog.WriteLog(string.Format("開始進行-非居住者基本工資1.5倍金額更新為{0}", FORSALBASD));

                            string UpdateTaxAmt = string.Format("update U_SYS9 set FORSALBASD={0}", FORSALBASD);
                            RefreshState(string.Format("開始進行-非居住者基本工資1.5倍金額更新為{0}", FORSALBASD), percent += 2);
                            try
                            {
                                var cc = db.ExecuteCommand(UpdateTaxAmt, new object[] { });
                                RefreshState("完成-非居住者基本工資1.5倍金額更新", percent += 2);
                            }
                            catch (Exception eex)
                            {
                                RefreshState("更新非居住者基本工資1.5倍金額失敗", percent += 2);
                                RefreshState(eex.Message, percent);
                                JBModule.Message.TextLog.WriteLog(eex);
                            } 
                        }
                    }
                    #endregion

                    RefreshState("執行完畢", 100);
                }
                catch (Exception ex)
                {
                    RefreshState("程式發生錯誤", toolStripProgressBar1.Value);
                    RefreshState(ex.Message, toolStripProgressBar1.Value);
                    JBModule.Message.TextLog.WriteLog(ex);
                }
            }
        }
        void RefreshState(string State, int Percent)
        {
            BW.ReportProgress(Percent, State + Environment.NewLine);
            JBModule.Message.TextLog.WriteLog(State);
            System.Threading.Thread.Sleep(1000);
        }

        private void BW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.UserState != null)
            { 
                textBox1.Text += e.UserState.ToString();
                
            }
            toolStripProgressBar1.Value = e.ProgressPercentage;
            StateMsg = textBox1.Text;
        }

        private void BW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (toolStripProgressBar1.Value == 100)
            {
                MessageBox.Show("資料更新完成", "完成", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this.Close();
            }
            else
            {
                MessageBox.Show("資料更新失敗", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
