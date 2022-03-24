using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.IO;

namespace JBHR
{
    public partial class U_SETDB : Form
    {
        //System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

        public U_SETDB()
        {
            InitializeComponent();
        }
        public static string Server = "";
        public static string DataBase = "";
        public static string UserId = "";
        public static string PassWord = "";
        private void U_SETDB_Load(object sender, EventArgs e)
        {
            //string[] keyValues = config.ConnectionStrings.ConnectionStrings["JBHR.Properties.Settings.JBHRConnectionString"].ConnectionString.Split(new char[] { ';' });
            //foreach (string keyValue in keyValues)
            //{
            //    string[] settings = keyValue.Split(new char[] { '=' });
            //    switch (settings[0].Trim())
            //    {
            //        case "Data Source":
            textBoxSERVER.Text = Server;
            //    break;
            //case "Initial Catalog":
            textBoxDATABASE.Text = DataBase;
            //    break;
            //case "User ID":
            textBoxUSERID.Text = UserId;
            //    break;
            //case "Password":
            textBoxPASSWORD.Text = PassWord;
            //break;
            //    }
            //}
        }

        private void bnTest_Click(object sender, EventArgs e)
        {
            string connectionString = string.Format("Data Source={0};Initial Catalog={1};Persist Security Info=True;User ID={2};Password={3}", textBoxSERVER.Text, textBoxDATABASE.Text, textBoxUSERID.Text, textBoxPASSWORD.Text);

            SqlConnection conn = new SqlConnection(connectionString);
            try
            {
                if (conn.State == System.Data.ConnectionState.Closed) conn.Open();
                MessageBox.Show("測試成功", "訊息", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
            catch
            {
                MessageBox.Show("資料庫測試失敗", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                conn.Close();
            }
        }

        private void bnSave_Click(object sender, EventArgs e)
        {
            if (textBoxSERVER.Text.Trim().Length == 0)
            {
                MessageBox.Show("主機位置不可以是空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (textBoxDATABASE.Text.Trim().Length == 0)
            {
                MessageBox.Show("資料庫名稱不可以是空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (textBoxUSERID.Text.Trim().Length == 0)
            {
                MessageBox.Show("使用者名稱不可以是空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //if (textBoxPASSWORD.Text.Trim().Length == 0)
            //{
            //    MessageBox.Show("密碼不可以是空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            Server = textBoxSERVER.Text;
            DataBase = textBoxDATABASE.Text;
            UserId = textBoxUSERID.Text;
            PassWord = textBoxPASSWORD.Text;
            this.Close();
        }
    }
}