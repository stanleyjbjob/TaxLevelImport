namespace JBHR
{
    partial class U_SETDB
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBoxSERVER = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.bnSave = new System.Windows.Forms.Button();
            this.bnTest = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxDATABASE = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxPASSWORD = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxUSERID = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // textBoxSERVER
            // 
            //this.textBoxSERVER.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange;
            //this.textBoxSERVER.CaptionLabel = this.label1;
            this.textBoxSERVER.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
            //this.textBoxSERVER.DecimalPlace = 2;
            //this.textBoxSERVER.IsEmpty = true;
            this.textBoxSERVER.Location = new System.Drawing.Point(131, 12);
            //this.textBoxSERVER.Mask = "";
            this.textBoxSERVER.Name = "textBoxSERVER";
            this.textBoxSERVER.PasswordChar = '\0';
            this.textBoxSERVER.ReadOnly = false;
            this.textBoxSERVER.Size = new System.Drawing.Size(157, 22);
            this.textBoxSERVER.TabIndex = 0;
            //this.textBoxSERVER.ValidType = JBControls.TextBox.EValidType.String;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(12, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "伺服器名稱或IP位址";
            // 
            // bnSave
            // 
            this.bnSave.Location = new System.Drawing.Point(213, 124);
            this.bnSave.Name = "bnSave";
            this.bnSave.Size = new System.Drawing.Size(75, 23);
            this.bnSave.TabIndex = 5;
            this.bnSave.Text = "儲存並離開";
            this.bnSave.UseVisualStyleBackColor = true;
            this.bnSave.Click += new System.EventHandler(this.bnSave_Click);
            // 
            // bnTest
            // 
            this.bnTest.Location = new System.Drawing.Point(132, 124);
            this.bnTest.Name = "bnTest";
            this.bnTest.Size = new System.Drawing.Size(75, 23);
            this.bnTest.TabIndex = 4;
            this.bnTest.Text = "連線測試";
            this.bnTest.UseVisualStyleBackColor = true;
            this.bnTest.Click += new System.EventHandler(this.bnTest_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(60, 45);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 13;
            this.label4.Text = "資料庫名稱";
            // 
            // textBoxDATABASE
            // 
            //this.textBoxDATABASE.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange;
            //this.textBoxDATABASE.CaptionLabel = this.label4;
            this.textBoxDATABASE.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
            //this.textBoxDATABASE.DecimalPlace = 2;
            //this.textBoxDATABASE.IsEmpty = true;
            this.textBoxDATABASE.Location = new System.Drawing.Point(131, 40);
            //this.textBoxDATABASE.Mask = "";
            this.textBoxDATABASE.Name = "textBoxDATABASE";
            this.textBoxDATABASE.PasswordChar = '\0';
            this.textBoxDATABASE.ReadOnly = false;
            this.textBoxDATABASE.Size = new System.Drawing.Size(157, 22);
            this.textBoxDATABASE.TabIndex = 1;
            //this.textBoxDATABASE.ValidType = JBControls.TextBox.EValidType.String;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(72, 101);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 17;
            this.label3.Text = "登入密碼";
            // 
            // textBoxPASSWORD
            // 
            //this.textBoxPASSWORD.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange;
            //this.textBoxPASSWORD.CaptionLabel = this.label3;
            this.textBoxPASSWORD.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
            //this.textBoxPASSWORD.DecimalPlace = 2;
            //this.textBoxPASSWORD.IsEmpty = true;
            this.textBoxPASSWORD.Location = new System.Drawing.Point(131, 96);
            //this.textBoxPASSWORD.Mask = "";
            this.textBoxPASSWORD.Name = "textBoxPASSWORD";
            this.textBoxPASSWORD.PasswordChar = '*';
            this.textBoxPASSWORD.ReadOnly = false;
            this.textBoxPASSWORD.Size = new System.Drawing.Size(157, 22);
            this.textBoxPASSWORD.TabIndex = 3;
            //this.textBoxPASSWORD.ValidType = JBControls.TextBox.EValidType.String;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(72, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 15;
            this.label2.Text = "登入帳號";
            // 
            // textBoxUSERID
            // 
            //this.textBoxUSERID.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange;
            //this.textBoxUSERID.CaptionLabel = this.label2;
            this.textBoxUSERID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
            //this.textBoxUSERID.DecimalPlace = 2;
            //this.textBoxUSERID.IsEmpty = true;
            this.textBoxUSERID.Location = new System.Drawing.Point(131, 68);
            //this.textBoxUSERID.Mask = "";
            this.textBoxUSERID.Name = "textBoxUSERID";
            this.textBoxUSERID.PasswordChar = '*';
            this.textBoxUSERID.ReadOnly = false;
            this.textBoxUSERID.Size = new System.Drawing.Size(157, 22);
            this.textBoxUSERID.TabIndex = 2;
            //this.textBoxUSERID.ValidType = JBControls.TextBox.EValidType.String;
            // 
            // U_SETDB
            // 
            this.AcceptButton = this.bnSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(304, 157);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxPASSWORD);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxUSERID);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBoxDATABASE);
            this.Controls.Add(this.bnTest);
            this.Controls.Add(this.bnSave);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxSERVER);
            //this.FormSize = JBControls.JBForm.FormSizeType.Custom;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "U_SETDB";
            this.Text = "設定資料伺服器連入參數";
            this.Load += new System.EventHandler(this.U_SETDB_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxSERVER;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button bnSave;
        private System.Windows.Forms.Button bnTest;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxDATABASE;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxPASSWORD;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxUSERID;
    }
}