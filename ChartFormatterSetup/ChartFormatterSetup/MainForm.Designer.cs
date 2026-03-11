namespace ChartFormatterSetup
{
    partial class MainForm
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
            btnInstallOrUpdate = new Button();
            btnClose = new Button();
            lblGuide = new Label();
            groupBox1 = new GroupBox();
            btnUninstall = new Button();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // btnInstallOrUpdate
            // 
            btnInstallOrUpdate.Location = new Point(345, 194);
            btnInstallOrUpdate.Name = "btnInstallOrUpdate";
            btnInstallOrUpdate.Size = new Size(75, 23);
            btnInstallOrUpdate.TabIndex = 0;
            btnInstallOrUpdate.Text = "インストール";
            btnInstallOrUpdate.UseVisualStyleBackColor = true;
            btnInstallOrUpdate.Click += btnInstallOrUpdate_Click;
            // 
            // btnClose
            // 
            btnClose.Location = new Point(426, 194);
            btnClose.Name = "btnClose";
            btnClose.Size = new Size(75, 23);
            btnClose.TabIndex = 1;
            btnClose.Text = "閉じる";
            btnClose.UseVisualStyleBackColor = true;
            btnClose.Click += btnClose_Click;
            // 
            // lblGuide
            // 
            lblGuide.Dock = DockStyle.Fill;
            lblGuide.Location = new Point(3, 19);
            lblGuide.Name = "lblGuide";
            lblGuide.Size = new Size(483, 154);
            lblGuide.TabIndex = 2;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(lblGuide);
            groupBox1.Location = new Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(489, 176);
            groupBox1.TabIndex = 3;
            groupBox1.TabStop = false;
            // 
            // btnUninstall
            // 
            btnUninstall.Location = new Point(245, 194);
            btnUninstall.Name = "btnUninstall";
            btnUninstall.Size = new Size(94, 23);
            btnUninstall.TabIndex = 4;
            btnUninstall.Text = "アンインストール";
            btnUninstall.UseVisualStyleBackColor = true;
            btnUninstall.Click += btnUninstall_Click;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(513, 228);
            Controls.Add(btnUninstall);
            Controls.Add(groupBox1);
            Controls.Add(btnClose);
            Controls.Add(btnInstallOrUpdate);
            Name = "MainForm";
            Text = "ChartFormatterSetup";
            Load += MainForm_Load;
            groupBox1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Button btnInstallOrUpdate;
        private Button btnClose;
        private Label lblGuide;
        private GroupBox groupBox1;
        private Button btnUninstall;
    }
}