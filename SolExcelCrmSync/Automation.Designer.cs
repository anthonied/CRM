namespace SolExcelCrmSync
{
    partial class Automation
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Automation));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cmdManualRun = new System.Windows.Forms.Button();
            this.cmdExportInvoice = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.lbInvoiceProgress = new System.Windows.Forms.ListBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.lblLastExport = new System.Windows.Forms.Label();
            this.selInvoiceExportOptions = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.pnlLastExportFile = new System.Windows.Forms.Panel();
            this.llExportedFile = new System.Windows.Forms.LinkLabel();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lbCustomerProgress = new System.Windows.Forms.ListBox();
            this.backgroundWorkerInvoices = new System.ComponentModel.BackgroundWorker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.pnlLastExportFile.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::SolExcelCrmSync.Properties.Resources.small_print;
            this.pictureBox1.Location = new System.Drawing.Point(535, -1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(108, 100);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 43;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(10, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(191, 30);
            this.label1.TabIndex = 44;
            this.label1.Text = "CRM Custom Tool";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(14, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 13);
            this.label2.TabIndex = 46;
            this.label2.Text = "v 1.3";
            // 
            // cmdManualRun
            // 
            this.cmdManualRun.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.cmdManualRun.Location = new System.Drawing.Point(3, 273);
            this.cmdManualRun.Name = "cmdManualRun";
            this.cmdManualRun.Size = new System.Drawing.Size(627, 48);
            this.cmdManualRun.TabIndex = 47;
            this.cmdManualRun.Text = "Manual Customer Sync";
            this.cmdManualRun.UseVisualStyleBackColor = true;
            this.cmdManualRun.Click += new System.EventHandler(this.cmdManualRun_Click);
            // 
            // cmdExportInvoice
            // 
            this.cmdExportInvoice.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.cmdExportInvoice.Location = new System.Drawing.Point(3, 273);
            this.cmdExportInvoice.Name = "cmdExportInvoice";
            this.cmdExportInvoice.Size = new System.Drawing.Size(627, 48);
            this.cmdExportInvoice.TabIndex = 48;
            this.cmdExportInvoice.Text = "Export Invoices To CSV";
            this.cmdExportInvoice.UseVisualStyleBackColor = true;
            this.cmdExportInvoice.Click += new System.EventHandler(this.ExportInvoice_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.tabControl1.Location = new System.Drawing.Point(0, 110);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(641, 350);
            this.tabControl1.TabIndex = 49;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.lbInvoiceProgress);
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.Controls.Add(this.pnlLastExportFile);
            this.tabPage1.Controls.Add(this.cmdExportInvoice);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(633, 324);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Invoice Management";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // lbInvoiceProgress
            // 
            this.lbInvoiceProgress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbInvoiceProgress.FormattingEnabled = true;
            this.lbInvoiceProgress.Location = new System.Drawing.Point(3, 69);
            this.lbInvoiceProgress.Name = "lbInvoiceProgress";
            this.lbInvoiceProgress.Size = new System.Drawing.Size(627, 172);
            this.lbInvoiceProgress.TabIndex = 49;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.lblLastExport);
            this.panel1.Controls.Add(this.selInvoiceExportOptions);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(627, 66);
            this.panel1.TabIndex = 51;            
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(4, 36);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(68, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Prev. Export:";
            // 
            // lblLastExport
            // 
            this.lblLastExport.AutoSize = true;
            this.lblLastExport.Location = new System.Drawing.Point(83, 36);
            this.lblLastExport.Name = "lblLastExport";
            this.lblLastExport.Size = new System.Drawing.Size(0, 13);
            this.lblLastExport.TabIndex = 2;
            // 
            // selInvoiceExportOptions
            // 
            this.selInvoiceExportOptions.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.selInvoiceExportOptions.FormattingEnabled = true;
            this.selInvoiceExportOptions.Items.AddRange(new object[] {
            "From Last Import Date"});
            this.selInvoiceExportOptions.Location = new System.Drawing.Point(85, 5);
            this.selInvoiceExportOptions.Name = "selInvoiceExportOptions";
            this.selInvoiceExportOptions.Size = new System.Drawing.Size(155, 21);
            this.selInvoiceExportOptions.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(79, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Export Options:";
            // 
            // pnlLastExportFile
            // 
            this.pnlLastExportFile.Controls.Add(this.llExportedFile);
            this.pnlLastExportFile.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlLastExportFile.Location = new System.Drawing.Point(3, 241);
            this.pnlLastExportFile.Name = "pnlLastExportFile";
            this.pnlLastExportFile.Size = new System.Drawing.Size(627, 32);
            this.pnlLastExportFile.TabIndex = 50;
            // 
            // llExportedFile
            // 
            this.llExportedFile.AutoSize = true;
            this.llExportedFile.Location = new System.Drawing.Point(7, 9);
            this.llExportedFile.Name = "llExportedFile";
            this.llExportedFile.Size = new System.Drawing.Size(55, 13);
            this.llExportedFile.TabIndex = 0;
            this.llExportedFile.TabStop = true;
            this.llExportedFile.Text = "linkLabel1";
            this.llExportedFile.Visible = false;
            this.llExportedFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llExportedFile_LinkClicked);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lbCustomerProgress);
            this.tabPage2.Controls.Add(this.cmdManualRun);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(633, 324);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Customer Management";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // lbCustomerProgress
            // 
            this.lbCustomerProgress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbCustomerProgress.FormattingEnabled = true;
            this.lbCustomerProgress.Location = new System.Drawing.Point(3, 3);
            this.lbCustomerProgress.Name = "lbCustomerProgress";
            this.lbCustomerProgress.Size = new System.Drawing.Size(627, 270);
            this.lbCustomerProgress.TabIndex = 50;
            // 
            // backgroundWorkerInvoices
            // 
            this.backgroundWorkerInvoices.WorkerReportsProgress = true;
            this.backgroundWorkerInvoices.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorkerInvoices.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Location = new System.Drawing.Point(0, 460);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(641, 23);
            this.progressBar1.TabIndex = 50;
            // 
            // Automation
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(641, 483);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Automation";
            this.Text = "CRM Automation";
            this.Load += new System.EventHandler(this.Automation_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.pnlLastExportFile.ResumeLayout(false);
            this.pnlLastExportFile.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button cmdManualRun;
        private System.Windows.Forms.Button cmdExportInvoice;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.ComponentModel.BackgroundWorker backgroundWorkerInvoices;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ListBox lbInvoiceProgress;
        private System.Windows.Forms.ListBox lbCustomerProgress;
        private System.Windows.Forms.Panel pnlLastExportFile;
        private System.Windows.Forms.LinkLabel llExportedFile;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblLastExport;
        private System.Windows.Forms.ComboBox selInvoiceExportOptions;
    }
}