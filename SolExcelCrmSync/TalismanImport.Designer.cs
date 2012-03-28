namespace SolExcelCrmSync
{
    partial class TalismanImport
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.cmdStartImport = new System.Windows.Forms.Button();
            this.cmdUpdateRole = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Top;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(470, 233);
            this.dataGridView1.TabIndex = 0;
            // 
            // cmdStartImport
            // 
            this.cmdStartImport.Location = new System.Drawing.Point(327, 247);
            this.cmdStartImport.Name = "cmdStartImport";
            this.cmdStartImport.Size = new System.Drawing.Size(132, 60);
            this.cmdStartImport.TabIndex = 1;
            this.cmdStartImport.Text = "Import";
            this.cmdStartImport.UseVisualStyleBackColor = true;
            this.cmdStartImport.Click += new System.EventHandler(this.cmdStartImport_Click);
            // 
            // cmdUpdateRole
            // 
            this.cmdUpdateRole.Location = new System.Drawing.Point(191, 251);
            this.cmdUpdateRole.Name = "cmdUpdateRole";
            this.cmdUpdateRole.Size = new System.Drawing.Size(120, 56);
            this.cmdUpdateRole.TabIndex = 2;
            this.cmdUpdateRole.Text = "Update Role";
            this.cmdUpdateRole.UseVisualStyleBackColor = true;
            this.cmdUpdateRole.Click += new System.EventHandler(this.cmdUpdateRole_Click);
            // 
            // TalismanImport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 319);
            this.Controls.Add(this.cmdUpdateRole);
            this.Controls.Add(this.cmdStartImport);
            this.Controls.Add(this.dataGridView1);
            this.Name = "TalismanImport";
            this.Text = "Talisman Import";
            this.Load += new System.EventHandler(this.TalismanImport_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button cmdStartImport;
        private System.Windows.Forms.Button cmdUpdateRole;
    }
}