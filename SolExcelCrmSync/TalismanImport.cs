using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SolExcelCrmSync.Source;

namespace SolExcelCrmSync
{
    public partial class TalismanImport : Form
    {
        public TalismanImport()
        {
            InitializeComponent();
            var a = new IAccountTalisman();
            
            dataGridView1.DataSource = a.FillAccountBaseFromSql();
        }

        private void TalismanImport_Load(object sender, EventArgs e)
        {

        }
    }
}
