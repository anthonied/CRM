using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SolExcelCrmSync.Source;
using SolExcelCrmSync.Services;
using SolExcelCrmSync.Containers.User;

namespace SolExcelCrmSync
{
    public partial class TalismanImport : Form
    {
        public TalismanImport()
        {
            InitializeComponent();
           
        }

        private void TalismanImport_Load(object sender, EventArgs e)
        {

        }

        private void cmdStartImport_Click(object sender, EventArgs e)
        {
             var Account = new IAccountTalisman();
             var importData = Account.FillAccountBaseFromSql();
             var CrmAccountService = new CRMCustomerService();

             CrmAccountService.importFromSqlSolCRM(importData);

            dataGridView1.DataSource =importData;
        }

        private void cmdUpdateRole_Click(object sender, EventArgs e)
        {
            var marketer = new Marketer();
            var CrmUserService = new CRMUserService();

            marketer.sCrmUsername = "CARBONCLEAR\\Erica";
            CrmUserService.getUserFromCRM(marketer);

            
        }
    }
}
