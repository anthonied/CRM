using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SolExcelCrmSync.Classes;
using System.Configuration;
using System.IO;
using System.Diagnostics;
using LINQtoCSV;

namespace SolExcelCrmSync
{
    public partial class Automation : Form
    {

        List<Invoice> lInvoiceToCsv;
        Export exExportInfo = new Export();
        bool bAuto = false;
        public Automation(string[] args)
        {
            InitializeComponent();
            if (args.Length > 0)
            {
                foreach (string sArg in args)
                {
                    if (sArg == "-a") // automation run
                    {
                        bAuto = true;
                    }
                }
            }
        }
        
        private void Automation_Load(object sender, EventArgs e)
        {

            UpdateUI();
            if (selInvoiceExportOptions.Items.Count > 0)
                selInvoiceExportOptions.SelectedIndex = 0;

            if (bAuto)
            {
                tabControl1.SelectedIndex = 1;
                var CustomerBackGroundWorkerBackgroundWorker = new BackgroundWorker();
                CustomerBackGroundWorkerBackgroundWorker.WorkerReportsProgress = true;

                CustomerBackGroundWorkerBackgroundWorker.DoWork += (o, ea) =>
                    {
                        CustomerBackGroundWorkerBackgroundWorker.ReportProgress(0, "Auto Sync Requested - " + DateTime.Now.ToString("HH:mm:ss"));
                        runSync("Auto", CustomerBackGroundWorkerBackgroundWorker);
                        
                    };
                CustomerBackGroundWorkerBackgroundWorker.RunWorkerCompleted += (o, ea) =>
                {
                    lbCustomerProgress.Items.Insert(0, "Sync Process Completed");
                    UseWaitCursor = false;
                };
                CustomerBackGroundWorkerBackgroundWorker.ProgressChanged += (o, ea) =>
                {
                    if (ea.UserState != null)
                    {
                        lbCustomerProgress.Items.Insert(0, ea.UserState);
                        if (ea.UserState.ToString().StartsWith("Starting"))
                        {
                            progressBar1.Maximum = ea.ProgressPercentage;
                        }
                    }
                    else
                    {
                        progressBar1.Value = ea.ProgressPercentage;
                    }
                };
                CustomerBackGroundWorkerBackgroundWorker.RunWorkerAsync();
            }
        }

        private void UpdateUI()
        {
            if (exExportInfo.LastExport.Year == 1970)
                lblLastExport.Text = "No previous export.";
            else
                lblLastExport.Text = exExportInfo.LastExport.ToString("yyyy-MM-dd HH:mm:ss");
        }

        private void cmdManualRun_Click(object sender, EventArgs e)
        {
            var CustomerBackGroundWorkerBackgroundWorker = new BackgroundWorker();
            CustomerBackGroundWorkerBackgroundWorker.WorkerReportsProgress = true;
            CustomerBackGroundWorkerBackgroundWorker.DoWork += (o, ea) =>
                {
                    runSync("Manual",  CustomerBackGroundWorkerBackgroundWorker);
                };

            CustomerBackGroundWorkerBackgroundWorker.RunWorkerCompleted += (o,ea) =>
                {
                    lbCustomerProgress.Items.Insert(0, "Sync Process Completed");
                   UseWaitCursor = false;
                };
            CustomerBackGroundWorkerBackgroundWorker.ProgressChanged += (o, ea) =>
                {
                    if (ea.UserState != null)
                    {
                        lbCustomerProgress.Items.Insert(0, ea.UserState);
                        if (ea.UserState.ToString().StartsWith("Starting"))
                        {
                            progressBar1.Maximum = ea.ProgressPercentage;
                        }
                    }
                    else
                    {
                        progressBar1.Value = ea.ProgressPercentage;
                    }
                };
            UseWaitCursor = true;
            CustomerBackGroundWorkerBackgroundWorker.RunWorkerAsync();
        }

        private void runSync(string sMethod,  BackgroundWorker CustomerBackGroundWorkerBackgroundWorker)
        {
            CustomerBackGroundWorkerBackgroundWorker.ReportProgress(0, "Syncing Customers");
            int iTransactionCounter = 0;
            string sDirectory = ConfigurationManager.AppSettings["ExcelPath"];
            var ExcelRead = new ExcelRead();
            var crmEntity = new CRMEntityMananger();
            int iUpdated = 0;
            int iAdded = 0;
            string sTransID;
            List<string> lNewCustomers = new List<string>();
            foreach (var sFile in Directory.GetFiles(sDirectory))
            {
                iUpdated = 0;
                iAdded = 0;
                CustomerBackGroundWorkerBackgroundWorker.ReportProgress(0, "Reading Excel File...");
                List<accountDetails> lAccountDetails = ExcelRead.readExcel(sFile);
                crmEntity.updateAccountDetails(lAccountDetails, ref iAdded, ref iUpdated, ref lNewCustomers, CustomerBackGroundWorkerBackgroundWorker);
                sTransID = Loging.addTransaction(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), sDirectory, sFile, iUpdated, iAdded, sMethod);
                CustomerBackGroundWorkerBackgroundWorker.ReportProgress(0, iUpdated.ToString() + " records updated.");
                CustomerBackGroundWorkerBackgroundWorker.ReportProgress(0, iAdded.ToString() + " records added.");
                Loging.addNewCustomers(sTransID, lNewCustomers);
                
                iTransactionCounter++;
            }
            if (iTransactionCounter == 0) //no transactions logged for this sync run
            {
                Loging.addTransaction(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), sDirectory, "", 0, 0, "No files found - " + sMethod);
            }
            if (sMethod == "Auto")
            {
                Application.Exit();
                try
                {
                    Process.GetCurrentProcess().Kill();
                }
                catch (Exception ex)
                {
                    
                }
            }
        }

        private void ExportInvoice_Click(object sender, EventArgs e)
        {
            this.UseWaitCursor = true;
            pnlLastExportFile.Visible = false;
            backgroundWorkerInvoices.RunWorkerAsync();
            backgroundWorkerInvoices.RunWorkerCompleted += (o, ea) =>
                {
                    this.UseWaitCursor = false;
                    if (exExportInfo.InvoicesExportCount > 0)
                    {
                        var outputFileDescription = new CsvFileDescription
                        {
                            SeparatorChar = ',',
                            FirstLineHasColumnNames = false,
                            EnforceCsvColumnAttribute = true
                        };
                        var cc = new CsvContext();

                        SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                        saveFileDialog1.Filter = "Invoice|*.001";
                        saveFileDialog1.FileName = "IUSER" + DateTime.Now.ToString("MMdd");
                        saveFileDialog1.Title = "Save Invoice";
                        saveFileDialog1.ShowDialog();

                        // If the file name is not an empty string open it for saving.
                        if (saveFileDialog1.FileName != "")
                        {
                            cc.Write(lInvoiceToCsv, saveFileDialog1.FileName, outputFileDescription);
                            CorrectInvoiceProductString(saveFileDialog1.FileName);
                        }
                        llExportedFile.Visible = true;
                        llExportedFile.Text = "Click here to open CSV file.";
                        llExportedFile.Links.Add(6, 4, saveFileDialog1.FileName);

                        progressBar1.Value = 0;
                        lbInvoiceProgress.Items.Insert(0, "Export Completed");
                        lbInvoiceProgress.Items.Insert(0, "");
                        exExportInfo.SetLastExport(DateTime.Now);

                        UpdateUI();
                    }
                    else
                    {
                        lbInvoiceProgress.Items.Insert(0, "Export Completed");
                        lbInvoiceProgress.Items.Insert(0, "No Records To Export");
                        lbInvoiceProgress.Items.Insert(0, "");
                    }
                };
        }

        private void CorrectInvoiceProductString(string filePath)
        {
            string CsvText = "";
            using (var streamReader = new StreamReader(filePath))
            {
                CsvText = streamReader.ReadToEnd();
                streamReader.Close();
                CsvText = CsvText.Replace("\"*remove*", "").Replace("*remove*\"", ",.").Replace("*quote*", "\"");                
            }
            using (var streamWriter = new StreamWriter(filePath))
            {
                streamWriter.Write(CsvText);
                streamWriter.Close();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {            
            var crmEntity = new CRMEntityMananger();
            lInvoiceToCsv = crmEntity.ListInvoices(ref backgroundWorkerInvoices, exExportInfo.LastExport);
            exExportInfo.InvoicesExportCount = lInvoiceToCsv.Count;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.UserState != null)
            {
                lbInvoiceProgress.Items.Insert(0, e.UserState);
                if (e.UserState.ToString() == "Starting export process...")
                {
                    progressBar1.Maximum = e.ProgressPercentage;
                }
            }
            else
            {
                progressBar1.Value = e.ProgressPercentage;
            }
            
        }

        private void llExportedFile_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Link.LinkData.ToString());
        }

             
        private void backgroundWorkerCustomer_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.UserState != null)
            {
                lbCustomerProgress.Items.Insert(0, e.UserState);
                if (e.UserState.ToString().StartsWith("Starting"))
                {
                    progressBar1.Maximum = e.ProgressPercentage;
                }
            }
            else
            {
                progressBar1.Value = e.ProgressPercentage;
            }
        }

        private void backgroundWorkerCustomer_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

    }
}
