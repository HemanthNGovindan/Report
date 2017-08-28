using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Report_Compare
{
    public partial class ReportCompare : Form
    {
        ReportData reportData = null;
        string AllowedExtensions = "xlsx.xlsm.xls.csv";

        #region Initialize WinForm
        public ReportCompare()
        {
            InitializeComponent();
            CustomInitialization();

        }

        private void CustomInitialization()
        {
            reportData = new ReportData();
            lblError.Text = "";
            lblResourceSheetFileName.Text = lblInvoiceSummaryFileName.Text = lblAribaMappingFileName.Text = "";
            btnResouceSheetClear.Visible = btnInvoiceSummaryClear.Visible = btnAribaMappingClear.Visible = false;
            reportData.ReportFileNameTemplate = Properties.Settings.Default.TemplatePath;
            if (string.IsNullOrEmpty(reportData.ReportFileNameTemplate))
            {
                btnSelectTemplate.Text = "Select";
            }
            else
            {
                btnSelectTemplate.Text = "Change";
                lblTempltePath.Text = "Template Path - " + reportData.ReportFileNameTemplate;

            }
            lblError.BackColor = System.Drawing.Color.Red;
            backgroundWorker1.WorkerReportsProgress = true;
            reportData.ReportCount = Convert.ToInt32(nUDReportCount.Value);
            // Start the BackgroundWorker.
        }
        #endregion

        #region Choosing Files
        private void btnResouceSheetChoose_Click(object sender, EventArgs e)
        {
            lblError.Text = "";
            lblError.BackColor = System.Drawing.Color.Red;
            DialogResult result = openFileDialog.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {

                try
                {
                    reportData.ResouceSheetFileName = openFileDialog.FileName;
                    if (!string.IsNullOrEmpty(reportData.ResouceSheetFileName) && !AllowedExtensions.Contains(reportData.ResouceSheetFileName.Split('.')[reportData.ResouceSheetFileName.Split('.').Length - 1].ToLower()))
                    {
                        reportData.ResouceSheetFileName = "";
                        lblError.Text = "Vaild for .xlsx, .xlsm, .xls, .csv files";
                        lblResourceSheetFileName.Text = "";
                        btnResouceSheetClear.Visible = false;
                        btnResouceSheetChoose.Visible = true;
                        return;
                    }
                    else if (!string.IsNullOrEmpty(reportData.InvoiceSummaryFileName) && reportData.ResouceSheetFileName.Equals(reportData.InvoiceSummaryFileName))
                    {
                        reportData.ResouceSheetFileName = "";
                        lblError.Text = "seems to be same as Invoice Summary File. Please verify";
                        lblResourceSheetFileName.Text = "";
                        btnResouceSheetClear.Visible = false;
                        btnResouceSheetChoose.Visible = true;
                        return;
                    }
                    else if (!string.IsNullOrEmpty(reportData.AribaMappingFileName) && reportData.ResouceSheetFileName.Equals(reportData.AribaMappingFileName))
                    {
                        reportData.ResouceSheetFileName = "";
                        lblError.Text = "seems to be same as Ariba Mapping File. Please verify";
                        lblResourceSheetFileName.Text = "";
                        btnResouceSheetClear.Visible = false;
                        btnResouceSheetChoose.Visible = true;
                        return;
                    }
                    lblResourceSheetFileName.Text = "File - " + new FileInfo(openFileDialog.FileName).Name;
                    btnResouceSheetChoose.Visible = false;
                    btnResouceSheetClear.Visible = true;
                }
                catch (Exception ex)
                {
                    reportData.ResouceSheetFileName = "";
                    lblError.Text = "Error : " + ex.Message;
                    lblResourceSheetFileName.Text = "";
                    btnResouceSheetClear.Visible = false;
                    btnResouceSheetChoose.Visible = true;
                }
            }

        }
        private void btnInvoiceSummaryChoose_Click(object sender, EventArgs e)
        {
            lblError.Text = "";
            lblError.BackColor = System.Drawing.Color.Red;
            DialogResult result = openFileDialog.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {

                try
                {
                    reportData.InvoiceSummaryFileName = openFileDialog.FileName;
                    if (!string.IsNullOrEmpty(reportData.InvoiceSummaryFileName) && !AllowedExtensions.Contains(reportData.InvoiceSummaryFileName.Split('.')[reportData.InvoiceSummaryFileName.Split('.').Length - 1].ToLower()))
                    {
                        reportData.InvoiceSummaryFileName = "";
                        lblError.Text = "Vaild for .xlsx, .xlsm, .xls, .csv files";
                        lblInvoiceSummaryFileName.Text = "";
                        btnInvoiceSummaryClear.Visible = false;
                        btnInvoiceSummaryChoose.Visible = true;
                        return;
                    }
                    else if (!string.IsNullOrEmpty(reportData.ResouceSheetFileName) && reportData.InvoiceSummaryFileName.Equals(reportData.ResouceSheetFileName))
                    {
                        reportData.ResouceSheetFileName = "";
                        lblError.Text = "seems to be same as Resouce Sheet File. Please verify";
                        lblInvoiceSummaryFileName.Text = "";
                        btnInvoiceSummaryClear.Visible = false;
                        btnInvoiceSummaryChoose.Visible = true;
                        return;
                    }
                    else if (!string.IsNullOrEmpty(reportData.AribaMappingFileName) && reportData.InvoiceSummaryFileName.Equals(reportData.AribaMappingFileName))
                    {
                        reportData.ResouceSheetFileName = "";
                        lblError.Text = "seems to be same as Ariba Mapping File. Please verify";
                        lblInvoiceSummaryFileName.Text = "";
                        btnInvoiceSummaryClear.Visible = false;
                        btnInvoiceSummaryChoose.Visible = true;
                        return;
                    }
                    lblInvoiceSummaryFileName.Text = "File - " + new FileInfo(openFileDialog.FileName).Name;
                    btnInvoiceSummaryChoose.Visible = false;
                    btnInvoiceSummaryClear.Visible = true;
                }
                catch (Exception ex)
                {
                    reportData.ResouceSheetFileName = "";
                    lblError.Text = "Error : " + ex.Message;
                    lblInvoiceSummaryFileName.Text = "";
                    btnInvoiceSummaryClear.Visible = false;
                    btnInvoiceSummaryChoose.Visible = true;
                }
            }
        }
        private void btnAribaMappingChoose_Click(object sender, EventArgs e)
        {
            lblError.Text = "";
            lblError.BackColor = System.Drawing.Color.Red;
            DialogResult result = openFileDialog.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {

                try
                {
                    reportData.AribaMappingFileName = openFileDialog.FileName;
                    if (!string.IsNullOrEmpty(reportData.AribaMappingFileName) && !AllowedExtensions.Contains(reportData.AribaMappingFileName.Split('.')[reportData.AribaMappingFileName.Split('.').Length - 1].ToLower()))
                    {
                        reportData.AribaMappingFileName = "";
                        lblError.Text = "Vaild for .xlsx, .xlsm, .xls, .csv files";
                        lblAribaMappingFileName.Text = "";
                        btnAribaMappingChoose.Visible = true;
                        btnAribaMappingClear.Visible = false;
                        return;
                    }
                    else if (!string.IsNullOrEmpty(reportData.ResouceSheetFileName) && reportData.AribaMappingFileName.Equals(reportData.ResouceSheetFileName))
                    {
                        reportData.AribaMappingFileName = "";
                        lblError.Text = "seems to be same as Resouce Sheet File. Please verify";
                        lblAribaMappingFileName.Text = "";
                        btnAribaMappingChoose.Visible = true;
                        btnAribaMappingClear.Visible = false;
                        return;
                    }
                    else if (!string.IsNullOrEmpty(reportData.InvoiceSummaryFileName) && reportData.AribaMappingFileName.Equals(reportData.InvoiceSummaryFileName))
                    {
                        reportData.AribaMappingFileName = "";
                        lblError.Text = "seems to be same as Invoice Summary File. Please verify";
                        lblAribaMappingFileName.Text = "";
                        btnAribaMappingChoose.Visible = true;
                        btnAribaMappingClear.Visible = false;
                        return;
                    }
                    lblAribaMappingFileName.Text = "File - " + new FileInfo(openFileDialog.FileName).Name;
                    btnAribaMappingChoose.Visible = false;
                    btnAribaMappingClear.Visible = true;
                }
                catch (Exception ex)
                {
                    reportData.AribaMappingFileName = "";
                    lblError.Text = "Error : " + ex.Message;
                    lblAribaMappingFileName.Text = "";
                    btnAribaMappingChoose.Visible = true;
                    btnAribaMappingClear.Visible = false;
                }
            }
        }
        #endregion

        #region Clear Selected Files

        private void btnResouceSheetClear_Click(object sender, EventArgs e)
        {
            reportData.ResouceSheetFileName = "";
            lblResourceSheetFileName.Text = "";
            btnResouceSheetClear.Visible = false;
            btnResouceSheetChoose.Visible = true;

        }

        private void btnInvoiceSummaryClear_Click(object sender, EventArgs e)
        {
            reportData.InvoiceSummaryFileName = "";
            lblInvoiceSummaryFileName.Text = "";
            btnInvoiceSummaryClear.Visible = false;
            btnInvoiceSummaryChoose.Visible = true;
            return;
        }

        private void btnAribaMappingClear_Click(object sender, EventArgs e)
        {
            reportData.AribaMappingFileName = "";
            lblAribaMappingFileName.Text = "";
            btnAribaMappingClear.Visible = false;
            btnAribaMappingChoose.Visible = true;
            return;
        }
        #endregion

        #region Generate Report
        private void btnGenerateReport_Click(object sender, EventArgs e)
        {
            try
            {
                lblError.Text = "";
                lblError.BackColor = System.Drawing.Color.Red;
                //  reportData.AribaReportFileName = @"C:\\Users\\Administrator\\Desktop\\arya\\invoice_report (10).csv";
                //  reportData.InfosysReportFileName = @"C:\\Users\\Administrator\\Desktop\\arya\\Invoice Summary Nov 2016.xlsx";

                if (string.IsNullOrEmpty(reportData.InvoiceSummaryFileName))
                {
                    lblError.Text = "Please choose valid Invoice Summary file.";
                    btnInvoiceSummaryChoose.Visible = true;
                    return;
                }
                else if (string.IsNullOrEmpty(reportData.ResouceSheetFileName))
                {
                    lblError.Text = "Please choose valid Resouce Sheet file.";
                    btnResouceSheetChoose.Visible = true;
                    return;
                }

                else if (string.IsNullOrEmpty(reportData.AribaMappingFileName))
                {
                    lblError.Text = "Please choose valid Ariba Mapping file.";
                    btnAribaMappingChoose.Visible = true;
                    return;
                }
                else if (reportData.ResouceSheetFileName.Equals(reportData.InvoiceSummaryFileName) || reportData.ResouceSheetFileName.Equals(reportData.AribaMappingFileName) || reportData.ResouceSheetFileName.Equals(reportData.AribaMappingFileName))
                {
                    reportData.ResouceSheetFileName = reportData.InvoiceSummaryFileName = reportData.ResouceSheetFileName = "";
                    lblError.Text = "Both files seems to be same. Please verify";
                    lblResourceSheetFileName.Text = lblInvoiceSummaryFileName.Text = lblAribaMappingFileName.Text = "";
                    btnResouceSheetChoose.Visible = btnInvoiceSummaryChoose.Visible = btnAribaMappingChoose.Visible = true;
                    return;
                }
                else if (string.IsNullOrEmpty(reportData.ReportFileNameTemplate) || !new FileInfo(reportData.ReportFileNameTemplate).Exists)
                {
                    lblError.Text = "Please choose report template path";
                }
                else
                {

                    if (new FileInfo(reportData.ResouceSheetFileName).Exists && new FileInfo(reportData.InvoiceSummaryFileName).Exists && new FileInfo(reportData.AribaMappingFileName).Exists)
                    {

                        btnInvoiceSummaryChoose.Visible = false;
                        btnAribaMappingChoose.Visible = false;
                        btnGenerateReport.Visible = false;

                        btnInvoiceSummaryClear.Visible = false;
                        btnResouceSheetClear.Visible = false;
                        btnAribaMappingClear.Visible = false;
                        btnSelectTemplate.Visible = false;
                        lblError.Visible = false;
                        this.progressBar1.Visible = true;
                        //backgroundWorker1.RunWorkerAsync();
                        timer1.Enabled = true;
                        timer1.Start();
                        var result = reportData.GenerateReport(reportData);
                        this.Text = "Report Compare";

                        if (result)
                        {
                            InitializeComponent();
                            CustomInitialization();
                            lblError.Text = "Reported generated successfully";
                            lblError.BackColor = System.Drawing.Color.Green;
                            lblError.ForeColor = System.Drawing.Color.White;
                            lblError.Visible = true;
                            progressBar1.Visible = false;
                            FileInfo reportFileInfo = new FileInfo(reportData.ReportDirectory);

                            ProcessStartInfo startInfo = new ProcessStartInfo
                            {
                                Arguments = reportFileInfo.DirectoryName,
                                FileName = "explorer.exe"
                            };
                            Process.Start(startInfo);
                            timer1.Enabled = false;
                            timer1.Stop();
                            return;
                        }
                    }

                    else
                    {
                        InitializeComponent();
                        CustomInitialization();
                        timer1.Enabled = false;
                        timer1.Stop();
                        lblError.Text = "Please choose valid files.";
                        lblError.Visible = true;
                        return;
                    }
                }

            }
            catch (Exception ex)
            {
                InitializeComponent();
                CustomInitialization();
                lblError.Text = "Please try again. Error: " + ex.Message;
                lblError.Visible = true;
                timer1.Enabled = false;
                timer1.Stop();
                return;
            }

        }
        #endregion

        #region Background Process
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            for (int i = 1; i <= 100; i++)
            {
                // Wait 100 milliseconds.
                Thread.Sleep(100);
                // Report progress.
                backgroundWorker1.ReportProgress(i);
            }
        }
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            progressBar1.Value = e.ProgressPercentage;
            if (progressBar1.Value == 100)
            {
                this.progressBar1.Visible = false;
                btnInvoiceSummaryChoose.Visible = true;
                btnAribaMappingChoose.Visible = true;
                btnGenerateReport.Visible = true;
                lblError.Visible = true;
                lblInvoiceSummaryFileName.Text = lblAribaMappingFileName.Text = "";
                progressBar1.Value = 0;
            }
            // Set the text.
            // this.Text = "Report Comapre ( " + e.ProgressPercentage.ToString() + "% )";
        }

        #endregion

        private void btnSelectTemplate_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK)
            {
                reportData.ReportFileNameTemplate = openFileDialog.FileName;
                if (!string.IsNullOrEmpty(reportData.ReportFileNameTemplate) && !AllowedExtensions.Contains(reportData.ReportFileNameTemplate.Split('.')[reportData.ReportFileNameTemplate.Split('.').Length - 1].ToLower()))
                {
                    reportData.ReportFileNameTemplate = "";
                    lblTempltePath.Text = "Template Path - ";
                    lblError.Text = "Vaild for .xlsx, .xlsm, .xls, .csv files";
                    btnSelectTemplate.Text = "Select";
                }

                Properties.Settings.Default.TemplatePath = reportData.ReportFileNameTemplate;
                Properties.Settings.Default.Save();
                lblTempltePath.Text = "Template Path - " + reportData.ReportFileNameTemplate;
                btnSelectTemplate.Text = "Change";
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Value = reportData.ProgressValue;
        }
    }
    public class ReportData
    {
        public string InvoiceSummaryFileName { get; set; }
        public string ResouceSheetFileName { get; set; }
        public string AribaMappingFileName { get; set; }
        public List<string> ReportFileName { get; set; }
        public string ReportFileNameTemplate { get; set; }
        public List<List<string>> ResultList { get; set; }
        public string ReportDirectory { get; set; }
        public int ReportCount { get; set; }
        public int ProgressValue { get; internal set; }

        public bool GenerateReport(ReportData reportData)
        {
            //Microsoft Excel 14 object in references-> COM tab
            try
            {

                //Create COM Objects. Create a COM object for everything that is referenced
                reportData.ResultList = new List<List<string>>();
                reportData.ReportFileName = new List<string>();
                Excel.Application xlApp = new Excel.Application();

                Excel.Workbook invoiceSummarySheetXLWorkbook = xlApp.Workbooks.Open(reportData.InvoiceSummaryFileName, ReadOnly: false);
                Excel._Worksheet invoiceSummarySheetXLWorksheet = invoiceSummarySheetXLWorkbook.Sheets[1];
                Excel.Range invoiceSummarySheetXLRange = invoiceSummarySheetXLWorksheet.UsedRange.Columns["B:B", Type.Missing];
                invoiceSummarySheetXLRange.Copy(Type.Missing);

                Excel.Workbook resourceSheetXLWorkbook = xlApp.Workbooks.Open(reportData.ResouceSheetFileName, ReadOnly: false);
                Excel._Worksheet resourceSheetXLWorksheet = resourceSheetXLWorkbook.Sheets[1];
                Excel.Range resourceSheetXLRange = resourceSheetXLWorksheet.UsedRange.Columns["B:B", Type.Missing];
                resourceSheetXLRange.Copy(Type.Missing);

                Excel.Workbook aribaSheetXLWorkbook = xlApp.Workbooks.Open(reportData.AribaMappingFileName, ReadOnly: false);
                Excel._Worksheet aribaDiscountSheetXLWorksheet = aribaSheetXLWorkbook.Sheets[1];
                Excel.Range aribaDiscountSheetXLRange = aribaDiscountSheetXLWorksheet.UsedRange;
                aribaDiscountSheetXLRange.Copy(Type.Missing);
                Excel._Worksheet aribaSuntrustLineSheetXLWorksheet = aribaSheetXLWorkbook.Sheets[2];
                Excel.Range aribaSuntrustLineXLRange = aribaSuntrustLineSheetXLWorksheet.UsedRange;
                aribaSuntrustLineXLRange.Copy(Type.Missing);

                //excel is not zero based!!
                //starting form index 2 ignoring header
                object misValue = System.Reflection.Missing.Value;
                for (int invoiceSummaryRow = 2; invoiceSummaryRow <= invoiceSummarySheetXLRange.Rows.Count; invoiceSummaryRow++)
                {
                    reportData.ProgressValue = (invoiceSummaryRow / invoiceSummarySheetXLRange.Rows.Count) * 100;
                    if (invoiceSummarySheetXLRange.Cells[invoiceSummaryRow, 1] == null || invoiceSummarySheetXLRange.Cells[invoiceSummaryRow, 1].Value == null || invoiceSummarySheetXLRange.Cells[invoiceSummaryRow, 1].Value.ToString() == string.Empty)
                    {
                        break;
                    }
                    string invoiceNumber = invoiceSummarySheetXLRange.Cells[invoiceSummaryRow, 1].Value.ToString();
                    if (!string.IsNullOrEmpty(invoiceNumber))
                    {
                        var contractID = invoiceSummarySheetXLWorksheet.get_Range("A" + invoiceSummaryRow, "A" + invoiceSummaryRow).Cells[1, 1].Value2.ToString();
                        ReadResourceList(reportData, resourceSheetXLRange, aribaDiscountSheetXLRange, aribaSuntrustLineXLRange, invoiceNumber, contractID, xlApp);
                    }

                }
                if (reportData.ResultList.Count > 0)
                {
                    SaveReport(xlApp, reportData);
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();


                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background

                Marshal.ReleaseComObject(invoiceSummarySheetXLRange);
                Marshal.ReleaseComObject(invoiceSummarySheetXLWorksheet);
                Marshal.ReleaseComObject(invoiceSummarySheetXLWorkbook);

                Marshal.ReleaseComObject(resourceSheetXLRange);
                Marshal.ReleaseComObject(resourceSheetXLWorksheet);
                Marshal.ReleaseComObject(resourceSheetXLWorkbook);

                Marshal.ReleaseComObject(aribaDiscountSheetXLRange);
                Marshal.ReleaseComObject(aribaDiscountSheetXLWorksheet);
                Marshal.ReleaseComObject(aribaSuntrustLineXLRange);
                Marshal.ReleaseComObject(aribaSuntrustLineSheetXLWorksheet);
                Marshal.ReleaseComObject(aribaSheetXLWorkbook);


                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return true;

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }


        private void ReadResourceList(ReportData reportData, Range resourceSheetXLRange, Range aribaDiscountSheetXLRange, Range aribaSuntrustLineXLRange, string invoiceNumber, string contractID, Excel.Application xlApp)
        {
            try
            {

                var activityDescription = string.Empty;
                var lineNumber = string.Empty;
                for (int resourceRow = 2; resourceRow <= resourceSheetXLRange.Rows.Count; resourceRow++)
                {
                    if (resourceSheetXLRange.Cells[resourceRow, 1] == null || resourceSheetXLRange.Cells[resourceRow, 1].Value == null || resourceSheetXLRange.Cells[resourceRow, 1].Value.ToString() == string.Empty)
                    {
                        continue;
                    }
                    string resourceInvoiceNumber = resourceSheetXLRange.Cells[resourceRow, 1].Value.ToString();
                    if (invoiceNumber.Equals(resourceInvoiceNumber))
                    {
                        var resourceFileReportList = new List<string>();
                        var discount = string.Empty;
                        resourceFileReportList.Add(resourceInvoiceNumber); // Invoice Number -> invoiceID
                        var tempDate = resourceSheetXLRange.get_Range("B" + resourceRow, "B" + resourceRow).Cells[1, 1].Value.ToString();
                        resourceFileReportList.Add(tempDate.Split('/')[0] + '/' + tempDate.Split('/')[1] + '/' + tempDate.Split('/')[2].Substring(0, 4));// Billing Date -> invoiceDate
                        resourceFileReportList.Add(contractID);  //Contract # -> contractNumber

                        resourceFileReportList.Add(resourceSheetXLRange.get_Range("I" + resourceRow, "I" + resourceRow).Cells[1, 1].Value2.ToString()); // Work Effort
                        resourceFileReportList.Add(resourceSheetXLRange.get_Range("K" + resourceRow, "K" + resourceRow).Cells[1, 1].Value2.ToString()); // Rate
                        if (!activityDescription.ToLower().Equals(resourceSheetXLRange.get_Range("H" + resourceRow, "H" + resourceRow).Cells[1, 1].Value2.ToString().ToLower()))
                        {
                            discount = string.Empty;
                            lineNumber = string.Empty;
                            activityDescription = resourceSheetXLRange.get_Range("H" + resourceRow, "H" + resourceRow).Cells[1, 1].Value2.ToString(); // Activity Description
                            ReadAribaList(reportData, aribaDiscountSheetXLRange, aribaSuntrustLineXLRange, activityDescription, contractID, ref discount, ref lineNumber);
                        }
                        resourceFileReportList.Add(lineNumber);
                        resourceFileReportList.Add(activityDescription);
                        resourceFileReportList.Add(resourceSheetXLRange.get_Range("F" + resourceRow, "F" + resourceRow).Cells[1, 1].Value2.ToString()); // Employee Name
                        tempDate = resourceSheetXLRange.get_Range("T" + resourceRow, "T" + resourceRow).Cells[1, 1].Value.ToString();
                        resourceFileReportList.Add(tempDate.Split('/')[0] + '/' + tempDate.Split('/')[1] + '/' + tempDate.Split('/')[2].Substring(0, 4));// Start Date
                        tempDate = resourceSheetXLRange.get_Range("U" + resourceRow, "U" + resourceRow).Cells[1, 1].Value.ToString();
                        resourceFileReportList.Add(tempDate.Split('/')[0] + '/' + tempDate.Split('/')[1] + '/' + tempDate.Split('/')[2].Substring(0, 4)); // End Date
                        resourceFileReportList.Add(resourceSheetXLRange.get_Range("E" + resourceRow, "E" + resourceRow).Cells[1, 1].Value2.ToString()); // Employee No
                        resourceFileReportList.Add(resourceSheetXLRange.get_Range("P" + resourceRow, "P" + resourceRow).Cells[1, 1].Value2.ToString()); // Amount in USD
                        reportData.ResultList.Add(resourceFileReportList);
                        if (reportData.ResultList.Count == reportData.ReportCount - 1)
                        {
                            SaveReport(xlApp, reportData);
                            reportData.ResultList = new List<List<string>>();
                        }

                    }
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void ReadAribaList(ReportData reportData, Range aribaDiscountSheetXLRange, Range aribaSuntrustLineXLRange, string activityDescription, string contractID, ref string discount, ref string lineNumber)
        {
            try
            {

                for (int aribaRow = 6; aribaRow <= aribaSuntrustLineXLRange.Rows.Count; aribaRow++)
                {
                    if (aribaSuntrustLineXLRange.Cells[aribaRow, 2] == null || aribaSuntrustLineXLRange.Cells[aribaRow, 2].Value == null || aribaSuntrustLineXLRange.Cells[aribaRow, 2].Value.ToString() == string.Empty)
                    {
                        continue;
                    }
                    else
                    {
                        string aribaContractID = aribaSuntrustLineXLRange.Cells[aribaRow, 2].Value.ToString();
                        if (contractID.Equals(aribaContractID))
                        {
                            if (aribaSuntrustLineXLRange.Cells[aribaRow, 4] == null || aribaSuntrustLineXLRange.Cells[aribaRow, 4].Value == null || aribaSuntrustLineXLRange.Cells[aribaRow, 4].Value.ToString() == string.Empty)
                            {
                                continue;
                            }
                            else
                            {
                                var test = aribaSuntrustLineXLRange.Cells[aribaRow, 4].Value.ToString();
                                if (activityDescription.ToLower().Equals(aribaSuntrustLineXLRange.Cells[aribaRow, 4].Value.ToString().ToLower()))
                                {
                                    lineNumber = aribaSuntrustLineXLRange.Cells[aribaRow, 8].Value.ToString();
                                    break;
                                }
                            }
                        }
                    }
                }
                //aribaDiscountSheetXLRange.AutoFilter(3, contractID, Excel.XlAutoFilterOperator.xlAnd, System.Reflection.Missing.Value, false);

                //////and get only visible cells after the filter.
                //var discountResult = aribaDiscountSheetXLRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing);
                //if (discountResult == null)
                //{
                //}
                //else
                //{
                //    foreach (Excel.Range row in discountResult.Rows)
                //    {

                //        if (!string.IsNullOrEmpty(row.Cells[1, 7].Value.ToString()))
                //        {
                //            discount = row.Cells[1, 7].Value.ToString();
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {

                throw ex;
            }


        }
        private void SaveReport(Excel.Application xlApp, ReportData reportData)
        {
            try
            {
                Excel.Workbook reportFileNameTemplateSheetXLWorkbook = xlApp.Workbooks.Open(reportData.ReportFileNameTemplate, ReadOnly: false);
                Excel._Worksheet reportFileNameTemplateSheetXLWorksheet = reportFileNameTemplateSheetXLWorkbook.Sheets[1];
                Excel.Range reportFileNameTemplateSheetXLRange = reportFileNameTemplateSheetXLWorksheet.UsedRange;
                reportFileNameTemplateSheetXLRange.Copy(Type.Missing);
                object misValue = System.Reflection.Missing.Value;
                var invoiceLineIDCount = 1;
                var tempInvoiceID = string.Empty;
                double subtotalAmount = 0;
                int reportCellStartIndex = 2;
                foreach (var list in reportData.ResultList)
                {
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 1] = list[0]; // a - invoiceID
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 2] = list[1]; // b - invoiceDate
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 3] = list[2]; // c - contractNumber
                    for (var iCount = 4; iCount < 34; iCount++)
                    {
                        reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, iCount] = reportFileNameTemplateSheetXLRange.Cells[2, iCount]; // d - ag
                    }
                    if (tempInvoiceID.Equals(list[0]))
                    {
                        invoiceLineIDCount++;
                    }
                    else
                    {
                        tempInvoiceID = list[0];
                        if (invoiceLineIDCount > 1)
                        {
                            var currentRow = reportCellStartIndex - 1;
                            for (var iCount = invoiceLineIDCount; iCount > 1; iCount--)
                            {
                                reportFileNameTemplateSheetXLRange.Cells[currentRow, 63] = subtotalAmount;
                                reportFileNameTemplateSheetXLRange.Cells[currentRow, 72] = subtotalAmount;
                                reportFileNameTemplateSheetXLRange.Cells[currentRow, 73] = subtotalAmount;
                            }
                        }
                        invoiceLineIDCount = 1;
                        subtotalAmount = 0;

                    }
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 40] = invoiceLineIDCount; // an - 41 - invoiceLineID count
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 41] = list[3]; // ao - quantity
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 42] = list[4]; // ap - unitOfMeasure
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 43] = "HUR"; // aq - unitPriceAmount
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 44] = list[5]; // ar - lineNumber
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 45] = list[6] + " - " + list[7] + " - " + list[8] + " TO " + list[9]; // as - itemDescription
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 46] = list[10]; // at - supplierPartID
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 47] = list[11]; // au - itemSubtotalAmount
                    subtotalAmount += Convert.ToDouble(list[11]);
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 63] = subtotalAmount; // as - subtotalAmount
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 64] = 0; // as - taxAmount
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex, 72] = subtotalAmount; // as - subtotalAmount
                    reportFileNameTemplateSheetXLRange.Cells[reportCellStartIndex++, 73] = subtotalAmount; // as - subtotalAmount

                }
                reportFileNameTemplateSheetXLRange.Columns.AutoFit();
                var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "Invoice_Report");
                DirectoryInfo directoryInfo = new DirectoryInfo(path);
                if (!directoryInfo.Exists)
                {
                    directoryInfo.Create();
                }
                reportData.ReportDirectory = directoryInfo.FullName;
                var tempFileName = directoryInfo.FullName + "\\Invoice_Report" + DateTime.Now.ToString("_yyyyMMddHHmmss");
                reportData.ReportFileName.Add(tempFileName);
                reportFileNameTemplateSheetXLWorkbook.SaveAs(tempFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                reportFileNameTemplateSheetXLWorkbook.SaveAs(tempFileName, Excel.XlFileFormat.xlCSV, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                Marshal.ReleaseComObject(reportFileNameTemplateSheetXLRange);
                Marshal.ReleaseComObject(reportFileNameTemplateSheetXLWorksheet);
                reportFileNameTemplateSheetXLWorkbook.Close(false);
                Marshal.ReleaseComObject(reportFileNameTemplateSheetXLWorkbook);
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }




    }
}
