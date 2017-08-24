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
                        lblError.Visible = false;
                        this.progressBar1.Visible = true;
                        backgroundWorker1.RunWorkerAsync();
                        var result = reportData.GenerateReport(reportData);
                        this.Text = "Report Compare";

                        if (result)
                        {
                            lblError.Text = "Reported generated successfully";
                            lblError.BackColor = System.Drawing.Color.Green;
                            FileInfo infyReportFileInfo = new FileInfo(reportData.ReportFileName);

                            ProcessStartInfo startInfo = new ProcessStartInfo
                            {
                                Arguments = infyReportFileInfo.DirectoryName + "\\MergeReport",
                                FileName = "explorer.exe"
                            };
                            Process.Start(startInfo);
                            return;
                        }
                    }

                    else
                    {
                        lblError.Text = "Please choose valid files.";
                        return;
                    }
                }

            }
            catch (Exception ex)
            {
                lblError.Text = "Please try again. Error: " + ex.Message;
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

    }
    public class ReportData
    {
        public string InvoiceSummaryFileName { get; set; }
        public string ResouceSheetFileName { get; set; }
        public string AribaMappingFileName { get; set; }
        public string ReportFileName { get; set; }
        public string ReportFileNameTemplate { get; set; }


        public bool GenerateReport(ReportData reportData)
        {
            //Microsoft Excel 14 object in references-> COM tab
            try
            {
                //FileInfo resourceSheetFileInfo = new FileInfo(reportData.ResouceSheetFileName);

                //Create COM Objects. Create a COM object for everything that is referenced
                List<List<string>> resultList = new List<List<string>>();
                Excel.Application xlApp = new Excel.Application();

                Excel.Workbook invoiceSummarySheetXLWorkbook = xlApp.Workbooks.Open(reportData.InvoiceSummaryFileName, ReadOnly: false);
                Excel._Worksheet invoiceSummarySheetXLWorksheet = invoiceSummarySheetXLWorkbook.Sheets[1];
                Excel.Range invoiceSummarySheetXLRange = invoiceSummarySheetXLWorksheet.UsedRange.Columns["B:B", Type.Missing];
                invoiceSummarySheetXLRange.Copy(Type.Missing);

                Excel.Workbook resourceSheetXLWorkbook = xlApp.Workbooks.Open(reportData.ResouceSheetFileName, ReadOnly: false);
                Excel._Worksheet resourceSheetXLWorksheet = resourceSheetXLWorkbook.Sheets[1];
                Excel.Range resourceSheetXLRange = resourceSheetXLWorksheet.UsedRange.Columns["B:B", Type.Missing];
                resourceSheetXLRange.Copy(Type.Missing);


                //Excel.Workbook aribaSheetXLWorkbook = xlApp.Workbooks.Open(reportData.AribaMappingFileName, ReadOnly: false);
                //Excel._Worksheet aribaSheetXLWorksheet = aribaSheetXLWorkbook.Sheets[1];
                //Excel.Range aribaSheetXLRange = aribaSheetXLWorksheet.UsedRange.Columns["B:B", Type.Missing];
                //resourceSheetXLRange.Copy(Type.Missing);



                //Excel.Workbook reportTemplateSheetXLWorkbook = xlApp.Workbooks.Open(reportData.AribaMappingFileName, ReadOnly: false);
                //Excel._Worksheet reportTemplateSheetXLWorksheet = reportTemplateSheetXLWorkbook.Sheets[1];

                //Excel.Range reportTemplateSheetXLRange = reportTemplateSheetXLWorksheet.UsedRange.Columns["B:B", Type.Missing];
                //resourceSheetXLRange.Copy(Type.Missing);

                //Excel.Range infosysReportXLStatusRange = resourceSheetXLWorksheet.UsedRange.Columns["J:J", Type.Missing];
                //infosysReportXLStatusRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                //infosysReportXLStatusRange.Cells[1, 1] = "Invoice ARIBA Status";
                //infosysReportXLStatusRange.Interior.Color = Excel.XlRgbColor.rgbYellowGreen;
                //Excel.Range infosysReportXLRemarksRange = resourceSheetXLWorksheet.UsedRange.Columns["K:K", Type.Missing];
                //infosysReportXLRemarksRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                //infosysReportXLRemarksRange.Cells[1, 1] = "Report Remarks";
                //infosysReportXLRemarksRange.Interior.Color = Excel.XlRgbColor.rgbYellowGreen;


                int invoiceSummaryRowCount = invoiceSummarySheetXLRange.Rows.Count;
                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                //starting form index 2 ignoring header
                List<string> list = null;
                for (int invoiceSummaryRow = 2; invoiceSummaryRow <= invoiceSummaryRowCount; invoiceSummaryRow++)
                {
                    list = new List<string>();
                    var iCount = 0;
                    if (resourceSheetXLRange.Cells[invoiceSummaryRow, 1] == null || resourceSheetXLRange.Cells[invoiceSummaryRow, 1].Value2 == null || resourceSheetXLRange.Cells[invoiceSummaryRow, 1].Value2.ToString() == string.Empty)
                    {
                        break;
                    }
                    string invoiceNumber = invoiceSummarySheetXLRange.Cells[invoiceSummaryRow, 1].Value2.ToString();
                    if (!string.IsNullOrEmpty(invoiceNumber))
                    {

                        invoiceNumber = invoiceNumber.Split('.')[0];
                        var xlRng = invoiceSummarySheetXLWorksheet.get_Range("A" + invoiceSummaryRow, "A" + invoiceSummaryRow).Cells[1, 1].Value2.ToString();
                        resourceSheetXLRange.AutoFilter(Field: 1, Criteria1: invoiceNumber);

                        //and get only visible cells after the filter.
                        var invoiceResult = resourceSheetXLRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing);
                        if (invoiceResult == null)
                        {
                        }
                        else
                        {
                            foreach (Excel.Range row in invoiceResult.Rows)
                            {
                                if (iCount++ > 0)
                                {
                                    list.Add(invoiceNumber); // Invoice Number -> invoiceID
                                    list.Add(row.Cells[1, 2].Value2.ToString()); // Billing Date -> invoiceDate
                                    list.Add(xlRng); //Contract # -> contractNumber
                                    list.Add(row.Cells[1, 9].Value2.ToString()); // Work Effort -> quantity
                                    list.Add(row.Cells[1, 11].Value2.ToString()); // Rate -> unitPriceAmount
                                    list.Add(row.Cells[1, 8].Value2.ToString() + " - " + row.Cells[1, 6].Value2.ToString() + " - " + row.Cells[1, 20].Value + " to " + row.Cells[1, 21].Value);
                                    list.Add(row.Cells[1, 5].Value2.ToString()); // Employee No -> supplierPartID
                                    list.Add(row.Cells[1, 19].Value2.ToString()); // Net Amount-> itemSubtotalAmount
                                }

                            }
                        }


                    }

                }
                //infosysReportXLStatusRange.Columns.AutoFit();
                //infosysReportXLRemarksRange.Columns.AutoFit();


                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();


                object misValue = System.Reflection.Missing.Value;
                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background

                Marshal.ReleaseComObject(invoiceSummarySheetXLRange);
                Marshal.ReleaseComObject(invoiceSummarySheetXLWorksheet);
                invoiceSummarySheetXLWorkbook.Close();
                Marshal.ReleaseComObject(resourceSheetXLWorkbook);

                Marshal.ReleaseComObject(resourceSheetXLRange);
                Marshal.ReleaseComObject(resourceSheetXLWorksheet);
                resourceSheetXLWorkbook.Close();
                Marshal.ReleaseComObject(resourceSheetXLWorkbook);

                //close and release
                DirectoryInfo directoryInfo = null;// new DirectoryInfo(resourceSheetFileInfo.DirectoryName + "\\MergeReport");
                if (!directoryInfo.Exists)
                {
                    directoryInfo.Create();
                }
                var fileName = directoryInfo.FullName + "\\MergeReport" + DateTime.Now.ToString("_yyyyMMddHHmmss");
                // infosysReportXLWorkbook.SaveAs(tee.DirectoryName + "hems1", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                //resourceSheetXLWorkbook.SaveAs(fileName, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);


                //Marshal.ReleaseComObject(aribaReportXLInvoiceRange);
                //Marshal.ReleaseComObject(aribaReportXLInvoiceStatusRange);
                //Marshal.ReleaseComObject(aribaReportXLWorksheet);
                //aribaReportXLWorkbook.Close(false);
                //Marshal.ReleaseComObject(aribaReportXLWorkbook);


                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return true;

            }
            catch (Exception ex)
            {
                return false;
            }

        }
    }
}
