﻿using System.ComponentModel;
namespace Report_Compare
{
    partial class ReportCompare
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportCompare));
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.lblChooseFiles = new System.Windows.Forms.Label();
            this.lblInvoiceSummaryText = new System.Windows.Forms.Label();
            this.lblAribaMappingText = new System.Windows.Forms.Label();
            this.btnInvoiceSummaryChoose = new System.Windows.Forms.Button();
            this.btnAribaMappingChoose = new System.Windows.Forms.Button();
            this.btnGenerateReport = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.lblInvoiceSummaryFileName = new System.Windows.Forms.Label();
            this.lblAribaMappingFileName = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.btnInvoiceSummaryClear = new System.Windows.Forms.Button();
            this.btnAribaMappingClear = new System.Windows.Forms.Button();
            this.lblResourceSheetText = new System.Windows.Forms.Label();
            this.btnResouceSheetChoose = new System.Windows.Forms.Button();
            this.lblResourceSheetFileName = new System.Windows.Forms.Label();
            this.btnResouceSheetClear = new System.Windows.Forms.Button();
            this.lblTempltePath = new System.Windows.Forms.Label();
            this.btnSelectTemplate = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.lblError = new System.Windows.Forms.Label();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.nUDReportCount = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nUDReportCount)).BeginInit();
            this.SuspendLayout();
            // 
            // lblChooseFiles
            // 
            this.lblChooseFiles.AutoSize = true;
            this.lblChooseFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblChooseFiles.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.lblChooseFiles.Location = new System.Drawing.Point(202, 10);
            this.lblChooseFiles.Name = "lblChooseFiles";
            this.lblChooseFiles.Size = new System.Drawing.Size(113, 20);
            this.lblChooseFiles.TabIndex = 0;
            this.lblChooseFiles.Text = "Choose Files";
            // 
            // lblInvoiceSummaryText
            // 
            this.lblInvoiceSummaryText.AutoSize = true;
            this.lblInvoiceSummaryText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInvoiceSummaryText.ForeColor = System.Drawing.Color.White;
            this.lblInvoiceSummaryText.Location = new System.Drawing.Point(26, 46);
            this.lblInvoiceSummaryText.Name = "lblInvoiceSummaryText";
            this.lblInvoiceSummaryText.Size = new System.Drawing.Size(130, 17);
            this.lblInvoiceSummaryText.TabIndex = 1;
            this.lblInvoiceSummaryText.Text = "Invoice Summary";
            this.lblInvoiceSummaryText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblAribaMappingText
            // 
            this.lblAribaMappingText.AutoSize = true;
            this.lblAribaMappingText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAribaMappingText.ForeColor = System.Drawing.Color.White;
            this.lblAribaMappingText.Location = new System.Drawing.Point(26, 163);
            this.lblAribaMappingText.Name = "lblAribaMappingText";
            this.lblAribaMappingText.Size = new System.Drawing.Size(112, 17);
            this.lblAribaMappingText.TabIndex = 2;
            this.lblAribaMappingText.Text = "Ariba Mapping";
            this.lblAribaMappingText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnInvoiceSummaryChoose
            // 
            this.btnInvoiceSummaryChoose.BackColor = System.Drawing.Color.White;
            this.btnInvoiceSummaryChoose.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInvoiceSummaryChoose.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnInvoiceSummaryChoose.Location = new System.Drawing.Point(395, 42);
            this.btnInvoiceSummaryChoose.Name = "btnInvoiceSummaryChoose";
            this.btnInvoiceSummaryChoose.Size = new System.Drawing.Size(73, 23);
            this.btnInvoiceSummaryChoose.TabIndex = 3;
            this.btnInvoiceSummaryChoose.Text = "Choose";
            this.btnInvoiceSummaryChoose.UseVisualStyleBackColor = false;
            this.btnInvoiceSummaryChoose.Click += new System.EventHandler(this.btnInvoiceSummaryChoose_Click);
            // 
            // btnAribaMappingChoose
            // 
            this.btnAribaMappingChoose.BackColor = System.Drawing.Color.White;
            this.btnAribaMappingChoose.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAribaMappingChoose.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnAribaMappingChoose.Location = new System.Drawing.Point(395, 160);
            this.btnAribaMappingChoose.Name = "btnAribaMappingChoose";
            this.btnAribaMappingChoose.Size = new System.Drawing.Size(73, 23);
            this.btnAribaMappingChoose.TabIndex = 4;
            this.btnAribaMappingChoose.Text = "Choose";
            this.btnAribaMappingChoose.UseVisualStyleBackColor = false;
            this.btnAribaMappingChoose.Click += new System.EventHandler(this.btnAribaMappingChoose_Click);
            // 
            // btnGenerateReport
            // 
            this.btnGenerateReport.BackColor = System.Drawing.Color.White;
            this.btnGenerateReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGenerateReport.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnGenerateReport.Location = new System.Drawing.Point(153, 333);
            this.btnGenerateReport.Name = "btnGenerateReport";
            this.btnGenerateReport.Size = new System.Drawing.Size(215, 23);
            this.btnGenerateReport.TabIndex = 5;
            this.btnGenerateReport.Text = "Compare and Generate Report";
            this.btnGenerateReport.UseVisualStyleBackColor = false;
            this.btnGenerateReport.Click += new System.EventHandler(this.btnGenerateReport_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.BackColor = System.Drawing.Color.Gainsboro;
            this.progressBar1.ForeColor = System.Drawing.Color.LimeGreen;
            this.progressBar1.Location = new System.Drawing.Point(153, 360);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(215, 23);
            this.progressBar1.Step = 1;
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar1.TabIndex = 1;
            this.progressBar1.UseWaitCursor = true;
            this.progressBar1.Visible = false;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            // 
            // lblInvoiceSummaryFileName
            // 
            this.lblInvoiceSummaryFileName.AutoSize = true;
            this.lblInvoiceSummaryFileName.ForeColor = System.Drawing.Color.White;
            this.lblInvoiceSummaryFileName.Location = new System.Drawing.Point(28, 71);
            this.lblInvoiceSummaryFileName.Name = "lblInvoiceSummaryFileName";
            this.lblInvoiceSummaryFileName.Size = new System.Drawing.Size(107, 13);
            this.lblInvoiceSummaryFileName.TabIndex = 8;
            this.lblInvoiceSummaryFileName.Text = "Invoice Summary File";
            // 
            // lblAribaMappingFileName
            // 
            this.lblAribaMappingFileName.AutoSize = true;
            this.lblAribaMappingFileName.ForeColor = System.Drawing.Color.White;
            this.lblAribaMappingFileName.Location = new System.Drawing.Point(29, 191);
            this.lblAribaMappingFileName.Name = "lblAribaMappingFileName";
            this.lblAribaMappingFileName.Size = new System.Drawing.Size(94, 13);
            this.lblAribaMappingFileName.TabIndex = 9;
            this.lblAribaMappingFileName.Text = "Ariba Mapping File";
            // 
            // btnInvoiceSummaryClear
            // 
            this.btnInvoiceSummaryClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInvoiceSummaryClear.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnInvoiceSummaryClear.Location = new System.Drawing.Point(408, 66);
            this.btnInvoiceSummaryClear.Name = "btnInvoiceSummaryClear";
            this.btnInvoiceSummaryClear.Size = new System.Drawing.Size(60, 23);
            this.btnInvoiceSummaryClear.TabIndex = 10;
            this.btnInvoiceSummaryClear.Text = "Clear";
            this.btnInvoiceSummaryClear.UseVisualStyleBackColor = true;
            this.btnInvoiceSummaryClear.Click += new System.EventHandler(this.btnInvoiceSummaryClear_Click);
            // 
            // btnAribaMappingClear
            // 
            this.btnAribaMappingClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAribaMappingClear.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnAribaMappingClear.Location = new System.Drawing.Point(408, 186);
            this.btnAribaMappingClear.Name = "btnAribaMappingClear";
            this.btnAribaMappingClear.Size = new System.Drawing.Size(60, 23);
            this.btnAribaMappingClear.TabIndex = 11;
            this.btnAribaMappingClear.Text = "Clear";
            this.btnAribaMappingClear.UseVisualStyleBackColor = true;
            this.btnAribaMappingClear.Click += new System.EventHandler(this.btnAribaMappingClear_Click);
            // 
            // lblResourceSheetText
            // 
            this.lblResourceSheetText.AutoSize = true;
            this.lblResourceSheetText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblResourceSheetText.ForeColor = System.Drawing.Color.White;
            this.lblResourceSheetText.Location = new System.Drawing.Point(27, 105);
            this.lblResourceSheetText.Name = "lblResourceSheetText";
            this.lblResourceSheetText.Size = new System.Drawing.Size(124, 17);
            this.lblResourceSheetText.TabIndex = 1;
            this.lblResourceSheetText.Text = "Resource Sheet";
            this.lblResourceSheetText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnResouceSheetChoose
            // 
            this.btnResouceSheetChoose.BackColor = System.Drawing.Color.White;
            this.btnResouceSheetChoose.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnResouceSheetChoose.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnResouceSheetChoose.Location = new System.Drawing.Point(395, 101);
            this.btnResouceSheetChoose.Name = "btnResouceSheetChoose";
            this.btnResouceSheetChoose.Size = new System.Drawing.Size(73, 23);
            this.btnResouceSheetChoose.TabIndex = 3;
            this.btnResouceSheetChoose.Text = "Choose";
            this.btnResouceSheetChoose.UseVisualStyleBackColor = false;
            this.btnResouceSheetChoose.Click += new System.EventHandler(this.btnResouceSheetChoose_Click);
            // 
            // lblResourceSheetFileName
            // 
            this.lblResourceSheetFileName.AutoSize = true;
            this.lblResourceSheetFileName.ForeColor = System.Drawing.Color.White;
            this.lblResourceSheetFileName.Location = new System.Drawing.Point(29, 130);
            this.lblResourceSheetFileName.Name = "lblResourceSheetFileName";
            this.lblResourceSheetFileName.Size = new System.Drawing.Size(103, 13);
            this.lblResourceSheetFileName.TabIndex = 8;
            this.lblResourceSheetFileName.Text = "Resource Sheet File";
            // 
            // btnResouceSheetClear
            // 
            this.btnResouceSheetClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnResouceSheetClear.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnResouceSheetClear.Location = new System.Drawing.Point(408, 125);
            this.btnResouceSheetClear.Name = "btnResouceSheetClear";
            this.btnResouceSheetClear.Size = new System.Drawing.Size(60, 23);
            this.btnResouceSheetClear.TabIndex = 10;
            this.btnResouceSheetClear.Text = "Clear";
            this.btnResouceSheetClear.UseVisualStyleBackColor = true;
            this.btnResouceSheetClear.Click += new System.EventHandler(this.btnResouceSheetClear_Click);
            // 
            // lblTempltePath
            // 
            this.lblTempltePath.AutoSize = true;
            this.lblTempltePath.ForeColor = System.Drawing.Color.White;
            this.lblTempltePath.Location = new System.Drawing.Point(3, 0);
            this.lblTempltePath.Name = "lblTempltePath";
            this.lblTempltePath.Padding = new System.Windows.Forms.Padding(2);
            this.lblTempltePath.Size = new System.Drawing.Size(89, 17);
            this.lblTempltePath.TabIndex = 12;
            this.lblTempltePath.Text = "Template Path : ";
            // 
            // btnSelectTemplate
            // 
            this.btnSelectTemplate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSelectTemplate.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnSelectTemplate.Location = new System.Drawing.Point(370, 3);
            this.btnSelectTemplate.Name = "btnSelectTemplate";
            this.btnSelectTemplate.Size = new System.Drawing.Size(61, 23);
            this.btnSelectTemplate.TabIndex = 13;
            this.btnSelectTemplate.Text = "Select";
            this.btnSelectTemplate.UseVisualStyleBackColor = true;
            this.btnSelectTemplate.Click += new System.EventHandler(this.btnSelectTemplate_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 84.75751F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.24249F));
            this.tableLayoutPanel1.Controls.Add(this.lblTempltePath, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.btnSelectTemplate, 1, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(34, 223);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(434, 32);
            this.tableLayoutPanel1.TabIndex = 14;
            // 
            // lblError
            // 
            this.lblError.AutoSize = true;
            this.lblError.BackColor = System.Drawing.Color.Red;
            this.lblError.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblError.ForeColor = System.Drawing.Color.White;
            this.lblError.Location = new System.Drawing.Point(3, 0);
            this.lblError.Name = "lblError";
            this.lblError.Size = new System.Drawing.Size(47, 15);
            this.lblError.TabIndex = 7;
            this.lblError.Text = "lblError";
            this.lblError.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Controls.Add(this.lblError, 0, 0);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(34, 298);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(434, 29);
            this.tableLayoutPanel2.TabIndex = 15;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 2;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 62F));
            this.tableLayoutPanel3.Controls.Add(this.nUDReportCount, 1, 0);
            this.tableLayoutPanel3.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel3.Location = new System.Drawing.Point(34, 262);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 1;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(434, 30);
            this.tableLayoutPanel3.TabIndex = 16;
            // 
            // nUDReportCount
            // 
            this.nUDReportCount.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.nUDReportCount.Location = new System.Drawing.Point(375, 3);
            this.nUDReportCount.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.nUDReportCount.Minimum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.nUDReportCount.Name = "nUDReportCount";
            this.nUDReportCount.Size = new System.Drawing.Size(56, 20);
            this.nUDReportCount.TabIndex = 17;
            this.nUDReportCount.Value = new decimal(new int[] {
            50,
            0,
            0,
            0});
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(142, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "Number of rows in report";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // ReportCompare
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(539, 424);
            this.Controls.Add(this.tableLayoutPanel3);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.btnAribaMappingClear);
            this.Controls.Add(this.btnResouceSheetClear);
            this.Controls.Add(this.btnInvoiceSummaryClear);
            this.Controls.Add(this.lblAribaMappingFileName);
            this.Controls.Add(this.lblResourceSheetFileName);
            this.Controls.Add(this.lblInvoiceSummaryFileName);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnGenerateReport);
            this.Controls.Add(this.btnAribaMappingChoose);
            this.Controls.Add(this.btnResouceSheetChoose);
            this.Controls.Add(this.btnInvoiceSummaryChoose);
            this.Controls.Add(this.lblAribaMappingText);
            this.Controls.Add(this.lblResourceSheetText);
            this.Controls.Add(this.lblInvoiceSummaryText);
            this.Controls.Add(this.lblChooseFiles);
            this.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "ReportCompare";
            this.Text = "Report Compare";
            this.TransparencyKey = System.Drawing.SystemColors.Desktop;
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nUDReportCount)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Label lblChooseFiles;
        private System.Windows.Forms.Label lblInvoiceSummaryText;
        private System.Windows.Forms.Label lblAribaMappingText;
        private System.Windows.Forms.Button btnInvoiceSummaryChoose;
        private System.Windows.Forms.Button btnAribaMappingChoose;
        private System.Windows.Forms.Button btnGenerateReport;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label lblInvoiceSummaryFileName;
        private System.Windows.Forms.Label lblAribaMappingFileName;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btnInvoiceSummaryClear;
        private System.Windows.Forms.Button btnAribaMappingClear;
        private System.Windows.Forms.Label lblResourceSheetText;
        private System.Windows.Forms.Button btnResouceSheetChoose;
        private System.Windows.Forms.Label lblResourceSheetFileName;
        private System.Windows.Forms.Button btnResouceSheetClear;
        private System.Windows.Forms.Label lblTempltePath;
        private System.Windows.Forms.Button btnSelectTemplate;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label lblError;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown nUDReportCount;
        private System.Windows.Forms.Timer timer1;
    }
}

