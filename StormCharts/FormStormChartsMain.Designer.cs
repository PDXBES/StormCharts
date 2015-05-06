namespace StormCharts
{
    partial class FormStormChartsMain
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
            this.buttonCreateStormCharts = new System.Windows.Forms.Button();
            this.dateTimePickerStartTime = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerEndTime = new System.Windows.Forms.DateTimePicker();
            this.labelStartTime = new System.Windows.Forms.Label();
            this.labelEndTime = new System.Windows.Forms.Label();
            this.numericUpDownH2Number = new System.Windows.Forms.NumericUpDown();
            this.labelH2Number = new System.Windows.Forms.Label();
            this.backgroundWorkerSingle = new System.ComponentModel.BackgroundWorker();
            this.lblModelRunStatus = new System.Windows.Forms.Label();
            this.pnlCancelBackgroundWorker = new System.Windows.Forms.Panel();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonCreateChartOneStormManyGauges = new System.Windows.Forms.Button();
            this.backgroundWorkerMultiple = new System.ComponentModel.BackgroundWorker();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownH2Number)).BeginInit();
            this.pnlCancelBackgroundWorker.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonCreateStormCharts
            // 
            this.buttonCreateStormCharts.Location = new System.Drawing.Point(280, 256);
            this.buttonCreateStormCharts.Name = "buttonCreateStormCharts";
            this.buttonCreateStormCharts.Size = new System.Drawing.Size(297, 53);
            this.buttonCreateStormCharts.TabIndex = 0;
            this.buttonCreateStormCharts.Text = "Create Storm Charts - Multiple Storms, One Gauge";
            this.buttonCreateStormCharts.UseVisualStyleBackColor = true;
            this.buttonCreateStormCharts.Click += new System.EventHandler(this.buttonCreateStormCharts_Click);
            // 
            // dateTimePickerStartTime
            // 
            this.dateTimePickerStartTime.CustomFormat = "MM/dd/yyyy HH:mm";
            this.dateTimePickerStartTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerStartTime.Location = new System.Drawing.Point(12, 67);
            this.dateTimePickerStartTime.Name = "dateTimePickerStartTime";
            this.dateTimePickerStartTime.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerStartTime.TabIndex = 1;
            // 
            // dateTimePickerEndTime
            // 
            this.dateTimePickerEndTime.CustomFormat = "MM/dd/yyyy HH:mm";
            this.dateTimePickerEndTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerEndTime.Location = new System.Drawing.Point(280, 67);
            this.dateTimePickerEndTime.Name = "dateTimePickerEndTime";
            this.dateTimePickerEndTime.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerEndTime.TabIndex = 2;
            // 
            // labelStartTime
            // 
            this.labelStartTime.AutoSize = true;
            this.labelStartTime.Location = new System.Drawing.Point(9, 41);
            this.labelStartTime.Name = "labelStartTime";
            this.labelStartTime.Size = new System.Drawing.Size(58, 13);
            this.labelStartTime.TabIndex = 3;
            this.labelStartTime.Text = "Start Time:";
            // 
            // labelEndTime
            // 
            this.labelEndTime.AutoSize = true;
            this.labelEndTime.Location = new System.Drawing.Point(277, 41);
            this.labelEndTime.Name = "labelEndTime";
            this.labelEndTime.Size = new System.Drawing.Size(55, 13);
            this.labelEndTime.TabIndex = 4;
            this.labelEndTime.Text = "End Time:";
            // 
            // numericUpDownH2Number
            // 
            this.numericUpDownH2Number.Location = new System.Drawing.Point(12, 124);
            this.numericUpDownH2Number.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.numericUpDownH2Number.Name = "numericUpDownH2Number";
            this.numericUpDownH2Number.Size = new System.Drawing.Size(120, 20);
            this.numericUpDownH2Number.TabIndex = 5;
            // 
            // labelH2Number
            // 
            this.labelH2Number.AutoSize = true;
            this.labelH2Number.Location = new System.Drawing.Point(9, 108);
            this.labelH2Number.Name = "labelH2Number";
            this.labelH2Number.Size = new System.Drawing.Size(60, 13);
            this.labelH2Number.TabIndex = 6;
            this.labelH2Number.Text = "h2 number:";
            // 
            // backgroundWorkerSingle
            // 
            this.backgroundWorkerSingle.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerSingle_DoWork);
            this.backgroundWorkerSingle.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerSingle_RunWorkerCompleted);
            // 
            // lblModelRunStatus
            // 
            this.lblModelRunStatus.Location = new System.Drawing.Point(9, 9);
            this.lblModelRunStatus.Name = "lblModelRunStatus";
            this.lblModelRunStatus.Size = new System.Drawing.Size(246, 41);
            this.lblModelRunStatus.TabIndex = 1;
            this.lblModelRunStatus.Text = "Generating Report";
            // 
            // pnlCancelBackgroundWorker
            // 
            this.pnlCancelBackgroundWorker.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnlCancelBackgroundWorker.Controls.Add(this.progressBar1);
            this.pnlCancelBackgroundWorker.Controls.Add(this.lblModelRunStatus);
            this.pnlCancelBackgroundWorker.Controls.Add(this.buttonCancel);
            this.pnlCancelBackgroundWorker.Location = new System.Drawing.Point(12, 150);
            this.pnlCancelBackgroundWorker.Name = "pnlCancelBackgroundWorker";
            this.pnlCancelBackgroundWorker.Size = new System.Drawing.Size(565, 100);
            this.pnlCancelBackgroundWorker.TabIndex = 24;
            this.pnlCancelBackgroundWorker.Visible = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(3, 35);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(555, 33);
            this.progressBar1.TabIndex = 2;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(3, 74);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(555, 23);
            this.buttonCancel.TabIndex = 0;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonCreateChartOneStormManyGauges
            // 
            this.buttonCreateChartOneStormManyGauges.Location = new System.Drawing.Point(280, 315);
            this.buttonCreateChartOneStormManyGauges.Name = "buttonCreateChartOneStormManyGauges";
            this.buttonCreateChartOneStormManyGauges.Size = new System.Drawing.Size(297, 51);
            this.buttonCreateChartOneStormManyGauges.TabIndex = 25;
            this.buttonCreateChartOneStormManyGauges.Text = "Create Storm Charts - One Storm, Multiple Gauges";
            this.buttonCreateChartOneStormManyGauges.UseVisualStyleBackColor = true;
            this.buttonCreateChartOneStormManyGauges.Click += new System.EventHandler(this.buttonCreateChartOneStormManyGauges_Click);
            // 
            // backgroundWorkerMultiple
            // 
            this.backgroundWorkerMultiple.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerMultiple_DoWork);
            this.backgroundWorkerMultiple.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerMultiple_RunWorkerCompleted);
            // 
            // FormStormChartsMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(591, 378);
            this.Controls.Add(this.buttonCreateChartOneStormManyGauges);
            this.Controls.Add(this.pnlCancelBackgroundWorker);
            this.Controls.Add(this.labelH2Number);
            this.Controls.Add(this.numericUpDownH2Number);
            this.Controls.Add(this.labelEndTime);
            this.Controls.Add(this.labelStartTime);
            this.Controls.Add(this.dateTimePickerEndTime);
            this.Controls.Add(this.dateTimePickerStartTime);
            this.Controls.Add(this.buttonCreateStormCharts);
            this.Name = "FormStormChartsMain";
            this.Text = "StormCharts";
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownH2Number)).EndInit();
            this.pnlCancelBackgroundWorker.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonCreateStormCharts;
        private System.Windows.Forms.DateTimePicker dateTimePickerStartTime;
        private System.Windows.Forms.DateTimePicker dateTimePickerEndTime;
        private System.Windows.Forms.Label labelStartTime;
        private System.Windows.Forms.Label labelEndTime;
        private System.Windows.Forms.NumericUpDown numericUpDownH2Number;
        private System.Windows.Forms.Label labelH2Number;
        private System.ComponentModel.BackgroundWorker backgroundWorkerSingle;
        private System.Windows.Forms.Label lblModelRunStatus;
        private System.Windows.Forms.Panel pnlCancelBackgroundWorker;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonCreateChartOneStormManyGauges;
        private System.ComponentModel.BackgroundWorker backgroundWorkerMultiple;
    }
}

