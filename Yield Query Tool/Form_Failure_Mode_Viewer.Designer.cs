namespace Yield_Query_Tool
{
    partial class Form_Failure_Mode_Viewer
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.Failure_Mode_chart = new System.Windows.Forms.DataVisualization.Charting.Chart();
            ((System.ComponentModel.ISupportInitialize)(this.Failure_Mode_chart)).BeginInit();
            this.SuspendLayout();
            // 
            // Failure_Mode_chart
            // 
            this.Failure_Mode_chart.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            chartArea1.Name = "ChartArea1";
            this.Failure_Mode_chart.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.Failure_Mode_chart.Legends.Add(legend1);
            this.Failure_Mode_chart.Location = new System.Drawing.Point(32, 46);
            this.Failure_Mode_chart.Name = "Failure_Mode_chart";
            this.Failure_Mode_chart.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Bright;
            series1.ChartArea = "ChartArea1";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.Failure_Mode_chart.Series.Add(series1);
            this.Failure_Mode_chart.Size = new System.Drawing.Size(967, 547);
            this.Failure_Mode_chart.TabIndex = 0;
            this.Failure_Mode_chart.Text = "chart1";
            // 
            // Form_Failure_Mode_Viewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1039, 605);
            this.Controls.Add(this.Failure_Mode_chart);
            this.Name = "Form_Failure_Mode_Viewer";
            this.Text = "Form_Failure_Mode_Viewer";
            ((System.ComponentModel.ISupportInitialize)(this.Failure_Mode_chart)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart Failure_Mode_chart;
    }
}