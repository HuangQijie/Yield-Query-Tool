using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Yield_Query_Tool
{
    public partial class Form_Failure_Mode_Viewer : Form
    {
        public Form_Failure_Mode_Viewer(string yieldtype, string ModelID, string DataSetName, string AllfailQTY, string weeknumber,
            string FirstFailureModeName, string FirstFailureModeQTY, string FirstFailureModeRatio,
            string SecondFailureModeName, string SecondFailureModeQTY, string SecondFailureModeRatio,
            string ThirdFailureModeName, string ThirdFailureModeQTY, string ThirdFailureModeRatio)
        {
            InitializeComponent();
            this.Text = ModelID + " " + DataSetName + " (Week " + weeknumber + ") " + yieldtype + " Failure Mode Ratio View";

            #region ChartArea
            Failure_Mode_chart.Series.Clear();
            ChartArea chartArea = Failure_Mode_chart.ChartAreas[0];
            chartArea.BorderDashStyle = ChartDashStyle.Solid;
            chartArea.BackColor = Color.WhiteSmoke;// Color.FromArgb(0, 0, 0, 0);       
            chartArea.ShadowColor = Color.FromArgb(0, 0, 0, 0);
            //chartArea.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dash;//设置网格为虚线
            //chartArea.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            //chartArea.AxisX.MinorGrid.LineDashStyle = ChartDashStyle.Dash;//设置网格为虚线
            //chartArea.AxisX.Minimum = Double.Parse(tempWeekNumber.Min());// set x axis start at min week number
            //chartArea.AxisX.Maximum = Double.Parse(tempWeekNumber.Max());// set x axis end at max week number

            chartArea.Area3DStyle.Enable3D = true;//开启三维模式;PointDepth:厚度BorderWidth:边框宽
            chartArea.Area3DStyle.Rotation = 180;//起始角度
            chartArea.Area3DStyle.Inclination = 30;//倾斜度(0～90)
            chartArea.Area3DStyle.LightStyle = LightStyle.Realistic;//表面光泽度
            //chartArea.AxisY.LabelStyle.Format = "0%";//格式化，为了显示百分号

            //newChart.ChartAreas.Add(chartArea);
            #endregion

            Failure_Mode_chart.Titles.Add("Total Failed QTY is " + AllfailQTY);


            Failure_Mode_chart.Series.Add("FailureModeChart");
            //Failure_Mode_chart.Series["FailureModeChart"].BorderWidth = 3;
            Failure_Mode_chart.Series["FailureModeChart"].ChartType = SeriesChartType.Pie;
            //Failure_Mode_chart.Series["FailureModeChart"].YAxisType = AxisType.Primary;
            Failure_Mode_chart.Series["FailureModeChart"].IsValueShownAsLabel = true;
            //Failure_Mode_chart.Series["FailureModeChart"].MarkerStyle = MarkerStyle.Square;
            Failure_Mode_chart.Series["FailureModeChart"]["PieLabelStyle"] = "Outside";
            Failure_Mode_chart.Series["FailureModeChart"]["PieLineColor"] = "Black";

            //Failure_Mode_chart.Series["FailureModeChart"].LegendText = "#PERCENT";





            if (ThirdFailureModeName != "")
            {
                double first = Double.Parse(FirstFailureModeRatio);
                double second = Double.Parse(SecondFailureModeRatio);
                double third = Double.Parse(ThirdFailureModeRatio);

                string[] xValue = { FirstFailureModeName, SecondFailureModeName, ThirdFailureModeName, "Others" };
                double[] yValue = { first, second, third, (1 - first - second - third) > 0 ? (1 - first - second - third) : 0 };

                Failure_Mode_chart.Series["FailureModeChart"].Points.DataBindXY(xValue, yValue);

            }
            else if (SecondFailureModeName != "")
            {
                double first = Double.Parse(FirstFailureModeRatio);
                double second = Double.Parse(SecondFailureModeRatio);
                string[] xValue = { FirstFailureModeName, SecondFailureModeName };
                double[] yValue = { first, second };

                Failure_Mode_chart.Series["FailureModeChart"].Points.DataBindXY(xValue, yValue);
            }
            else if (FirstFailureModeName != "")
            {
                double first = Double.Parse(FirstFailureModeRatio);

                string[] xValue = { FirstFailureModeName };
                double[] yValue = { first };

                Failure_Mode_chart.Series["FailureModeChart"].Points.DataBindXY(xValue, yValue);
            }
            else
            {
                MessageBox.Show("All PASSED!!");

            }



            //Failure_Mode_chart.Series["FailureModeChart"].Points.DataBindXY(xValue, yValue);
            //SecondFailureModeName, SecondFailureModeRatio, 
            //ThirdFailureModeName, ThirdFailureModeRatio);


        }
    }
}
