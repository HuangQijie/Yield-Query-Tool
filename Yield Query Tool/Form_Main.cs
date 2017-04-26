using System;
using System.Data;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
//using System.Windows.Forms.DataVisualization.Charting;

//using System.Data.SqlClient;
//using System.Threading;

namespace Yield_Query_Tool
{
    public partial class MainForm : Form
    {
        Function fc = new Function();


        public string dbPath = System.AppDomain.CurrentDomain.BaseDirectory + "Config.mdb";
        public string OracleConnectString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=shg-oracle)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=ateshg)));User Id=extviewer;Password=extviewer";
        public string StartTime;
        public string EndTime;
        public string EnvString = System.AppDomain.CurrentDomain.BaseDirectory + "Oracle_OCI";

        #region
        //This section is for ER CPK used only
        public string StartTime_ERCPK;
        public string EndTime_ERCPK;
        //
        #endregion


        //Change SWVersion here
        public string SWVersion = "4.4";

        //




        public MainForm()
        {
            InitializeComponent();

            //OracleConnection OracleConn = oracledbo.OpenOracleConnection2(OracleConnectString);



        }






        private void StartTimePicker_ValueChanged(object sender, EventArgs e)
        {
            StartTime = fc.TimePickerFormat(StartTimePicker);

        }

        private void EndTimePicker_ValueChanged(object sender, EventArgs e)
        {
            EndTime = fc.TimePickerFormat(EndTimePicker);
        }



        private void MainForm_Load(object sender, EventArgs e)
        {
            // These lines are used to solve the heap corruption error
            //MessageBox.Show("Yield Query Tool Alpha Version" + SWVersion + "\n Any question please contanct Andy Huang");
            Environment.SetEnvironmentVariable("PATH", EnvString, EnvironmentVariableTarget.Process);
            //MessageBox.Show(Environment.GetEnvironmentVariable("PATH"));
            fc.OpenOracleConnection_ThenCLosed2(OracleConnectString);
            // These lines are used to solve the heap corruption error

            //Function fc_temp = new Function();
            fc.InitialGetFromConfig(dbPath, Step_listBox, DataSet_listBox, DataName);
            fc.InitialStatusTab(DataSetStatus, DataNameStatus, Comp_Type_listBox);
            fc.InitialTimePicker(StartTimePicker, EndTimePicker);

            #region
            //This section is for ER CPK used only
            fc.InitialTimePicker(dateTimePicker_ERCPK_Start, dateTimePicker_ERCPK_End);
            #endregion



            MainPanelToolTip.AutoPopDelay = 3000;
            MainPanelToolTip.InitialDelay = 800;
            MainPanelToolTip.ReshowDelay = 500;
            MainPanelToolTip.ShowAlways = true;
            //MainPanelToolTip.SetToolTip(JobOrder, "支持SQL通配符");
            //MainPanelToolTip.SetToolTip(DataSet, "支持SQL通配符");
            MainPanelToolTip.SetToolTip(DataName, "支持%通配符");
            MainPanelToolTip.SetToolTip(YieldEnablecheckBox, "查询Yield时勾上");
            MainPanelToolTip.SetToolTip(CPKEnablecheckBox, "查询CPK时勾上");
            MainPanelToolTip.SetToolTip(ReadSNFromFile, "txt中用换行键分隔");
            MainPanelToolTip.SetToolTip(ReadJOFromFile, "txt中用换行键分隔");
            MainPanelToolTip.SetToolTip(ReadBOMPNFromFile, "txt中用换行键分隔");
            MainPanelToolTip.SetToolTip(ReadModelIDFromFile, "txt中用换行键分隔");
            MainPanelToolTip.SetToolTip(NoPathConfig, "如果查询报错，请勾上");

            MainPanelToolTip.SetToolTip(ReadCpkLimitFromMPL, "Get from t_limits");
            MainPanelToolTip.SetToolTip(FPYRawdataGridView, "FPY Raw data,仅显示最早的测试记录");
            MainPanelToolTip.SetToolTip(FYRawdataGridView, "FY Raw data,仅显示最新的测试记录");
            MainPanelToolTip.SetToolTip(FPYTabledataGridView, "FPY table，计算并显示良率");
            MainPanelToolTip.SetToolTip(FYTabledataGridView, "FY table，计算并显示良率");

            //MainPanelToolTip.SetToolTip(Step_listBox, "双击全选");
            //MainPanelToolTip.SetToolTip(DataSet_listBox, "双击全选");

            this.Text = "Yield Query Tool Alpha " + SWVersion + " (Production_DB)";

            YieldTab.Parent = null;
            CPKtabPage.Parent = null;
            ERCPKtabPage.Parent = null;
            WhereUsedTab.Parent = null;
            Dataset_Testtime_tab.Parent = null;
            Yield_Plot_Tab.Parent = null;
            Component_Edata_Tab.Parent = null;
            WIP_Status_tabPage.Parent = null;

            Selected_Step_InforLabel.ForeColor = System.Drawing.Color.Blue;
            Selected_Dataset_InforLabel.ForeColor = System.Drawing.Color.Blue;
            Comp_Type_InforLabel.ForeColor = System.Drawing.Color.Blue;
        }



        private void Step_SelectedIndexChanged(object sender, EventArgs e)
        {


            //fc.DataSetChangeWithStep(dbPath, Step.Text.Trim(), DataSet, DataName,Step_listBox, DataSet_listBox, DataName_listBox);

        }



        private void DataSet_SelectedIndexChanged(object sender, EventArgs e)
        {

            //fc.DataNameChangeWithDataSet(dbPath, Step.Text.Trim(), DataSet.Text.Trim(), DataName, DataSet_listBox);
        }


        private void DataSetStatus_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            int ContainerNumber;

            Logger("Search");//trace into log 

            MainPanelInformLabel.Text = "Searching...";
            MainPanelInformLabel.ForeColor = System.Drawing.Color.Red;
            Application.DoEvents();
            //SearchDataGridView.Rows.Clear();
            string Search_Record_Type = "AllRecord";

            if (Search_FirstRecord_checkBox.Checked)
                Search_Record_Type = "FirstRecord";
            if (Search_LastRecord_checkBox.Checked)
                Search_Record_Type = "LastRecord";

            DateTime start = DateTime.Now;
            ds = fc.MainOracleQuery(OracleConnectString, SerialNumber.Text.Trim(), JobOrder.Text.Trim(),
                BOMPN.Text.Trim(), BOMPNRev.Text.Trim(), ModelID.Text.Trim(),
                Selected_Dataset_InforLabel.Text.Trim(), DataName.Text.Trim(), DataNameVal.Text.Trim(),
                DataSetStatus.Text.Trim(), DataNameStatus.Text.Trim(), StartTime, EndTime, MainPanelInformLabel,
                Comp_SN_textBox.Text.Trim(), Comp_Type_InforLabel.Text.Trim(), Comp_eDataName_comboBox.Text.Trim(), Comp_eData_Value_textBox.Text.Trim(),
                Comp_eData_Include_checkBox.Checked, Search_Record_Type,Comp_PN_textBox.Text.Trim(),Comp_Already_Removed_checkBox.Checked);
            SearchDataGridView.DataSource = ds;
            SearchDataGridView.DataMember = ds.Tables[0].TableName;
            //SearchDataGridView = fc.ChangeDataGridColor(SearchDataGridView);
            MainPanelInformLabel.ForeColor = System.Drawing.Color.Black;
            ContainerNumber = fc.ContainerNumberCounter(ds.Tables[0]);
            DateTime end = DateTime.Now;
            MainPanelInformLabel.Text = "Find total " + ds.Tables[0].Rows.Count + " recorders...The active container number is " + ContainerNumber.ToString() + "...Elapse time is "
                + (end - start).TotalSeconds.ToString() + " seconds.";
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(SearchDataGridView);
        }

        private void FPYExport_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(FPYRawdataGridView);
        }
        private void FYExport_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(FYRawdataGridView);
        }

        private void FPYResultTableExport_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(FPYTabledataGridView);
        }

        private void FYResultTableExport_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(FYTabledataGridView);
        }




        private void GetYieldButton_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            DataSet rs = new DataSet();
            DataSet ts = new DataSet();
            DataSet hs = new DataSet();
            DataSet ls = new DataSet();
            int ContainerNumber;

            Logger("FPY&FY");//trace into log 

            MainPanelInformLabel.Text = "Searching...";
            MainPanelInformLabel.ForeColor = System.Drawing.Color.Red;
            Application.DoEvents();
            DateTime start = DateTime.Now;
            //SearchDataGridView.Rows.Clear();
            //FPYRawdataGridView.Rows.Clear();
            //FYRawdataGridView.Rows.Clear();
            ds = fc.YieldOracleQuery(OracleConnectString, SerialNumber.Text.Trim(), JobOrder.Text.Trim(), BOMPN.Text.Trim(),
                BOMPNRev.Text.Trim(), ModelID.Text.Trim(), Selected_Dataset_InforLabel.Text.Trim(),
                DataName.Text.Trim(), StartTime, EndTime, MainPanelInformLabel);
            DateTime end = DateTime.Now;
            //ds = fc.MainOracleQuery(OracleConnectString, SerialNumber.Text.Trim(), JobOrder.Text.Trim(), BOMPN.Text.Trim(), BOMPNRev.Text.Trim(), ModelID.Text.Trim(), DataSet.Text.Trim(), DataName.Text.Trim(), DataSetStatus.Text.Trim(), DataNameStatus.Text.Trim(), StartTime, EndTime, MainPanelInformLabel);
            if (ds.Tables[0].Rows.Count > 0)
            {
                rs = fc.FPYandFYDataFilter(ds.Tables[0]);
                SearchDataGridView.DataSource = rs;
                SearchDataGridView.DataMember = rs.Tables[0].TableName;
                //SearchDataGridView = fc.ChangeDataGridColor(SearchDataGridView);
                MainPanelInformLabel.ForeColor = System.Drawing.Color.Black;
                ContainerNumber = fc.ContainerNumberCounter(rs.Tables[0]);
                MainPanelInformLabel.Text = "Find total yield raw data " + ds.Tables[0].Rows.Count + " recorders...The active container number is " + ContainerNumber.ToString() +
                "...Elapse time is " + (end - start).TotalSeconds.ToString() + " seconds.";


                FPYRawdataGridView.DataSource = rs;
                FPYRawdataGridView.DataMember = rs.Tables[1].TableName;
                //FPYRawdataGridView = fc.ChangeDataGridColor(FPYRawdataGridView);
                ContainerNumber = fc.ContainerNumberCounter(rs.Tables[1]);
                FPYInformLabel.Text = "Find FPY raw data " + rs.Tables[1].Rows.Count + " recorders...The active container number is " + ContainerNumber.ToString();

                FYRawdataGridView.DataSource = rs;
                FYRawdataGridView.DataMember = rs.Tables[2].TableName;
                //FYRawdataGridView = fc.ChangeDataGridColor(FYRawdataGridView);
                ContainerNumber = fc.ContainerNumberCounter(rs.Tables[2]);
                FYInformLabel.Text = "Find FY raw data " + rs.Tables[2].Rows.Count + " recorders...The active container number is " + ContainerNumber.ToString();

                ts = fc.FillYieldTable(rs.Tables[1]);
                FPYTabledataGridView.DataSource = ts;
                FPYTabledataGridView.DataMember = ts.Tables[0].TableName;

                ts = fc.FillYieldTable(rs.Tables[2]);
                FYTabledataGridView.DataSource = ts;
                FYTabledataGridView.DataMember = ts.Tables[0].TableName;


                hs = fc.Yield_Plot_Table(rs.Tables[1]);
                FPY_Yield_Plot_dataGridView.DataSource = hs;
                FPY_Yield_Plot_dataGridView.DataMember = hs.Tables[0].TableName;

                ls = fc.Yield_Plot_Table(rs.Tables[2]);
                FY_Yield_Plot_dataGridView.DataSource = ls;
                FY_Yield_Plot_dataGridView.DataMember = ls.Tables[0].TableName;

                fc.Yield_Chart_Plot(hs.Tables[0], FPY_Yield_Plot_chart, "FPY Ratio");
                fc.Yield_Chart_Plot(ls.Tables[0], FY_Yield_Plot_chart, "FY Ratio");


            }
            else
                MainPanelInformLabel.Text = "No result...";

        }

        private void FYInformLabel_Click(object sender, EventArgs e)
        {

        }

        private void YieldTab_Click(object sender, EventArgs e)
        {

        }



        private void YieldEnablecheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (YieldEnablecheckBox.Checked)
            {
                DataSetStatus.Enabled = false;
                DataNameStatus.Enabled = false;
                DataName.Enabled = false;
                ReadDataNameFromFile.Enabled = false;
                //SearchButton.Enabled = false;
                GetYieldButton.Enabled = true;
                YieldTab.Parent = MainTabControl;
                Yield_Plot_Tab.Parent = MainTabControl;

                CPKEnablecheckBox.Enabled = false;
                DataSet_TestTime_checkBox.Enabled = false;
                Search_Enable_checkBox.Enabled = false;


            }
            else
            {
                {
                    DataSetStatus.Enabled = true;
                    DataNameStatus.Enabled = true;
                    DataName.Enabled = true;
                    ReadDataNameFromFile.Enabled = true;
                    //SearchButton.Enabled = true;
                    GetYieldButton.Enabled = false;
                    YieldTab.Parent = null;
                    Yield_Plot_Tab.Parent = null;


                    CPKEnablecheckBox.Enabled = true;
                    DataSet_TestTime_checkBox.Enabled = true;
                    Search_Enable_checkBox.Enabled = true;



                }
            }


        }



        private void NoPathConfig_CheckedChanged(object sender, EventArgs e)
        {
            if (NoPathConfig.Checked)
            {
                Environment.SetEnvironmentVariable("PATH", EnvString, EnvironmentVariableTarget.Process);
                MessageBox.Show("Environment Path configuration finished\n" + EnvString);
            }
            else
            {
                Environment.SetEnvironmentVariable("PATH", null, EnvironmentVariableTarget.Process);
                MessageBox.Show("Environment Path configuration cleared\n" + EnvString);
            }


        }


        private void ReadSNFromFile_Click(object sender, EventArgs e)
        {
            SerialNumber.Text = fc.ReadListFromTxt();

        }

        private void ReadJOFromFile_Click(object sender, EventArgs e)
        {
            JobOrder.Text = fc.ReadListFromTxt();
        }

        private void ReadBOMPNFromFile_Click(object sender, EventArgs e)
        {
            BOMPN.Text = fc.ReadListFromTxt();
        }

        private void ReadModelIDFromFile_Click(object sender, EventArgs e)
        {
            ModelID.Text = fc.ReadListFromTxt();
        }

        private void Test_DB_CheckedChanged(object sender, EventArgs e)
        {
            if (Test_DB_Enable_checkBox.Checked)
            {
                OracleConnectString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=shg-oracle)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=ateshg)));User Id=test;Password=test";
                this.Text = "Yield Query Tool Alpha " + SWVersion + " (Test_DB)";
            }
            else
            {
                OracleConnectString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=shg-oracle)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=ateshg)));User Id=extviewer;Password=extviewer";
                this.Text = "Yield Query Tool Alpha " + SWVersion + " (Production_DB)";
            }

        }

        private void CPK_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            DataSet ts = new DataSet();
            int ContainerNumber;

            Logger("CPK");//trace into log 

            //Check USL and LSL is not Blank
            if (USL_Val.Text == "" || LSL_Val.Text == "")
                MessageBox.Show("USL and LSL can not be blank!");
            else if (ModelID.Text.Contains(","))
                MessageBox.Show("ModelID number cannot larger than one");
            else if (DataSet_listBox.SelectedItems.Count > 1)
                MessageBox.Show("DataSet number cannot larger than one");
            else
            {
                MainPanelInformLabel.Text = "Searching...";
                MainPanelInformLabel.ForeColor = System.Drawing.Color.Red;
                Application.DoEvents();
                DateTime start = DateTime.Now;
                //SearchDataGridView.Rows.Clear();
                ds = fc.CPKOracleQuery(OracleConnectString, SerialNumber.Text.Trim(), JobOrder.Text.Trim(),
                    BOMPN.Text.Trim(), BOMPNRev.Text.Trim(), ModelID.Text.Trim(), Selected_Dataset_InforLabel.Text.Trim(), DataName.Text.Trim(), DataNameVal.Text.Trim(),
                    DataSetStatus.Text.Trim(), DataNameStatus.Text.Trim(), StartTime, EndTime, MainPanelInformLabel);
                SearchDataGridView.DataSource = ds;
                SearchDataGridView.DataMember = ds.Tables[0].TableName;
                //SearchDataGridView = fc.ChangeDataGridColor(SearchDataGridView);
                MainPanelInformLabel.ForeColor = System.Drawing.Color.Black;
                ContainerNumber = fc.ContainerNumberCounter(ds.Tables[0]);
                DateTime end = DateTime.Now;
                MainPanelInformLabel.Text = "Find total CPK raw data" + ds.Tables[0].Rows.Count + " recorders...The active container number is " + ContainerNumber.ToString() + "...Elapse time is "
                + (end - start).TotalSeconds.ToString() + " seconds.";




                //double[] CPkRawData = new double[ds.Tables[0].Rows.Count];
                double[] CPkRawData = new double[ds.Tables[0].Rows.Count];
                if (ds.Tables[0].Rows.Count > 0)
                {

                    for (var i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        try
                        {
                            CPkRawData[i] = Double.Parse(ds.Tables[0].Rows[i][9].ToString());
                        }
                        catch (Exception ex)
                        {
                            //Console.WriteLine(ex.ToString());
                            throw new Exception(CPkRawData[i] + " is not a number? You may need to adjust a better DataName in SQL" + ex.Message);
                        }

                    }

                    ts = fc.CPKCalculator(DataName.Text.Trim(), CPkRawData, ContainerNumber, Double.Parse(USL_Val.Text.Trim()), Double.Parse(LSL_Val.Text.Trim()));
                    CPKdataGridView.DataSource = ts;
                    CPKdataGridView.DataMember = ts.Tables[0].TableName;
                    CPKInformLabel.Text = "Use Non-normal distribution to calculate CPK";

                    fc.CPKChartPlot(CPkRawData, CPKChart, Double.Parse(ts.Tables[0].Rows[0]["CPK"].ToString()), Double.Parse(LSL_Val.Text.Trim()), Double.Parse(USL_Val.Text.Trim()), 40);
                }
                else
                    MessageBox.Show("NO result is query out.");


            }

        }

        private void CPK_Select_CheckedChanged(object sender, EventArgs e)
        {
            if (CPKEnablecheckBox.Checked)
            {
                CPKButton.Enabled = true;
                //SearchButton.Enabled = false;
                DataSetStatus.Text = "PASS";
                USL_Val.Enabled = true;
                LSL_Val.Enabled = true;
                ReadCpkLimitFromMPL.Enabled = true;
                //MessageBox.Show("");
                CPKtabPage.Parent = MainTabControl;

                YieldEnablecheckBox.Enabled = false;
                DataSet_TestTime_checkBox.Enabled = false;
                Search_Enable_checkBox.Enabled = false;



            }
            else
            {
                CPKButton.Enabled = false;
                //SearchButton.Enabled = true;
                DataSetStatus.Text = "";
                USL_Val.Enabled = false;
                LSL_Val.Enabled = false;
                ReadCpkLimitFromMPL.Enabled = false;
                //MessageBox.Show("");
                CPKtabPage.Parent = null;

                YieldEnablecheckBox.Enabled = true;
                DataSet_TestTime_checkBox.Enabled = true;
                Search_Enable_checkBox.Enabled = true;
            }

        }

        private void ER_CPK_checkBox_CheckedChanged(object sender, EventArgs e)
        {

            if (ER_CPK_checkBox.Checked)
            {
                ERCPKtabPage.Parent = MainTabControl;

            }
            else
            {
                ERCPKtabPage.Parent = null;
            }

        }





        private void ReadCpkLimit_From_MPL_tlimit_Click(object sender, EventArgs e)
        {

            if (BOMPN.Text.Length == 0 && ModelID.Text.Length == 0 && DataSet_listBox.SelectedItems.Count == 0 && DataName.Text.Length == 0)
                MessageBox.Show("BOMPN, ModelID, DataSet, DataName cannot be all blank");
            else if (DataSet_listBox.SelectedItems.Count == 0)
                MessageBox.Show("DataSet cannot be  blank");
            else if (DataName.Text.Length == 0)
                MessageBox.Show("DataName cannot be  blank");
            else if (BOMPN.Text.Length == 0 && ModelID.Text.Length == 0)
                MessageBox.Show("BOMPN, ModelID cannot be all blank");
            else if (DataSet_listBox.SelectedItems.Count > 1)
                MessageBox.Show("Cannot select more than one DataSet");
            else
            {
                //remove [] and % smybol from dataname, then input to MPL query
                string TempDataName = DataName.Text.Trim();
                //TempDataName = TempDataName.Replace("[", "");
                //TempDataName = TempDataName.Replace("]", "");
                //TempDataName = TempDataName.Replace("%", "");

                string[] RemmvalStringList = { "[", "%", "_chan" };//this is for further removal string updating
                for (var i = 0; i < RemmvalStringList.Length; i++)
                {
                    if (TempDataName.LastIndexOf(RemmvalStringList[i]) > 0)
                        TempDataName = TempDataName.Substring(0, TempDataName.LastIndexOf(RemmvalStringList[i]));
                }
                //    if (TempDataName.LastIndexOf("[") > 0)
                //        TempDataName = TempDataName.Substring(0, TempDataName.LastIndexOf("["));

                //if (TempDataName.LastIndexOf("%") > 0)
                //    TempDataName = TempDataName.Substring(0, TempDataName.LastIndexOf("%"));
                //if (TempDataName.LastIndexOf("_chan") > 0)
                //    TempDataName = TempDataName.Substring(0, TempDataName.LastIndexOf("_chan_"));


                TempDataName = DataSet_listBox.SelectedItems[0].ToString().Trim() + ":" + TempDataName;
                try
                {
                    //MessageBox.Show("");
                    DataSet ds = fc.MPLOracleQuery(OracleConnectString, BOMPN.Text.Trim(), BOMPNRev.Text.Trim(), ModelID.Text.Trim(), TempDataName, "t_limits");

                    LSL_Val.Text = ds.Tables[0].Rows[0][7].ToString();//get LSL from first row in dataset
                    USL_Val.Text = ds.Tables[0].Rows[0][8].ToString();//get USL from first row in dataset
                    LimitUnit.Text = ds.Tables[0].Rows[0][9].ToString();//get limit unit from first row in dataset
                    MessageBox.Show("Successfully get specification limit from MPL, please check it");
                }
                catch (Exception ee)
                {
                    MessageBox.Show("Cannot get Specification Limit from MPL, Please input it manually.\n Error information:" + ee.ToString());
                }
            }

        }

        private void CPKTableExport_Click(object sender, EventArgs e)
        {


            fc.DataGridViewToExcel(CPKdataGridView);

        }

        private void CPKChart_Click(object sender, EventArgs e)
        {

        }

        private void CPKChartExport_Click(object sender, EventArgs e)
        {
            fc.ExportChart(CPKChart);
        }

        private void ModelIDImport_ERCPK_Click(object sender, EventArgs e)
        {
            ModelID_ERCPK.Text = fc.ReadListFromTxt();
        }

        #region
        //This section is for ER CPK used only
        private void dateTimePicker_ERCPK_Start_ValueChanged(object sender, EventArgs e)
        {
            StartTime_ERCPK = fc.TimePickerFormat(dateTimePicker_ERCPK_Start);

        }

        private void dateTimePicker_ERCPK_End_ValueChanged(object sender, EventArgs e)
        {
            EndTime_ERCPK = fc.TimePickerFormat(dateTimePicker_ERCPK_End);
        }

        private void ER_CPK_button_Click(object sender, EventArgs e)
        {
            if (ModelID_ERCPK.Text == "")
            {
                MessageBox.Show("ModelID cannot be blank!");
            }
            else
            {
                DataSet ds;
                DataSet ts;

                ds = fc.ERCPK_OracleQuery(OracleConnectString, ModelID_ERCPK.Text.Trim(), StartTime_ERCPK, EndTime_ERCPK);
                ERCPK_rawdata_dataGridView.DataSource = ds;
                ERCPK_rawdata_dataGridView.DataMember = ds.Tables[0].TableName;

                ts = fc.ERCPK_Data_Processor(ds.Tables[0]);
                ERCPK_Table_dataGridView.DataSource = ts;
                ERCPK_Table_dataGridView.DataMember = ts.Tables[0].TableName;

            }
        }

        private void ERCPK_Table_Export_button_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(ERCPK_Table_dataGridView);
        }

        private void ERCPK_Rawdata_Export_button_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(ERCPK_rawdata_dataGridView);
        }

        #endregion

        # region WhereUsed query section start here

        private void AlreadyRemoved_checkBox_No_CheckedChanged(object sender, EventArgs e)
        {
            if (AlreadyRemoved_checkBox_No.Checked)
                AlreadyRemoved_checkBox_Yes.Checked = false;
            else
                AlreadyRemoved_checkBox_Yes.Checked = true;
        }

        private void AlreadyRemoved_checkBox_Yes_CheckedChanged(object sender, EventArgs e)
        {
            if (AlreadyRemoved_checkBox_Yes.Checked)
                AlreadyRemoved_checkBox_No.Checked = false;
            else
                AlreadyRemoved_checkBox_No.Checked = true;
        }

        private void Where_Used_button_Click(object sender, EventArgs e)
        {
            DataSet ds;

            ds = fc.WhereUsedOracleQuery(OracleConnectString, AlreadyRemoved_checkBox_Yes.Checked, SN_WhereUsed.Text.Trim());

            WhereUsed_dataGridView.DataSource = ds;
            WhereUsed_dataGridView.DataMember = ds.Tables[0].TableName;
        }

        private void WhereUsedQuerycheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (WhereUsedQuerycheckBox.Checked)

                WhereUsedTab.Parent = MainTabControl;
            else
                WhereUsedTab.Parent = null;


        }
        #endregion



        private void DataName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void DataName_TextUpdate(object sender, EventArgs e)
        {
            if (DataName.Text.Trim() != "")
            {
                DataNameVal_Label.Enabled = true;
                DataNameVal.Enabled = true;
            }
            else
            {
                DataNameVal_Label.Enabled = false;
                DataNameVal.Enabled = false;
            }

        }

        private void Step_listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            fc.DataSetChangeWithStep(dbPath, Step_listBox, DataSet_listBox, DataName);
            Selected_Step_InforLabel.Text = fc.ListBox2SQL_in_Query_String(Step_listBox);
            Selected_Dataset_InforLabel.Text = fc.ListBox2SQL_in_Query_String(DataSet_listBox);

            //Selected_Step_InforLabel.ForeColor = System.Drawing.Color.Blue;
            //Selected_Dataset_InforLabel.ForeColor = System.Drawing.Color.Blue;




        }

        private void DataSet_listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            fc.DataNameChangeWithDataSet(dbPath, Step_listBox, DataSet_listBox, DataName);
            Selected_Step_InforLabel.Text = fc.ListBox2SQL_in_Query_String(Step_listBox);
            Selected_Dataset_InforLabel.Text = fc.ListBox2SQL_in_Query_String(DataSet_listBox);

            //Selected_Step_InforLabel.ForeColor = System.Drawing.Color.Blue;
            //Selected_Dataset_InforLabel.ForeColor = System.Drawing.Color.Blue;

        }



        private void Step_listBox_MouseEnter(object sender, EventArgs e)
        {
            Step_listBox.Height = 160;
            Step_listBox.Select();
        }

        private void Step_listBox_MouseLeave(object sender, EventArgs e)
        {
            Step_listBox.Height = 20;
        }

        private void DataSet_listBox_MouseEnter(object sender, EventArgs e)
        {
            DataSet_listBox.Height = 300;
            DataSet_listBox.Select();
        }

        private void DataSet_listBox_MouseLeave(object sender, EventArgs e)
        {
            DataSet_listBox.Height = 20;
        }

        private void DataSet_listBox_DoubleClick(object sender, EventArgs e)
        {


            if (DataSet_listBox.Items.Count != DataSet_listBox.SelectedItems.Count + 1)
            {
                for (int i = 0; i < DataSet_listBox.Items.Count; i++)
                    DataSet_listBox.SetSelected(i, true);

            }
            else
            {
                for (int i = 0; i < DataSet_listBox.Items.Count; i++)
                    DataSet_listBox.SetSelected(i, false);
            }

        }



        private void Step_listBox_DoubleClick(object sender, EventArgs e)
        {

            if (Step_listBox.Items.Count != Step_listBox.SelectedItems.Count + 1)
            {
                for (int i = 0; i < Step_listBox.Items.Count; i++)
                    Step_listBox.SetSelected(i, true);

            }
            else
            {
                for (int i = 0; i < Step_listBox.Items.Count; i++)
                    Step_listBox.SetSelected(i, false);
            }

        }

        private void Comp_Type_listBox_MouseEnter(object sender, EventArgs e)
        {
            Comp_Type_listBox.Height = 80;
            Comp_Type_listBox.Select();
        }

        private void Comp_Type_listBox_MouseLeave(object sender, EventArgs e)
        {
            Comp_Type_listBox.Height = 20;
        }

        private void Comp_Type_listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Comp_Type_InforLabel.Text = fc.ListBox2SQL_in_Query_String(Comp_Type_listBox);
            //Comp_Type_InforLabel.ForeColor = System.Drawing.Color.Blue;
        }

        private void Comp_Type_listBox_DoubleClick(object sender, EventArgs e)
        {
            if (Comp_Type_listBox.Items.Count != Comp_Type_listBox.SelectedItems.Count + 1)
            {
                for (int i = 0; i < Comp_Type_listBox.Items.Count; i++)
                    Comp_Type_listBox.SetSelected(i, true);

            }
            else
            {
                for (int i = 0; i < Comp_Type_listBox.Items.Count; i++)
                    Comp_Type_listBox.SetSelected(i, false);
            }
        }


        private void DataSet_TestTime_button_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            DataSet ts = new DataSet();
            int ContainerNumber;

            Logger("TestTime");//trace into log 

            MainPanelInformLabel.Text = "Searching...";
            MainPanelInformLabel.ForeColor = System.Drawing.Color.Red;
            Application.DoEvents();
            //SearchDataGridView.Rows.Clear();
            ds = fc.DataSet_TestTime_OracleQuery(OracleConnectString, SerialNumber.Text.Trim(), JobOrder.Text.Trim(),
                BOMPN.Text.Trim(), BOMPNRev.Text.Trim(), ModelID.Text.Trim(),
                Selected_Dataset_InforLabel.Text.Trim(), DataSetStatus.Text.Trim(), StartTime, EndTime, MainPanelInformLabel);

            MainPanelInformLabel.Text = "Result is shown in DataSet TestTime Tab";
            MainPanelInformLabel.ForeColor = System.Drawing.Color.Red;

            DataSet_TestTime_RawData_dataGridView.DataSource = ds;
            DataSet_TestTime_RawData_dataGridView.DataMember = ds.Tables[0].TableName;
            //SearchDataGridView = fc.ChangeDataGridColor(SearchDataGridView);
            DataSet_TestTime_InforLabel.ForeColor = System.Drawing.Color.Black;
            ContainerNumber = fc.ContainerNumberCounter(ds.Tables[0]);
            DataSet_TestTime_InforLabel.Text = "Find total " + ds.Tables[0].Rows.Count + " recorders...The active container number is " + ContainerNumber.ToString();


            if (ds.Tables[0].Rows.Count > 0)
            {
                ts = fc.DataSet_TestTime_Calculator(ds.Tables[0]);
                DataSet_TestTime_Result_dataGridView.DataSource = ts;
                DataSet_TestTime_Result_dataGridView.DataMember = ts.Tables[0].TableName;

            }

            else
                MainPanelInformLabel.Text = "No Result";



        }

        private void DataSet_TestTime_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (DataSet_TestTime_checkBox.Checked)
            {

                DataNameStatus.Enabled = false;
                DataName.Enabled = false;
                ReadDataNameFromFile.Enabled = false;

                //SearchButton.Enabled = false;
                DataSet_TestTime_button.Enabled = true;
                Dataset_Testtime_tab.Parent = MainTabControl;

                YieldEnablecheckBox.Enabled = false;
                CPKEnablecheckBox.Enabled = false;
                Search_Enable_checkBox.Enabled = false;




            }
            else
            {
                DataNameStatus.Enabled = true;
                DataName.Enabled = true;
                ReadDataNameFromFile.Enabled = true;

                //SearchButton.Enabled = true;
                DataSet_TestTime_button.Enabled = false;
                Dataset_Testtime_tab.Parent = null;

                YieldEnablecheckBox.Enabled = true;
                CPKEnablecheckBox.Enabled = true;
                Search_Enable_checkBox.Enabled = true;

            }
        }

        private void DataSet_TestTime_Result_table_export_button_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(DataSet_TestTime_Result_dataGridView);
        }

        private void DataSet_TestTime_Rawdata_table_export_button_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(DataSet_TestTime_RawData_dataGridView);
        }

        private void FPY_Yield_Plot_chart_Click(object sender, EventArgs e)
        {

        }

        private void FPY_Yield_Plot_data_button_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(FPY_Yield_Plot_dataGridView);
        }

        private void FY_Yield_Plot_data_button_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(FY_Yield_Plot_dataGridView);
        }

        private void FPY_Plot_Export_button_Click(object sender, EventArgs e)
        {
            fc.ExportChart(FPY_Yield_Plot_chart);
        }

        private void FY_Plot_Export_button_Click(object sender, EventArgs e)
        {
            fc.ExportChart(FY_Yield_Plot_chart);
        }

        private void ReadDataNameFromFile_Click(object sender, EventArgs e)
        {
            DataName.Text = fc.ReadListFromTxt();
        }



        private void FPY_Yield_Plot_dataGridView_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //double first = Double.Parse(FPY_Yield_Plot_dataGridView.CurrentRow.Cells[9].Value.ToString());
            //double second = Double.Parse(FPY_Yield_Plot_dataGridView.CurrentRow.Cells[11].Value.ToString());
            //double third = Double.Parse(FPY_Yield_Plot_dataGridView.CurrentRow.Cells[13].Value.ToString());
            if (FPY_Yield_Plot_dataGridView.CurrentRow.Cells[8].Value.ToString() != "")
            {
                Form_Failure_Mode_Viewer FPY_FailureMode_Form = new Form_Failure_Mode_Viewer(
                    "FPY",
                    FPY_Yield_Plot_dataGridView.CurrentRow.Cells[0].Value.ToString(),//module ID
                    FPY_Yield_Plot_dataGridView.CurrentRow.Cells[1].Value.ToString(),//dataset name
                    FPY_Yield_Plot_dataGridView.CurrentRow.Cells[3].Value.ToString(),// failed qty
                    FPY_Yield_Plot_dataGridView.CurrentRow.Cells[5].Value.ToString(),// weeknumber
                  FPY_Yield_Plot_dataGridView.CurrentRow.Cells[8].Value.ToString(), FPY_Yield_Plot_dataGridView.CurrentRow.Cells[9].Value.ToString(), FPY_Yield_Plot_dataGridView.CurrentRow.Cells[10].Value.ToString(),// first failure mode, qty,ratio
                  FPY_Yield_Plot_dataGridView.CurrentRow.Cells[11].Value.ToString(), FPY_Yield_Plot_dataGridView.CurrentRow.Cells[12].Value.ToString(), FPY_Yield_Plot_dataGridView.CurrentRow.Cells[13].Value.ToString(),// second failure mode, qty,ratio
                FPY_Yield_Plot_dataGridView.CurrentRow.Cells[14].Value.ToString(), FPY_Yield_Plot_dataGridView.CurrentRow.Cells[15].Value.ToString(), FPY_Yield_Plot_dataGridView.CurrentRow.Cells[16].Value.ToString());// third failure mode, qty,ratio

                FPY_FailureMode_Form.Show();
            }

        }

        private void FY_Yield_Plot_dataGridView_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (FY_Yield_Plot_dataGridView.CurrentRow.Cells[8].Value.ToString() != "")
            {
                Form_Failure_Mode_Viewer FY_FailureMode_Form = new Form_Failure_Mode_Viewer(
                    "FY",
                    FY_Yield_Plot_dataGridView.CurrentRow.Cells[0].Value.ToString(),//module ID
                    FY_Yield_Plot_dataGridView.CurrentRow.Cells[1].Value.ToString(),//dataset name
                    FY_Yield_Plot_dataGridView.CurrentRow.Cells[3].Value.ToString(),// failed qty
                    FY_Yield_Plot_dataGridView.CurrentRow.Cells[5].Value.ToString(),// weeknumber
                  FY_Yield_Plot_dataGridView.CurrentRow.Cells[8].Value.ToString(), FY_Yield_Plot_dataGridView.CurrentRow.Cells[9].Value.ToString(), FY_Yield_Plot_dataGridView.CurrentRow.Cells[10].Value.ToString(),// first failure mode, qty,ratio
                  FY_Yield_Plot_dataGridView.CurrentRow.Cells[11].Value.ToString(), FY_Yield_Plot_dataGridView.CurrentRow.Cells[12].Value.ToString(), FY_Yield_Plot_dataGridView.CurrentRow.Cells[13].Value.ToString(),// second failure mode, qty,ratio
                FY_Yield_Plot_dataGridView.CurrentRow.Cells[14].Value.ToString(), FY_Yield_Plot_dataGridView.CurrentRow.Cells[15].Value.ToString(), FY_Yield_Plot_dataGridView.CurrentRow.Cells[16].Value.ToString());// third failure mode, qty,ratio

                FY_FailureMode_Form.Show();
            }

        }




        private void Logger(string QueryType)
        {
            string LogFilePath = System.AppDomain.CurrentDomain.BaseDirectory + "Log.txt";
            string LogHeadLine = "Qurey_Type\tSN\tJO#\tBOM_PN\tRev\tModel_ID\tStep\tDataSet\tDataSet_Status\tDataName\tData_Status\tDataVal\tStart_Time\tEnd_Time\tLSL\tUSL\tQuery_Time\t"
                + "Component_SN\tComponent_Type\tComponent_Edata_Name\tComponent_Edata_Value\tTest_DB_Checked\tComp_eData_Checked\tSearch_First_Record_Checked\tSearch_Last_Record_Checked"
                + "\tComponent_PN";

            if (File.Exists(LogFilePath) == false)
            {
                //不存在文件 或者存在文件，则对比文件首行LogHeadLine,如果不一致,则删除txt所有内容，然后新增首行LogHeadLine
                //File.Create(LogFilePath);//创建该文件
                FileStream fs = new FileStream(LogFilePath, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);



                //开始写入
                sw.Write(LogHeadLine);
                //清空缓冲区
                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();



            }
            else
            {
                //存在文件，则对比文件首行LogHeadLine,如果不一致,则删除txt所有内容，然后新增首行LogHeadLine
                FileStream fs = new FileStream(LogFilePath, FileMode.Open);
                StreamReader sr = new StreamReader(fs);

                if (sr.ReadLine().ToString() != LogHeadLine)
                {
                    fs.Close();
                    sr.Close();
                    FileStream ffs = new FileStream(LogFilePath, FileMode.Create);
                    StreamWriter sw = new StreamWriter(ffs);
                    sw.Write(LogHeadLine);
                    sw.Flush();
                    sw.Close();
                }


                fs.Close();
                sr.Close();



            }



            FileStream Fs = new FileStream(LogFilePath, FileMode.Append);
            StreamWriter Sw = new StreamWriter(Fs);
            Sw.Write("\r\n" + QueryType + "\t" + SerialNumber.Text.Trim() + "\t" + JobOrder.Text.Trim() + "\t" + BOMPN.Text.Trim() + "\t" + BOMPNRev.Text.Trim() + "\t" + ModelID.Text.Trim() + "\t" + Selected_Step_InforLabel.Text.Trim() + "\t" +
                fc.ListBox2SQL_in_Query_String(DataSet_listBox).Trim() + "\t" + DataSetStatus.Text.Trim() + "\t" + DataName.Text.Trim() + "\t" + DataNameStatus.Text.Trim() + "\t" + DataNameVal.Text.Trim() + "\t"
               + StartTime.ToString() + "\t" + EndTime.ToString() + "\t" + LSL_Val.Text.Trim() + "\t" + USL_Val.Text.Trim() + "\t" + DateTime.Now.ToString() + "\t" + Comp_SN_textBox.Text.Trim()
               + "\t" + fc.ListBox2SQL_in_Query_String(Comp_Type_listBox).Trim() + "\t" + Comp_eDataName_comboBox.Text.ToString().Trim() + "\t" + Comp_eData_Value_textBox.Text.Trim() + "\t" +
               Test_DB_Enable_checkBox.Checked.ToString() + "\t" + Comp_eData_Include_checkBox.Checked.ToString() + "\t" + Search_FirstRecord_checkBox.Checked.ToString() + "\t" + Search_LastRecord_checkBox.Checked.ToString()+ "\t" +
               Comp_PN_textBox.Text.Trim() + "\t" +Comp_Already_Removed_checkBox.Checked.ToString());
            Sw.Flush();
            //关闭流
            Sw.Close();
            Fs.Close();

        }
        private void Export_Query_button_Click(object sender, EventArgs e)
        {
            string LogFilePath = System.AppDomain.CurrentDomain.BaseDirectory + "Log.txt";

            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "TXT files (*.txt)|*.txt";
            dlg.FilterIndex = 0;
            dlg.RestoreDirectory = true;
            dlg.CreatePrompt = true;
            dlg.Title = "保存为TXT文件";
            dlg.FilterIndex = 2;//记忆上次保存路径  


            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Stream myStream;
                myStream = dlg.OpenFile();
                StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));
                try
                {
                    FileStream Fs = new FileStream(LogFilePath, FileMode.Open);
                    StreamReader LogerSw = new StreamReader(Fs);
                    string st = string.Empty;
                    while (!LogerSw.EndOfStream)
                    {
                        st = LogerSw.ReadLine();//get last line of log file
                    }

                    sw.WriteLine(st);

                    Fs.Close();
                    sw.Close();
                    myStream.Close();
                }
                catch (Exception exxxx)
                {
                    MessageBox.Show(exxxx.ToString());
                }
                finally
                {
                    
                    sw.Close();
                    myStream.Close();
                }
            }
        }

        private void Import_Query_button_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "TXT files|*.txt";
            dlg.FilterIndex = 0;
            dlg.RestoreDirectory = true;
            //dlg. = true;
            dlg.Title = "Open txt file";





            String line = "";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                try
                {   // Open the text file using a stream reader.
                    Stream myStream = dlg.OpenFile();
                    StreamReader sr = new StreamReader(myStream);

                    // Read the stream to a string
                    line = sr.ReadToEnd();// DO NOT use Trim() here

                    ReadQuryHistory(line);

                    myStream.Close();
                    sr.Close();

                    dlg.Dispose();




                }
                catch (Exception exxxxxx)
                {
                    MessageBox.Show("The file could not be read:\n" + exxxxxx.Message);
                
                    dlg.Dispose();
                }


            }


        }

        private void ReadQuryHistory(string QueryHistroyString)
        {
            int QueryType = 0;
            int SN = 1;
            int JobID = 2;
            int BOM_PN = 3;
            int BOM_PN_Rev = 4;
            int Model_ID = 5;
            int Step = 6;
            int DataSet_Name = 7;
            int DataSet_Status = 8;
            int Data_Name = 9;
            int Data_Name_Status = 10;
            int Data_Name_Val = 11;
            int Start_Time = 12;
            int End_Time = 13;
            int LSL_Value = 14;
            int USL_Value = 15;
            int Qurery_Time = 16;
            int Comp_SN = 17;
            int Comp_Type = 18;
            int Comp_eDataName = 19;
            int Comp_eData_Value = 20;
            int Test_DB_Checked = 21;
            int Comp_eData_Include_checkBox_Checked = 22;
            int Search_First_Record_ChecBox_Checked = 23;
            int Search_Last_Record_ChecBox_Checked = 24;
            int Comp_PN = 25;
            int Comp_Already_Removed_checkBox_Checked = 26;


            string[] FieldName = QueryHistroyString.Split('\t');


            SerialNumber.Text = FieldName[SN];
            JobOrder.Text = FieldName[JobID];
            BOMPN.Text = FieldName[BOM_PN];
            BOMPNRev.Text = FieldName[BOM_PN_Rev];
            ModelID.Text = FieldName[Model_ID];
            Selected_Step_InforLabel.Text = FieldName[Step];
            Selected_Dataset_InforLabel.Text = FieldName[DataSet_Name];
            DataSetStatus.Text = FieldName[DataSet_Status];
            DataName.Text = FieldName[Data_Name];
            DataNameStatus.Text = FieldName[Data_Name_Status];
            DataNameVal.Text = FieldName[Data_Name_Val];

            //StartTime = FieldName[Start_Time];
            //EndTime = FieldName[End_Time];

            //string yyyymmdd073000;
            ////datetime.Format=DateTimePickerFormat.Custom;
            ////datetime.CustomFormat="yyyyMMdd";
            //yyyymmdd073000 = datetime.Value.ToString("yyyyMMdd") + "073000";
            //return yyyymmdd073000;

            StartTimePicker.Value = DateTime.ParseExact(FieldName[Start_Time], "yyyyMMddHHmmss", System.Globalization.CultureInfo.CurrentCulture);
            EndTimePicker.Value = DateTime.ParseExact(FieldName[End_Time], "yyyyMMddHHmmss", System.Globalization.CultureInfo.CurrentCulture);
            //StartTime_HHMMSS_textBox.Text = FieldName[Start_Time].Substring(8, FieldName[Start_Time].Length - 8);
            //EndTime_HHMMSS_textBox.Text = FieldName[End_Time].Substring(8, FieldName[End_Time].Length - 8);

            LSL_Val.Text = FieldName[LSL_Value];
            USL_Val.Text = FieldName[USL_Value];

            Comp_SN_textBox.Text = FieldName[Comp_SN];
            Comp_PN_textBox.Text = FieldName[Comp_PN];
            Comp_Type_InforLabel.Text = FieldName[Comp_Type];
            Comp_eDataName_comboBox.Text = FieldName[Comp_eDataName];
            Comp_eData_Value_textBox.Text = FieldName[Comp_eData_Value];
            Test_DB_Enable_checkBox.Checked = bool.Parse(FieldName[Test_DB_Checked]);
            Comp_eData_Include_checkBox.Checked = bool.Parse(FieldName[Comp_eData_Include_checkBox_Checked]);
            Search_FirstRecord_checkBox.Checked = bool.Parse(FieldName[Search_First_Record_ChecBox_Checked]);
            Search_LastRecord_checkBox.Checked = bool.Parse(FieldName[Search_Last_Record_ChecBox_Checked]);
            Comp_Already_Removed_checkBox.Checked = bool.Parse(FieldName[Comp_Already_Removed_checkBox_Checked]);

            YieldEnablecheckBox.Checked = false;
            CPKEnablecheckBox.Checked = false;
            Search_Enable_checkBox.Checked = false;
            DataSet_TestTime_checkBox.Checked = false;

            if (FieldName[QueryType] == "Search")
                Search_Enable_checkBox.Checked = true;

            if (FieldName[QueryType] == "FPY&FY")
                YieldEnablecheckBox.Checked = true;

            if (FieldName[QueryType] == "CPK")
                CPKEnablecheckBox.Checked = true;

            if (FieldName[QueryType] == "TestTime")
                DataSet_TestTime_checkBox.Checked = true;



            



        }




        private void FPYRawdataGridView_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            fc.Search_Open_TestFile_From_Production_PC(FPYRawdataGridView.CurrentRow.Cells[0].Value.ToString(), FPYRawdataGridView.CurrentRow.Cells[8].Value.ToString(), FPYRawdataGridView.CurrentRow.Cells[9].Value.ToString());
        }

        private void FYRawdataGridView_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            fc.Search_Open_TestFile_From_Production_PC(FYRawdataGridView.CurrentRow.Cells[0].Value.ToString(), FYRawdataGridView.CurrentRow.Cells[8].Value.ToString(), FYRawdataGridView.CurrentRow.Cells[9].Value.ToString());
        }

        private void SearchDataGridView_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (YieldEnablecheckBox.Checked)
                fc.Search_Open_TestFile_From_Production_PC(SearchDataGridView.CurrentRow.Cells[0].Value.ToString(), SearchDataGridView.CurrentRow.Cells[8].Value.ToString(), SearchDataGridView.CurrentRow.Cells[9].Value.ToString());

            else
                fc.Search_Open_TestFile_From_Production_PC(SearchDataGridView.CurrentRow.Cells[0].Value.ToString(), SearchDataGridView.CurrentRow.Cells[12].Value.ToString(), SearchDataGridView.CurrentRow.Cells[13].Value.ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SN_WhereUsed.Text = fc.ReadListFromTxt();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Component_Edata_SN_textBox.Text = fc.ReadListFromTxt();
        }

        private void Component_Edata_Search_button_Click(object sender, EventArgs e)
        {
            Component_Edata_search_Infor_label.Text = "Searching...";
            Component_Edata_search_Infor_label.ForeColor = System.Drawing.Color.Red;
            Application.DoEvents();
            DataSet ds;

            ds = fc.Component_Edata_Search(OracleConnectString, Component_Edata_SN_textBox.Text.Trim(), Component_Edata_DataName_textBox.Text.Trim(), Component_Edata_DataVal_textBox.Text.Trim());

            Component_Edata_dataGridView.DataSource = ds;

            Component_Edata_dataGridView.DataMember = ds.Tables[0].TableName;

            Component_Edata_search_Infor_label.ForeColor = System.Drawing.Color.Black;
            int ContainerNumber = fc.ContainerNumberCounter(ds.Tables[0]);
            Component_Edata_search_Infor_label.Text = "Find total " + ds.Tables[0].Rows.Count + " recorders...The active container number is " + ContainerNumber.ToString();


        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Component_Edata_DataName_textBox.Text = fc.ReadListFromTxt();
        }

        private void Component_Edata_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Component_Edata_checkBox.Checked)

                Component_Edata_Tab.Parent = MainTabControl;

            else
                Component_Edata_Tab.Parent = null;


        }




        private void Component_Edata_DataName_textBox_TextChanged(object sender, EventArgs e)
        {

            if (Component_Edata_DataName_textBox.Text.Trim() != "")
            {
                Component_Edata_DataVal_textBox_Label.Enabled = true;
                Component_Edata_DataVal_textBox.Enabled = true;
            }
            else
            {
                Component_Edata_DataVal_textBox_Label.Enabled = false;
                Component_Edata_DataVal_textBox.Enabled = false;
            }


        }

        private void Component_Edata_export_button_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(Component_Edata_dataGridView);
        }

        private void Comp_SN_Import_button_Click(object sender, EventArgs e)
        {
            Comp_SN_textBox.Text = fc.ReadListFromTxt();

        }

        private void Comp_eDataName_button_Click(object sender, EventArgs e)
        {

            Comp_eDataName_comboBox.Text = fc.ReadListFromTxt();
        }

        private void Comp_eDataName_comboBox_TextUpdate(object sender, EventArgs e)
        {
            if (Comp_eDataName_comboBox.Text.Trim() != "")
            {
                Comp_eData_Value_label.Enabled = true;
                Comp_eData_Value_textBox.Enabled = true;
            }
            else
            {
                Comp_eData_Value_label.Enabled = false;
                Comp_eData_Value_textBox.Enabled = false;
            }
        }

        private void Comp_eData_Include_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Comp_eData_Include_checkBox.Checked)
            {
                Comp_SN_label.Enabled = true;
                Comp_SN_textBox.Enabled = true;
                Comp_SN_Import_button.Enabled = true;

                Comp_PN_label.Enabled = true;
                Comp_PN_textBox.Enabled = true;
                Comp_PN_Import_button.Enabled = true;

                Comp_Type_label.Enabled = true;
                Comp_Type_listBox.Enabled = true;


                Comp_Edata_Name_label.Enabled = true;
                Comp_eDataName_comboBox.Enabled = true;
                Comp_eDataName_button.Enabled = true;

                Comp_Already_Removed_checkBox.Enabled = true;

            }
            else
            {
                Comp_SN_label.Enabled = false; ;
                Comp_SN_textBox.Enabled = false;
                Comp_SN_Import_button.Enabled = false;

                Comp_PN_label.Enabled = false;
                Comp_PN_textBox.Enabled = false;
                Comp_PN_Import_button.Enabled = false;

                Comp_Type_label.Enabled = false;
                Comp_Type_listBox.Enabled = false;

                Comp_Edata_Name_label.Enabled = false;
                Comp_eDataName_comboBox.Enabled = false;
                Comp_eDataName_button.Enabled = false;

                Comp_Already_Removed_checkBox.Enabled = false;
                Comp_Already_Removed_checkBox.Checked = false;

            }


        }

        private void Search_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Search_Enable_checkBox.Checked)
            {
                YieldEnablecheckBox.Enabled = false;
                CPKEnablecheckBox.Enabled = false;
                DataSet_TestTime_checkBox.Enabled = false;

                SearchButton.Enabled = true;
                Comp_eData_Include_checkBox.Enabled = true;
                Search_FirstRecord_checkBox.Enabled = true;
                Search_LastRecord_checkBox.Enabled = true;

            }
            else
            {
                YieldEnablecheckBox.Enabled = true;
                CPKEnablecheckBox.Enabled = true;
                DataSet_TestTime_checkBox.Enabled = true;

                SearchButton.Enabled = false;
                Comp_eData_Include_checkBox.Enabled = false;
                Comp_eData_Include_checkBox.Checked = false;
                Search_FirstRecord_checkBox.Enabled = false;
                Search_LastRecord_checkBox.Enabled = false;
                // if Search_ checkbox is unchecked, then search fisrt or search last checkbox should be unchecked as well
                Search_FirstRecord_checkBox.Checked = false;
                Search_LastRecord_checkBox.Checked = false;


            }



        }

        private void ModelID_TextChanged(object sender, EventArgs e)
        {
            if (ModelID.Text != "")
            {
                //string tempmModerlID=ModelID.Text;
                //ModelID.Text = tempmModerlID.ToUpper();

                string tempBOMPN = BOMPN.Text;
                DataSet ds = new DataSet();
                string sql = "Select DISTINCT f.BOM_PN FROM BOM_CONTEXT_ID f Where f.MODEL_ID in ('" + ModelID.Text.ToString().Replace(",", "','") + "')";

                try
                {
                    ds = fc.GetOracleDataSet2(OracleConnectString, sql);
                    BOMPN.Text = "";
                    for (var i = 0; i < ds.Tables[0].Rows.Count - 1; i++)
                        BOMPN.Text += ds.Tables[0].Rows[i][0].ToString() + ",";
                    BOMPN.Text += ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1][0].ToString();//last element should be withour ","
                }
                catch (Exception exxx)
                {
                    BOMPN.Text = tempBOMPN;
                }



            }
        }

        private void DataSetName_Import_button_Click(object sender, EventArgs e)
        {
            Selected_Dataset_InforLabel.Text = fc.ReadListFromTxt();
            //Selected_Dataset_InforLabel.ForeColor = System.Drawing.Color.Blue;
        }

        

        private void Search_FirstRecord_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Search_FirstRecord_checkBox.Checked)
                Search_LastRecord_checkBox.Checked = false;

        }

        private void Search_LastRecord_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Search_LastRecord_checkBox.Checked)
                Search_FirstRecord_checkBox.Checked = false;
        }

        private void Comp_PN_Import_button_Click(object sender, EventArgs e)
        {
            Comp_PN_textBox.Text = fc.ReadListFromTxt();
        }

        private void WhereUsed_dataGridView_Export_button_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(WhereUsed_dataGridView);
        }

        private void WIP_Status_JobID_Import_button_Click(object sender, EventArgs e)
        {
            WIP_Status_JobID_textBox.Text = fc.ReadListFromTxt();
            
        }

        private void WIP_Status_Export_button_Click(object sender, EventArgs e)
        {
            fc.DataGridViewToExcel(WIP_Status_dataGridView);
        }

        private void WIP_Status_Search_button_Click(object sender, EventArgs e)
        {

            WIP_Status_Infor_label.Text = "Searching...";
            WIP_Status_Infor_label.ForeColor = System.Drawing.Color.Red;
            Application.DoEvents();
            DataSet ds;

            ds = fc.WIP_Status_Search(OracleConnectString, WIP_Status_JobID_textBox.Text.Trim());


            WIP_Status_dataGridView.DataSource = ds;

            WIP_Status_dataGridView.DataMember = ds.Tables[0].TableName;

            WIP_Status_Infor_label.ForeColor = System.Drawing.Color.Black;
            int ContainerNumber = fc.ContainerNumberCounter(ds.Tables[0]);
            WIP_Status_Infor_label.Text = "Find total " + ds.Tables[0].Rows.Count + " recorders...The active container number is " + ContainerNumber.ToString();

        }

        private void WIP_Status_Enable_checkBox_CheckedChanged(object sender, EventArgs e)
        {

            if (WIP_Status_Enable_checkBox.Checked)


                WIP_Status_tabPage.Parent = MainTabControl;
            else
                WIP_Status_tabPage.Parent = null;
        }



    }
















}

