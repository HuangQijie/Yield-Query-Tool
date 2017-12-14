using System;
using System.Data.OracleClient;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Windows.Forms.DataVisualization.Charting;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Globalization;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Threading;





namespace Yield_Query_Tool
{
    class Function
    {

        //public AccessDataBaseOperator accessdbo = new AccessDataBaseOperator();
        //public OracleDataBaseOperator oracledbo = new OracleDataBaseOperator();


        String sql = "";
        DataSet ds = new DataSet();



        public void InitialGetFromConfig(string dbPath, ListBox Step_listBox, ListBox DataSet_listBox, ComboBox DataName)
        {


            //OleDbConnection AccessConnection = accessdbo.OpenAccessConnection2(dbPath);


            //sql = "select distinct [Part Number] from [PN_List]";
            //ds = accessdbo.GetAccessDataset2(sql, AccessConnection);
            //PartNumber.Items.Add("");
            //for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            //{
            //    PartNumber.Items.Add(ds.Tables[0].Rows[i][0]).ToString().Trim();
            //}



            sql = "select distinct [Step] from [Step_DataSet_DataName]";


            //oracledbo.OpenOracleConnection2("Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=shg-oracle)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=ateshg)));User Id=extviewer;Password=extviewer");


            //ds = accessdbo.GetAccessDataset2(sql, AccessConnection);
            ds = GetAccessDataset2(dbPath, sql);


            //Step_listBox.Items.Add("");
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {

                Step_listBox.Items.Add(ds.Tables[0].Rows[i][0]).ToString().Trim();
            }
            ds.Dispose();


            sql = "select distinct [DataSet] from [Step_DataSet_DataName]";
            //ds = accessdbo.GetAccessDataset2(sql, AccessConnection);
            ds = GetAccessDataset2(dbPath, sql);

            //DataSet_listBox.Items.Add("");
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {

                DataSet_listBox.Items.Add(ds.Tables[0].Rows[i][0]).ToString().Trim();
            }
            ds.Dispose();







            sql = "select distinct [DataName] from [Step_DataSet_DataName]";
            //ds = accessdbo.GetAccessDataset2(sql, AccessConnection);
            ds = GetAccessDataset2(dbPath, sql);

            //DataName_listBox.Items.Add("");
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {

                DataName.Items.Add(ds.Tables[0].Rows[i][0]).ToString().Trim();
            }
            ds.Dispose();





        }
        public void InitialStatusTab(ComboBox DataSetStatus, ComboBox DataStatus, ListBox Comp_Type_listBox)
        {
            DataSetStatus.Items.Add("");
            DataSetStatus.Items.Add("PASS");
            DataSetStatus.Items.Add("FAIL");
            DataSetStatus.Items.Add("INFO");

            DataStatus.Items.Add("");
            DataStatus.Items.Add("NULL");
            DataStatus.Items.Add("PASS");
            DataStatus.Items.Add("FAIL");
            DataStatus.Items.Add("INFO");
            DataStatus.Items.Add("ERROR");

            Comp_Type_listBox.Items.Add("daughterBdSN");
            Comp_Type_listBox.Items.Add("laserSN");
            Comp_Type_listBox.Items.Add("PWBA");
            Comp_Type_listBox.Items.Add("receiverSN");

            Comp_Type_listBox.Items.Add("ROSA");
            Comp_Type_listBox.Items.Add("TOSA");
        }

        public void DataSetChangeWithStep(string dbPath, ListBox Step_listBox, ListBox DataSet_listBox, ComboBox DataName)
        {
            //OleDbConnection AccessConnection = accessdbo.OpenAccessConnection2(dbPath);
            //if (StepSelectedConext.Length==0)
            //    sql = "select distinct [DataSet] from [Step_DataSet_DataName]";
            //else


            DataSet_listBox.Items.Clear();
            DataName.Items.Clear();
            //DataName_listBox.Items.Add("");


            if (Step_listBox.SelectedItems.Count != 0)// first row is for blank
            {

                sql = "select distinct [DataSet] from [Step_DataSet_DataName] where [Step] in ('" + ListBox2SQL_in_Query_String(Step_listBox).Replace(",", "','") + "')";

                //DataSet_listBox.Items.Add("All");

            }
            else
                sql = "select distinct [DataSet] from [Step_DataSet_DataName]";
            //ds = accessdbo.GetAccessDataset2(sql, AccessConnection);
            ds = GetAccessDataset2(dbPath, sql);
            //DataSet.Items.Clear();

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {

                DataSet_listBox.Items.Add(ds.Tables[0].Rows[i][0]).ToString().Trim();
            }

            ds.Dispose();
            //AccessConnection.Dispose();
            //AccessConnection.Close();
            //accessdbo = null;
            //accessdbo.AccessConnectionClose();
        }

        public void DataNameChangeWithDataSet(string dbPath, ListBox Step_listBox, ListBox DataSet_listBox, ComboBox DataName)
        {
            //OleDbConnection AccessConnection = accessdbo.OpenAccessConnection2(dbPath);
            if (Step_listBox.SelectedItems.Count != 0 && DataSet_listBox.SelectedItems.Count != 0)
                sql = "select distinct [DataName] from [Step_DataSet_DataName] where [DataSet] in ('" + ListBox2SQL_in_Query_String(DataSet_listBox).Replace(",", "','")
                    + "') and [Step] in ('" + ListBox2SQL_in_Query_String(Step_listBox).Replace(",", "','") + "')";
            else if (Step_listBox.SelectedItems.Count == 0 && DataSet_listBox.SelectedItems.Count != 0)
                sql = "select distinct [DataName] from [Step_DataSet_DataName] where [DataSet] in ('" + ListBox2SQL_in_Query_String(DataSet_listBox).Replace(",", "','") + "')";
            else if (Step_listBox.SelectedItems.Count != 0 && DataSet_listBox.SelectedItems.Count == 0)
                sql = "select distinct [DataName] from [Step_DataSet_DataName] where [Step] in ('" + ListBox2SQL_in_Query_String(Step_listBox).Replace(",", "','") + "')";
            else if (Step_listBox.SelectedItems.Count == 0 && DataSet_listBox.SelectedItems.Count == 0)
                sql = "select distinct [DataName] from [Step_DataSet_DataName]";


            //ds = accessdbo.GetAccessDataset2(sql, AccessConnection);
            ds = GetAccessDataset2(dbPath, sql);
            DataName.Items.Clear();
            //DataName.Items.Add("");
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                DataName.Items.Add(ds.Tables[0].Rows[i][0]).ToString().Trim();
            }

            ds.Dispose();
            //AccessConnection.Dispose();
            // AccessConnection.Close();
            //accessdbo = null;
            //accessdbo.AccessConnectionClose();
        }

        public DataSet CPKOracleQuery(string connectionstring, string SerialNumber, string JobOrder,
            string BOMPN, string BOMPNRev, string ModelID, string DataSet, string DataName, string DataNameVal,
            string DataSetStatus, string DataStatus, string StartTime, string EndTime, Label InformLabel)
        {

            sql = "SELECT a.MFR_SN AS SN,b.JOB_ID  AS JOB_ID,  c.MODEL_ID  AS Model_ID,c.BOM_PN  AS BOM_PN,c.BOM_PN_REV BOM_PN_Rev,d.DATASET_NAME AS DataSet_Name,d.STATUS AS DataSetStatus,d.END_TIME AS DataSet_EndTime,e.DATA_NAME AS Data_Name,e.DATA_VAL1  AS Data_Val,e.DATA_VAL2 AS Data_Status,e.DATA_VAL3 AS Data_Spec, d.STATION as Station, j.Computer as Computer, d.ROUTE_ID as Route_ID" +
                   " FROM parts a INNER JOIN routes b ON b.PART_INDEX = a.OPT_INDEX INNER JOIN bom_context_id c ON c.BOM_CONTEXT_ID = b.BOM_CONTEXT_ID INNER JOIN datasets d ON d.ROUTE_ID = b.ROUTE_ID INNER JOIN STATIONS j on J.Name=D.Station INNER JOIN route_data e ON e.DATASET_ID = d.DATASET_ID  WHERE '1' = '1'";
            if (SerialNumber.Length != 0)
            {
                SerialNumber = SerialNumber.Replace(",", "','");
                sql += " and " + "a.MFR_SN in ('" + SerialNumber + "')";
            }
            if (JobOrder.Length != 0)
            {
                JobOrder = JobOrder.Replace(",", "','");
                sql += "and " + "b.JOB_ID in ('" + JobOrder + "')";
            }
            if (BOMPN.Length != 0)
            {
                BOMPN = BOMPN.Replace(",", "','");
                sql += "and " + "c.BOM_PN in ('" + BOMPN + "')";
                //sql += " and " + "c.BOM_PN='" + BOMPN + "'";
            }

            if (BOMPNRev.Length != 0)
                sql += " and " + "c.BOM_PN_REV='" + BOMPNRev + "'";

            if (ModelID.Length != 0)
            {
                ModelID = ModelID.Replace(",", "','");
                sql += "and " + "c.MODEL_ID in ('" + ModelID + "')";
                //sql += " and " + "c.MODEL_ID='" + ModelID + "'";
            }

            if (DataSet.Length != 0)
            {
                //if (DataSet == "All")
                //{
                //    DataSet = "";
                //    for (var i = 2; i < DataSet_Combx.Items.Count - 1; i++)//i=0, items=blank;i=1, items=All,so i start from 2
                //        DataSet += DataSet_Combx.Items[i].ToString() + ",";

                //    DataSet += DataSet_Combx.Items[DataSet_Combx.Items.Count - 1];
                //    DataSet = DataSet.Replace(",", "','");

                //    sql += " and " + "d.DATASET_NAME in ('" + DataSet + "')";
                //}
                //else
                sql += " and " + "d.DATASET_NAME in ('" + DataSet.Replace(",", "','") + "')";
            }

            if (DataName.Length != 0)
                if (DataName.Split(',').Length > 1)
                {
                    string[] tempstring = DataName.Split(',');
                    if (tempstring[0].Contains("%"))
                        sql += " and " + "(e.DATA_NAME like '" + tempstring[0] + "'";
                    else
                        sql += " and " + "(e.DATA_NAME = '" + tempstring[0] + "'";
                    for (var i = 1; i < tempstring.Length; i++)
                    {
                        if (tempstring[0].Contains("%"))
                            sql += " or " + "e.DATA_NAME like '" + tempstring[i] + "'";
                        else
                            sql += " or " + "e.DATA_NAME = '" + tempstring[i] + "'";
                    }
                    sql += ')';
                }
                else if (DataName.Contains("%"))
                    sql += " and " + "e.DATA_NAME like '" + DataName + "'";
                else
                    sql += " and " + "e.DATA_NAME = '" + DataName + "'";

            if (DataNameVal.Length != 0)
                sql += " and " + "e.DATA_VAL1 " + DataNameVal;

            if (DataSetStatus.Length != 0)
                sql += " and " + "d.STATUS='" + DataSetStatus + "'";

            if (DataStatus.Length != 0 && DataStatus != "NULL")
                sql += " and " + "e.DATA_VAL2='" + DataStatus + "'";
            if (DataStatus.Length != 0 && DataStatus == "NULL")
                sql += " and " + "e.DATA_VAL2 is NULL";


            if (StartTime.Length != 0 && EndTime.Length != 0)
                sql += " and " + "d.END_TIME between '" + StartTime + "' and '" + EndTime + "'";


            //----------------------------------------------
            //sql += " and D.Ate_Version like 'ATE_Code-Wuxi-1.00-%_DL'";
            //----------------------------------------------

            sql += "ORDER BY SN,  d.DATASET_NAME, d.END_TIME, e.DATA_NAME";

            //OracleConnection OracleConn = oracledbo.OpenOracleConnection2(connectionstring);
            //InformLabel.Text = "Searching...";
            //ds = oracledbo.GetOracleDataSet2(sql, OracleConn);
            ds = GetOracleDataSet2(connectionstring, sql);
            //oracledbo.OracleConnectionClose();
            //OracleConn.Close();
            return ds;
        }

        public DataSet YieldOracleQuery(string connectionstring, string SerialNumber, string JobOrder,
            string BOMPN, string BOMPNRev, string ModelID, string DataSet, string DataName,
            string StartTime, string EndTime, Label InformLabel)
        {
            DataSet pass;
            DataSet fail;
            // DataSet qs= new DataSet();
            DataSet ts = new DataSet();
            DataView vs;

            string passsql;
            string failsql;


            sql = "SELECT a.MFR_SN AS SN,b.JOB_ID  AS JOB_ID,  c.MODEL_ID  AS Model_ID,c.BOM_PN  AS BOM_PN,c.BOM_PN_REV BOM_PN_Rev," +
            "d.DATASET_NAME AS DataSet_Name,d.STATUS AS DataSetStatus,d.END_TIME AS DataSet_EndTime," +
            "e.DATA_NAME AS Data_Name,e.DATA_VAL1  AS Data_Val,e.DATA_VAL2 AS Data_Status,e.DATA_VAL3 AS Data_Spec," +
            "d.STATION as Station, j.Computer as Computer,d.ROUTE_ID as Route_ID," +
            "d.DATASET_ID as DataSet_ID, d.ATE_VERSION as ATE_Version," +
            "b.STATE as Current_State, b.STATUS as Current_Status" +
        " FROM parts a INNER JOIN routes b ON b.PART_INDEX = a.OPT_INDEX INNER JOIN bom_context_id c ON c.BOM_CONTEXT_ID = b.BOM_CONTEXT_ID" +
        " INNER JOIN datasets d ON d.ROUTE_ID = b.ROUTE_ID INNER JOIN STATIONS j on J.Name=D.Station" +
        " INNER JOIN route_data e ON e.DATASET_ID = d.DATASET_ID  WHERE '1' = '1'";
            if (SerialNumber.Length != 0)
            {
                SerialNumber = SerialNumber.Replace(",", "','");
                sql += " and " + "a.MFR_SN in ('" + SerialNumber + "')";
            }
            if (JobOrder.Length != 0)
            {
                JobOrder = JobOrder.Replace(",", "','");
                sql += "and " + "b.JOB_ID in ('" + JobOrder + "')";
            }
            if (BOMPN.Length != 0)
            {
                BOMPN = BOMPN.Replace(",", "','");
                sql += "and " + "c.BOM_PN in ('" + BOMPN + "')";
                //sql += " and " + "c.BOM_PN='" + BOMPN + "'";
            }

            if (BOMPNRev.Length != 0)
                sql += " and " + "c.BOM_PN_REV='" + BOMPNRev + "'";

            if (ModelID.Length != 0)
            {
                ModelID = ModelID.Replace(",", "','");
                sql += "and " + "c.MODEL_ID in ('" + ModelID + "')";
                //sql += " and " + "c.MODEL_ID='" + ModelID + "'";
            }
            if (DataSet.Length != 0)
            {

                sql += " and " + "d.DATASET_NAME in ('" + DataSet.Replace(",", "','") + "')";
            }
            //if (DataName.Length != 0)
            //    sql += " and " + "e.DATA_NAME like '" + DataName + "'";

            //if (DataSetStatus.Length != 0)
            //    sql += " and " + "d.STATUS='" + DataSetStatus + "'";

            //if (DataStatus.Length != 0 && DataStatus != "NULL")
            //    sql += " and " + "e.DATA_VAL2='" + DataStatus + "'";
            //if (DataStatus.Length != 0 && DataStatus == "NULL")
            //    sql += " and " + "e.DATA_VAL2 is NULL";


            if (StartTime.Length != 0 && EndTime.Length != 0)
                sql += " and " + "d.END_TIME between '" + StartTime + "' and '" + EndTime + "'";

            //  to sql out the pass conditon, only get the dataset level in oder to save time
            passsql = sql.Replace(",e.DATA_NAME AS Data_Name,e.DATA_VAL1  AS Data_Val,e.DATA_VAL2 AS Data_Status,e.DATA_VAL3 AS Data_Spec", "") + "and d.STATUS <>'FAIL'";
            passsql = passsql.Replace(" INNER JOIN route_data e ON e.DATASET_ID = d.DATASET_ID", "");
            //passsql += "ORDER BY SN,  d.DATASET_NAME, d.END_TIME";

            // to sql out the failed condition, get to dataname level
            failsql = sql + "and d.STATUS='FAIL' and (e.DATA_NAME = 'failed' or e.DATA_VAL2='FAIL'or e.DATA_VAL2='ERROR' or e.DATA_NAME like '%ERROR%') and e.DATA_VAL1 <> 'not run'";
            //failsql += "ORDER BY SN,  d.DATASET_NAME, d.END_TIME";


            ////------------------------------------------------------------
            //passsql += " and D.Ate_Version like 'ATE_Code-Wuxi-1.00-%_DL'";
            //failsql += " and D.Ate_Version like 'ATE_Code-Wuxi-1.00-%_DL'";
            ////------------------------------------------------------------





            //Task pass_task = Task.Run(() => { pass = GetOracleDataSet2(connectionstring, passsql); });

            Task<DataSet> pass_task = Task<DataSet>.Factory.StartNew(() =>
            {
                pass = GetOracleDataSet2(connectionstring, passsql);
                return pass;
            });
           

            Task<DataSet> fail_task = Task<DataSet>.Factory.StartNew(() =>
            {
                fail = GetOracleDataSet2(connectionstring, failsql);
                return fail;
            });

            pass = pass_task.Result;
            fail = fail_task.Result;
            
            pass_task.Dispose();
            fail_task.Dispose();
            



            ////------------------------------------------------------------
            //pass = GetOracleDataSet2(connectionstring, passsql);

            //fail = GetOracleDataSet2(connectionstring, failsql);
            ////------------------------------------------------------------

            pass.Tables[0].Columns.Add("DATA_NAME");
            pass.Tables[0].Columns.Add("DATA_VAL");
            pass.Tables[0].Columns.Add("DATA_STATUS");
            pass.Tables[0].Columns.Add("DATA_SPEC");

            pass.Merge(fail);
            vs = pass.Tables[0].DefaultView;
            vs.Sort = "SN, DATASET_NAME, DATASET_ENDTIME";
            ts.Tables.Add(vs.ToTable());
            //ts = vs.ToTable();
            //qs.Tables.Add(ts.Copy());
            //oracledbo.OracleConnectionClose();
            //OracleConn.Close();
            return ts;

        }

        public DataSet MPLOracleQuery(string connectionstring, string BOMPN, string BOMPNRev, string ModelID, string DataName, string TypeDesc)
        {
            sql = "SELECT   c.MODEL_ID  AS Model_ID,c.BOM_PN  AS BOM_PN,c.BOM_PN_REV AS BOM_PN_Rev, c.TIME AS TIME,f.*, g.* " +
                   " FROM BOM_CONTEXT_ID c inner join BOM_CONTEXT_DATA f on c.BOM_CONTEXT_ID=f.BOM_CONTEXT_ID inner join DATA_TYPES g on f.DATA_TYPE_ID=g.TYPE_ID WHERE '1' = '1'";

            if (BOMPN.Length != 0)
            {
                BOMPN = BOMPN.Replace(",", "','");
                sql += "and " + "c.BOM_PN in ('" + BOMPN + "')";
                //sql += " and " + "c.BOM_PN='" + BOMPN + "'";
            }

            if (BOMPNRev.Length != 0)
                sql += " and " + "c.BOM_PN_REV='" + BOMPNRev + "'";

            if (ModelID.Length != 0)
            {
                ModelID = ModelID.Replace(",", "','");
                sql += "and " + "c.MODEL_ID in ('" + ModelID + "')";
                //sql += " and " + "c.MODEL_ID='" + ModelID + "'";
            }


            if (DataName.Length != 0)
                sql += " and " + "f.DATA_NAME = '" + DataName + "'";

            if (TypeDesc.Length != 0)
                sql += " and " + "g.TYPE_DESC = '" + TypeDesc + "'";

            //sort from new to old
            sql += "ORDER BY c.TIME DESC";

            //OracleConnection OracleConn = oracledbo.OpenOracleConnection2(connectionstring);

            //ds = oracledbo.GetOracleDataSet2(sql, OracleConn);
            ds = GetOracleDataSet2(connectionstring, sql);
            //OracleConn.Close();
            //oracledbo.OracleConnectionClose(OracleConn);
            return ds;


        }

        public DataSet MainOracleQuery(string connectionstring, string SerialNumber, string JobOrder,
                    string BOMPN, string BOMPNRev, string ModelID, string DataSet, string DataName, string DataNameVal,
                    string DataSetStatus, string DataStatus, string StartTime, string EndTime, Label InformLabel, string Comp_SN, string Comp_Type,
            string Comp_edata_name, string Comp_edata_name_Val, bool Comp_eData_Include_checkBox_Checked, string Search_Record_Type, string Comp_PN,
            bool Comp_Already_Removed_checkBox_Checked, bool Ignore_PN_Rev)
        {
            if (Comp_eData_Include_checkBox_Checked)
            {

                sql = "SELECT a.MFR_SN AS Module_SN,b.JOB_ID  AS JOB_ID,  c.MODEL_ID  AS Model_ID,c.BOM_PN  AS BOM_PN,c.BOM_PN_REV BOM_PN_Rev," +
                "d.DATASET_NAME AS DataSet_Name,d.STATUS AS DataSetStatus,d.END_TIME AS DataSet_EndTime," +
                "e.DATA_NAME AS Data_Name,e.DATA_VAL1  AS Data_Val,e.DATA_VAL2 AS Data_Status,e.DATA_VAL3 AS Data_Spec, " +
                "d.STATION as Station, j.Computer as Computer, d.ROUTE_ID as Route_ID, d.DATASET_ID as DataSet_ID, d.ATE_VERSION as ATE_Version, " +
                "y.MFR_SN AS Component_SN, y.MFR_PN AS Component_PN, x.NOTES as Component_Type, v.DATA_NAME as Component_Edata_Name,v.DATA_VAL1 as Component_Edata_Val,v.TIME as Component_Edata_Time" +

                " FROM parts a INNER JOIN routes b ON b.PART_INDEX = a.OPT_INDEX" +
                " INNER JOIN bom_context_id c ON c.BOM_CONTEXT_ID = b.BOM_CONTEXT_ID" +
                " INNER JOIN datasets d ON d.ROUTE_ID = b.ROUTE_ID" +
                " INNER JOIN STATIONS j on J.Name=D.Station" +
                " INNER JOIN route_data e ON e.DATASET_ID = d.DATASET_ID" +
                " INNER JOIN Assembly x on x.PARENT_INDEX=a.OPT_INDEX" +
                " INNER JOIN PARTS y on y.OPT_INDEX=x.CHILD_INDEX" +
                " INNER JOIN PART_DATA v ON v.OPT_INDEX=x.CHILD_INDEX" +
                "  WHERE '1' = '1'";

                if (Comp_Already_Removed_checkBox_Checked)
                    sql = sql.Replace("INNER JOIN Assembly x on x.PARENT_INDEX=a.OPT_INDEX", "INNER JOIN Assembly_History x on x.ROUTE_ID=b.ROUTE_ID");
            }
            else
            {
                sql = "SELECT a.MFR_SN AS Module_SN,b.JOB_ID  AS JOB_ID,  c.MODEL_ID  AS Model_ID,c.BOM_PN  AS BOM_PN,c.BOM_PN_REV BOM_PN_Rev," +
                  "d.DATASET_NAME AS DataSet_Name,d.STATUS AS DataSetStatus,d.END_TIME AS DataSet_EndTime," +
                  "e.DATA_NAME AS Data_Name,e.DATA_VAL1  AS Data_Val,e.DATA_VAL2 AS Data_Status,e.DATA_VAL3 AS Data_Spec, " +
                  "d.STATION as Station, j.Computer as Computer, d.ROUTE_ID as Route_ID, d.DATASET_ID as DataSet_ID, d.ATE_VERSION as ATE_Version" +
                    //"y.MFR_SN AS Component_SN, y.MFR_PN AS Component_PN, x.NOTES as Component_Type, v.DATA_NAME as Component_Edata_Name,v.DATA_VAL1 as Component_Edata_Val,v.TIME as Component_Edata_Time" +

                  " FROM parts a INNER JOIN routes b ON b.PART_INDEX = a.OPT_INDEX" +
                  " INNER JOIN bom_context_id c ON c.BOM_CONTEXT_ID = b.BOM_CONTEXT_ID" +
                  " INNER JOIN datasets d ON d.ROUTE_ID = b.ROUTE_ID" +
                  " INNER JOIN STATIONS j on J.Name=D.Station" +
                  " INNER JOIN route_data e ON e.DATASET_ID = d.DATASET_ID" +
                    //" INNER JOIN Assembly x on x.PARENT_INDEX=a.OPT_INDEX" +
                    //" INNER JOIN PARTS y on y.OPT_INDEX=x.CHILD_INDEX" +
                    //" INNER JOIN PART_DATA v ON v.OPT_INDEX=x.CHILD_INDEX" +
                  "  WHERE '1' = '1'";
            }

            if (SerialNumber.Length != 0)
            {
                SerialNumber = SerialNumber.Replace(",", "','");
                sql += " and " + "a.MFR_SN in ('" + SerialNumber + "')";
            }
            if (JobOrder.Length != 0)
            {
                JobOrder = JobOrder.Replace(",", "','");
                sql += "and " + "b.JOB_ID in ('" + JobOrder + "')";
            }
            if (BOMPN.Length != 0)
            {
                BOMPN = BOMPN.Replace(",", "','");
                sql += "and " + "c.BOM_PN in ('" + BOMPN + "')";
                //sql += " and " + "c.BOM_PN='" + BOMPN + "'";
            }

            if (BOMPNRev.Length != 0)
                sql += " and " + "c.BOM_PN_REV='" + BOMPNRev + "'";

            if (ModelID.Length != 0)
            {
                ModelID = ModelID.Replace(",", "','");
                sql += "and " + "c.MODEL_ID in ('" + ModelID + "')";
                //sql += " and " + "c.MODEL_ID='" + ModelID + "'";
            }

            if (DataSet.Length != 0)
            {
                //if (DataSet == "All")
                //{
                //    DataSet = "";
                //    for (var i = 2; i < DataSet_Combx.Items.Count - 1; i++)//i=0, items=blank;i=1, items=All,so i start from 2
                //        DataSet += DataSet_Combx.Items[i].ToString() + ",";

                //    DataSet += DataSet_Combx.Items[DataSet_Combx.Items.Count - 1];
                //    DataSet = DataSet.Replace(",", "','");

                //    sql += " and " + "d.DATASET_NAME in ('" + DataSet + "')";
                //}
                //else
                sql += " and " + "d.DATASET_NAME in ('" + DataSet.Replace(",", "','") + "')";
            }

            if (DataName.Length != 0)
            {
                if (DataName.Split(',').Length > 1)
                {
                    string[] tempstring = DataName.Split(',');
                    if (tempstring[0].Contains("%"))
                        sql += " and " + "(e.DATA_NAME like '" + tempstring[0] + "'";
                    else
                        sql += " and " + "(e.DATA_NAME = '" + tempstring[0] + "'";
                    for (var i = 1; i < tempstring.Length; i++)
                    {
                        if (tempstring[i].Contains("%"))
                            sql += " or " + "e.DATA_NAME like '" + tempstring[i] + "'";
                        else
                            sql += " or " + "e.DATA_NAME = '" + tempstring[i] + "'";
                    }
                    sql += ')';
                }

                else if (DataName.Contains("%"))
                    sql += " and " + "e.DATA_NAME like '" + DataName + "'";
                else
                    sql += " and " + "e.DATA_NAME = '" + DataName + "'";

            }

            if (DataNameVal.Length != 0)
                sql += " and " + "e.DATA_VAL1 " + DataNameVal;

            if (DataSetStatus.Length != 0)
                sql += " and " + "d.STATUS='" + DataSetStatus + "'";

            if (DataStatus.Length != 0 && DataStatus != "NULL")
                sql += " and " + "e.DATA_VAL2='" + DataStatus + "'";
            if (DataStatus.Length != 0 && DataStatus == "NULL")
                sql += " and " + "e.DATA_VAL2 is NULL";


            if (StartTime.Length != 0 && EndTime.Length != 0)
                sql += " and " + "d.END_TIME between '" + StartTime + "' and '" + EndTime + "'";

            if (Comp_eData_Include_checkBox_Checked)
            {
                if (Comp_SN.Length != 0)
                {
                    Comp_SN = Comp_SN.Replace(",", "','");
                    sql += " and " + "y.MFR_SN in ('" + Comp_SN + "')";
                }

                if (Comp_PN.Length != 0)
                {
                    Comp_PN = Comp_PN.Replace(",", "','");
                    sql += " and " + "y.MFR_PN in ('" + Comp_PN + "')";
                }


                if (Comp_Type.Length != 0)
                {
                    Comp_Type = Comp_Type.Replace(",", "','");
                    sql += " and " + "x.NOTES in ('" + Comp_Type + "')";
                }


                if (Comp_edata_name.Length != 0)
                {
                    if (Comp_edata_name.Split(',').Length > 1)
                    {
                        string[] tempstring = Comp_edata_name.Split(',');
                        if (tempstring[0].Contains("%"))
                            sql += " and " + "(v.DATA_NAME like '" + tempstring[0] + "'";
                        else
                            sql += " and " + "(v.DATA_NAME = '" + tempstring[0] + "'";
                        for (var i = 1; i < tempstring.Length; i++)
                        {
                            if (tempstring[0].Contains("%"))
                                sql += " or " + "v.DATA_NAME like '" + tempstring[i] + "'";
                            else
                                sql += " or " + "v.DATA_NAME = '" + tempstring[i] + "'";
                        }
                        sql += ')';
                    }

                    else
                        sql += " and " + "v.DATA_NAME like '" + Comp_edata_name + "'";
                }

                if (Comp_edata_name_Val.Length != 0)
                    sql += " and " + "v.DATA_VAL1 " + Comp_edata_name_Val;
            }


            if (Search_Record_Type == "FirstRecord")
                sql += "ORDER BY a.MFR_SN,  d.DATASET_NAME, e.DATA_NAME, d.END_TIME";

            if (Search_Record_Type == "LastRecord")
                sql += "ORDER BY a.MFR_SN,  d.DATASET_NAME, e.DATA_NAME, d.END_TIME DESC";


            //----------------------------------------------
            //sql += " and D.Ate_Version like 'ATE_Code-Wuxi-1.00-%_DL'";
            //----------------------------------------------

            //sql += "ORDER BY SN,  d.DATASET_NAME, d.END_TIME, e.DATA_NAME";

            //OracleConnection OracleConn = oracledbo.OpenOracleConnection2(connectionstring);
            //InformLabel.Text = "Searching...";
            //ds = oracledbo.GetOracleDataSet2(sql, OracleConn);

            //Thread t = new Thread(() =>
            //{
            //    ds = GetOracleDataSet2(connectionstring, sql);

            //});

            //t.IsBackground = true;

            //t.Start();


            ds = GetOracleDataSet2(connectionstring, sql);

            if (Search_Record_Type == "FirstRecord" || Search_Record_Type == "LastRecord")
                ds = RemoveDuplicateRows(ds.Tables[0], Ignore_PN_Rev);


            return ds;
        }



        public string TimePickerFormat(DateTimePicker datetime)
        {
            string yyyymmddHHmmss;
            //datetime.Format=DateTimePickerFormat.Custom;
            //datetime.CustomFormat="yyyyMMdd";
            yyyymmddHHmmss = datetime.Value.ToString("yyyyMMddHHmmss");
            return yyyymmddHHmmss;

        }

        public void InitialTimePicker(DateTimePicker StartTime, DateTimePicker EndTime)
        {// pcik up now time as default starttimepicker and endtime picker
            //StartTime.Format = DateTimePickerFormat.Long;
            //StartTime.ShowUpDown = true;
            var date = DateTime.Now.Date.Add(new TimeSpan(8, 0, 0));


            StartTime.Text = date.ToString();
            EndTime.Text = date.ToString();


        }




        #region DateGridView导出到csv格式的Excel,use data stream
        /// <summary>   
        /// 常用方法，列之间加\t，一行一行输出，此文件其实是csv文件，不过默认可以当成Excel打开。   
        /// </summary>   
        /// <remarks>   

        /// </remarks>   
        /// <param name="dgv"></param>   
        public void DataGridViewToExcel(DataGridView dgv)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "CSV files (*.csv)|*.csv";
            dlg.FilterIndex = 0;
            dlg.RestoreDirectory = true;
            dlg.CreatePrompt = true;
            dlg.Title = "保存为CSV文件";
            dlg.FilterIndex = 2;//记忆上次保存路径  


            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Stream myStream;
                myStream = dlg.OpenFile();
                StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));
                string columnTitle = "";
                try
                {
                    //写入列标题   
                    for (int i = 0; i < dgv.ColumnCount; i++)
                    {
                        if (i > 0)
                        {
                            columnTitle += ",";
                        }
                        columnTitle += dgv.Columns[i].HeaderText;
                    }
                    sw.WriteLine(columnTitle);

                    //写入列内容   
                    for (int j = 0; j < dgv.Rows.Count; j++)
                    {
                        string columnValue = "";
                        for (int k = 0; k < dgv.Columns.Count; k++)
                        {
                            if (k > 0)
                            {
                                columnValue += ",";
                            }
                            if (dgv.Rows[j].Cells[k].Value == null)
                                columnValue += "";
                            else
                            {

                                columnValue += dgv.Rows[j].Cells[k].Value.ToString().Trim();
                                columnValue = columnValue.Replace("\n", "\\n");// replace newline symbol
                                columnValue = columnValue.Replace("\t", "\\t");// replace tab symbol
                                columnValue = columnValue.Replace("\r", "\\r");// replace return symbol
                            }

                        }

                        sw.WriteLine(columnValue);
                    }
                    sw.Close();
                    myStream.Close();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
                finally
                {
                    sw.Close();
                    myStream.Close();
                }
            }
        }
        #endregion

        //#region  DataGridView导出到Excel，use OLEDB

        //public void DataGridViewToExcel(DataGridView dataGridView)
        //{
        //    string filepath = Environment.GetEnvironmentVariable("TEMP");
        //    string[] col = new string[dataGridView.ColumnCount]; //用于存放datagridview中的列名
        //    for (int i = 0; i < dataGridView.ColumnCount; i++)
        //    {
        //        col[i] = dataGridView.Columns[i].HeaderText;
        //    }
        //    string file = filepath + "\\" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xls";

        //    string OLEDBConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + file + ";";
        //    OLEDBConnStr += "Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
        //    string strCreateTableSQL = @" CREATE TABLE ";
        //    strCreateTableSQL += " [导出数据] ";
        //    strCreateTableSQL += " ( ";
        //    for (int i = 0; i < dataGridView.ColumnCount - 1; i++)
        //    {
        //        strCreateTableSQL += "[" + col[i] + "]" + " varchar(50),";
        //    }
        //    strCreateTableSQL += "[" + col[dataGridView.ColumnCount - 1] + "]" + " varchar(50)";
        //    strCreateTableSQL += " ) ";

        //    OleDbConnection oConn = new OleDbConnection();

        //    oConn.ConnectionString = OLEDBConnStr;
        //    OleDbCommand oCreateComm = new OleDbCommand();
        //    oCreateComm.Connection = oConn;
        //    oCreateComm.CommandText = strCreateTableSQL;

        //    oConn.Open();
        //    oCreateComm.ExecuteNonQuery(); //执行创建表操作

        //    //将result中的数据插入excel中
        //    string insertstr = "Insert into [导出数据] (";
        //    for (int i = 0; i < dataGridView.ColumnCount - 1; i++)
        //    {
        //        insertstr += "[" + col[i] + "]" + ",";
        //    }
        //    insertstr += "[" + col[dataGridView.ColumnCount - 1] + "]" + ") values (";

        //    for (int i = 0; i < dataGridView.RowCount; i++)
        //    {
        //        string temp = insertstr;
        //        for (int j = 0; j < dataGridView.ColumnCount - 1; j++)
        //        {
        //            temp += "'" + Convert.ToString((dataGridView.Rows[i]).Cells[j].Value).Replace("'", "''") + "',";
        //        }
        //        temp += "'" + Convert.ToString((dataGridView.Rows[i]).Cells[dataGridView.ColumnCount - 1].Value).Replace("'", "''") + "')";
        //        oCreateComm.CommandText = temp;
        //        oCreateComm.ExecuteNonQuery(); //执行插入数据操作
        //    }


        //    //for (int i = 0; i < dataGridView.RowCount; i++)
        //    //{
        //    //    if (dataGridView.Rows[i].Cells[0].Value == null)
        //    //    {
        //    //        break;
        //    //    }
        //    //    string temp = insertstr;
        //    //    for (int j = 0; j < dataGridView.ColumnCount - 1; j++)
        //    //    {
        //    //        temp += @"'" + dataGridView.Rows[i].Cells[j].Value.ToString() + "',";
        //    //    }
        //    //    temp += @"'" + dataGridView.Rows[i].Cells[dataGridView.ColumnCount - 1].Value.ToString() + "')";
        //    //    oCreateComm.CommandText = temp;
        //    //    oCreateComm.ExecuteNonQuery(); //执行插入数据操作
        //    //}

        //    oConn.Close();

        //    System.Diagnostics.Process.Start(@"excel.EXE", file);

        //}

        //#endregion

        //#region DataGridView导出到Excel，有一定的判断性
        ///// <summary>    
        /////方法，导出DataGridView中的数据到Excel文件    
        ///// </summary>    
        ///// <remarks>   

        ///// </remarks>   
        ///// <param name= "dgv"> DataGridView </param>    
        //public void DataGridViewToExcel(DataGridView dgv)
        //{


        //    #region   验证可操作性

        //    //申明保存对话框    
        //    SaveFileDialog dlg = new SaveFileDialog();
        //    //默然文件后缀    
        //    dlg.DefaultExt = "xlsx ";
        //    //文件后缀列表    
        //    dlg.Filter = "Excel文件(*.xlsx)|*.xlsx ";
        //    //默然路径是系统当前路径    
        //    dlg.InitialDirectory = Directory.GetCurrentDirectory();
        //    //打开保存对话框    
        //    if (dlg.ShowDialog() == DialogResult.Cancel) return;
        //    //返回文件路径    
        //    string fileNameString = dlg.FileName;
        //    //验证strFileName是否为空或值无效    
        //    if (fileNameString.Trim() == " ")
        //    { return; }
        //    //定义表格内数据的行数和列数    
        //    int rowscount = dgv.Rows.Count;
        //    int colscount = dgv.Columns.Count;
        //    //行数必须大于0    
        //    if (rowscount <= 0)
        //    {
        //        MessageBox.Show("没有数据可供保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        return;
        //    }

        //    //列数必须大于0    
        //    if (colscount <= 0)
        //    {
        //        MessageBox.Show("没有数据可供保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        return;
        //    }

        //    //行数不可以大于65536    
        //    if (rowscount > 65536)
        //    {
        //        MessageBox.Show("数据记录数太多(最多不能超过65536条)，不能保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        return;
        //    }

        //    //列数不可以大于255    
        //    if (colscount > 255)
        //    {
        //        MessageBox.Show("数据记录行数太多，不能保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        return;
        //    }

        //    //验证以fileNameString命名的文件是否存在，如果存在删除它    
        //    FileInfo file = new FileInfo(fileNameString);
        //    if (file.Exists)
        //    {
        //        try
        //        {
        //            file.Delete();
        //        }
        //        catch (Exception error)
        //        {
        //            MessageBox.Show(error.Message, "删除失败 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //            return;
        //        }
        //    }
        //    #endregion
        //    Excel.Application objExcel = null;
        //    Excel.Workbook objWorkbook = null;
        //    Excel.Worksheet objsheet = null;
        //    try
        //    {
        //        //申明对象    
        //        objExcel = new Microsoft.Office.Interop.Excel.Application();
        //        objWorkbook = objExcel.Workbooks.Add(Missing.Value);
        //        objsheet = (Excel.Worksheet)objWorkbook.ActiveSheet;
        //        //设置EXCEL不可见    
        //        objExcel.Visible = false;

        //        //向Excel中写入表格的表头    
        //        int displayColumnsCount = 1;
        //        for (int i = 0; i <= dgv.ColumnCount - 1; i++)
        //        {
        //            if (dgv.Columns[i].Visible == true)
        //            {
        //                objExcel.Cells[1, displayColumnsCount] = dgv.Columns[i].HeaderText.Trim();
        //                displayColumnsCount++;
        //            }
        //        }
        //        //设置进度条    
        //        //tempProgressBar.Refresh();    
        //        //tempProgressBar.Visible   =   true;    
        //        //tempProgressBar.Minimum=1;    
        //        //tempProgressBar.Maximum=dgv.RowCount;    
        //        //tempProgressBar.Step=1;    
        //        //向Excel中逐行逐列写入表格中的数据    
        //        for (int row = 0; row <= dgv.RowCount - 1; row++)
        //        {
        //            //tempProgressBar.PerformStep();    

        //            displayColumnsCount = 1;
        //            for (int col = 0; col < colscount; col++)
        //            {
        //                if (dgv.Columns[col].Visible == true)
        //                {
        //                    try
        //                    {
        //                        objExcel.Cells[row + 2, displayColumnsCount] = dgv.Rows[row].Cells[col].Value.ToString().Trim();
        //                        displayColumnsCount++;
        //                    }
        //                    catch (Exception error)
        //                    {
        //                        MessageBox.Show(error.Message, "警告 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                        return;
        //                    }

        //                }
        //            }
        //        }
        //        //隐藏进度条    
        //        //tempProgressBar.Visible   =   false;    
        //        //保存文件    
        //        objWorkbook.SaveAs(fileNameString, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
        //                Missing.Value, Excel.XlSaveAsAccessMode.xlShared, Missing.Value, Missing.Value, Missing.Value,
        //                Missing.Value, Missing.Value);
        //    }
        //    catch (Exception error)
        //    {
        //        MessageBox.Show(error.Message, "警告 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }
        //    finally
        //    {
        //        //关闭Excel应用    
        //        if (objWorkbook != null) objWorkbook.Close(Missing.Value, Missing.Value, Missing.Value);
        //        if (objExcel.Workbooks != null) objExcel.Workbooks.Close();
        //        if (objExcel != null) objExcel.Quit();

        //        objsheet = null;
        //        objWorkbook = null;
        //        objExcel = null;
        //    }
        //    MessageBox.Show(fileNameString + "\n\n导出完毕! ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);

        //}

        //#endregion


        public DataSet FPYandFYDataFilter(DataTable QueryResultTable)
        {
            //SN is column 1
            //DataSetName is column 6
            //DataSetStatus is column 7
            //DataStatus is column 11
            int SN = 1;
            int JobID = 2;
            int ModelID = 3;
            int BOMPN = 4;
            int BOMPNRev = 5;
            int DataSetName = 6;
            int DataSetStatus = 7;
            int DataSetEndtime = 8;
            int Station = 9;
            int Computer = 10;
            int RouteID = 11;
            int DataName = 12;
            int DataValue = 13;
            int DataStatus = 14;
            int DataSpec = 15;



            int Flag = QueryResultTable.Columns.Count + 1;
            //int RowNumber;

            DataSet FPYandFYRawData = new DataSet("FPYandFYRawData");
            //DataTable TotalData = FPYandFYRawData.Tables.Add("TotalData");
            //DataTable FPYRawData = FPYandFYRawData.Tables.Add("FPYRawData");
            //DataTable FYRawData = FPYandFYRawData.Tables.Add("FYRawData");
            DataTable TotalData;
            DataTable FPYRawData;
            DataTable FYRawData;

            QueryResultTable.Columns.Add("DeleteFlag", typeof(string));




            QueryResultTable.AcceptChanges();
            TotalData = QueryResultTable.Copy();
            FPYRawData = QueryResultTable.Copy();
            FYRawData = QueryResultTable.Copy();
            TotalData.TableName = "TotalData";
            FPYRawData.TableName = "FPYRawData";
            FYRawData.TableName = "FYRawData";



            #region filt and get FPY data
            for (int i = 0; i <= FPYRawData.Rows.Count - 1; i++)//Filter out the FIRST record of each datasetname
            {


                string tempSN = FPYRawData.Rows[i][SN - 1].ToString();
                string tempdatasetname = FPYRawData.Rows[i][DataSetName - 1].ToString();
                string tempdatasettime = FPYRawData.Rows[i][DataSetEndtime - 1].ToString();
                string tempBOMPN = FPYRawData.Rows[i][BOMPN - 1].ToString();
                string tempBOMPNRev = FPYRawData.Rows[i][BOMPNRev - 1].ToString();


                for (int j = i + 1; j <= FPYRawData.Rows.Count - 1; j++)
                {
                    if (FPYRawData.Rows[j][SN - 1].ToString() == tempSN &&
                        FPYRawData.Rows[j][BOMPN - 1].ToString() == tempBOMPN &&
                        FPYRawData.Rows[j][BOMPNRev - 1].ToString() == tempBOMPNRev &&
                        FPYRawData.Rows[j][DataSetName - 1].ToString() == tempdatasetname)
                    {
                        if (FPYRawData.Rows[j][DataName - 1].ToString().LastIndexOf("ERROR") > 0 &&
                            FPYRawData.Rows[j][DataSetEndtime - 1].ToString() == tempdatasettime)
                            // if dataname consist of ERROR plus datasettime is the same, then use it into yield report data and delete the first or last record row, because this %_ERROR row usually contain more useful information
                            FPYRawData.Rows[i][Flag - 1] = "Del";

                        else
                            FPYRawData.Rows[j][Flag - 1] = "Del";



                    }
                }




            }


            for (int i = FPYRawData.Rows.Count - 1; i >= 0; i--)
            {
                if (FPYRawData.Rows[i][Flag - 1] == "Del")
                    FPYRawData.Rows[i].Delete();
            }

            FPYRawData.AcceptChanges();



            #endregion

            #region filt and get FY  data
            for (int i = FYRawData.Rows.Count - 1; i >= 0; i--)//get the LAST record for each dataset name
            {

                string tempSN = FYRawData.Rows[i][SN - 1].ToString();
                string tempdatasetname = FYRawData.Rows[i][DataSetName - 1].ToString();
                string tempdatasettime = FYRawData.Rows[i][DataSetEndtime - 1].ToString();
                string tempBOMPN = FYRawData.Rows[i][BOMPN - 1].ToString();
                string tempBOMPNRev = FYRawData.Rows[i][BOMPNRev - 1].ToString();



                for (int j = i - 1; j >= 0; j--)
                {
                    if (FYRawData.Rows[j][SN - 1].ToString() == tempSN &&
                        FYRawData.Rows[j][BOMPN - 1].ToString() == tempBOMPN &&
                        FYRawData.Rows[j][BOMPNRev - 1].ToString() == tempBOMPNRev &&
                        FYRawData.Rows[j][DataSetName - 1].ToString() == tempdatasetname)
                    {
                        if (FYRawData.Rows[j][DataName - 1].ToString().LastIndexOf("ERROR") > 0 &&
                             FYRawData.Rows[j][DataSetEndtime - 1].ToString() == tempdatasettime)

                            // if dataname consist of ERROR, then use it into yield report data and delete the first or last record row, because this %_ERROR row usually contain more useful information
                            FYRawData.Rows[i][Flag - 1] = "Del";
                        else
                            FYRawData.Rows[j][Flag - 1] = "Del";
                    }

                }




            }
            for (int i = FYRawData.Rows.Count - 1; i >= 0; i--)
            {
                if (FYRawData.Rows[i][Flag - 1] == "Del")
                    FYRawData.Rows[i].Delete();
            }

            FYRawData.AcceptChanges();

            #endregion

            FPYandFYRawData.Tables.Add(TotalData.Copy());
            FPYandFYRawData.Tables.Add(FPYRawData.Copy());
            FPYandFYRawData.Tables.Add(FYRawData.Copy());


            return FPYandFYRawData;
        }






        public DataSet FillYieldTable(DataTable sourceyieldtable)//read and generate FPY and FY table
        {
            DataTable resultyieldtable = new DataTable();
            DataSet returndataset = new DataSet();

            int SN = 1;
            int JobID = 2;
            int ModelID = 3;
            int BOMPN = 4;
            int BOMPNRev = 5;
            int DataSetName = 6;
            int DataSetStatus = 7;
            int DataSetEndtime = 8;
            int Station = 9;
            int Computer = 10;
            int RouteID = 11;
            int DataName = 12;
            int DataVal = 13;
            int DataStatus = 14;
            int DataSpec = 15;

            int Flag = 16;

            resultyieldtable.Columns.Add("ModelID");//0   
            resultyieldtable.Columns.Add("BOMPN");//1
            resultyieldtable.Columns.Add("BOMPNRev");//2
            resultyieldtable.Columns.Add("DataSetName");//3
            resultyieldtable.Columns.Add("PASS QTY", typeof(float));//4
            resultyieldtable.Columns.Add("FAIL QTY", typeof(float));//5
            resultyieldtable.Columns.Add("Ratio");//6


            resultyieldtable.Rows.Add(sourceyieldtable.Rows[0][ModelID - 1], sourceyieldtable.Rows[0][BOMPN - 1], sourceyieldtable.Rows[0][BOMPNRev - 1], sourceyieldtable.Rows[0][DataSetName - 1], 0, 0, 0);

            for (int i = 1; i <= sourceyieldtable.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= resultyieldtable.Rows.Count - 1; j++)
                {
                    if (resultyieldtable.Rows[j][0].ToString() == sourceyieldtable.Rows[i][ModelID - 1].ToString() &&
                        resultyieldtable.Rows[j][1].ToString() == sourceyieldtable.Rows[i][BOMPN - 1].ToString() &&
                        resultyieldtable.Rows[j][2].ToString() == sourceyieldtable.Rows[i][BOMPNRev - 1].ToString() &&
                        resultyieldtable.Rows[j][3].ToString() == sourceyieldtable.Rows[i][DataSetName - 1].ToString())
                    {
                        break;
                    }
                    else if (j == resultyieldtable.Rows.Count - 1)//last row
                    {
                        resultyieldtable.Rows.Add(sourceyieldtable.Rows[i][ModelID - 1], sourceyieldtable.Rows[i][BOMPN - 1], sourceyieldtable.Rows[i][BOMPNRev - 1], sourceyieldtable.Rows[i][DataSetName - 1], 0, 0, 0);

                    }

                }

            }


            // caculate the yield ratio
            for (int j = 0; j <= resultyieldtable.Rows.Count - 1; j++)
            {
                for (int i = 0; i <= sourceyieldtable.Rows.Count - 1; i++)
                {
                    if (resultyieldtable.Rows[j][0].ToString() == sourceyieldtable.Rows[i][ModelID - 1].ToString() &&
                        resultyieldtable.Rows[j][1].ToString() == sourceyieldtable.Rows[i][BOMPN - 1].ToString() &&
                        resultyieldtable.Rows[j][2].ToString() == sourceyieldtable.Rows[i][BOMPNRev - 1].ToString() &&
                        resultyieldtable.Rows[j][3].ToString() == sourceyieldtable.Rows[i][DataSetName - 1].ToString())
                    {
                        if (sourceyieldtable.Rows[i][DataSetStatus - 1].ToString() != "FAIL")//dataset status is pass
                            resultyieldtable.Rows[j][4] = (float)resultyieldtable.Rows[j][4] + 1;

                        else //dataset status is fail
                            resultyieldtable.Rows[j][5] = (float)resultyieldtable.Rows[j][5] + 1;


                    }


                }

            }

            for (int j = 0; j <= resultyieldtable.Rows.Count - 1; j++)
            {
                resultyieldtable.Rows[j][6] = ((float)resultyieldtable.Rows[j][4] / ((float)resultyieldtable.Rows[j][4] + (float)resultyieldtable.Rows[j][5])).ToString("0.00");
                // resultyieldtable.Rows[j][6] = ((float)resultyieldtable.Rows[j][6] * 100).ToString() + "%";

            }

            returndataset.Tables.Add(resultyieldtable.Copy());
            return returndataset;

        }
        public int ContainerNumberCounter(DataTable QueryResultTable)
        {
            // count container number
            int ContainerNumberCount = 0;
            int SN = 1;


            for (int i = 0; i <= QueryResultTable.Rows.Count - 1; i++)
            {

                if (i <= QueryResultTable.Rows.Count - 2 && QueryResultTable.Rows[i][SN - 1].ToString() != QueryResultTable.Rows[i + 1][SN - 1].ToString())
                    ContainerNumberCount++;


            }
            if (QueryResultTable.Rows.Count > 0)
                ContainerNumberCount++;
            else
                ContainerNumberCount = 0;




            return ContainerNumberCount;
        }

        public string ReadListFromTxt()
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

                    // Read the stream to a string, and write the string to the console.
                    line = sr.ReadToEnd().Trim();
                    sr.Dispose();
                    //Console.WriteLine(line);




                }
                catch (Exception e)
                {
                    //Console.WriteLine("The file could not be read:");
                    //Console.WriteLine(e.Message);
                    MessageBox.Show("The file could not be read:\n" + e.Message);

                }




            }
            line = line.Trim();//去除首尾空格
            line = line.Replace("\t", "");//去除制表符
            line = line.Replace("\r\n", ",");//回车替换为逗号

            ///替换空格+逗号，这样保留字段内部的空格，同时删除逗号和字符之间的多余空格
            string templine;
            do
            {
                templine = line;
                line = line.Replace(" ,", ",");
            }

            while (line != templine);
            ///
            return line;



        }

        public DataGridView ChangeDataGridColor(DataGridView dg)
        {
            for (var i = 0; i < dg.Rows.Count; i++)
            {
                if (dg.Rows[i].Cells["DATASETSTATUS"].Value.ToString() == "PASS")
                {
                    dg.Rows[i].DefaultCellStyle.ForeColor = Color.Green;


                }

                if (dg.Rows[i].Cells["DATASETSTATUS"].Value.ToString() == "FAIL" || dg.Rows[i].Cells["DATASETSTATUS"].Value.ToString() == "ERROR")
                {
                    dg.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                }



            }

            return dg;
        }


        public DataSet CPKCalculator(string DataName, double[] rawdata, int ContainerNumber, double USL, double LSL)
        {
            DataSet CPkDataSet = new DataSet();
            DataTable CPkDataTable = new DataTable();

            Array.Sort(rawdata);


            var x05 = rawdata[((int)Math.Floor(rawdata.Count() * 0.5))];
            var x0995 = rawdata[(int)Math.Floor(rawdata.Count() * 0.995)];
            var x0005 = rawdata[(int)Math.Floor(rawdata.Count() * 0.005)];

            double xSum = rawdata.Sum();
            double xAvg = xSum / rawdata.Count();
            double xAvgd = rawdata.Average();
            double xStDevSum = 0;
            for (var i = 0; i < rawdata.Count(); i++)
            {
                xStDevSum += (rawdata[i] - xAvg) * (rawdata[i] - xAvg);
            }
            double xStDev = Math.Sqrt(xStDevSum / rawdata.Count());
            //double CPU = (USL - xAvg) / (3 * xStDev);
            //double CPL = (xAvg - LSL) / (3 * xStDev);
            // use non normal distribution to calculate CPk
            var CP = (USL - LSL) / (x0995 - x0005);
            var CPU = (USL - x05) / (x0995 - x05);
            var CPL = (x05 - LSL) / (x05 - x0005);
            double CPk = Math.Min(CPU, CPL);

            CPkDataTable.Columns.Add("DataName");
            CPkDataTable.Columns.Add("ContainerNumber");
            CPkDataTable.Columns.Add("DataCounts");
            CPkDataTable.Columns.Add("LSL");
            CPkDataTable.Columns.Add("USL");
            CPkDataTable.Columns.Add("Average");
            CPkDataTable.Columns.Add("StDev");
            CPkDataTable.Columns.Add("0.5% percentile");
            CPkDataTable.Columns.Add("50% percentile");
            CPkDataTable.Columns.Add("99.5% percentile");
            CPkDataTable.Columns.Add("CP");
            CPkDataTable.Columns.Add("CPU");
            CPkDataTable.Columns.Add("CPL");
            CPkDataTable.Columns.Add("CPK");

            CPkDataTable.Rows.Add(DataName, ContainerNumber, rawdata.Count(), LSL, USL, xAvg.ToString("0.0000"), xStDev.ToString("0.0000"), x0005, x05, x0995, CP.ToString("0.0000"), CPU.ToString("0.0000"), CPL.ToString("0.0000"), CPk.ToString("0.0000"));

            CPkDataSet.Tables.Add(CPkDataTable.Copy());

            return CPkDataSet;

        }

        public void CPKChartPlot(double[] rawdata, Chart CPKChart, double CPK, double LSL, double USL, int partition_number)
        {
            var xstart = Math.Min(rawdata.Min(), LSL);
            var xend = Math.Max(rawdata.Max(), USL);

            var step = (xend - xstart) / partition_number;

            double[] x = new double[partition_number]; ;

            for (var i = 0; i < partition_number; i++)
            {
                x[i] = xstart + (i + 0.5) * step;
            }


            int[] y = new int[partition_number];// can Initial all cells as 0??

            for (var i = 0; i < rawdata.Count(); i++)
            {
                for (var j = 0; j < partition_number; j++)
                {
                    if (rawdata[i] >= (xstart + j * step) && rawdata[i] < (xstart + (j + 1) * step))
                        y[j]++;
                }


            }
            CPKChart.Series.Clear();


            #region chart properies
            //CPKChart.Width = 990;
            //CPKChart.Height =160;
            CPKChart.BackColor = Color.White;// Color.FromArgb(211, 223, 240);
            //CPKHistgram.ID = chartName;
            //CPKChart.CssClass = "chartInfo";
            CPKChart.Palette = ChartColorPalette.BrightPastel;
            //newChart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;
            #endregion

            #region ChartArea
            ChartArea chartArea = CPKChart.ChartAreas[0];
            chartArea.BorderDashStyle = ChartDashStyle.Solid;
            chartArea.BackColor = Color.WhiteSmoke;// Color.FromArgb(0, 0, 0, 0);       
            chartArea.ShadowColor = Color.FromArgb(0, 0, 0, 0);
            chartArea.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dash;//设置网格为虚线
            chartArea.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            //newChart.ChartAreas.Add(chartArea);
            #endregion
            //CPKChart.Series["CPKHistgram"].ChartType = SeriesChartType.Line;


            CPKChart.Series.Add("CPKHistogram");
            //CPKChart.Series["CPKHistogram"].Label = "LSL=" + LSL.ToString() + "\nUSL=" + USL.ToString() + "\nCPK=" + CPK.ToString();
            CPKChart.Series["CPKHistogram"].LegendText = "CPK=" + CPK.ToString();

            for (int i = 0; i < x.Count(); i++)
            {
                CPKChart.Series["CPKHistogram"].Points.AddXY(x[i], y[i]);

            }

            CPKChart.Series.Add("LSL");
            CPKChart.Series["LSL"].ChartType = SeriesChartType.Line;
            CPKChart.Series["LSL"].BorderWidth = 3;
            CPKChart.Series["LSL"].BorderDashStyle = ChartDashStyle.DashDot;
            //CPKChart.Series["LSL"].Label = "LSL";
            CPKChart.Series["LSL"].Points.AddXY(LSL, 0);
            CPKChart.Series["LSL"].Points.AddXY(LSL, y.Max());
            CPKChart.Series["LSL"].LegendText = "LSL=" + LSL.ToString();


            CPKChart.Series.Add("USL");
            CPKChart.Series["USL"].ChartType = SeriesChartType.Line;
            CPKChart.Series["USL"].BorderWidth = 3;
            CPKChart.Series["USL"].BorderDashStyle = ChartDashStyle.DashDot;
            //CPKChart.Series["USL"].Label = "USL";
            CPKChart.Series["USL"].Points.AddXY(USL, 0);
            CPKChart.Series["USL"].Points.AddXY(USL, y.Max());
            CPKChart.Series["USL"].LegendText = "USL=" + USL.ToString();


            ////Add Box Plot
            //double[] yValues= {55.62, 45.54, 73.45, 9.73, 88.42, 45.9, 63.6, 85.1, 67.2, 23.6};

            //CPKChart.Series.Add("BoxPlot");
            //CPKChart.Series["BoxPlot"].ChartType = SeriesChartType.BoxPlot;
            ////CPKChart.Series["BoxPlot"]["BoxPlotSeries"] = "Price:Y2;Volume";
            //CPKChart.Series["BoxPlot"].Points.DataBindY(yValues);



        }


        public void ExportChart(Chart chart1)
        {
            //SystemInformation.UserInteractive = true;  
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Save Files";
            saveFileDialog.Filter = "bmp文件(*.bmp*)|*.bmp";
            //saveFileDialog.Filter = "JPG文件(*.jpg)|*.jpg|bmp文件(*.bmp*)|*.bmp";
            //默认文件类型显示顺序  
            saveFileDialog.FilterIndex = 2;
            //记忆上次保存路径  
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (!string.IsNullOrEmpty(saveFileDialog.FileName))
                {
                    string localFilePath = saveFileDialog.FileName.ToString();
                    chart1.SaveImage(localFilePath, ChartImageFormat.Bmp);

                }
            }

        }


        #region //This section is for ER CPK used only

        public DataSet ERCPK_OracleQuery(string connectionstring, string modelID, string starttime, string endtime)
        {
            sql = "SELECT a.MFR_SN AS SN,b.JOB_ID  AS JOB_ID,  c.MODEL_ID  AS Model_ID,c.BOM_PN  AS BOM_PN,c.BOM_PN_REV BOM_PN_Rev,d.DATASET_NAME AS DataSet_Name,d.STATUS AS DataSetStatus,d.END_TIME AS DataSet_EndTime,e.DATA_NAME AS Data_Name,e.DATA_VAL1  AS Data_Val,e.DATA_VAL2 AS Data_Status,e.DATA_VAL3 AS Data_Spec" +
                   " FROM parts a INNER JOIN routes b ON b.PART_INDEX = a.OPT_INDEX INNER JOIN bom_context_id c ON c.BOM_CONTEXT_ID = b.BOM_CONTEXT_ID INNER JOIN datasets d ON d.ROUTE_ID = b.ROUTE_ID INNER JOIN route_data e ON e.DATASET_ID = d.DATASET_ID WHERE '1' = '1'";

            if (modelID.Length != 0)
            {
                modelID = modelID.Replace(",", "','");
                sql += "and " + "c.MODEL_ID in ('" + modelID + "')";
                //sql += " and " + "c.MODEL_ID='" + ModelID + "'";
            }



            sql += " and d.STATUS='PASS' and d.DATASET_NAME = 'final_tx-rm' and (e.DATA_NAME like 'tx_eye_params_mc%:extinctionRatio_chan_%' or e.DATA_NAME like 'tx_eye_params_mc%:extinctionRatio[%]')";

            if (starttime.Length != 0 && endtime.Length != 0)
                sql += " and " + "d.END_TIME between '" + starttime + "' and '" + endtime + "'";

            sql += " ORDER BY Model_ID, SN, d.END_TIME";

            //OracleConnection OracleConn = oracledbo.OpenOracleConnection2(connectionstring);
            //InformLabel.Text = "Searching...";
            //ds = oracledbo.GetOracleDataSet2(sql, OracleConn);
            ds = GetOracleDataSet2(connectionstring, sql);
            //oracledbo.OracleConnectionClose();
            //OracleConn.Close();
            return ds;
        }

        public DataSet ERCPK_Data_Processor(DataTable rawdata)
        {
            int SN = 1;
            int JobID = 2;
            int ModelID = 3;
            int BOMPN = 4;
            int BOMPNRev = 5;
            int DataSetName = 6;
            int DataSetStatus = 7;
            int DataSetEndtime = 8;
            int DataName = 9;
            int DataValue = 10;
            int DataStatus = 11;
            int DataSpec = 12;

            int Flag = 13;

            DataSet CPkDataSet = new DataSet();
            DataTable CPkDataTable = new DataTable();


            CPkDataTable.Columns.Add("ModelID");
            CPkDataTable.Columns.Add("DataCounts");
            CPkDataTable.Columns.Add("LSL");
            CPkDataTable.Columns.Add("USL");
            CPkDataTable.Columns.Add("Average");
            CPkDataTable.Columns.Add("StDev");
            CPkDataTable.Columns.Add("0.5% percentile");
            CPkDataTable.Columns.Add("50% percentile");
            CPkDataTable.Columns.Add("99.5% percentile");
            CPkDataTable.Columns.Add("CP");
            CPkDataTable.Columns.Add("CPU");
            CPkDataTable.Columns.Add("CPL");
            CPkDataTable.Columns.Add("CPK");

            for (var i = 0; i < rawdata.Rows.Count; i++)
            {
                List<double> tempdata = new List<double>();

                //double[] tempdata =new double[] {} ;

                var tempModelID = rawdata.Rows[i][ModelID - 1].ToString();

                //var k = 0;
                tempdata.Add(Double.Parse(rawdata.Rows[i][DataValue - 1].ToString()));

                //string[] tempspec = rawdata.Rows[i][DataSpec - 1].ToString().Split(" <= ");
                string[] tempspec = Regex.Split(rawdata.Rows[i][DataSpec - 1].ToString(), " <= ");
                double LSL = Double.Parse(tempspec[0]);
                double USL = Double.Parse(tempspec[2]);


                for (var j = i + 1; j < rawdata.Rows.Count; j++)
                {

                    if (rawdata.Rows[j][ModelID - 1].ToString() == tempModelID)
                    {
                        //k++;
                        tempdata.Add(Double.Parse(rawdata.Rows[j][DataValue - 1].ToString()));
                    }
                    else
                    {
                        i = j - 1;
                        break;

                    }

                    if (j == rawdata.Rows.Count - 1)
                        i = j;
                }
                tempdata.Sort();
                // Array.Sort(tempdata);
                var x05 = tempdata[((int)Math.Floor(tempdata.Count() * 0.5))];
                var x0995 = tempdata[(int)Math.Floor(tempdata.Count() * 0.995)];
                var x0005 = tempdata[(int)Math.Floor(tempdata.Count() * 0.005)];

                double xSum = tempdata.Sum();
                double xAvg = xSum / tempdata.Count();
                //double xAvgd = tempdata.Average();
                double xStDevSum = 0;
                for (var m = 0; m < tempdata.Count(); m++)
                {
                    xStDevSum += (tempdata[m] - xAvg) * (tempdata[m] - xAvg);
                }
                double xStDev = Math.Sqrt(xStDevSum / tempdata.Count());
                //double CPU = (USL - xAvg) / (3 * xStDev);
                //double CPL = (xAvg - LSL) / (3 * xStDev);
                // use non normal distribution to calculate CPk
                var CP = (USL - LSL) / (x0995 - x0005);
                var CPU = (USL - x05) / (x0995 - x05);
                var CPL = (x05 - LSL) / (x05 - x0005);
                double CPk = Math.Min(CPU, CPL);
                CPkDataTable.Rows.Add(tempModelID, tempdata.Count(), LSL, USL, xAvg, xStDev, x0005, x05, x0995, CP, CPU, CPL, CPk);

            }
            CPkDataSet.Tables.Add(CPkDataTable.Copy());
            return CPkDataSet;



        }

        #endregion



        #region this region for Access and Oracle database operator
        public void OpenOracleConnection_ThenCLosed2(string ConnectionStr)
        {

            OracleConnection Oracleconn = new OracleConnection(ConnectionStr);
            //LogInfo.writeLog(String.Format("open ORA DB connection: {0}", ConnectionStr));



            try
            {
                // Thread.Sleep(3000);

                //string temp = Oracleconn.State.ToString();
                if (Oracleconn.State == ConnectionState.Closed)
                //if(temp=="Closed")
                {

                    //Application.DoEvents();
                    // previus I often got curruption of the heap error and dont know why, maybe it is because Oracle_OCI has not been loaded in time? 2016.7.13
                    //MessageBox.Show(Oracleconn.State.ToString());

                    Oracleconn.Open();
                    Oracleconn.Close();


                }
                else if (Oracleconn.State == ConnectionState.Broken)
                {
                    //LogInfo.writeLog(String.Format("open ORA DB connection fail: {0}", ConnectionStr));
                    Oracleconn.Close();
                    Oracleconn.Open();
                    Oracleconn.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                throw new Exception(e.Message);
            }

            //return Oracleconn;

        }

        public DataSet GetOracleDataSet2(string ConnectionStr, string SQLString)
        {



            OracleConnection Oracleconn = new OracleConnection(ConnectionStr);


            try
            {

                if (Oracleconn.State == ConnectionState.Closed)
                {

                    Oracleconn.Open();

                }
                else if (Oracleconn.State == ConnectionState.Broken)
                {
                    Oracleconn.Close();
                    Oracleconn.Open();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                throw new Exception(e.Message);
            }


            try
            {
                DataSet ds = new DataSet();

                OracleDataAdapter sda = new OracleDataAdapter(SQLString, Oracleconn);
                //Form_InformBox form = new Form_InformBox();

                //form.Enabled = true;





                sda.Fill(ds, "ds");
                //form.Enabled = false;
                //form.Dispose();
                //form.Dispose_Form_and_Thread();




                Oracleconn.Close();
                Oracleconn.Dispose();
                return ds;
            }
            catch (OleDbException ex)
            {
                //DataSet ds = new DataSet();

                Oracleconn.Close();
                Oracleconn.Dispose();
                return ds;
                throw new Exception(ex.Message);


            }




        }
        //private delegate DataSet MyDelegate(string ConnectionStr, string SQLString);// this is for speed up the SQL query by multi threading




        public DataSet GetAccessDataset2(string dbPath, string SQLString)
        {

            string connectionString;
            OleDbConnection AccessConnection;
            //OleDbCommand cmd;
            //this.dbPath = dbPath;
            connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbPath + ";User Id=admin;Password=";
            AccessConnection = new OleDbConnection(connectionString);
            try
            {
                AccessConnection.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            // return AccessConnection;

            DataSet ds = new DataSet();
            try
            {
                OleDbDataAdapter sda = new OleDbDataAdapter(SQLString, AccessConnection);
                sda.Fill(ds, "ds");
                AccessConnection.Close();
                AccessConnection.Dispose();

                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            AccessConnection.Close();
            AccessConnection.Dispose();
            return ds;

        }


        public void AccessConnectionClose(OleDbConnection AccessConnection)
        {
            AccessConnection.Close();
            AccessConnection.Dispose();


        }


        #endregion



        public DataSet WhereUsedOracleQuery(string connectionstring, bool AlreadyRemoved, string SN)
        {
            DataSet ds;
            string sql;
            string sql_AlreadyRemovedYes = "select a.MFR_SN as Module_SN,a.MFR_PN as Module_PN, f.MODEL_ID as Model_ID, y.MFR_SN as Component_SN,y.MFR_PN as Component_PN,Z.ROUTE_ID,Z.Operation,Z.Notes, Z.TIME as Operation_Time " +
                "from PARTS a inner join ROUTES b on A.Opt_Index=B.Part_Index inner join BOM_CONTEXT_ID f on b.BOM_CONTEXT_ID=f.BOM_CONTEXT_ID inner join Assembly_History z on Z.Route_Id=B.Route_Id inner join PARTS y on y.Opt_Index=Z.Child_Index where z.OPERATION='rm' and y.MFR_SN ";
            string sql_AlradyRemovedNo = "select a.MFR_SN as Module_SN,a.MFR_PN as Module_PN, f.MODEL_ID as Model_ID, y.MFR_SN as Component_SN,y.MFR_PN as Component_PN, b.ROUTE_ID, x.NOTES as NOTES " +
                "from PARTS a inner join ROUTES b on A.Opt_Index=B.Part_Index inner join BOM_CONTEXT_ID f on b.BOM_CONTEXT_ID=f.BOM_CONTEXT_ID inner join " +
                "ASSEMBLY x on a.OPT_INDEX=x.PARENT_INDEX  inner join PARTS y on y.OPT_INDEX=x.CHILD_INDEX where y.MFR_SN ";
            if (SN.Length != 0)
            {

                SN = SN.Replace(",", "','");

                if (AlreadyRemoved)
                {
                    sql = sql_AlreadyRemovedYes + " in ('" + SN + "')";
                }
                else
                {
                    sql = sql_AlradyRemovedNo + " in ('" + SN + "')";
                }

                ds = GetOracleDataSet2(connectionstring, sql);
                return ds;
            }
            else
            {
                MessageBox.Show("Please input SN");
                return null;
            }



        }

        public string ListBox2SQL_in_Query_String(ListBox listbox)
        {
            string temp = "";
            if (listbox.SelectedItems.Count != 0)
            {
                for (var i = 0; i < listbox.SelectedItems.Count - 1; i++)
                    temp += listbox.SelectedItems[i] + ",";

                temp += listbox.SelectedItems[listbox.SelectedItems.Count - 1];
                //temp = temp.Replace(",", "','");


            }
            return temp;
        }

        public DataSet DataSet_TestTime_OracleQuery(string connectionstring, string SerialNumber, string JobOrder,
            string BOMPN, string BOMPNRev, string ModelID, string DataSet, string DataSetStatus, string StartTime, string EndTime, Label InformLabel)
        {
            sql = "SELECT a.MFR_SN AS SN,b.JOB_ID  AS JOB_ID,  c.MODEL_ID  AS Model_ID,c.BOM_PN  AS BOM_PN,c.BOM_PN_REV BOM_PN_Rev,d.DATASET_NAME AS DataSet_Name,d.STATUS AS DataSetStatus,d.START_TIME AS DataSet_StartTime,d.END_TIME AS DataSet_EndTime, D.Dataset_Id AS DataSet_ID, e.DATA_NAME, e.DATA_VAL1" +
                   " FROM parts a INNER JOIN routes b ON b.PART_INDEX = a.OPT_INDEX INNER JOIN bom_context_id c ON c.BOM_CONTEXT_ID = b.BOM_CONTEXT_ID INNER JOIN datasets d ON d.ROUTE_ID = b.ROUTE_ID  INNER JOIN ROUTE_DATA e ON e.DATASET_ID = d.DATASET_ID" +
                   " WHERE '1' = '1' AND (E.Data_Name like '%:TestSeconds' OR E.Data_Name = 'elapsed_seconds')";
            if (SerialNumber.Length != 0)
            {
                SerialNumber = SerialNumber.Replace(",", "','");
                sql += " and " + "a.MFR_SN in ('" + SerialNumber + "')";
            }
            if (JobOrder.Length != 0)
            {
                JobOrder = JobOrder.Replace(",", "','");
                sql += "and " + "b.JOB_ID in ('" + JobOrder + "')";
            }
            if (BOMPN.Length != 0)
            {
                BOMPN = BOMPN.Replace(",", "','");
                sql += "and " + "c.BOM_PN in ('" + BOMPN + "')";
                //sql += " and " + "c.BOM_PN='" + BOMPN + "'";
            }

            if (BOMPNRev.Length != 0)
                sql += " and " + "c.BOM_PN_REV='" + BOMPNRev + "'";

            if (ModelID.Length != 0)
            {
                ModelID = ModelID.Replace(",", "','");
                sql += "and " + "c.MODEL_ID in ('" + ModelID + "')";
                //sql += " and " + "c.MODEL_ID='" + ModelID + "'";
            }

            if (DataSet.Length != 0)
            {

                sql += " and " + "d.DATASET_NAME in ('" + DataSet.Replace(",", "','") + "')";
            }



            if (DataSetStatus.Length != 0)
                sql += " and " + "d.STATUS='" + DataSetStatus + "'";



            if (StartTime.Length != 0 && EndTime.Length != 0)
                sql += " and " + "d.END_TIME between '" + StartTime + "' and '" + EndTime + "'";

            //----------------------------------------------
            //sql += " and D.Ate_Version like 'ATE_Code-Wuxi-1.00-%_DL'";
            //----------------------------------------------

            sql += "ORDER BY  SN, d.DATASET_NAME, d.DATASET_ID";

            //OracleConnection OracleConn = oracledbo.OpenOracleConnection2(connectionstring);
            //InformLabel.Text = "Searching...";
            //ds = oracledbo.GetOracleDataSet2(sql, OracleConn);
            ds = GetOracleDataSet2(connectionstring, sql);
            //oracledbo.OracleConnectionClose();
            //OracleConn.Close();
            return ds;

        }

        public DataSet DataSet_TestTime_Calculator(DataTable rawdata)
        {
            int SN = 1;
            int JobID = 2;
            int ModelID = 3;
            int BOMPN = 4;
            int BOMPNRev = 5;
            int DataSetName = 6;
            int DataSetStatus = 7;
            int DataSetStartTIme = 8;
            int DataSetEndtime = 9;
            int DataSetID = 10;
            int DataName = 11;
            int Data_Val = 12;
            int TestTime = 13;


            DataSet TestTimeDataSet = new DataSet();
            DataTable TestTimeDataTable = new DataTable();


            TestTimeDataTable.Columns.Add("ModelID");
            TestTimeDataTable.Columns.Add("DataSet");
            //TestTimeDataTable.Columns.Add("ContainerCounts");
            TestTimeDataTable.Columns.Add("DataCounts");

            TestTimeDataTable.Columns.Add("Average");
            TestTimeDataTable.Columns.Add("StDev");
            TestTimeDataTable.Columns.Add("Min");
            TestTimeDataTable.Columns.Add("50% percentile");
            TestTimeDataTable.Columns.Add("Max");

            rawdata.Columns.Add("Test_Time");





            //for (var k = 0; k < rawdata.Rows.Count; k++)
            //{

            //    if (rawdata.Rows[k][TestTime - 1].ToString() == "0.0")//If test time is zero, then remove it, otherwise it will affect the test time static result
            //        rawdata.Rows[k].Delete();


            //}

            //rawdata.AcceptChanges();



            #region get the ModelID and DataSetName number,DataSetID, this method can be used in other place
            List<string> tempModelID = new List<string>();
            List<string> tempDataSet = new List<string>();
            List<string> tempDataSetID = new List<string>();
            tempModelID.Add(rawdata.Rows[0][ModelID - 1].ToString());
            tempDataSet.Add(rawdata.Rows[0][DataSetName - 1].ToString());
            tempDataSetID.Add(rawdata.Rows[0][DataSetID - 1].ToString());

            for (var i = 1; i < rawdata.Rows.Count; i++)
            {
                for (var j = 0; j < tempModelID.Count; j++)
                {
                    if (tempModelID[j] == rawdata.Rows[i][ModelID - 1].ToString())
                        break; // if contain, then break
                    else if (j == tempModelID.Count - 1)// if last cell still don't contain, then add it
                        tempModelID.Add(rawdata.Rows[i][ModelID - 1].ToString());
                }

            }


            for (var i = 1; i < rawdata.Rows.Count; i++)
            {
                for (var j = 0; j < tempDataSet.Count; j++)
                {
                    if (tempDataSet[j] == rawdata.Rows[i][DataSetName - 1].ToString())
                        break; // if contain, then break
                    else if (j == tempDataSet.Count - 1)// if last cell still don't contain, then add it
                        tempDataSet.Add(rawdata.Rows[i][DataSetName - 1].ToString());
                }

            }

            for (var i = 1; i < rawdata.Rows.Count; i++)
            {
                for (var j = 0; j < tempDataSetID.Count; j++)
                {
                    if (tempDataSetID[j] == rawdata.Rows[i][DataSetID - 1].ToString())
                        break; // if contain, then break
                    else if (j == tempDataSetID.Count - 1)// if last cell still don't contain, then add it
                        tempDataSetID.Add(rawdata.Rows[i][DataSetID - 1].ToString());
                }

            }
            #endregion


            for (var i = 0; i < tempModelID.Count; i++)
            {
                for (var j = 0; j < tempDataSet.Count; j++)
                {
                    for (var l = 0; l < tempDataSetID.Count; l++)
                    {
                        double temptime = 0;
                        for (var k = 0; k < rawdata.Rows.Count; k++)
                        {

                            if (rawdata.Rows[k][ModelID - 1].ToString() == tempModelID[i]
                                && rawdata.Rows[k][DataSetName - 1].ToString() == tempDataSet[j]
                                && rawdata.Rows[k][DataSetID - 1].ToString() == tempDataSetID[l])
                                temptime += Double.Parse(rawdata.Rows[k][Data_Val - 1].ToString()) / 60;// Get the sum of each testcase testtime (has the samme datasetID), and convert it into Minute

                        }
                        for (var k = 0; k < rawdata.Rows.Count; k++)
                        {
                            if (rawdata.Rows[k][ModelID - 1].ToString() == tempModelID[i]
                               && rawdata.Rows[k][DataSetName - 1].ToString() == tempDataSet[j]
                               && rawdata.Rows[k][DataSetID - 1].ToString() == tempDataSetID[l])
                                rawdata.Rows[k][TestTime - 1] = temptime;// Add sum of each testcase testtime into a new column
                        }

                    }
                }
            }

            List<double> tempdata = new List<double>();
            for (var i = 0; i < tempModelID.Count; i++)
            {
                for (var j = 0; j < tempDataSet.Count; j++)
                {
                    for (var l = 0; l < tempDataSetID.Count; l++)
                    {
                        for (var k = 0; k < rawdata.Rows.Count; k++)
                        {

                            if (rawdata.Rows[k][ModelID - 1].ToString() == tempModelID[i]
                                && rawdata.Rows[k][DataSetName - 1].ToString() == tempDataSet[j]
                                 && rawdata.Rows[k][DataSetID - 1].ToString() == tempDataSetID[l])
                            {
                                tempdata.Add(Double.Parse(rawdata.Rows[k][TestTime - 1].ToString()));
                                break;//Just need to add the test time one time
                            }
                        }
                    }

                    tempdata.Sort();
                    var x05 = tempdata[((int)Math.Floor(tempdata.Count() * 0.5))];
                    //var x0995 = tempdata[(int)Math.Floor(tempdata.Count() * 0.995)];
                    //var x0005 = tempdata[(int)Math.Floor(tempdata.Count() * 0.005)];

                    double xSum = tempdata.Sum();
                    double xAvg = xSum / tempdata.Count();
                    ////double xAvgd = tempdata.Average();
                    double xStDevSum = 0;
                    for (var m = 0; m < tempdata.Count(); m++)
                    {
                        xStDevSum += (tempdata[m] - xAvg) * (tempdata[m] - xAvg);
                    }
                    double xStDev = Math.Sqrt(xStDevSum / tempdata.Count());
                    //double CPU = (USL - xAvg) / (3 * xStDev);
                    //double CPL = (xAvg - LSL) / (3 * xStDev);
                    // use non normal distribution to calculate CPk
                    //var CP = (USL - LSL) / (x0995 - x0005);
                    //var CPU = (USL - x05) / (x0995 - x05);
                    //var CPL = (x05 - LSL) / (x05 - x0005);
                    //double CPk = Math.Min(CPU, CPL);
                    TestTimeDataTable.Rows.Add(tempModelID[i], tempDataSet[j], tempdata.Count(), xAvg.ToString("0.00"), xStDev.ToString("0.00"), tempdata.Min().ToString("0.00"), x05.ToString("0.00"), tempdata.Max().ToString("0.00"));
                    tempdata.Clear();
                }

            }

            TestTimeDataSet.Tables.Add(TestTimeDataTable.Copy());
            return TestTimeDataSet;

        }

        private double EndTime_StartTime_Calculator(string starttime, string endtime)
        {


            DateTime StartDate = DateTime.ParseExact(starttime, "yyyyMMddHHmmss", System.Globalization.CultureInfo.CurrentCulture);
            DateTime EndDate = DateTime.ParseExact(endtime, "yyyyMMddHHmmss", System.Globalization.CultureInfo.CurrentCulture);

            TimeSpan TestTime = EndDate - StartDate;

            return TestTime.TotalMinutes;

        }


        public DataSet Yield_Plot_Table(DataTable rawdata)
        {
            int SN = 1;
            int JobID = 2;
            int ModelID = 3;
            int BOMPN = 4;
            int BOMPNRev = 5;
            int DataSetName = 6;
            int DataSetStatus = 7;
            int DataSetEndtime = 8;
            int Station = 9;
            int Computer = 10;
            int RouteID = 11;

            int DataName = 16;
            int DataValue = 17;
            int DataStatus = 18;
            int DataSpec = 19;

            //int Flag = 14;



            int WeekNumber = rawdata.Columns.Count;
            rawdata.Columns[rawdata.Columns.Count - 1].ColumnName = "Week_Number";

            DataTable Yield_Plot_Table = new DataTable(); ;
            DataSet Yied_Plot_DataSet = new DataSet();

            Yield_Plot_Table.Columns.Add("ModelID");
            Yield_Plot_Table.Columns.Add("DataSet");
            Yield_Plot_Table.Columns.Add("Pass Quantity");
            Yield_Plot_Table.Columns.Add("Fail Quantity");
            Yield_Plot_Table.Columns.Add("Ratio");
            Yield_Plot_Table.Columns.Add("Week Number");
            Yield_Plot_Table.Columns.Add("Start Time");
            Yield_Plot_Table.Columns.Add("End Time");
            Yield_Plot_Table.Columns.Add("First Failure Mode");
            Yield_Plot_Table.Columns.Add("First Failure Mode Quantity");
            Yield_Plot_Table.Columns.Add("First Failure Mode Ratio");
            Yield_Plot_Table.Columns.Add("Second Failure Mode");
            Yield_Plot_Table.Columns.Add("Second Failure Mode Quantity");
            Yield_Plot_Table.Columns.Add("Second Failure Mode Ratio");
            Yield_Plot_Table.Columns.Add("Third Failure Mode");
            Yield_Plot_Table.Columns.Add("Third Failure Mode Quantity");
            Yield_Plot_Table.Columns.Add("Third Failure Mode Ratio");



            for (var i = 0; i < rawdata.Rows.Count; i++)
                rawdata.Rows[i][WeekNumber - 1] = GetWeek_Number_From_DateTime(rawdata.Rows[i][DataSetEndtime - 1].ToString());

            #region get the ModelID , DataSetName number, tempWeekNumber, tempFailureMode, this method can be used in other place
            List<string> tempModelID = new List<string>();
            List<string> tempDataSet = new List<string>();
            List<string> tempWeekNumber = new List<string>();
            List<string> tempFailureMode = new List<string>();
            //List<List<int>> tempFailureMode=new List<List<int>>;

            tempModelID.Add(rawdata.Rows[0][ModelID - 1].ToString());
            tempDataSet.Add(rawdata.Rows[0][DataSetName - 1].ToString());
            tempWeekNumber.Add(rawdata.Rows[0][WeekNumber - 1].ToString());

            for (var i = 1; i < rawdata.Rows.Count; i++)
            {
                for (var j = 0; j < tempModelID.Count; j++)
                {
                    if (tempModelID[j] == rawdata.Rows[i][ModelID - 1].ToString())
                        break; // if contain, then break
                    else if (j == tempModelID.Count - 1)// if last cell still don't contain, then add it
                        tempModelID.Add(rawdata.Rows[i][ModelID - 1].ToString());
                }

            }


            for (var i = 1; i < rawdata.Rows.Count; i++)
            {
                for (var j = 0; j < tempDataSet.Count; j++)
                {
                    if (tempDataSet[j] == rawdata.Rows[i][DataSetName - 1].ToString())
                        break; // if contain, then break
                    else if (j == tempDataSet.Count - 1)// if last cell still don't contain, then add it
                        tempDataSet.Add(rawdata.Rows[i][DataSetName - 1].ToString());
                }

            }

            for (var i = 1; i < rawdata.Rows.Count; i++)
            {
                for (var j = 0; j < tempWeekNumber.Count; j++)
                {
                    if (tempWeekNumber[j] == rawdata.Rows[i][WeekNumber - 1].ToString())
                        break; // if contain, then break
                    else if (j == tempWeekNumber.Count - 1)// if last cell still don't contain, then add it
                        tempWeekNumber.Add(rawdata.Rows[i][WeekNumber - 1].ToString());
                }

            }

            for (var i = 1; i < rawdata.Rows.Count; i++)
            {
                if (rawdata.Rows[i][DataSetStatus - 1].ToString() == "FAIL")
                {
                    string combineStr = "";
                    if (rawdata.Rows[i][DataName - 1].ToString() == "failed")
                    {
                        string[] tempArray = rawdata.Rows[i][DataValue - 1].ToString().Split(' ', ';');

                        for (var xx = 0; xx < tempArray.Length; xx = xx + 2)
                            combineStr = combineStr + " " + tempArray[xx];
                        // if dataname is failed, usually it appear like final_tx-rm 18554477 final_rx-rm 18554547 final_ber_p1440-rm 18554548 final_ber_btob-rm 18554549 
                        // The number is refer to dataset ID which have affection on failure mode statistics, so only dataset name is pull out into failure mode.
                        // if dataname is failed, in coherent 100G ACO product, usually it appear like mca_test_mca-nn 23836320;mca_test_temp_mca-cn 23837462
                        // The number together with ";" is not need to pull out into report table
                    }
                    else
                        combineStr = rawdata.Rows[i][DataValue - 1].ToString();



                    string tempstring = rawdata.Rows[i][DataName - 1].ToString() + " \t " + combineStr +
                            " \t " + rawdata.Rows[i][DataStatus - 1].ToString() + " \t " + rawdata.Rows[i][DataSpec - 1].ToString();
                    if (tempFailureMode.Count == 0)
                        tempFailureMode.Add(tempstring);
                    else
                    {
                        for (var j = 0; j < tempFailureMode.Count; j++)
                        {
                            if (tempFailureMode[j] == tempstring)
                                break; // if contain, then break
                            else if (j == tempFailureMode.Count - 1)// if last cell still don't contain, then add it
                                tempFailureMode.Add(tempstring);
                        }
                    }
                }

            }

            #endregion


            //List<int> tempPassNumber = new List<int>();
            //List<int> tempFailNumber = new List<int>();
            int tempPassNumber = 0;
            int tempFailNumber = 0;
            float ratio = 0;
            List<string> tempTime = new List<string>();
            //List<List<int>> FailureMode_Count = new List<List<int>>();

            string[,] FailureMode_Count = new string[tempFailureMode.Count, 2];
            //pass the tempFailureMode to FailureMode_Count, FailureMode_Count has 2 columun, fist is the failure mode, the second is for the count.
            for (var i = 0; i < tempFailureMode.Count; i++)
            {
                FailureMode_Count[i, 0] = tempFailureMode[i];
                FailureMode_Count[i, 1] = "0";
            }


            for (var i = 0; i < tempModelID.Count; i++)
            {
                for (var j = 0; j < tempDataSet.Count; j++)
                {
                    for (var h = 0; h < tempWeekNumber.Count; h++)
                    {
                        tempPassNumber = 0;
                        tempFailNumber = 0;

                        for (var k = 0; k < rawdata.Rows.Count; k++)
                        {

                            if (rawdata.Rows[k][ModelID - 1].ToString() == tempModelID[i] &&
                                rawdata.Rows[k][DataSetName - 1].ToString() == tempDataSet[j] &&
                                rawdata.Rows[k][WeekNumber - 1].ToString() == tempWeekNumber[h])
                            {
                                if (rawdata.Rows[k][DataSetStatus - 1].ToString() != "FAIL")
                                {
                                    tempPassNumber++;
                                    tempTime.Add(rawdata.Rows[k][DataSetEndtime - 1].ToString());
                                }

                                else
                                {
                                    string combineStr = "";
                                    if (rawdata.Rows[k][DataName - 1].ToString() == "failed")
                                    {
                                        string[] tempArray = rawdata.Rows[k][DataValue - 1].ToString().Split(' ', ';');

                                        for (var xx = 0; xx < tempArray.Length; xx = xx + 2)
                                            combineStr = combineStr + " " + tempArray[xx];
                                        // if dataname is failed, in tunable 10G product, usually it appear like final_tx-rm 18554477 final_rx-rm 18554547 final_ber_p1440-rm 18554548 final_ber_btob-rm 18554549 
                                        // The number is refer to dataset ID which have affection on failure mode statistics, so only dataset name is pull out into failure mode.
                                        // if dataname is failed, in coherent 100G ACO product, usually it appear like mca_test_mca-nn 23836320;mca_test_temp_mca-cn 23837462
                                        // The number together with ";" is not need to pull out into report table

                                    }
                                    else
                                        combineStr = rawdata.Rows[k][DataValue - 1].ToString();




                                    tempFailNumber++;
                                    tempTime.Add(rawdata.Rows[k][DataSetEndtime - 1].ToString());

                                    string tempstring = rawdata.Rows[k][DataName - 1].ToString() + " \t " + combineStr +
                            " \t " + rawdata.Rows[k][DataStatus - 1].ToString() + " \t " + rawdata.Rows[k][DataSpec - 1].ToString();
                                    for (var x = 0; x < tempFailureMode.Count; x++)
                                    {
                                        if (tempstring == FailureMode_Count[x, 0])
                                        {

                                            FailureMode_Count[x, 1] = (Int32.Parse(FailureMode_Count[x, 1]) + 1).ToString();
                                            //FailureMode_Count[x][1] count +1
                                        }

                                    }

                                }
                            }


                        }


                        for (var y = 0; y < tempFailureMode.Count - 1; y++)// sort based on failure mode qty
                        {
                            for (var z = 0; z < tempFailureMode.Count - y - 1; z++)
                            {
                                if (Int32.Parse(FailureMode_Count[z, 1]) < Int32.Parse(FailureMode_Count[z + 1, 1]))
                                {
                                    var tempdataset = FailureMode_Count[z, 0];
                                    var tempcnt = FailureMode_Count[z, 1];
                                    FailureMode_Count[z, 0] = FailureMode_Count[z + 1, 0];
                                    FailureMode_Count[z, 1] = FailureMode_Count[z + 1, 1];
                                    FailureMode_Count[z + 1, 0] = tempdataset;
                                    FailureMode_Count[z + 1, 1] = tempcnt;

                                }

                            }

                        }



                        ratio = (float)tempPassNumber / (tempPassNumber + tempFailNumber);

                        string firstfailmode;
                        string firstfailmodeQty;
                        double firstfailmodeRatio;
                        string secondfailmode;
                        string secondfailmodeQty;
                        double secondfailmodeRatio;
                        string thirdfailmode;
                        string thirdfailmodeQty;
                        double thirdfailmodeRatio;

                        if (tempFailNumber != 0)
                        {
                            if (FailureMode_Count.Length / 2 > 2)// failure mode number more than 3 
                            {
                                if (FailureMode_Count[2, 1] != "0")
                                {
                                    firstfailmode = FailureMode_Count[0, 0];
                                    firstfailmodeQty = FailureMode_Count[0, 1];
                                    secondfailmode = FailureMode_Count[1, 0];
                                    secondfailmodeQty = FailureMode_Count[1, 1];
                                    thirdfailmode = FailureMode_Count[2, 0];
                                    thirdfailmodeQty = FailureMode_Count[2, 1];

                                    firstfailmodeRatio = Double.Parse(firstfailmodeQty.ToString()) / tempFailNumber;
                                    secondfailmodeRatio = Double.Parse(secondfailmodeQty.ToString()) / tempFailNumber;
                                    thirdfailmodeRatio = Double.Parse(thirdfailmodeQty.ToString()) / tempFailNumber;

                                    Yield_Plot_Table.Rows.Add(tempModelID[i], tempDataSet[j], tempPassNumber, tempFailNumber, ratio.ToString("0.00"), tempWeekNumber[h], tempTime.Min(), tempTime.Max(),
                                        firstfailmode, firstfailmodeQty, firstfailmodeRatio.ToString("0.00"),
                                        secondfailmode, secondfailmodeQty, secondfailmodeRatio.ToString("0.00"),
                                        thirdfailmode, thirdfailmodeQty, thirdfailmodeRatio.ToString("0.00"));
                                }
                                else if (FailureMode_Count[1, 1] != "0")
                                {
                                    firstfailmode = FailureMode_Count[0, 0];
                                    firstfailmodeQty = FailureMode_Count[0, 1];
                                    secondfailmode = FailureMode_Count[1, 0];
                                    secondfailmodeQty = FailureMode_Count[1, 1];


                                    firstfailmodeRatio = Double.Parse(firstfailmodeQty.ToString()) / tempFailNumber;
                                    secondfailmodeRatio = Double.Parse(secondfailmodeQty.ToString()) / tempFailNumber;


                                    Yield_Plot_Table.Rows.Add(tempModelID[i], tempDataSet[j], tempPassNumber, tempFailNumber, ratio.ToString("0.00"), tempWeekNumber[h], tempTime.Min(), tempTime.Max(),
                                        firstfailmode, firstfailmodeQty, firstfailmodeRatio.ToString("0.00"),
                                        secondfailmode, secondfailmodeQty, secondfailmodeRatio.ToString("0.00"));
                                }
                                else if (FailureMode_Count[0, 1] != "0")
                                {
                                    firstfailmode = FailureMode_Count[0, 0];
                                    firstfailmodeQty = FailureMode_Count[0, 1];



                                    firstfailmodeRatio = Double.Parse(firstfailmodeQty.ToString()) / tempFailNumber;



                                    Yield_Plot_Table.Rows.Add(tempModelID[i], tempDataSet[j], tempPassNumber, tempFailNumber, ratio.ToString("0.00"), tempWeekNumber[h], tempTime.Min(), tempTime.Max(),
                                        firstfailmode, firstfailmodeQty, firstfailmodeRatio.ToString("0.00"));
                                }


                            }

                            else if (FailureMode_Count.Length / 2 == 2)// failure mode number is 2
                            {
                                if (FailureMode_Count[1, 1] != "0")
                                {

                                    firstfailmode = FailureMode_Count[0, 0];
                                    firstfailmodeQty = FailureMode_Count[0, 1];
                                    secondfailmode = FailureMode_Count[1, 0];
                                    secondfailmodeQty = FailureMode_Count[1, 1];


                                    firstfailmodeRatio = Double.Parse(firstfailmodeQty.ToString()) / tempFailNumber;
                                    secondfailmodeRatio = Double.Parse(secondfailmodeQty.ToString()) / tempFailNumber;


                                    Yield_Plot_Table.Rows.Add(tempModelID[i], tempDataSet[j], tempPassNumber, tempFailNumber, ratio.ToString("0.00"), tempWeekNumber[h], tempTime.Min(), tempTime.Max(),
                                        firstfailmode, firstfailmodeQty, firstfailmodeRatio.ToString("0.00"),
                                        secondfailmode, secondfailmodeQty, secondfailmodeRatio.ToString("0.00"));
                                }
                                else if (FailureMode_Count[0, 1] != "0")
                                {
                                    firstfailmode = FailureMode_Count[0, 0];
                                    firstfailmodeQty = FailureMode_Count[0, 1];



                                    firstfailmodeRatio = Double.Parse(firstfailmodeQty.ToString()) / tempFailNumber;



                                    Yield_Plot_Table.Rows.Add(tempModelID[i], tempDataSet[j], tempPassNumber, tempFailNumber, ratio.ToString("0.00"), tempWeekNumber[h], tempTime.Min(), tempTime.Max(),
                                        firstfailmode, firstfailmodeQty, firstfailmodeRatio.ToString("0.00"));
                                }

                            }

                            else if (FailureMode_Count.Length / 2 == 1)// just have 1 failure mode
                            {
                                firstfailmode = FailureMode_Count[0, 0];
                                firstfailmodeQty = FailureMode_Count[0, 1];



                                firstfailmodeRatio = Double.Parse(firstfailmodeQty.ToString()) / tempFailNumber;



                                Yield_Plot_Table.Rows.Add(tempModelID[i], tempDataSet[j], tempPassNumber, tempFailNumber, ratio.ToString("0.00"), tempWeekNumber[h], tempTime.Min(), tempTime.Max(),
                                    firstfailmode, firstfailmodeQty, firstfailmodeRatio.ToString("0.00"));
                            }


                        }
                        else
                        {

                            // for the pass, leave failuremode as blank
                            Yield_Plot_Table.Rows.Add(tempModelID[i], tempDataSet[j], tempPassNumber, tempFailNumber, ratio.ToString("0.00"), tempWeekNumber[h], tempTime.Min(), tempTime.Max());

                        }







                        tempTime.Clear();
                        for (var y = 0; y < tempFailureMode.Count - 1; y++)
                        {
                            FailureMode_Count[y, 1] = "0";
                        }


                    }

                }
            }



            Yied_Plot_DataSet.Tables.Add(Yield_Plot_Table.Copy());
            return Yied_Plot_DataSet;






        }

        private string GetWeek_Number_From_DateTime(string DataSetTime)
        {
            DateTime DateTime = DateTime.ParseExact(DataSetTime, "yyyyMMddHHmmss", System.Globalization.CultureInfo.CurrentCulture);
            //int i=DateTime

            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            //DateTime date1 = new DateTime(2011, 1, 1);
            Calendar cal = dfi.Calendar;
            string i = DateTime.Year.ToString().Substring(2, 2) + cal.GetWeekOfYear(DateTime, dfi.CalendarWeekRule, dfi.FirstDayOfWeek).ToString("00");
            // cal.ToString().Substring(cal.ToString().LastIndexOf(".") + 1));       
            return i;

        }

        public void Yield_Chart_Plot(DataTable rawdata, Chart Yield_Chart, string YieldType)
        {


            int ModelID = 1;
            int DataSetName = 2;
            int PassQTY = 3;
            int FailQTY = 4;
            int Ratio = 5;
            int WeekNumber = 6;
            int StartTime = 7;
            int EndTime = 8;


            #region get the ModelID , DataSetName number, tempWeekNumber,this method can be used in other place
            List<string> tempModelID = new List<string>();
            List<string> tempDataSet = new List<string>();
            List<string> tempWeekNumber = new List<string>();
            tempModelID.Add(rawdata.Rows[0][ModelID - 1].ToString());
            tempDataSet.Add(rawdata.Rows[0][DataSetName - 1].ToString());
            tempWeekNumber.Add(rawdata.Rows[0][WeekNumber - 1].ToString());

            for (var i = 1; i < rawdata.Rows.Count; i++)
            {
                for (var j = 0; j < tempModelID.Count; j++)
                {
                    if (tempModelID[j] == rawdata.Rows[i][ModelID - 1].ToString())
                        break; // if contain, then break
                    else if (j == tempModelID.Count - 1)// if last cell still don't contain, then add it
                        tempModelID.Add(rawdata.Rows[i][ModelID - 1].ToString());
                }

            }


            for (var i = 1; i < rawdata.Rows.Count; i++)
            {
                for (var j = 0; j < tempDataSet.Count; j++)
                {
                    if (tempDataSet[j] == rawdata.Rows[i][DataSetName - 1].ToString())
                        break; // if contain, then break
                    else if (j == tempDataSet.Count - 1)// if last cell still don't contain, then add it
                        tempDataSet.Add(rawdata.Rows[i][DataSetName - 1].ToString());
                }

            }

            for (var i = 1; i < rawdata.Rows.Count; i++)
            {
                for (var j = 0; j < tempWeekNumber.Count; j++)
                {
                    if (tempWeekNumber[j] == rawdata.Rows[i][WeekNumber - 1].ToString())
                        break; // if contain, then break
                    else if (j == tempWeekNumber.Count - 1)// if last cell still don't contain, then add it
                        tempWeekNumber.Add(rawdata.Rows[i][WeekNumber - 1].ToString());
                }

            }

            #endregion

            Yield_Chart.Series.Clear();
            #region ChartArea
            ChartArea chartArea = Yield_Chart.ChartAreas[0];
            chartArea.BorderDashStyle = ChartDashStyle.Solid;
            chartArea.BackColor = Color.WhiteSmoke;// Color.FromArgb(0, 0, 0, 0);       
            chartArea.ShadowColor = Color.FromArgb(0, 0, 0, 0);
            chartArea.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dash;//设置网格为虚线
            chartArea.AxisX.MinorGrid.LineDashStyle = ChartDashStyle.Dash;//设置网格为虚线

            chartArea.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash;

            //chartArea.AxisX.Minimum = Double.Parse(tempWeekNumber.Min());// set x axis start at min week number
            //chartArea.AxisX.Maximum = Double.Parse(tempWeekNumber.Max());// set x axis end at max week number
            chartArea.AxisX.Title = "Week Number";
            chartArea.AxisY.Title = YieldType;

            chartArea.AxisY2.Enabled = AxisEnabled.True;
            Yield_Chart.ChartAreas[0].AxisY2.IsStartedFromZero = Yield_Chart.ChartAreas[0].AxisY.IsStartedFromZero;// 横坐标对齐？

            chartArea.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            chartArea.AxisY2.Title = "Quantity";
            //newChart.ChartAreas.Add(chartArea);

            chartArea.AxisX.TitleFont = new Font("Arial", 16.0F, FontStyle.Regular);
            chartArea.AxisY.TitleFont = new Font("Arial", 16.0F, FontStyle.Regular);
            chartArea.AxisY2.TitleFont = new Font("Arial", 16.0F, FontStyle.Regular);
            #endregion



            tempWeekNumber.Sort();// sort the week number so that weeknum in plot is from small to big

            string tempstring;
            string passString;
            string failString;
            for (var i = 0; i < tempModelID.Count; i++)
            {
                for (var j = 0; j < tempDataSet.Count; j++)
                {


                    passString = tempModelID[i].ToString() + " " + tempDataSet[j].ToString() + " Pass Quantity";
                    Yield_Chart.Series.Add(passString);
                    Yield_Chart.Series[passString].Font = new Font("Arial", 10.0F, FontStyle.Regular);
                    Yield_Chart.Series[passString].BorderWidth = 1;
                    Yield_Chart.Series[passString].YAxisType = AxisType.Secondary;
                    Yield_Chart.Series[passString].IsValueShownAsLabel = true;




                    failString = tempModelID[i].ToString() + " " + tempDataSet[j].ToString() + " Fail Quantity";
                    Yield_Chart.Series.Add(failString);
                    Yield_Chart.Series[failString].Font = new Font("Arial", 10.0F, FontStyle.Regular);
                    Yield_Chart.Series[failString].BorderWidth = 1;
                    Yield_Chart.Series[failString].YAxisType = AxisType.Secondary;
                    Yield_Chart.Series[failString].IsValueShownAsLabel = true;



                    tempstring = tempModelID[i].ToString() + " " + tempDataSet[j].ToString() + " Ratio";
                    Yield_Chart.Series.Add(tempstring);
                    Yield_Chart.Series[tempstring].Font = new Font("Arial", 10.0F, FontStyle.Regular);
                    Yield_Chart.Series[tempstring].BorderWidth = 3;
                    Yield_Chart.Series[tempstring].ChartType = SeriesChartType.Line;
                    Yield_Chart.Series[tempstring].YAxisType = AxisType.Primary;
                    Yield_Chart.Series[tempstring].IsValueShownAsLabel = true;
                    Yield_Chart.Series[tempstring].MarkerStyle = MarkerStyle.Square;




                    for (var h = 0; h < tempWeekNumber.Count; h++)
                    {

                        for (var k = 0; k < rawdata.Rows.Count; k++)
                        {
                            if (rawdata.Rows[k][ModelID - 1].ToString() == tempModelID[i] &&
                                rawdata.Rows[k][DataSetName - 1].ToString() == tempDataSet[j] &&
                                rawdata.Rows[k][WeekNumber - 1].ToString() == tempWeekNumber[h])
                            {
                                Yield_Chart.Series[passString].Points.AddXY(Double.Parse(rawdata.Rows[k][WeekNumber - 1].ToString()), rawdata.Rows[k][PassQTY - 1]);
                                Yield_Chart.Series[failString].Points.AddXY(Double.Parse(rawdata.Rows[k][WeekNumber - 1].ToString()), rawdata.Rows[k][FailQTY - 1]);
                                Yield_Chart.Series[tempstring].Points.AddXY(Double.Parse(rawdata.Rows[k][WeekNumber - 1].ToString()), rawdata.Rows[k][Ratio - 1]);
                            }


                        }
                    }
                }
            }



        }

        public void Search_Open_TestFile_From_Production_PC(string SerialNumber, string StationName, string ComputerName)
        {
            //sql = "select  [Computer Name] from [StationName_ComputerName] where [Station Name]=" + "'" + StationName + "'";
            ////ds = accessdbo.GetAccessDataset2(sql, AccessConnection);
            //try
            //{
            //    ds = GetAccessDataset2(dbPath, sql);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Cannot find " + StationName + "From Access Config File " + ex);

            //    //ds.Dispose();
            //}

            //DataSet.Items.Clear();
            //string ComputerName;

            //ComputerName = ds.Tables[0].Rows[0][0].ToString().Trim();
            //ds.Dispose();
            try
            {
                //This is for tunable 10G log file, txt file name is end with _S1
                string temp1 = "\\\\" + ComputerName + "\\testfiles\\" + SerialNumber + "_S1.txt";
                string temp2 = System.AppDomain.CurrentDomain.BaseDirectory + "testfiles\\" + StationName + "_" + SerialNumber + "_S1.txt";
                File.Copy(temp1, temp2, true);//Copy the log txe file to local folder
                //File.Open(temp2,FileMode.Open,FileAccess.ReadWrite);
                Process.Start(temp2);//Open the txt file
            }

            catch
            {
                try
                {
                    //This is for CFP2 100G log file, txt file name is end with _S0
                    string temp1 = "\\\\" + ComputerName + "\\testfiles\\" + SerialNumber + "_S0.txt";
                    string temp2 = System.AppDomain.CurrentDomain.BaseDirectory + "testfiles\\" + StationName + "_" + SerialNumber + "_S0.txt";
                    File.Copy(temp1, temp2, true);//Copy the log txe file to local folder
                    //File.Open(temp2,FileMode.Open,FileAccess.ReadWrite);
                    Process.Start(temp2);//Open the txt file
                }
                catch (Exception e)
                {
                    MessageBox.Show("Maybe Access of " + ComputerName + " is Denied?\n" + e.ToString());
                }
            }









        }


        public DataSet Component_Edata_Search(string connectionstring, string SN, string DataName, string DataNameVal)
        {
            DataSet ds;

            string sql = "SELECT Y.Mfr_Sn,Y.Mfr_Pn,V.Data_Name,V.Data_Val1,V.Data_Val2,V.Data_Val3,V.Data_Val4,V.Time FROM PARTS y INNER JOIN PART_DATA v ON y.OPT_INDEX = v.opt_INDEX WHERE y.mfr_sn";

            if (SN.Length != 0)
            {

                SN = SN.Replace(",", "','");
                sql += " in ('" + SN + "')";



                if (DataName.Length != 0)
                {
                    if (DataName.Split(',').Length > 1)
                    {
                        string[] tempstring = DataName.Split(',');
                        if (tempstring[0].Contains("%"))
                            sql += " and " + "(v.DATA_NAME like '" + tempstring[0] + "'";
                        else
                            sql += " and " + "(v.DATA_NAME = '" + tempstring[0] + "'";
                        for (var i = 1; i < tempstring.Length; i++)
                        {
                            if (tempstring[0].Contains("%"))
                                sql += " or " + "v.DATA_NAME like '" + tempstring[i] + "'";
                            else
                                sql += " or " + "v.DATA_NAME = '" + tempstring[i] + "'";
                        }
                        sql += ')';
                    }
                    else
                        sql += " and " + "v.DATA_NAME like '" + DataName + "'";
                }

                if (DataNameVal.Length != 0)
                    sql += " and " + "v.DATA_VAL1 " + DataNameVal;


                sql += "order by Y.Mfr_Sn, V.DATA_NAME";
                ds = GetOracleDataSet2(connectionstring, sql);
                return ds;
            }

            else
            {
                MessageBox.Show("Please input SN");
                return null;
            }

        }




        public DataSet RemoveDuplicateRows(DataTable FPYRawData, bool Ignor_PN_Rev)
        {

            //"SELECT a.MFR_SN AS Module_SN,b.JOB_ID  AS JOB_ID,  c.MODEL_ID  AS Model_ID,c.BOM_PN  AS BOM_PN,c.BOM_PN_REV BOM_PN_Rev," +
            //    "d.DATASET_NAME AS DataSet_Name,d.STATUS AS DataSetStatus,d.END_TIME AS DataSet_EndTime," +
            //    "e.DATA_NAME AS Data_Name,e.DATA_VAL1  AS Data_Val,e.DATA_VAL2 AS Data_Status,e.DATA_VAL3 AS Data_Spec, "

            DataSet returnDataset = new DataSet();


            int SN = 1;
            int JobID = 2;
            int ModelID = 3;
            int BOMPN = 4;
            int BOMPNRev = 5;
            int DataSetName = 6;
            int DataSetStatus = 7;
            int DataSetEndtime = 8;
            int DataName = 9;


            int Flag = FPYRawData.Columns.Count + 1;

            FPYRawData.Columns.Add("Flag");


            #region filt and get first column data
            for (int i = 0; i <= FPYRawData.Rows.Count - 1; i++)//Filter out the FIRST record of each datasetname
            {


                string tempSN = FPYRawData.Rows[i][SN - 1].ToString();
                string tempdatasetname = FPYRawData.Rows[i][DataSetName - 1].ToString();
                string tempdatasettime = FPYRawData.Rows[i][DataSetEndtime - 1].ToString();
                string tempBOMPN = FPYRawData.Rows[i][BOMPN - 1].ToString();
                string tempBOMPNRev = FPYRawData.Rows[i][BOMPNRev - 1].ToString();
                string tempDataName = FPYRawData.Rows[i][DataName - 1].ToString();


                for (int j = i + 1; j <= FPYRawData.Rows.Count - 1; j++)
                {
                    if (FPYRawData.Rows[j][SN - 1].ToString() == tempSN &&
                        FPYRawData.Rows[j][BOMPN - 1].ToString() == tempBOMPN &&
                        (FPYRawData.Rows[j][BOMPNRev - 1].ToString() == tempBOMPNRev || Ignor_PN_Rev) &&
                        FPYRawData.Rows[j][DataSetName - 1].ToString() == tempdatasetname &&
                        FPYRawData.Rows[j][DataName - 1].ToString() == tempDataName)
                        FPYRawData.Rows[j][Flag - 1] = "Del";


                }




            }

            for (int i = FPYRawData.Rows.Count - 1; i >= 0; i--)
            {
                if (FPYRawData.Rows[i][Flag - 1] == "Del")
                    FPYRawData.Rows[i].Delete();
            }

            FPYRawData.AcceptChanges();




            returnDataset.Tables.Add(FPYRawData.Copy());
            return returnDataset;
            //return FPYRawData;


            #endregion





        }


        public DataSet WIP_Status_Search(string connectionstring, string JobID)
        {
            DataSet ds;


            //This sql process is as below:
            //1. Pull out all SN according to Input Job ID list
            //2. Pull out the latest ROUTE_ID whose SN is in SN list of above step ---MAX(b.ROUTE_ID)
            //3. Pull out the all needed information according to ROUTE_ID in ROUTE_ID list of above step AND Job_ID in Input Job_ID list 
            //4. In summary, job ID--> SN--> Latest Route ID--> JobID & Latest Route ID 
            string sql = " SELECT a.MFR_SN,  f.BOM_PN,  f.BOM_PN_REV,  f.MODEL_ID,  f.LIFECYCLE,  b.JOB_ID,  b.STATE,  b.STATUS,  b.ROUTE_ID,  f.TIME " +
" FROM PARTS a INNER JOIN ROUTES b ON a.OPT_INDEX = b.PART_INDEX INNER JOIN BOM_CONTEXT_ID f ON f.BOM_CONTEXT_ID = b.BOM_CONTEXT_ID " +
" WHERE b.JOB_ID     Replace_JobID_Here " +
" AND b.ROUTE_ID     IN  (SELECT ROUTE_ID   FROM     (SELECT a.MFR_SN,       MAX(b.ROUTE_ID) AS ROUTE_ID    FROM PARTS a " +
" INNER JOIN ROUTES b    ON a.OPT_INDEX  = b.PART_INDEX    WHERE a.MFR_SN IN " +
     " (SELECT a.MFR_SN      FROM PARTS a      INNER JOIN ROUTES b      ON a.OPT_INDEX  = b.PART_INDEX " +
     " WHERE b.JOB_ID  Replace_JobID_Here )" +
         "GROUP BY a.MFR_SN    )  )";

            if (JobID.Length != 0)
            {

                JobID = JobID.Replace(",", "','");
                string temp = " in ('" + JobID + "')";
                sql = sql.Replace("Replace_JobID_Here", temp);

                //sql += " order by a.MFR_SN, f.time DESC";

                ds = GetOracleDataSet2(connectionstring, sql);



                return ds;


            }

            else
            {
                MessageBox.Show("Please input JobID");
                return null;
            }

        }





    }


}

