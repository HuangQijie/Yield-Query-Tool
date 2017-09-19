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
using System.Threading;

namespace Yield_Query_Tool
{
    public partial class Form_InformBox : Form
    {

        public Form_InformBox()
        {
            InitializeComponent();
            this.Show();
            //Thread t = new Thread(Query_Time_Remaining_Form);
            //t.Start();
            //Query_Time_Remaining_Form( ConnectionStr);
            //this.Hide();

        }



        //private void Query_Time_Remaining_Form(string ConnectionStr)
        //{
        //    //string ConnectionStr = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=shg-oracle)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=ateshg)));User Id=extviewer;Password=extviewer";

        //    Function fn = new Function();


        //    DataSet ds = new DataSet();
        //    DateTime start = DateTime.Now;
        //    double Elaspe_Seconds = 0;
        //    while (Elaspe_Seconds < 600 && this.IsDisposed == false && Thread.CurrentThread.IsAlive)
        //    {
        //        DateTime end = DateTime.Now;
        //        Elaspe_Seconds = (end - start).TotalSeconds;
        //        string Elapse_Time = Elaspe_Seconds.ToString("0.00");
        //        string sql = "select Sum(time_remaining) from V$SESSION_LONGOPS where time_remaining>0 and username in ('EXTVIEWER','TEST')";
        //        string Query_Remaining_Time = fn.GetOracleDataSet2(ConnectionStr, sql).Tables[0].Rows[0][0].ToString();
        //        string Infor_Msg = "Elapse Time is " + Elapse_Time + " Seconds\n Estimated Query Remaining Time is " + Query_Remaining_Time + " Seconds";

        //        Label_Infor_Msg.Text = Infor_Msg;
        //        Application.DoEvents();
        //        Thread.Sleep(500);
        //    }


        //}
        public void Change_Label_Infor_MSG(string Infor_Msg)
        {
            Label_Infor_Msg.Text = Infor_Msg;
        }


    }
}
