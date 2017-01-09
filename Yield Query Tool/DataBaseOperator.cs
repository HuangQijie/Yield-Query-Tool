using System;
using System.Data.OracleClient;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.Threading;

namespace Yield_Query_Tool
{
    class OracleDataBaseOperator
    {

        //OleDbConnection AccessConnection;
        public OracleConnection OpenOracleConnection2(string ConnectionStr)
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


                }
                else if (Oracleconn.State == ConnectionState.Broken)
                {
                    //LogInfo.writeLog(String.Format("open ORA DB connection fail: {0}", ConnectionStr));
                    Oracleconn.Close();
                    Oracleconn.Open();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                throw new Exception(e.Message);
            }

            return Oracleconn;

        }

        public DataSet GetOracleDataSet2(string SQLString, OracleConnection conn)
        {
            DataSet ds = new DataSet();
            try
            {
                OracleDataAdapter sda = new OracleDataAdapter(SQLString, conn);
                sda.Fill(ds, "ds");
            }
            catch (OleDbException ex)
            {
                throw new Exception(ex.Message);
            }
            return ds;
        }


        public void OracleConnectionClose(OracleConnection Oracleconn)
        {
            Oracleconn.Close();


        }


    }


    class AccessDataBaseOperator
    {

        public OleDbConnection OpenAccessConnection2(string dbPath)
        {
            //string dbPath;

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
            return AccessConnection;
        }
        public DataSet GetAccessDataset2(string SQLString, OleDbConnection connection)
        {
            DataSet ds = new DataSet();
            try
            {
                OleDbDataAdapter sda = new OleDbDataAdapter(SQLString, connection);
                sda.Fill(ds, "ds");
                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return ds;

        }


        public void AccessConnectionClose(OleDbConnection AccessConnection)
        {
            AccessConnection.Close();


        }
    }
}
