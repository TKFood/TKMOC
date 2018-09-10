using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using System.Globalization;
using Calendar.NET;

namespace TKMOC
{
    public partial class FrmMOCWEEK : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();


        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();


        int result;

        public FrmMOCWEEK()
        {
            InitializeComponent();

            SETTODAY();
            SETFIRSTDAY();
        }

        #region FUNCTION
        public void SETTODAY()
        {
            dateTimePicker1.Value = DateTime.Now;
        }

        public void SETFIRSTDAY()
        {
            DateTime dt = dateTimePicker1.Value;

            dt.AddDays(-((int)dt.DayOfWeek));
            dateTimePicker2.Value = GetWeekFirstDayMon(dt); 
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = dateTimePicker1.Value;

            dateTimePicker2.Value = GetWeekFirstDayMon(dt);
            dateTimePicker3.Value = dateTimePicker2.Value.AddDays(1);
            dateTimePicker4.Value = dateTimePicker2.Value.AddDays(2);
            dateTimePicker5.Value = dateTimePicker2.Value.AddDays(3);
            dateTimePicker6.Value = dateTimePicker2.Value.AddDays(4);
            dateTimePicker7.Value = dateTimePicker2.Value.AddDays(5);
            dateTimePicker8.Value = dateTimePicker2.Value.AddDays(6);
        }

        public DateTime GetWeekFirstDayMon(DateTime datetime)
        {
            //星期一为第一天
            int weeknow = Convert.ToInt32(datetime.DayOfWeek);

            //因为是以星期一为第一天，所以要判断weeknow等于0时，要向前推6天。
            weeknow = (weeknow == 0 ? (7 - 1) : (weeknow - 1));
            int daydiff = (-1) * weeknow;

            //本周第一天
            string FirstDay = datetime.AddDays(daydiff).ToString("yyyy-MM-dd");
            return Convert.ToDateTime(FirstDay);
        }
        public void search()
        {
            SETDGNULL();

            search1();
            search2();
            search3();
            search4();
            search5();
            search6();
            search7();
        }

        public void SETDGNULL()
        {
            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            dataGridView4.DataSource = null;
            dataGridView5.DataSource = null;
            dataGridView6.DataSource = null;
            dataGridView7.DataSource = null;
        }
        public void search1()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002]  AS '品名',CONVERT(NVARCHAR,(CONVERT(int,[BOX])))  AS '數量','箱'  AS '箱', CLINET AS '客戶'");
                sbSql.AppendFormat(@"  ,[ID],[SERNO],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029]  ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE] ");
                sbSql.AppendFormat(@"  WHERE [MANU]='新廠包裝線'  ");
                sbSql.AppendFormat(@"  AND [MANUDATE]='{0}'",dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ORDER BY [MANUDATE]  ");
                sbSql.AppendFormat(@"   ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        public void search2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002]  AS '品名',CONVERT(NVARCHAR,(CONVERT(int,[BOX])))  AS '數量','箱'  AS '箱', CLINET AS '客戶'");
                sbSql.AppendFormat(@"  ,[ID],[SERNO],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029]  ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE] ");
                sbSql.AppendFormat(@"  WHERE [MANU]='新廠包裝線'  ");
                sbSql.AppendFormat(@"  AND [MANUDATE]='{0}'", dateTimePicker3.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ORDER BY [MANUDATE]  ");
                sbSql.AppendFormat(@"   ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        public void search3()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002]  AS '品名',CONVERT(NVARCHAR,(CONVERT(int,[BOX])))  AS '數量','箱'  AS '箱', CLINET AS '客戶'");
                sbSql.AppendFormat(@"  ,[ID],[SERNO],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029]  ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE] ");
                sbSql.AppendFormat(@"  WHERE [MANU]='新廠包裝線'  ");
                sbSql.AppendFormat(@"  AND [MANUDATE]='{0}'", dateTimePicker4.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ORDER BY [MANUDATE]  ");
                sbSql.AppendFormat(@"   ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "TEMPds3");
                sqlConn.Close();


                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds3.Tables["TEMPds3"];
                        dataGridView3.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        public void search4()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002]  AS '品名',CONVERT(NVARCHAR,(CONVERT(int,[BOX])))  AS '數量','箱'  AS '箱', CLINET AS '客戶'");
                sbSql.AppendFormat(@"  ,[ID],[SERNO],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029]  ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE] ");
                sbSql.AppendFormat(@"  WHERE [MANU]='新廠包裝線'  ");
                sbSql.AppendFormat(@"  AND [MANUDATE]='{0}'", dateTimePicker5.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ORDER BY [MANUDATE]  ");
                sbSql.AppendFormat(@"   ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds4.Tables["TEMPds4"];
                        dataGridView4.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        public void search5()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002]  AS '品名',CONVERT(NVARCHAR,(CONVERT(int,[BOX])))  AS '數量','箱'  AS '箱', CLINET AS '客戶'");
                sbSql.AppendFormat(@"  ,[ID],[SERNO],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029]  ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE] ");
                sbSql.AppendFormat(@"  WHERE [MANU]='新廠包裝線'  ");
                sbSql.AppendFormat(@"  AND [MANUDATE]='{0}'", dateTimePicker6.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ORDER BY [MANUDATE]  ");
                sbSql.AppendFormat(@"   ");

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "TEMPds5");
                sqlConn.Close();


                if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds5.Tables["TEMPds5"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds5.Tables["TEMPds5"];
                        dataGridView5.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        public void search6()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002]  AS '品名',CONVERT(NVARCHAR,(CONVERT(int,[BOX])))  AS '數量','箱'  AS '箱', CLINET AS '客戶'");
                sbSql.AppendFormat(@"  ,[ID],[SERNO],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029]  ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE] ");
                sbSql.AppendFormat(@"  WHERE [MANU]='新廠包裝線'  ");
                sbSql.AppendFormat(@"  AND [MANUDATE]='{0}'", dateTimePicker7.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ORDER BY [MANUDATE]  ");
                sbSql.AppendFormat(@"   ");

                adapter6 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder6 = new SqlCommandBuilder(adapter6);
                sqlConn.Open();
                ds6.Clear();
                adapter6.Fill(ds6, "TEMPds6");
                sqlConn.Close();


                if (ds6.Tables["TEMPds6"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds6.Tables["TEMPds6"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView6.DataSource = ds6.Tables["TEMPds6"];
                        dataGridView6.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        public void search7()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002]  AS '品名',CONVERT(NVARCHAR,(CONVERT(int,[BOX])))  AS '數量','箱'  AS '箱', CLINET AS '客戶'");
                sbSql.AppendFormat(@"  ,[ID],[SERNO],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029]  ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE] ");
                sbSql.AppendFormat(@"  WHERE [MANU]='新廠包裝線'  ");
                sbSql.AppendFormat(@"  AND [MANUDATE]='{0}'", dateTimePicker8.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ORDER BY [MANUDATE]  ");
                sbSql.AppendFormat(@"   ");

                adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                sqlConn.Open();
                ds7.Clear();
                adapter7.Fill(ds7, "TEMPds7");
                sqlConn.Close();


                if (ds7.Tables["TEMPds7"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds7.Tables["TEMPds7"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView7.DataSource = ds7.Tables["TEMPds7"];
                        dataGridView7.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        #endregion




        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            search();
        }
        #endregion

        
    }
}
