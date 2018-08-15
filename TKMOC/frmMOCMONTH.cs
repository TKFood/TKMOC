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
    public partial class frmMOCMONTH : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();  
        SqlDataAdapter adapterCALENDAR = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderCALENDAR = new SqlCommandBuilder();
        SqlDataAdapter adapterCALENDAR2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderCALENDAR2 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet dsCALENDAR = new DataSet();
        DataSet dsCALENDAR2 = new DataSet();

        int result;

        public frmMOCMONTH()
        {
            InitializeComponent();

            SETCALENDAR();
            SETCALENDAR2();
        }
        #region FUNCTION
        public void SETCALENDAR()
        {
            string EVENT;
            DateTime dtEVENT;
            var ce2 = new CustomEvent();


            calendar1.RemoveAllEvents();
            calendar1.CalendarDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
            calendar1.CalendarView = CalendarViews.Month;
            calendar1.AllowEditingEvents = true;




            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  SELECT [EVENTDATE],[MOCLINE],[EVENT]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[CALENDARMONTH]");
                sbSql.AppendFormat(@"  WHERE [EVENTDATE]>='{0}'", DateTime.Now.ToString("yyyy") + "0101");
                sbSql.AppendFormat(@"  AND [MOCLINE]<>'包裝線'");
                sbSql.AppendFormat(@"  ORDER BY [EVENTDATE]");
                sbSql.AppendFormat(@"  ");

                adapterCALENDAR = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderCALENDAR = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                dsCALENDAR.Clear();
                adapterCALENDAR.Fill(dsCALENDAR, "TEMPdsCALENDAR");
                sqlConn.Close();


                if (dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows.Count >= 1)
                    {
                        foreach (DataRow od in dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows)
                        {
                            EVENT = od["MOCLINE"].ToString() + "-" + od["EVENT"].ToString();
                            dtEVENT = Convert.ToDateTime(od["EVENTDATE"].ToString());

                            ce2 = new CustomEvent
                            {
                                IgnoreTimeComponent = false,
                                EventText = EVENT,
                                Date = new DateTime(dtEVENT.Year, dtEVENT.Month, dtEVENT.Day),
                                EventLengthInHours = 2f,
                                RecurringFrequency = RecurringFrequencies.None,
                                EventFont = new Font("Verdana", 12, FontStyle.Regular),
                                Enabled = true,
                                EventColor = Color.FromArgb(120, 255, 120),
                                EventTextColor = Color.Black,
                                ThisDayForwardOnly = true
                            };

                            calendar1.AddEvent(ce2);
                        }



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
        public void ADDCALENDAR()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[CALENDARMONTH]");
                sbSql.AppendFormat(" ([EVENTDATE],[MOCLINE],[EVENT])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", dateTimePicker11.Value.ToString("yyyy/MM/dd"), comboBox10.Text, comboBox9.Text + "-" + textBox48.Text+" 小時");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DELCALENDAR()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[CALENDARMONTH]");
                sbSql.AppendFormat(" WHERE convert(varchar, [EVENTDATE], 112)='{0}' AND [MOCLINE]='{1}'", dateTimePicker11.Value.ToString("yyyyMMdd"), comboBox10.Text);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SETCALENDAR2()
        {
            string EVENT;
            DateTime dtEVENT;
            var ce2 = new CustomEvent();


            calendar2.RemoveAllEvents();
            calendar2.CalendarDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
            calendar2.CalendarView = CalendarViews.Month;
            calendar2.AllowEditingEvents = true;




            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  SELECT [EVENTDATE],[MOCLINE],[EVENT]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[CALENDARMONTH]");
                sbSql.AppendFormat(@"  WHERE [EVENTDATE]>='{0}'", DateTime.Now.ToString("yyyy") + "0101");
                sbSql.AppendFormat(@"  AND [MOCLINE]='包裝線'");
                sbSql.AppendFormat(@"  ORDER BY [EVENTDATE]");
                sbSql.AppendFormat(@"  ");

                adapterCALENDAR2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderCALENDAR2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                dsCALENDAR2.Clear();
                adapterCALENDAR2.Fill(dsCALENDAR2, "TEMPdsCALENDAR2");
                sqlConn.Close();


                if (dsCALENDAR2.Tables["TEMPdsCALENDAR2"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsCALENDAR2.Tables["TEMPdsCALENDAR2"].Rows.Count >= 1)
                    {
                        foreach (DataRow od in dsCALENDAR2.Tables["TEMPdsCALENDAR2"].Rows)
                        {
                            EVENT = od["MOCLINE"].ToString() + "-" + od["EVENT"].ToString();
                            dtEVENT = Convert.ToDateTime(od["EVENTDATE"].ToString());

                            ce2 = new CustomEvent
                            {
                                IgnoreTimeComponent = false,
                                EventText = EVENT,
                                Date = new DateTime(dtEVENT.Year, dtEVENT.Month, dtEVENT.Day),
                                EventLengthInHours = 2f,
                                RecurringFrequency = RecurringFrequencies.None,
                                EventFont = new Font("Verdana", 12, FontStyle.Regular),
                                Enabled = true,
                                EventColor = Color.FromArgb(120, 255, 120),
                                EventTextColor = Color.Black,
                                ThisDayForwardOnly = true
                            };

                            calendar2.AddEvent(ce2);
                        }



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
        public void ADDCALENDAR2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[CALENDARMONTH]");
                sbSql.AppendFormat(" ([EVENTDATE],[MOCLINE],[EVENT])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", dateTimePicker1.Value.ToString("yyyy/MM/dd"), comboBox2.Text, comboBox1.Text + "-" + textBox1.Text + " 小時");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DELCALENDAR2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[CALENDARMONTH]");
                sbSql.AppendFormat(" WHERE convert(varchar, [EVENTDATE], 112)='{0}' AND [MOCLINE]='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox2.Text);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        #endregion

        #region BUTTON

        private void button40_Click(object sender, EventArgs e)
        {
            ADDCALENDAR();
            SETCALENDAR();
        }
        private void button41_Click(object sender, EventArgs e)
        {
            DELCALENDAR();
            SETCALENDAR();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADDCALENDAR2();
            SETCALENDAR2();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DELCALENDAR2();
            SETCALENDAR2();
        }
        #endregion


    }
}
