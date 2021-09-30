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
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKMOC
{
    public partial class frmMOCMONTHVIEW : Form
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
        Report report1 = new Report();

        public frmMOCMONTHVIEW()
        {
            InitializeComponent();

            SETCALENDAR();
           
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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                //sbSql.AppendFormat(@"  SELECT [EVENTDATE],[MOCLINE],[EVENT]");
                //sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[CALENDARMONTH]");
                //sbSql.AppendFormat(@"  WHERE [EVENTDATE]>='{0}'", DateTime.Now.ToString("yyyy") + "0101");
                //sbSql.AppendFormat(@"  AND [MOCLINE]<>'包裝線'");
                //sbSql.AppendFormat(@"  ORDER BY [EVENTDATE]");
                //sbSql.AppendFormat(@"  ");

                sbSql.AppendFormat(@"  SELECT  [MOCMANULINE].[MANUDATE]");
                sbSql.AppendFormat(@"  ,CONVERT(decimal(16,2),SUM([MOCMANULINE].[PACKAGE]/[MOCSTDTIME].PROCESSNUM*[MOCSTDTIME].PROCESSTIME/60))  AS TIMES");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MOCSTDTIME] ON [MOCMANULINE].[MB001]=[MOCSTDTIME].[MB001]");
                sbSql.AppendFormat(@"  WHERE [MANU]='包裝線' AND ([MANUDATE]>=CONVERT(NVARCHAR,datepart(yyyy, getdate()))+'/1/1')");
                sbSql.AppendFormat(@"  GROUP BY [MANUDATE]");
                sbSql.AppendFormat(@"  ORDER BY [MANUDATE]");
                sbSql.AppendFormat(@"  ");
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
                            EVENT = "包裝:" + od["TIMES"].ToString()+"小時";
                            dtEVENT = Convert.ToDateTime(od["MANUDATE"].ToString());

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
                                EventTextColor = Color.Red,
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

        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();

            SQL = SETSQL();

            report1.Load(@"REPORT\包裝時數明細.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {


            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" SELECT  [MOCMANULINE].[MANU] AS '線別',CONVERT(nvarchar,[MOCMANULINE].[MANUDATE],112) AS '日期',[MOCMANULINE].[MB001] AS '品號',[MOCMANULINE].[MB002] AS '品名',[MOCMANULINE].[BOX] AS '箱數',[MOCMANULINE].[PACKAGE] AS '包裝數'");
            SB.AppendFormat(@" ,CONVERT(decimal(16,2),([MOCMANULINE].[PACKAGE]/[MOCSTDTIME].PROCESSNUM*[MOCSTDTIME].PROCESSTIME/60)) AS '包裝時數'");
            SB.AppendFormat(@" ,[MOCMANULINE].[ID],[MOCMANULINE].[SERNO]");
            SB.AppendFormat(@" ,[MOCSTDTIME].PROCESSNUM,[MOCSTDTIME].PROCESSTIME");
            SB.AppendFormat(@" FROM [TKMOC].[dbo].[MOCMANULINE]");
            SB.AppendFormat(@" LEFT JOIN [TKMOC].[dbo].[MOCSTDTIME] ON [MOCMANULINE].[MB001]=[MOCSTDTIME].[MB001]");
            SB.AppendFormat(@" WHERE [MANU]='包裝線' ");
            SB.AppendFormat(@" AND [MANUDATE]>='{0}'  AND  [MANUDATE]<='{1}' ",dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
            SB.AppendFormat(@" ORDER BY [MANUDATE],[MOCMANULINE].[MB001]");
            SB.AppendFormat(@" ");

            return SB;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion


    }
}
