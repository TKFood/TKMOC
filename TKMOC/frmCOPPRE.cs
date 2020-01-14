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
    public partial class frmCOPPRE : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();

        int result;

        

        public class ADDITEM
        {
            public string ORDERNO { get; set; }
            public string MB001 { get; set; }
            public int AMOUNT { get; set; }
            public int PRIORITYS { get; set; }
            public string MANU { get; set; }
            public decimal TIMES { get; set; }
            public int HRS { get; set; }
            public string WDT { get; set; }
            public int WHRS { get; set; }
        }
        public frmCOPPRE()
        {
            InitializeComponent();
        }


        #region FUNCTION

        public void PRESCHEDULE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [PREORDER].[ORDERNO],[PREORDER].[MB001],[PREORDER].[AMOUNT],[PREORDER].[PRIORITYS],[PREINVMBMANU].MANU,[PREINVMBMANU].TIMES");
                sbSql.AppendFormat(@"  ,CONVERT(INT,ROUND([PREORDER].[AMOUNT]/[PREINVMBMANU].TIMES,0)) AS HRS");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[PREORDER],[TKMOC].[dbo].[PREINVMBMANU]");
                sbSql.AppendFormat(@"  WHERE [PREORDER].MB001=[PREINVMBMANU].MB001");
                sbSql.AppendFormat(@"  ORDER BY [PREINVMBMANU].MANU,[PREORDER].[PRIORITYS] DESC,[ORDERNO]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    //dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        ADDNEWTARGET(ds1.Tables["TEMPds1"]);

                        ////dataGridView1.Rows.Clear();
                        //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        //dataGridView1.AutoResizeColumns();
                        ////dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
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

        public void ADDNEWTARGET(DataTable dt)
        {
            List<ADDITEM> ADDTARGET = new List<ADDITEM>();

            int WORKHRS = 0;
            int TWORKHRS = 0;
            int CHECKHRS = 0;
            int day = 0;
            DateTime wdt = new DateTime();
            wdt = DateTime.Now;

            bool result;
            string LASTMANU = "START";

            foreach (DataRow od in dt.Rows)
            {
                TWORKHRS = 0;
                CHECKHRS = Convert.ToInt16(od["HRS"].ToString());

                //MessageBox.Show(LASTMANU+" "+ od["MANU"].ToString());
                
                if(!LASTMANU.Equals(od["MANU"].ToString()))
                {
                    WORKHRS = 0;
                    TWORKHRS = 0;
                    day = 0;
                    wdt = DateTime.Now;

                    //MessageBox.Show(LASTMANU + " " + od["MANU"].ToString());
                }

                while ((CHECKHRS) >= TWORKHRS)
                {
                    if (WORKHRS <= 16)
                    {
                        ADDTARGET.Add(new ADDITEM { ORDERNO = od["ORDERNO"].ToString(), MB001 = od["MB001"].ToString(), AMOUNT = Convert.ToInt16(od["AMOUNT"].ToString()), PRIORITYS = Convert.ToInt16(od["PRIORITYS"].ToString()), MANU = od["MANU"].ToString(), TIMES = Convert.ToDecimal(od["TIMES"].ToString()), HRS = Convert.ToInt16(od["HRS"].ToString()), WDT = wdt.ToString("yyyyMMdd"), WHRS = (WORKHRS + 2) });
                        WORKHRS = WORKHRS + 2;
                    }
                    else if (WORKHRS > 16)
                    {
                        WORKHRS = 0;
                        day = day + 1;
                        wdt = wdt.AddDays(day);

                        ADDTARGET.Add(new ADDITEM { ORDERNO = od["ORDERNO"].ToString(), MB001 = od["MB001"].ToString(), AMOUNT = Convert.ToInt16(od["AMOUNT"].ToString()), PRIORITYS = Convert.ToInt16(od["PRIORITYS"].ToString()), MANU = od["MANU"].ToString(), TIMES = Convert.ToDecimal(od["TIMES"].ToString()), HRS = Convert.ToInt16(od["HRS"].ToString()), WDT = wdt.ToString("yyyyMMdd"), WHRS = (WORKHRS + 2) });
                        WORKHRS = WORKHRS + 2;
                    }

                    TWORKHRS = TWORKHRS + 2;
                }
                

               

                LASTMANU = od["MANU"].ToString();
            }

            
            var bindingList = new BindingList<ADDITEM>(ADDTARGET);
            var source = new BindingSource(bindingList, null);
            dataGridView1.DataSource = source;
            dataGridView1.AutoResizeColumns();
        }

        public void SEARCHPREORDER()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                
                sbSql.AppendFormat(@" SELECT [ORDERNO] AS '訂單',[CLIENT] AS '客戶',[OUTDATES] AS '預交日',[MB001] AS '品號',[MB002] AS '品名',[AMOUNT] AS '數量',[UNIT] AS '單位',[PRIORITYS] AS '優先權'");
                sbSql.AppendFormat(@" FROM [TKMOC].[dbo].[PREORDER] ");
                sbSql.AppendFormat(@" ORDER BY [CLIENT],[OUTDATES] ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {

                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["ds2"];
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
                sqlConn.Close();
            }
        }

        public void SEARCHERPCOPTD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TD001+TD002+TD003 AS '訂單',TC053 AS '客戶',TD013 AS '預交日',TD004 AS '品號',TD005 AS '品名',CASE WHEN ISNULL(MD001,'')<>'' THEN (TD008+TD024-TD009-TD025)*MD004 ELSE (TD008+TD024-TD009-TD025) END  AS '數量',MB004 AS '單位'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMD] ON MD001=TD004 AND MD002=TD010");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=TD004");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD016='N' AND TD021='Y'");
                sbSql.AppendFormat(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@"  AND (TD008+TD024-TD009-TD025)>0");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY TC053,TD013");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {

                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds3.Tables["ds3"];
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

        #endregion

            #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            PRESCHEDULE();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHPREORDER();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCHERPCOPTD();
        }

        #endregion


    }
}
